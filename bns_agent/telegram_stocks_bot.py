from __future__ import annotations

import json
import os
import time
import uuid
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

from .stocks import REGIONS, StockStore, parse_stock_report


class TelegramBot:
    def __init__(self) -> None:
        token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
        if not token:
            raise RuntimeError("Укажите TELEGRAM_BOT_TOKEN.")
        self.base = f"https://api.telegram.org/bot{token}"
        self.files = f"https://api.telegram.org/file/bot{token}"
        self.admin_chat_id = int(os.environ.get("ADMIN_CHAT_ID", "0") or 0)
        self.store = StockStore(Path(os.environ.get("STOCKS_DB", "data/stocks.sqlite3")))
        self.downloads = Path(os.environ.get("STOCKS_UPLOAD_DIR", "data/stocks_uploads"))
        self.downloads.mkdir(parents=True, exist_ok=True)
        self.offset = 0
        self.last_reminder_key = ""

    def api(self, method: str, **data):
        request = Request(
            f"{self.base}/{method}",
            data=urlencode(data).encode("utf-8"),
            method="POST",
        )
        with urlopen(request, timeout=70) as response:
            payload = json.loads(response.read().decode("utf-8"))
        if not payload.get("ok"):
            raise RuntimeError(payload.get("description", "Ошибка Telegram API"))
        return payload["result"]

    def send(self, chat_id: int, text: str) -> None:
        self.api("sendMessage", chat_id=chat_id, text=text)

    def send_document(self, chat_id: int, path: Path, caption: str = "") -> None:
        boundary = f"----AgroPulse{uuid.uuid4().hex}"
        parts: list[bytes] = []
        content_types = {
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xls": "application/vnd.ms-excel",
            ".zip": "application/zip",
        }
        content_type = content_types.get(path.suffix.lower(), "application/octet-stream")

        def add_field(name: str, value: object) -> None:
            parts.append(
                (
                    f"--{boundary}\r\n"
                    f"Content-Disposition: form-data; name=\"{name}\"\r\n\r\n"
                    f"{value}\r\n"
                ).encode("utf-8")
            )

        add_field("chat_id", chat_id)
        if caption:
            add_field("caption", caption)
        parts.append(
            (
                f"--{boundary}\r\n"
                f"Content-Disposition: form-data; name=\"document\"; filename=\"{path.name}\"\r\n"
                f"Content-Type: {content_type}\r\n\r\n"
            ).encode("utf-8")
        )
        parts.append(path.read_bytes())
        parts.append(f"\r\n--{boundary}--\r\n".encode("utf-8"))
        body = b"".join(parts)
        request = Request(
            f"{self.base}/sendDocument",
            data=body,
            method="POST",
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        )
        with urlopen(request, timeout=70) as response:
            payload = json.loads(response.read().decode("utf-8"))
        if not payload.get("ok"):
            raise RuntimeError(payload.get("description", "Ошибка отправки файла"))

    def is_admin(self, chat_id: int) -> bool:
        return bool(self.admin_chat_id and chat_id == self.admin_chat_id)

    def status_text(self) -> str:
        latest = self.store.latest_by_region()
        missing = [region for region in REGIONS if region not in latest]
        lines = [f"Получено: {len(latest)} из {len(REGIONS)} регионов."]
        lines.extend(
            f"✅ {region}: {row['report_date']}" for region, row in latest.items()
        )
        if missing:
            lines.append("\nНе поступили:\n" + "\n".join(f"⏳ {name}" for name in missing))
        return "\n".join(lines)

    def handle_text(self, chat_id: int, text: str) -> None:
        command, *rest = text.strip().split(maxsplit=1)
        command = command.split("@", 1)[0]
        if command in {"/start", "/help"}:
            self.send(
                chat_id,
                "Отправьте XLSX-файл с запасами.\n"
                "/register <область> — зарегистрировать контакт\n"
                "/status — статус сдачи отчётов\n"
                "/export — Excel-свод для администратора\n"
                "/myid — показать ваш chat ID\n"
                "/regions — допустимые названия регионов",
            )
        elif command == "/myid":
            self.send(chat_id, f"Ваш chat ID: {chat_id}")
        elif command == "/regions":
            self.send(chat_id, "\n".join(REGIONS))
        elif command == "/register":
            requested = rest[0].strip() if rest else ""
            region = next((name for name in REGIONS if name.casefold() == requested.casefold()), None)
            if not region:
                self.send(chat_id, "Укажите регион точно как в /regions.")
                return
            self.store.register_contact(chat_id, region)
            self.send(chat_id, f"Контакт зарегистрирован: {region}.")
        elif command in {"/status", "/report"}:
            self.send(chat_id, self.status_text())
        elif command == "/export":
            if not self.is_admin(chat_id):
                self.send(chat_id, "Команда доступна только администратору.")
                return
            path = Path("output") / f"stocks_summary_{datetime.now():%Y%m%d_%H%M}.xlsx"
            archive_path = Path("output") / f"stocks_sources_{datetime.now():%Y%m%d_%H%M}.zip"
            self.store.export_latest_summary(path)
            self.store.export_latest_sources_archive(archive_path, self.downloads)
            self.send_document(chat_id, path, "1/2 Сводная таблица по последним отчётам регионов.")
            self.send_document(chat_id, archive_path, "2/2 Архив исходных файлов областей для сверки. Внутри есть manifest.csv.")
        else:
            self.send(chat_id, "Пришлите XLSX-файл или используйте /help.")

    def handle_document(self, chat_id: int, document: dict) -> None:
        name = Path(document.get("file_name") or "report.xlsx").name
        suffix = Path(name).suffix.lower()
        if suffix not in {".xlsx", ".xls"}:
            self.send(chat_id, "Нужен файл XLSX. Старый XLS сначала сохраните как XLSX.")
            return
        info = self.api("getFile", file_id=document["file_id"])
        with urlopen(f"{self.files}/{quote(info['file_path'])}", timeout=70) as response:
            content = response.read()
        target = self.downloads / f"{datetime.now():%Y%m%d_%H%M%S}_{chat_id}_{name}"
        target.write_bytes(content)
        try:
            report = parse_stock_report(target)
            report = replace(report, source_file=name)
            self.store.save(report, chat_id)
            warning = "\n⚠️ " + "\n⚠️ ".join(report.warnings) if report.warnings else ""
            self.send(
                chat_id,
                f"✅ Принято: {report.region}, {report.report_date:%d.%m.%Y}\n"
                f"Товаров: {len(report.rows)}{warning}",
            )
            if self.admin_chat_id and self.admin_chat_id != chat_id:
                self.send(
                    self.admin_chat_id,
                    f"Получен отчёт: {report.region}, {report.report_date:%d.%m.%Y}, "
                    f"{len(report.rows)} товаров.",
                )
        except Exception as exc:
            self.send(chat_id, f"❌ Файл не принят: {exc}")

    def handle_update(self, update: dict) -> None:
        message = update.get("message") or update.get("channel_post")
        if not message:
            return
        chat_id = message["chat"]["id"]
        if message.get("document"):
            self.handle_document(chat_id, message["document"])
        elif message.get("text"):
            self.handle_text(chat_id, message["text"])

    def reminders(self) -> None:
        now = datetime.now()
        weekday = int(os.environ.get("REMINDER_WEEKDAY", "3"))
        hour = int(os.environ.get("REMINDER_HOUR", "10"))
        key = now.strftime("%Y-%m-%d-%H")
        if now.weekday() != weekday or now.hour != hour or key == self.last_reminder_key:
            return
        latest = self.store.latest_by_region()
        contacts = self.store.contacts()
        for contact in contacts:
            region = contact["region"]
            if region and region not in latest:
                self.send(contact["chat_id"], f"⏰ Напоминание: ожидается еженедельный отчёт по запасам — {region}.")
        if self.admin_chat_id:
            self.send(self.admin_chat_id, self.status_text())
        self.last_reminder_key = key

    def run(self) -> None:
        if self.admin_chat_id:
            self.store.register_contact(self.admin_chat_id, None, admin=True)
        while True:
            try:
                updates = self.api("getUpdates", offset=self.offset, timeout=50)
                for update in updates:
                    self.offset = update["update_id"] + 1
                    self.handle_update(update)
                self.reminders()
            except KeyboardInterrupt:
                return
            except Exception as exc:
                print(json.dumps({"error": str(exc)}, ensure_ascii=False))
                time.sleep(5)


def main() -> None:
    TelegramBot().run()


if __name__ == "__main__":
    main()
