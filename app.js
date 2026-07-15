const state = { data: null, selected: null, query: "", tableFilter: "all" };
const $ = selector => document.querySelector(selector);
const fmt = new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 });
const fmt1 = new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 1 });
const money = new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 });
const pct = value => value == null ? "н/д" : `${fmt1.format(value)}%`;
const signed = value => value == null ? "н/д" : `${value > 0 ? "+" : ""}${fmt1.format(value)}%`;
const num1 = value => value == null ? "н/д" : fmt1.format(value);
const pp = value => value == null ? "н/д" : `${fmt1.format(value)} п.п.`;
const dateFmt = iso => new Intl.DateTimeFormat("ru-RU", { day: "numeric", month: "long", year: "numeric" }).format(new Date(`${iso}T00:00:00`));
const deltaClass = value => value > 0 ? "up" : value < 0 ? "down" : "flat";
const score = value => value == null ? "н/д" : fmt1.format(value);
const unitShort = unit => {
  const text = String(unit || "тонн").toLowerCase();
  if (text.includes("штук")) return "тыс. шт.";
  if (text.includes("тонн")) return "т";
  return unit || "";
};
const quantityText = (value, unit = "тонн") => value == null ? "н/д" : `${fmt.format(value)} ${unitShort(unit)}`;
const maybeSetText = (selector, value) => {
  const element = $(selector);
  if (element) element.textContent = value;
};
const priceDelta = (label, value) => `<span>${label}</span><em class="delta ${deltaClass(value)}">${signed(value)}</em>`;
const stockCategoryLabels = {
  agricultural_enterprises: "сельхозпредприятия",
  farms: "КФХ",
  other_enterprises: "торговые точки и ИП",
  warehouses: "склады/ОРЦ/ТЛЦ",
  stabilization: "стабфонды",
  vegetable_storage: "овощехранилища",
  fruit_storage: "фруктохранилища",
};

async function init() {
  state.data = await (await fetch("data/bns.json?v=20260711-card")).json();
  state.selected = state.data.products[0]?.id;
  fillStatic();
  bind();
  renderAll();
}

function product() {
  return state.data.products.find(item => item.id === state.selected);
}

function visibleProducts() {
  let items = state.data.products.filter(item => !state.query || item.name.toLowerCase().includes(state.query));
  if (state.tableFilter === "risk") items = items.filter(item => item.riskScore >= 28);
  if (state.tableFilter === "growth") items = items.filter(item => (item.annualChange ?? 0) > 0 || (item.weekChange ?? 0) > 0);
  return [...items].sort((a, b) => b.riskScore - a.riskScore);
}

function productOptions() {
  return state.data.products
    .map(item => `<option value="${item.id}">${String(item.number).padStart(2, "0")} · ${item.name}</option>`)
    .join("");
}

function coverageStatus(value) {
  if (value == null) return ["Нет расчёта", "neutral"];
  if (value >= 100) return ["Самообеспечение", "good"];
  if (value >= 80) return ["Умеренный риск", "warn"];
  return ["Импортозависимость", "risk"];
}

function riskClass(score) {
  if (score >= 55) return "risk";
  if (score >= 28) return "warn";
  return "good";
}

function importShareText(value) {
  return value == null ? "н/д" : pct(value);
}

function coverageSourceNote(item) {
  if (!item.coverageSourceProduct) return "";
  const source = item.coverageSourceProduct.toLowerCase();
  const own = item.name.toLowerCase();
  return source === own ? "" : `Баланс сопоставлен с группой: ${item.coverageSourceProduct}.`;
}

function productionPeriodLabel(item) {
  if (item.production == null) return "годовое значение не найдено";
  const unit = item.productionBns?.unit;
  return [item.productionPeriod, unit].filter(Boolean).join(" · ");
}

function fillStatic() {
  const macro = state.data.macro || {};
  const old = macro.comparison2025 || {};
  $("#headlineInflation").textContent = pct(macro.inflation);
  $("#macroPeriod").textContent = [macro.period, macro.publicationDate ? `опубликовано ${dateFmt(macro.publicationDate)}` : ""].filter(Boolean).join(" · ");
  $("#contributionShare").textContent = pp(macro.foodContribution);
  $("#inflation").textContent = pct(macro.inflation);
  $("#foodInflation").textContent = pct(macro.foodInflation);
  $("#foodContribution").textContent = num1(macro.foodContribution);
  $("#monthlyFood").textContent = pct(macro.monthlyFoodInflation);
  $("#inflationCompare").textContent = macro.inflationFoot || `июнь 2025: ${pct(old.inflation)}`;
  $("#foodCompare").textContent = macro.foodFoot || `продовольственные товары: ${pct(macro.foodInflation)}`;
  $("#contributionCompare").textContent = macro.contributionFoot || "вклад в годовую инфляцию";
  const monthlyFoodFoot = document.querySelector("#monthlyFood")?.closest(".kpi-card")?.querySelector(".kpi-foot");
  if (monthlyFoodFoot) monthlyFoodFoot.textContent = macro.monthlyFoodFoot || "продовольственные товары";
  maybeSetText("#updatedDate", dateFmt(state.data.meta.updated));
  maybeSetText("#productCountLabel", `${state.data.meta.productCount} СЗПТ`);
  $("#detailTotal").textContent = state.data.meta.productCount;

  const options = productOptions();
  $("#marketProductSelect").innerHTML = options;
  $("#detailProductSelect").innerHTML = options;
  renderInflationGroups();
}

function bind() {
  $("#themeToggle").addEventListener("click", () => document.body.classList.toggle("dark"));
  ["marketProductSelect", "detailProductSelect"].forEach(id => {
    $(`#${id}`).addEventListener("change", event => selectProduct(event.target.value, id === "detailProductSelect"));
  });
  const searchInput = $("#searchInput");
  if (searchInput) {
    searchInput.addEventListener("input", event => {
      state.query = event.target.value.trim().toLowerCase();
      renderTable();
    });
  }
  document.querySelectorAll("[data-table-filter]").forEach(button => {
    button.addEventListener("click", () => {
      document.querySelectorAll("[data-table-filter]").forEach(item => item.classList.remove("active"));
      button.classList.add("active");
      state.tableFilter = button.dataset.tableFilter;
      renderTable();
    });
  });
}

function selectProduct(id, scroll = false) {
  state.selected = id;
  $("#marketProductSelect").value = id;
  $("#detailProductSelect").value = id;
  renderMarket();
  renderDetail();
  if (scroll) location.hash = "product-detail";
}

function renderAll() {
  renderMarket();
  renderRanking();
  renderTable();
  renderStocksOverview();
  renderDetail();
}

function renderInflationGroups() {
  const groups = state.data.macro.groups || [];
  if (!groups.length) {
    $("#inflationGroups").innerHTML = `<div class="empty-state">Публикация БНС по инфляции не загружена.</div>`;
    return;
  }
  const maxContribution = Math.max(...groups.map(item => Math.abs(item.contribution || 0)), 1);
  $("#inflationGroups").innerHTML = groups.map(item => `
    <div class="group-row ${item.contribution < 0 ? "negative" : ""}">
      <div><strong>${item.name}</strong><span>${pct(item.inflation)} за год</span></div>
      <div class="group-track"><i style="width:${Math.min(Math.abs(item.contribution || 0) / maxContribution * 100, 100)}%"></i></div>
      <b>${pp(item.contribution)}</b>
    </div>
  `).join("");
}

function safeDate(iso) {
  return iso ? dateFmt(iso) : "н/д";
}

function stockMainCategory(row) {
  const entries = Object.entries(row.categoryTotals || {}).filter(([, value]) => value > 0);
  if (!entries.length) return "не указано";
  entries.sort((a, b) => b[1] - a[1]);
  const [key, value] = entries[0];
  return `${stockCategoryLabels[key] || key}: ${fmt.format(value)}`;
}

function renderStocksOverview() {
  const stocks = state.data.stocks || {};
  const enabled = Boolean(stocks.enabled);
  $("#stocksUpdated").textContent = enabled ? safeDate(stocks.updated) : "нет данных";
  $("#stocksRegionCount").textContent = enabled ? `${stocks.receivedRegions || 0}/${stocks.expectedRegions || 0} регионов` : "0 регионов";
  $("#stocksReceived").textContent = enabled ? `${stocks.receivedRegions || 0}/${stocks.expectedRegions || 0}` : "0";
  $("#stocksTotalTons").textContent = stocks.totalTons == null ? "н/д" : fmt.format(stocks.totalTons);
  $("#stocksEggs").textContent = stocks.eggThousandPieces == null ? "н/д" : fmt.format(stocks.eggThousandPieces);
  $("#stocksMissingCount").textContent = fmt.format((stocks.missingRegions || []).length);

  const rows = stocks.productRows || [];
  if (!rows.length) {
    $("#stocksTable").innerHTML = `<tr><td colspan="5"><div class="empty-state">Файлы регионов ещё не поступили через Telegram-бот.</div></td></tr>`;
  } else {
    $("#stocksTable").innerHTML = rows.map(row => `
      <tr>
        <td><b>${row.product}</b><small class="muted-line">${row.sourceScope || "оперативная группа МИО"}</small></td>
        <td>${fmt.format(row.total || 0)} <small>${row.unit || "тонн"}</small></td>
        <td>${fmt.format(row.regionCount || 0)}</td>
        <td>${safeDate(row.latestDate)}</td>
        <td>${stockMainCategory(row)}</td>
      </tr>
    `).join("");
  }

  const missing = stocks.missingRegions || [];
  $("#stocksMissing").innerHTML = missing.length
    ? missing.map(region => `<span>${region}</span>`).join("")
    : `<div class="empty-state small">Все ожидаемые регионы сдали файл.</div>`;
}

function renderMarket() {
  const item = product();
  const [label, cls] = coverageStatus(item.coverage);
  const balanceUnit = item.balanceUnit || item.tradeUnit || "тонн";
  $("#coverageProduct").textContent = item.name;
  $("#coverageStatus").textContent = label;
  $("#coverageStatus").className = `status-badge ${cls}`;
  $("#coverageValue").textContent = pct(item.coverage);
  $("#coverageBar").style.width = `${Math.min(item.coverage || 0, 130) / 1.3}%`;
  $("#importShare").textContent = importShareText(item.importShare);
  $("#consumption").textContent = quantityText(item.consumption, balanceUnit);
  $("#production").textContent = item.production == null ? "н/д" : fmt.format(item.production);
  $("#productionPeriod").textContent = productionPeriodLabel(item);
  $("#export2025").textContent = quantityText(item.export2025, balanceUnit);
  $("#import2025").textContent = quantityText(item.import2025, balanceUnit);
  const flowUnitNote = document.querySelector("#flowUnitNote");
  if (flowUnitNote) flowUnitNote.textContent = unitShort(balanceUnit);
  $("#tradeTitle").textContent = `${item.name}: экспорт, импорт и средняя цена`;
  renderTrade(item);
}

function renderTrade(item) {
  const trade = item.trade2026;
  const months = trade.months || [];
  const exportQty = trade.export || [];
  const importQty = trade.import || [];
  const exportPrice = trade.exportUsdPerTon || [];
  const importPrice = trade.importUsdPerTon || [];
  const tradeUnit = trade.unit || item.tradeUnit || "тонн";
  const priceUnit = trade.priceUnit || item.tradePriceUnit || "$/т";
  const max = Math.max(...exportQty, ...importQty, 1);
  if (!months.length) {
    $("#tradeChart").innerHTML = `<div class="empty-state">По текущему году нет сопоставленных строк ТН ВЭД.</div>`;
    const note = document.querySelector("#tradeSourceNote");
    if (note) note.textContent = "Текущий экспорт/импорт появится после сопоставления кода ТН ВЭД.";
    return;
  }
  $("#tradeChart").innerHTML = months.map((month, i) => `
    <div class="trade-month">
      <div class="trade-bars">
        <i class="export-bar" style="height:${Math.max(3, (exportQty[i] || 0) / max * 100)}%"></i>
        <i class="import-bar" style="height:${Math.max(3, (importQty[i] || 0) / max * 100)}%"></i>
      </div>
      <strong>${month}</strong>
      <small>${fmt.format(exportQty[i] || 0)} / ${fmt.format(importQty[i] || 0)} ${unitShort(tradeUnit)}</small>
      <em>${priceUnit}: ${exportPrice[i] ? money.format(exportPrice[i]) : "н/д"} / ${importPrice[i] ? money.format(importPrice[i]) : "н/д"}</em>
    </div>
  `).join("");
  const note = document.querySelector("#tradeSourceNote");
  if (note) {
    const productionNote = item.operationalProduction == null
      ? ""
      : ` · производство: ${quantityText(item.operationalProduction, item.productionBns?.unit)} (${item.operationalProductionPeriod})`;
    note.textContent = `${item.tradeCurrentPeriod || state.data.meta.tradePeriod2026 || "2026"} · ${unitShort(tradeUnit)} · средняя цена ${priceUnit}${productionNote}`;
  }
}

function renderRanking() {
  const items = [...state.data.products].sort((a, b) => b.riskScore - a.riskScore).slice(0, 8);
  $("#coverageRanking").innerHTML = items.map(item => {
    const cls = riskClass(item.riskScore);
    return `<button data-id="${item.id}">
      <span>${item.name}</span>
      <div class="rank-track"><i class="${cls}" style="width:${Math.min(item.riskScore, 100)}%"></i></div>
      <strong>${score(item.riskScore)}</strong>
    </button>`;
  }).join("");
  document.querySelectorAll("#coverageRanking button").forEach(button => {
    button.addEventListener("click", () => selectProduct(button.dataset.id, true));
  });
}

function renderTable() {
  if (!$("#productsTable")) return;
  const items = visibleProducts();
  $("#tableCount").textContent = `${items.length} позиций · сортировка по риску`;
  $("#productsTable").innerHTML = items.map(item => {
    const [status, coverageCls] = coverageStatus(item.coverage);
    const rCls = riskClass(item.riskScore);
    return `<tr data-id="${item.id}">
      <td><b>${String(item.number).padStart(2, "0")}</b> ${item.name}<small class="muted-line">${item.riskLabel}</small></td>
      <td>${item.price == null ? "нет котировки" : `${fmt.format(item.price)} <small>${item.unit}</small>`}</td>
      <td><span class="delta ${deltaClass(item.weekChange)}">${signed(item.weekChange)}</span><small class="metric-note">год: ${signed(item.annualChange)}</small></td>
      <td><span class="coverage-pill ${coverageCls}">${item.coverage == null ? "н/д" : pct(item.coverage)}</span><small>${status}</small></td>
      <td><span class="risk-pill ${rCls}">${score(item.riskScore)}</span></td>
      <td>${quantityText(item.import2025, item.balanceUnit || item.tradeUnit || "тонн")}</td>
      <td><button class="open-row" type="button">Открыть</button></td>
    </tr>`;
  }).join("");
  document.querySelectorAll("#productsTable tr").forEach(row => {
    row.addEventListener("click", () => selectProduct(row.dataset.id, true));
  });
}

function renderDetail() {
  const item = product();
  const tradeBalance = (item.export2025 ?? 0) - (item.import2025 ?? 0);
  $("#detailProductSelect").value = item.id;
  $("#marketProductSelect").value = item.id;
  $("#detailNumber").textContent = String(item.number).padStart(2, "0");
  $("#detailName").textContent = item.name;
  $("#detailSummary").textContent = detailSummary(item);
  $("#detailPrice").textContent = item.price == null ? "н/д" : fmt.format(item.price);
  $("#detailUnit").textContent = item.price == null ? "еженедельная котировка отсутствует" : item.unit;
  $("#detailWeekChange").innerHTML = priceDelta("неделя", item.weekChange);
  $("#detailYtdChange").innerHTML = priceDelta("с начала года", item.yearChange);
  $("#detailAnnualChange").innerHTML = priceDelta("за год", item.annualChange);
  $("#detailCoverage").textContent = pct(item.coverage);
  $("#detailProduction").textContent = item.production == null ? "н/д" : fmt.format(item.production);
  const productionUnit = document.querySelector("#detailProductionUnit");
  if (productionUnit) productionUnit.textContent = productionPeriodLabel(item);
  $("#detailBalance").textContent = item.export2025 == null || item.import2025 == null
    ? "н/д"
    : `${tradeBalance >= 0 ? "+" : "−"}${quantityText(Math.abs(tradeBalance), item.balanceUnit || item.tradeUnit || "тонн")}`;
  renderPriceTrend(item);
  renderProductionCompare(item);
  renderStockDetail(item);
  renderRegions(item);
}

function detailSummary(item) {
  const parts = [`Риск-рейтинг: ${score(item.riskScore)} (${item.riskLabel.toLowerCase()}).`];
  if (item.coverage != null) parts.push(`Обеспеченность по балансу 2025: ${fmt1.format(item.coverage)}%, импорт к потреблению: ${importShareText(item.importShare)}.`);
  if (item.annualChange != null) parts.push(`Годовое изменение цены: ${signed(item.annualChange)}.`);
  const sourceNote = coverageSourceNote(item);
  if (sourceNote) parts.push(sourceNote);
  if ((item.importShare ?? 0) > 100) parts.push("Показатель выше 100% возможен в балансовых данных из-за запасов, переработки или расхождений периодов.");
  return parts.join(" ");
}

function renderPriceTrend(item) {
  const years = ["2024", "2025", "2026"];
  const series = years.map(year => ({
    year,
    points: (item.priceByYear?.[year] || []).filter(point => point.price != null),
  })).filter(row => row.points.length);

  if (!series.length) return renderWeeklyTrend(item);

  const allPrices = series.flatMap(row => row.points.map(point => point.price));
  const min = Math.min(...allPrices) * 0.97;
  const max = Math.max(...allPrices) * 1.03;
  const width = 700;
  const height = 260;
  const padX = 60;
  const padRight = 48;
  const padY = 30;
  const padBottom = 34;
  const colors = { "2024": "#2563eb", "2025": "#15803d", "2026": "#f97316" };
  const plotRight = width - padRight;
  const plotBottom = height - padBottom;
  const x = day => padX + ((day || 1) - 1) * ((plotRight - padX) / 365);
  const y = value => padY + (max - value) * ((plotBottom - padY) / Math.max(max - min, 1));
  const yTicks = [0, 1, 2, 3, 4].map(i => max - i * (max - min) / 4);
  const monthStarts = [
    [1, "01"], [32, "02"], [60, "03"], [91, "04"], [121, "05"], [152, "06"],
    [182, "07"], [213, "08"], [244, "09"], [274, "10"], [305, "11"], [335, "12"]
  ];
  const lines = series.map(row => {
    const points = row.points.map((point, index) => [x(point.dayOfYear || (index + 1) * 14), y(point.price), point]);
    const last = points[points.length - 1];
    const sparse = points.filter((_, index) => index === 0 || index === points.length - 1 || index % 8 === 0);
    return `<polyline class="line y${row.year}" stroke="${colors[row.year]}" points="${points.map(point => `${point[0]},${point[1]}`).join(" ")}"/>
      ${sparse.map(point => `<circle class="dot y${row.year}" fill="#fff" stroke="${colors[row.year]}" cx="${point[0]}" cy="${point[1]}" r="4"><title>${row.year}: ${point[2].label} — ${fmt.format(point[2].price)}</title></circle>`).join("")}
      ${points.map(point => `<circle class="hit-dot" cx="${point[0]}" cy="${point[1]}" r="8"><title>${row.year}: ${point[2].label} — ${fmt.format(point[2].price)}</title></circle>`).join("")}
      <text class="year-tag" x="${Math.min(last[0] + 8, width - 28)}" y="${last[1] + 4}" fill="${colors[row.year]}">${row.year}</text>`;
  }).join("");
  const legend = series.map(row => `<span><i style="background:${colors[row.year]}"></i>${row.year}</span>`).join("");
  const missingYears = years.filter(year => !series.some(row => row.year === year));
  const missingNote = missingYears.length ? ` Нет точек БНС за: ${missingYears.join(", ")}.` : "";

  $("#priceTrend").innerHTML = `
    <div class="chart-legend">${legend}</div>
    <svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
      <line class="axis-line" x1="${padX}" x2="${padX}" y1="${padY}" y2="${plotBottom}"/>
      <line class="axis-line" x1="${padX}" x2="${plotRight}" y1="${plotBottom}" y2="${plotBottom}"/>
      <text class="axis-unit" x="${padX}" y="14">${item.unit || "тг/кг"}</text>
      ${yTicks.map(value => `<line class="grid-line" x1="${padX}" x2="${plotRight}" y1="${y(value)}" y2="${y(value)}"/><text class="axis-label" x="${padX - 8}" y="${y(value) + 4}" text-anchor="end">${fmt.format(value)}</text>`).join("")}
      ${monthStarts.map(([day, label]) => `<text x="${x(day)}" y="${height - 8}" text-anchor="middle">${label}</text>`).join("")}
      ${lines}
    </svg>
    <p class="chart-note">Сравнение строится по недельным публикациям БНС Т-15-05-Н. Левая ось показывает цену (${item.unit || "тг/кг"}); подпись справа показывает год линии.${missingNote} Точки оставлены выборочно, чтобы график не превращался в “бисер”.</p>
  `;
}

function renderWeeklyTrend(item) {
  const points = item.history.filter(point => point.price != null);
  if (!points.length) {
    $("#priceTrend").innerHTML = `<div class="empty-state">Цена БНС по этой позиции не найдена в загруженных файлах.</div>`;
    return;
  }
  const values = points.map(point => point.price);
  const min = Math.min(...values) * 0.98;
  const max = Math.max(...values) * 1.02;
  const width = 650;
  const height = 235;
  const padX = 60;
  const padRight = 28;
  const padY = 30;
  const padBottom = 34;
  const plotRight = width - padRight;
  const plotBottom = height - padBottom;
  const x = index => padX + index * ((plotRight - padX) / Math.max(points.length - 1, 1));
  const y = value => padY + (max - value) * ((plotBottom - padY) / Math.max(max - min, 1));
  const yTicks = [0, 1, 2, 3].map(i => max - i * (max - min) / 3);
  const coords = points.map((point, index) => [x(index), y(point.price)]);
  const note = state.data.meta.priceSourceNote || "График строится по загруженным публикациям БНС.";
  $("#priceTrend").innerHTML = `<svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
    <line class="axis-line" x1="${padX}" x2="${padX}" y1="${padY}" y2="${plotBottom}"/>
    <line class="axis-line" x1="${padX}" x2="${plotRight}" y1="${plotBottom}" y2="${plotBottom}"/>
    <text class="axis-unit" x="${padX}" y="14">${item.unit || "тг/кг"}</text>
    ${yTicks.map(value => `<line class="grid-line" x1="${padX}" x2="${plotRight}" y1="${y(value)}" y2="${y(value)}"/><text class="axis-label" x="${padX - 8}" y="${y(value) + 4}" text-anchor="end">${fmt.format(value)}</text>`).join("")}
    <polyline class="line" points="${coords.map(coord => coord.join(",")).join(" ")}"/>
    ${coords.map((coord, index) => `<circle class="dot" cx="${coord[0]}" cy="${coord[1]}" r="5"/><text x="${coord[0]}" y="${coord[1] - 13}" text-anchor="middle">${fmt.format(points[index].price)}</text><text x="${coord[0]}" y="${height - 8}" text-anchor="middle">${new Date(points[index].date + "T00:00:00").toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit" })}</text>`).join("")}
  </svg><p class="chart-note">${note} Для сравнения с 2025/2024 добавьте архивные недельные файлы БНС в формате .xlsx.</p>`;
}

function renderProductionCompare(item) {
  const prod = item.productionBns;
  if (!prod) {
    $("#productionCompare").innerHTML = `<div class="empty-state">Официальная строка БНС/Талдау по производству для этого СЗПТ пока не подключена.</div>`;
    return;
  }
  const current = prod.current || {};
  const annual = prod.annual || {};
  const rows = [];
  if (annual["2024"] != null) rows.push({ label: "2024 год", value: annual["2024"], cls: "" });
  if (annual["2025"] != null) rows.push({ label: "2025 год", value: annual["2025"], cls: "current" });
  if (current.value != null && current.period && !rows.some(row => row.label === current.period)) {
    rows.push({ label: current.period, value: current.value, cls: annual["2025"] == null ? "current" : "" });
  }
  if (current.previousYearPeriod != null) {
    const previousLabel = comparablePeriodLabel(current.period);
    if (!rows.some(row => row.label === previousLabel)) {
      rows.push({ label: previousLabel, value: current.previousYearPeriod, cls: "" });
    }
  }
  if (!rows.length) {
    $("#productionCompare").innerHTML = `<div class="empty-state">Официальное значение производства не найдено.</div>`;
    return;
  }
  const max = Math.max(...rows.map(row => row.value || 0), 1);
  const tempo = current.changePct == null ? null : current.changePct - 100;
  $("#productionCompare").innerHTML = `
    ${rows.map(row => `<div class="compare-row ${row.cls}"><span>${row.label}</span><div><i style="width:${row.value / max * 100}%"></i></div><strong>${fmt.format(row.value)}</strong></div>`).join("")}
    <p>БНС · ${prod.unit || "тонн"} · темп к аналогичному периоду: ${pct(tempo)}</p>
    <p class="chart-note">${(prod.sourceRows || []).join("; ")}</p>`;
}

function renderStockDetail(item) {
  const stock = item.operationalStocksMio;
  if (!stock) {
    $("#stockDetail").innerHTML = `<div class="empty-state">По выбранному СЗПТ пока нет сопоставленной строки в оперативных запасах МИО.</div>`;
    return;
  }
  const regions = (stock.regions || []).slice(0, 6);
  const categories = Object.entries(stock.categoryTotals || {})
    .filter(([, value]) => value > 0)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 4);
  $("#stockDetail").innerHTML = `
    <div class="stock-detail-total">
      <span>${stock.product}</span>
      <strong>${fmt.format(stock.total || 0)}</strong>
      <small>${stock.unit || "тонн"} · ${fmt.format(stock.regionCount || 0)} регионов · ${safeDate(stock.latestDate)}</small>
    </div>
    <div class="stock-detail-grid">
      ${categories.map(([key, value]) => `<div><span>${stockCategoryLabels[key] || key}</span><strong>${fmt.format(value)}</strong></div>`).join("")}
    </div>
    <div class="stock-region-list">
      ${regions.map(row => `<p><span>${row.region}</span><b>${fmt.format(row.total || 0)} ${stock.unit || "тонн"}</b></p>`).join("")}
    </div>
    <p class="chart-note">Оперативные данные МИО из Telegram-бота; не смешиваются с официальной статистикой БНС/Талдау.</p>
  `;
}

function comparablePeriodLabel(period) {
  if (!period) return "аналогичный период прошлого года";
  const match = String(period).match(/(20\d{2})/);
  if (!match) return "аналогичный период прошлого года";
  return String(period).replace(match[1], String(Number(match[1]) - 1));
}

function renderRegions(item) {
  const entries = Object.entries(item.regions).filter(([, value]) => value != null).sort((a, b) => b[1] - a[1]);
  if (!entries.length) {
    $("#regionBars").innerHTML = `<div class="empty-state">Региональная цена БНС по этой позиции отсутствует.</div>`;
    return;
  }
  const max = Math.max(...entries.map(row => row[1]));
  $("#regionBars").innerHTML = entries.map(([region, value]) => `<div class="region-row"><span title="${region}">${region}</span><div class="bar-track"><div class="bar-fill" style="width:${value / max * 100}%"></div></div><strong>${fmt.format(value)}</strong></div>`).join("");
}

init().catch(error => {
  document.body.innerHTML = `<main style="padding:40px"><h1>Не удалось загрузить данные</h1><p>${error.message}</p></main>`;
});
