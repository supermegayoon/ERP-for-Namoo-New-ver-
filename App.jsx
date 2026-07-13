
import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const ROLE_LABELS = {
  Admin: "관리자",
  Manager: "매니저",
  Sales: "영업",
  Sourcing: "소싱/원가",
  Production: "생산",
  QA: "QA",
  Viewer: "조회전용",
};

const MENU_GROUPS = [
  { title: "메인", items: [{ id: "dashboard", label: "대시보드", icon: "📊", roles: Object.keys(ROLE_LABELS) }] },
  { title: "영업 관리", items: [{ id: "orders", label: "주문 관리", icon: "📋", roles: ["Admin", "Manager", "Sales"] }, { id: "customers", label: "거래처 관리", icon: "👥", roles: ["Admin", "Manager", "Sales"] }] },
  { title: "생산 관리", items: [{ id: "styles", label: "상품 관리", icon: "👕", roles: ["Admin", "Manager", "Sales", "Sourcing"] }, { id: "production", label: "생산 관리", icon: "🏭", roles: ["Admin", "Manager", "Production", "QA"] }, { id: "inventory", label: "재고 관리", icon: "📦", roles: ["Admin", "Manager", "Production", "Sourcing"] }] },
  { title: "원가 · 재무", items: [{ id: "costing", label: "원가 관리", icon: "💰", roles: ["Admin", "Manager", "Sourcing"] }, { id: "sales", label: "매출 · 수익", icon: "📈", roles: ["Admin", "Manager"] }] },
  { title: "구매", items: [{ id: "suppliers", label: "공급업체", icon: "🏬", roles: ["Admin", "Manager", "Sourcing"] }, { id: "purchase", label: "발주 관리", icon: "🛒", roles: ["Admin", "Manager", "Sourcing"] }] },
  { title: "관리자", items: [{ id: "users", label: "사용자 관리", icon: "🔐", roles: ["Admin"] }] },
];

const STAGES = [
  { key: "quote", label: "최초 견적" },
  { key: "nego1", label: "1차 네고" },
  { key: "final", label: "최종 원가" },
];

const uid = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const num = (value) => Number.isFinite(Number(value)) ? Number(value) : 0;
const usd = (value) => `$${num(value).toFixed(2)}`;
const krw = (value) => `₩${Math.round(num(value)).toLocaleString()}`;

function emptyStage(currency = "USD") {
  return { usage: 0, unitPrice: 0, currency };
}

function createCostRow(category = "원자재") {
  const defaultUnit = category === "원자재" ? "YD" : category === "공정" ? "EA" : "EA";
  return {
    id: uid(),
    category,
    item: "",
    description: "",
    supplier: "",
    spec: "",
    unit: defaultUnit,
    quote: emptyStage("USD"),
    nego1: emptyStage("USD"),
    final: emptyStage("USD"),
    remark: "",
  };
}

const SAMPLE_COST_SHEET = {
  id: "COST-NJ5JS53JKL",
  header: {
    division: "Garment",
    brand: "WHITE LABEL",
    productionType: "완사입",
    styleNo: "NJ5JS53JKL",
    itemName: "COMFY TRAINING ZIP UP",
    quantity: 4800,
    factory: "NAMOO",
    date: "2026-04-14",
    sizeRange: "90(S) ~ 110(2XL)",
    exchangeRate: 1470,
    retailPrice: "",
    country: "Vietnam",
  },
  rows: [
    {
      id: uid(), category: "원자재", item: "SHELL A", description: "ST-TH-3745_YOC",
      supplier: "SOLTTEX", spec: '58"', unit: "YD",
      quote: { usage: 1.45, unitPrice: 7500, currency: "KRW" },
      nego1: { usage: 1.185, unitPrice: 7500, currency: "KRW" },
      final: { usage: 1.45, unitPrice: 7500, currency: "KRW" },
      remark: "요척 재확인",
    },
    {
      id: uid(), category: "원자재", item: "LINING A", description: "20039-3_YOC (T-24)",
      supplier: "DAEKYUNG", spec: '58"', unit: "YD",
      quote: { usage: 0.24, unitPrice: 1470, currency: "KRW" },
      nego1: { usage: 0.206, unitPrice: 1470, currency: "KRW" },
      final: { usage: 0.24, unitPrice: 1470, currency: "KRW" },
      remark: "",
    },
    {
      id: uid(), category: "부자재", item: "부속자재", description: "N#3 REVERSE 2-WAY ZIPPER",
      supplier: "YKK", spec: "", unit: "EA",
      quote: { usage: 1, unitPrice: 1810, currency: "KRW" },
      nego1: { usage: 1, unitPrice: 1810, currency: "KRW" },
      final: { usage: 1, unitPrice: 847, currency: "KRW" },
      remark: "지퍼 단가 업데이트",
    },
    {
      id: uid(), category: "부자재", item: "봉사", description: "THREAD TEX 40 (40/3)",
      supplier: "NAMOO", spec: "5000M", unit: "CONE",
      quote: { usage: 0.15, unitPrice: 2, currency: "USD" },
      nego1: { usage: 0.15, unitPrice: 2, currency: "USD" },
      final: { usage: 0.15, unitPrice: 2, currency: "USD" },
      remark: "",
    },
    {
      id: uid(), category: "부자재", item: "자수/프린트/와펜", description: "LOGO EMBO (FRONT / BACK)",
      supplier: "NAMOO", spec: "5 CM", unit: "EA",
      quote: { usage: 2, unitPrice: 0.4, currency: "USD" },
      nego1: { usage: 2, unitPrice: 0.4, currency: "USD" },
      final: { usage: 2, unitPrice: 0.4, currency: "USD" },
      remark: "",
    },
    {
      id: uid(), category: "부자재", item: "포장자재", description: "POLY BAG",
      supplier: "NAMOO", spec: "", unit: "EA",
      quote: { usage: 1, unitPrice: 0.05, currency: "USD" },
      nego1: { usage: 1, unitPrice: 0.05, currency: "USD" },
      final: { usage: 1, unitPrice: 0.05, currency: "USD" },
      remark: "",
    },
  ],
  stageCosts: {
    quote: { localCharge: 0, pureSewing: 6, marginRate: 15, dutyFreightRate: 3, suppliedMaterial: 0, vatRate: 10 },
    nego1: { localCharge: 0, pureSewing: 5.5, marginRate: 15, dutyFreightRate: 3, suppliedMaterial: 0, vatRate: 10 },
    final: { localCharge: 0, pureSewing: 6, marginRate: 15, dutyFreightRate: 3, suppliedMaterial: 0, vatRate: 10 },
  },
};

const INITIAL_DATA = {
  users: [
    { id: "admin", password: "1234", name: "Yoon", role: "Admin", department: "Management", status: "Active" },
    { id: "qa01", password: "1111", name: "QA User", role: "QA", department: "QA", status: "Active" },
  ],
  styles: [
    { styleNo: "TG-26001", buyer: "Kohl's", item: "Men's Fleece Hoodie", season: "SP26", factory: "Minh Duc", status: "개발중" },
    { styleNo: "HCO-1188", buyer: "A&F", item: "Women's Rib Tank", season: "SU26", factory: "EDC Sample", status: "샘플진행" },
    { styleNo: "FLX-7720", buyer: "Kohl's", item: "Girls Dress", season: "FA26", factory: "TDT", status: "생산중" },
  ],
  orders: [
    { po: "45001234", styleNo: "TG-26001", buyer: "Kohl's", qty: 12000, delivery: "2026-08-15", status: "Open" },
    { po: "45001235", styleNo: "HCO-1188", buyer: "A&F", qty: 8500, delivery: "2026-07-20", status: "Review" },
    { po: "45001236", styleNo: "FLX-7720", buyer: "Kohl's", qty: 22000, delivery: "2026-09-10", status: "Production" },
  ],
  production: [
    { styleNo: "TG-26001", factory: "Minh Duc", cutting: 35, sewing: 20, packing: 0, inspection: "대기", issue: "원단 입고 확인중" },
    { styleNo: "HCO-1188", factory: "EDC Sample", cutting: 100, sewing: 100, packing: 45, inspection: "진행중", issue: "Buyer comment 대기" },
    { styleNo: "FLX-7720", factory: "TDT", cutting: 80, sewing: 62, packing: 20, inspection: "대기", issue: "부자재 일부 지연" },
  ],
  inventory: [
    { material: "Cotton Poly Fleece", type: "원단", stock: 35000, unit: "YDS", status: "충분" },
    { material: "Main Label", type: "부자재", stock: 12000, unit: "PCS", status: "주의" },
  ],
  customers: [
    { buyer: "Kohl's", brand: "Tek Gear / FLX / Sonoma", country: "USA", status: "Active" },
    { buyer: "A&F", brand: "HCO / ANF", country: "USA", status: "Active" },
  ],
  suppliers: [
    { name: "Nice Dye", type: "원단", country: "Vietnam", status: "Active" },
    { name: "R-pac", type: "Label/Trim", country: "Vietnam", status: "Active" },
  ],
  purchase: [
    { po: "MAT-001", supplier: "Nice Dye", material: "Cotton Poly Fleece", qty: 35000, status: "입고 예정" },
  ],
  costSheets: [SAMPLE_COST_SHEET],
  salesData: {},
};

const STORAGE_KEY = "fashion_erp_namoo_v7";
function loadData() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return INITIAL_DATA;
    const parsed = JSON.parse(saved);
    return { ...INITIAL_DATA, ...parsed, salesData: parsed.salesData || {} };
  } catch {
    return INITIAL_DATA;
  }
}
function persist(data) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); } catch {}
}
function canAccess(role, item) { return item.roles.includes(role); }

function stageAmount(row, stageKey, exchangeRate) {
  const stage = row[stageKey] || emptyStage();
  const raw = num(stage.usage) * num(stage.unitPrice);
  const usdAmount = stage.currency === "KRW" ? raw / Math.max(num(exchangeRate), 1) : raw;
  const krwAmount = stage.currency === "KRW" ? raw : raw * num(exchangeRate);
  return { usd: usdAmount, krw: krwAmount };
}

function calculateStage(sheet, stageKey) {
  const exchangeRate = num(sheet.header.exchangeRate) || 1;
  const sumCategory = (category) => sheet.rows
    .filter((r) => r.category === category)
    .reduce((sum, r) => sum + stageAmount(r, stageKey, exchangeRate).usd, 0);

  const material = sumCategory("원자재");
  const trim = sumCategory("부자재");
  const process = sumCategory("공정");
  const other = sumCategory("기타비용");
  const costs = sheet.stageCosts[stageKey];
  const materialTotal = material + trim;
  const laborBase = process + num(costs.pureSewing);
  const marginAmount = (materialTotal + num(costs.localCharge) + laborBase) * num(costs.marginRate) / 100;
  const cm = laborBase + marginAmount;
  const fobValue = materialTotal + num(costs.localCharge) + cm + other;
  const dutyFreight = fobValue * num(costs.dutyFreightRate) / 100;
  const purchaseVatMinus = fobValue + dutyFreight;
  const manufacturingVatMinus = purchaseVatMinus + num(costs.suppliedMaterial);
  const manufacturingVatPlus = manufacturingVatMinus * (1 + num(costs.vatRate) / 100);
  return {
    material, trim, process, other, materialTotal, marginAmount, cm, fobValue,
    dutyFreight, purchaseVatMinus, manufacturingVatMinus, manufacturingVatPlus,
  };
}


function excelDateToString(value) {
  if (!value) return "";
  if (typeof value === "string") return value;
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const month = String(parsed.m).padStart(2, "0");
      const day = String(parsed.d).padStart(2, "0");
      return `${parsed.y}-${month}-${day}`;
    }
  }
  return String(value);
}

function normalizeCurrency(value) {
  const text = String(value || "USD").trim().toUpperCase();
  return text.includes("KRW") || text.includes("원") ? "KRW" : "USD";
}

function parseBuyerCostWorkbook(workbook) {
  return workbook.SheetNames.map((sheetName) => {
    const ws = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
    const row4 = matrix[3] || [];
    const row5 = matrix[4] || [];
    const row6 = matrix[5] || [];
    let detectedSellingFob = 0;
    for (let ri = 0; ri < Math.min(matrix.length, 30); ri += 1) {
      const row = matrix[ri] || [];
      for (let ci = 0; ci < row.length; ci += 1) {
        const label = String(row[ci] || "").replace(/\s/g, "").toUpperCase();
        if (label.includes("SELLINGFOB") || label.includes("SELLFOB") || label.includes("판매FOB") || label.includes("바이어FOB")) {
          for (let offset = 1; offset <= 3; offset += 1) {
            const candidate = num(row[ci + offset]);
            if (candidate > 0) {
              detectedSellingFob = candidate;
              break;
            }
          }
        }
        if (detectedSellingFob > 0) break;
      }
      if (detectedSellingFob > 0) break;
    }

    const header = {
      division: row4[2] || "",
      brand: row5[2] || "",
      buyer: row5[2] || "",
      productionType: row4[6] || "완사입",
      styleNo: row6[2] || sheetName,
      itemName: row5[6] || "",
      quantity: num(row6[6]),
      factory: row5[11] || "",
      date: excelDateToString(row4[11]),
      sizeRange: row4[15] || "",
      exchangeRate: num(row5[15]) || 1470,
      retailPrice: row6[11] || "",
      sellingFob: detectedSellingFob,
      country: row6[15] || "",
      sourceSheet: sheetName,
    };

    let currentCategory = "";
    const rows = [];
    for (let i = 9; i < matrix.length; i += 1) {
      const r = matrix[i] || [];
      const categoryCell = String(r[0] || "").replace(/\s/g, "");
      const item = String(r[1] || "").trim();
      const description = String(r[2] || "").trim();

      if (categoryCell.includes("원자재")) currentCategory = "원자재";
      if (categoryCell.includes("부자재")) currentCategory = "부자재";
      if (item.includes("원자재 소계") || item.includes("부자재 소계") || item.includes("원부자재계") || categoryCell.includes("원가사항")) {
        if (item.includes("원부자재계") || categoryCell.includes("원가사항")) break;
        continue;
      }
      if (!currentCategory || (!item && !description)) continue;

      rows.push({
        id: uid(),
        category: currentCategory,
        item,
        description,
        supplier: r[3] || "",
        spec: r[4] || "",
        unit: r[5] || "EA",
        quote: { usage: num(r[6]), unitPrice: num(r[7]), currency: normalizeCurrency(r[8]) },
        nego1: { usage: num(r[11]), unitPrice: num(r[12]), currency: normalizeCurrency(r[13]) },
        final: { usage: num(r[16]), unitPrice: num(r[17]), currency: normalizeCurrency(r[18]) },
        remark: r[21] || "",
      });
    }

    const findCostRow = (label) => matrix.find((r) => String((r || [])[1] || "").replace(/\s/g, "").includes(label));
    const sewingRow = findCostRow("순수봉제비") || [];
    const marginRow = findCostRow("업체마진") || [];
    const dutyRow = findCostRow("관세및운송비") || [];
    const suppliedRow = findCostRow("공급분자재") || [];
    const vatMinusRow = findCostRow("제조원가(VAT-)") || [];
    const vatPlusRow = findCostRow("제조원가(VAT+)") || [];

    const stageCosts = {};
    [["quote", 6, 7, 9], ["nego1", 11, 12, 14], ["final", 16, 17, 19]].forEach(([key, rateCol, priceCol, usdCol]) => {
      const vatMinus = num(vatMinusRow[usdCol]);
      const vatPlus = num(vatPlusRow[usdCol]);
      const inferredVat = vatMinus > 0 && vatPlus > 0 ? Math.max(0, (vatPlus / vatMinus - 1) * 100) : 10;
      const marginCell = num(marginRow[rateCol]);
      const dutyCell = num(dutyRow[rateCol]);
      stageCosts[key] = {
        localCharge: 0,
        pureSewing: num(sewingRow[priceCol]),
        marginRate: marginCell <= 1 ? marginCell * 100 : marginCell,
        dutyFreightRate: dutyCell <= 1 ? dutyCell * 100 : dutyCell,
        suppliedMaterial: num(suppliedRow[usdCol]),
        vatRate: Number(inferredVat.toFixed(2)),
      };
    });

    return {
      id: `IMPORT-${sheetName}-${Date.now()}-${Math.random().toString(16).slice(2)}`,
      header,
      rows,
      stageCosts,
      importedAt: new Date().toISOString(),
      source: "Buyer Excel",
    };
  }).filter((sheet) => sheet.header.styleNo);
}

function finalCostMetrics(sheet, sellingFob) {
  const final = calculateStage(sheet, "final");
  const manufacturingCost = final.manufacturingVatPlus;
  const selling = num(sellingFob);
  const unitProfit = selling - manufacturingCost;
  const marginRate = selling > 0 ? unitProfit / selling * 100 : 0;
  const markupRate = manufacturingCost > 0 ? unitProfit / manufacturingCost * 100 : 0;
  const quantity = num(sheet.header.quantity);
  return {
    manufacturingCost,
    manufacturingVatMinus: final.manufacturingVatMinus,
    finalFob: final.fobValue,
    selling,
    unitProfit,
    marginRate,
    markupRate,
    quantity,
    totalSales: selling * quantity,
    totalCost: manufacturingCost * quantity,
    totalProfit: unitProfit * quantity,
  };
}


function latestCostSheets(costSheets) {
  const grouped = new Map();
  (costSheets || []).forEach((sheet, index) => {
    const styleNo = String(sheet.header?.styleNo || sheet.id || `STYLE-${index}`);
    const current = grouped.get(styleNo);
    const candidateOrder = Number(sheet.version || 0) * 10000000000000 +
      new Date(sheet.importedAt || sheet.header?.date || 0).getTime() + index;
    if (!current || candidateOrder >= current.order) {
      grouped.set(styleNo, { sheet, order: candidateOrder });
    }
  });
  return Array.from(grouped.values()).map((entry) => entry.sheet);
}

function marginStatus(rate) {
  if (rate < 0) return { label: "LOSS", className: "status-loss" };
  if (rate < 15) return { label: "CRITICAL", className: "status-critical" };
  if (rate < 20) return { label: "NEED REVIEW", className: "status-review" };
  if (rate >= 30) return { label: "EXCELLENT", className: "status-excellent" };
  return { label: "GOOD", className: "status-good" };
}

function calculateSalesRows(data) {
  const salesData = data.salesData || {};
  return latestCostSheets(data.costSheets).map((sheet) => {
    const styleNo = sheet.header.styleNo;
    const sellingFob = salesData[styleNo]?.sellingFob ?? sheet.header.sellingFob ?? "";
    const metrics = finalCostMetrics(sheet, sellingFob);
    return {
      sheet,
      styleNo,
      sellingFob,
      metrics,
      status: marginStatus(metrics.marginRate),
    };
  });
}

function Badge({ children }) { return <span className="badge">{children}</span>; }

function Login({ data, onLogin }) {
  const [id, setId] = useState("admin");
  const [password, setPassword] = useState("1234");
  const [error, setError] = useState("");
  function submit(e) {
    e.preventDefault();
    const found = data.users.find((u) => u.id === id && u.password === password && u.status === "Active");
    if (!found) return setError("ID 또는 비밀번호를 확인해 주세요.");
    onLogin(found);
  }
  return (
    <div className="login-page">
      <style>{CSS}</style>
      <form className="login-card" onSubmit={submit}>
        <div className="login-logo">👗</div><h1>Fashion ERP</h1><p>관리자 / 직원 로그인</p>
        <label>ID</label><input value={id} onChange={(e) => setId(e.target.value)} />
        <label>Password</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} />
        {error && <div className="login-error">{error}</div>}
        <button>로그인</button><small>관리자: admin / 1234</small>
      </form>
    </div>
  );
}

function SimpleTable({ columns, rows }) {
  return (
    <div className="table-wrap"><table><thead><tr>{columns.map((c) => <th key={c.key}>{c.label}</th>)}</tr></thead>
      <tbody>{rows.length ? rows.map((row, index) => <tr key={row.id || row.styleNo || row.po || row.name || index}>{columns.map((c) => <td key={c.key}>{c.render ? c.render(row) : row[c.key]}</td>)}</tr>) : <tr><td colSpan={columns.length} className="empty">데이터가 없습니다.</td></tr>}</tbody>
    </table></div>
  );
}

function Dashboard({ data, search }) {
  const matched = data.styles.filter((s) => [s.styleNo, s.buyer, s.item, s.factory].join(" ").toLowerCase().includes(search.toLowerCase()));
  const styleSet = new Set(matched.map((s) => s.styleNo));
  const orders = data.orders.filter((o) => !search || styleSet.has(o.styleNo));
  const qty = orders.reduce((sum, o) => sum + num(o.qty), 0);

  const salesRows = calculateSalesRows(data).filter((row) => !search || styleSet.has(row.styleNo));
  const totals = salesRows.reduce((sum, row) => ({
    sales: sum.sales + row.metrics.totalSales,
    cost: sum.cost + row.metrics.totalCost,
    profit: sum.profit + row.metrics.totalProfit,
  }), { sales: 0, cost: 0, profit: 0 });
  const avgMargin = totals.sales > 0 ? totals.profit / totals.sales * 100 : 0;
  const riskRows = [...salesRows]
    .filter((row) => row.metrics.selling > 0)
    .sort((a, b) => a.metrics.marginRate - b.metrics.marginRate)
    .slice(0, 5);

  return (
    <div className="content-grid">
      <div className="stats-grid">
        <div className="stat-card"><span>💵</span><div><p>총 매출</p><h3>{usd(totals.sales)}</h3><small>Selling FOB × Qty</small></div></div>
        <div className="stat-card"><span>📈</span><div><p>총 이익</p><h3 className={totals.profit < 0 ? "negative" : "positive"}>{usd(totals.profit)}</h3><small>매출 - 제조원가</small></div></div>
        <div className="stat-card"><span>％</span><div><p>평균 Margin</p><h3 className={avgMargin < 0 ? "negative" : "positive"}>{avgMargin.toFixed(1)}%</h3><small>총이익 ÷ 총매출</small></div></div>
        <div className="stat-card"><span>📋</span><div><p>총 주문수량</p><h3>{qty.toLocaleString()}</h3><small>PO 기준 전체 수량</small></div></div>
      </div>

      <div className="dashboard-two-column">
        <div className="panel">
          <h2>매출·마진 현황</h2>
          <div className="dashboard-profit-list">
            {salesRows.length ? salesRows.slice(0, 8).map((row) => (
              <div className="dashboard-profit-row" key={row.sheet.id}>
                <div><strong>{row.styleNo}</strong><span>{row.sheet.header.buyer || row.sheet.header.brand} · {row.sheet.header.itemName}</span></div>
                <div className="dashboard-profit-values">
                  <span>{usd(row.metrics.selling)}</span>
                  <strong className={row.metrics.totalProfit < 0 ? "negative" : "positive"}>{row.metrics.marginRate.toFixed(1)}%</strong>
                  <em className={`margin-status ${row.status.className}`}>{row.status.label}</em>
                </div>
              </div>
            )) : <div className="empty">Selling FOB를 입력하면 매출·마진이 표시됩니다.</div>}
          </div>
        </div>

        <div className="panel">
          <h2>마진 점검 필요</h2>
          <div className="risk-list">
            {riskRows.length ? riskRows.map((row) => (
              <div className="risk-row" key={row.sheet.id}>
                <div><strong>{row.styleNo}</strong><span>제조원가 {usd(row.metrics.manufacturingCost)}</span></div>
                <div><strong className={row.metrics.marginRate < 0 ? "negative" : ""}>{row.metrics.marginRate.toFixed(1)}%</strong><em className={`margin-status ${row.status.className}`}>{row.status.label}</em></div>
              </div>
            )) : <div className="empty">마진 점검 데이터가 없습니다.</div>}
          </div>
        </div>
      </div>

      <div className="panel"><h2>최근 주문 현황</h2><SimpleTable rows={orders} columns={[
        {key:"po",label:"PO No."},{key:"buyer",label:"Buyer"},{key:"styleNo",label:"Style No."},
        {key:"qty",label:"수량",render:r=>num(r.qty).toLocaleString()},{key:"delivery",label:"납기"},
        {key:"status",label:"상태",render:r=><Badge>{r.status}</Badge>}
      ]}/></div>
    </div>
  );
}

function TextInput({ value, onChange, type = "text", className = "" }) {
  return <input className={className} type={type} value={value ?? ""} onChange={(e) => onChange(e.target.value)} />;
}


function CostSection({
  title,
  subtitle,
  category,
  rows,
  canEdit,
  exchangeRate,
  changeRow,
  changeStage,
  addRow,
  removeRow,
}) {
  const [open, setOpen] = useState(true);

  return (
    <div className="panel costing-table-panel">
      <div className="section-title row-between clickable-title" onClick={() => setOpen(!open)}>
        <div>
          <h2>{open ? "▾" : "▸"} {title}</h2>
          <span>{subtitle}</span>
        </div>
        {canEdit && (
          <div className="inline-actions" onClick={(e) => e.stopPropagation()}>
            <button onClick={() => addRow(category)}>+ 항목 추가</button>
          </div>
        )}
      </div>

      {open && (
        <div className="compact-cost-list">
          {rows.length === 0 && (
            <div className="empty-section">
              등록된 항목이 없습니다. “+ 항목 추가”를 눌러 직접 입력하세요.
            </div>
          )}

          {rows.map((row, rowIndex) => (
            <div className="compact-cost-row" key={row.id}>
              <div className="compact-row-top">
                <div className="row-number">{rowIndex + 1}</div>
                <label>품목<TextInput value={row.item} onChange={(v) => changeRow(row.id, "item", v)} /></label>
                <label className="wide-field">상세품명<TextInput value={row.description} onChange={(v) => changeRow(row.id, "description", v)} /></label>
                <label>업체/공정처<TextInput value={row.supplier} onChange={(v) => changeRow(row.id, "supplier", v)} /></label>
                <label>규격<TextInput value={row.spec} onChange={(v) => changeRow(row.id, "spec", v)} /></label>
                <label className="unit-field">단위<TextInput value={row.unit} onChange={(v) => changeRow(row.id, "unit", v)} /></label>
                {canEdit && <button className="row-delete" onClick={() => removeRow(row.id)}>×</button>}
              </div>

              <div className="stage-cards">
                {STAGES.map((stage) => {
                  const amount = stageAmount(row, stage.key, exchangeRate);
                  return (
                    <div className={`stage-card stage-card-${stage.key}`} key={stage.key}>
                      <div className="stage-card-title">{stage.label}</div>
                      <div className="stage-input-grid">
                        <label>요척/수량<TextInput type="number" value={row[stage.key].usage} onChange={(v) => changeStage(row.id, stage.key, "usage", v)} /></label>
                        <label>단가<TextInput type="number" value={row[stage.key].unitPrice} onChange={(v) => changeStage(row.id, stage.key, "unitPrice", v)} /></label>
                        <label>화폐
                          <select value={row[stage.key].currency} onChange={(e) => changeStage(row.id, stage.key, "currency", e.target.value)}>
                            <option>USD</option><option>KRW</option>
                          </select>
                        </label>
                      </div>
                      <div className="stage-result">
                        <span>{usd(amount.usd)}</span>
                        <strong>{krw(amount.krw)}</strong>
                      </div>
                    </div>
                  );
                })}
              </div>

              <label className="remark-field">비고<TextInput value={row.remark} onChange={(v) => changeRow(row.id, "remark", v)} /></label>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function CostingPage({ data, updateData, canEdit }) {
  const [selectedId, setSelectedId] = useState(data.costSheets[0]?.id || "");
  const selected = data.costSheets.find((s) => s.id === selectedId) || data.costSheets[0];

  function replaceSheet(nextSheet) {
    updateData({ ...data, costSheets: data.costSheets.map((s) => s.id === nextSheet.id ? nextSheet : s) });
  }

  function newSheet() {
    const sheet = {
      id: uid(),
      header: { division: "", brand: "", productionType: "완사입", styleNo: "", itemName: "", quantity: 0, factory: "", date: new Date().toISOString().slice(0,10), sizeRange: "", exchangeRate: 1470, retailPrice: "", country: "Vietnam" },
      rows: [createCostRow("원자재"), createCostRow("부자재"), createCostRow("공정"), createCostRow("기타비용")],
      stageCosts: Object.fromEntries(STAGES.map((s) => [s.key, { localCharge: 0, pureSewing: 0, marginRate: 15, dutyFreightRate: 3, suppliedMaterial: 0, vatRate: 10 }])),
    };
    updateData({ ...data, costSheets: [sheet, ...data.costSheets] });
    setSelectedId(sheet.id);
  }

  function deleteSheet() {
    if (!selected || !confirm("이 원가견적서를 삭제할까요?")) return;
    const remain = data.costSheets.filter((s) => s.id !== selected.id);
    updateData({ ...data, costSheets: remain });
    setSelectedId(remain[0]?.id || "");
  }

  function changeHeader(key, value) {
    replaceSheet({ ...selected, header: { ...selected.header, [key]: value } });
  }

  function changeRow(rowId, key, value) {
    replaceSheet({ ...selected, rows: selected.rows.map((r) => r.id === rowId ? { ...r, [key]: value } : r) });
  }

  function changeStage(rowId, stageKey, key, value) {
    replaceSheet({ ...selected, rows: selected.rows.map((r) => r.id === rowId ? { ...r, [stageKey]: { ...r[stageKey], [key]: value } } : r) });
  }

  function changeStageCost(stageKey, key, value) {
    replaceSheet({ ...selected, stageCosts: { ...selected.stageCosts, [stageKey]: { ...selected.stageCosts[stageKey], [key]: value } } });
  }

  function addRow(category) {
    replaceSheet({ ...selected, rows: [...selected.rows, createCostRow(category)] });
  }

  function removeRow(rowId) {
    replaceSheet({ ...selected, rows: selected.rows.filter((r) => r.id !== rowId) });
  }


  function importBuyerExcel(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (loadEvent) => {
      try {
        const workbook = XLSX.read(loadEvent.target.result, { type: "array", cellDates: false });
        const importedSheets = parseBuyerCostWorkbook(workbook);
        if (!importedSheets.length) {
          alert("불러올 수 있는 원가 시트를 찾지 못했습니다.");
          return;
        }
        const importedStyles = importedSheets.map((sheet) => ({
          styleNo: sheet.header.styleNo,
          buyer: sheet.header.buyer || sheet.header.brand,
          item: sheet.header.itemName,
          season: "",
          factory: sheet.header.factory,
          status: "원가진행",
        }));
        const existingStyleNos = new Set(data.styles.map((style) => style.styleNo));
        const newStyles = importedStyles.filter((style) => !existingStyleNos.has(style.styleNo));
        const versionedSheets = importedSheets.map((sheet) => {
          const previousCount = data.costSheets.filter((old) => old.header.styleNo === sheet.header.styleNo).length;
          return { ...sheet, version: previousCount + 1 };
        });
        const importedSales = { ...(data.salesData || {}) };
        versionedSheets.forEach((sheet) => {
          if (num(sheet.header.sellingFob) > 0) {
            importedSales[sheet.header.styleNo] = {
              ...(importedSales[sheet.header.styleNo] || {}),
              sellingFob: num(sheet.header.sellingFob),
            };
          }
        });
        const next = {
          ...data,
          costSheets: [...versionedSheets, ...data.costSheets],
          styles: [...newStyles, ...data.styles],
          salesData: importedSales,
        };
        updateData(next);
        setSelectedId(versionedSheets[0].id);
        alert(`${versionedSheets.length}개 시트를 새 원가 버전으로 저장하고 매출·수익에 연결했습니다.`);
      } catch (error) {
        console.error(error);
        alert("Excel을 불러오지 못했습니다. TNF 원가견적서 형식을 확인해 주세요.");
      }
    };
    reader.readAsArrayBuffer(file);
    event.target.value = "";
  }

  function exportCostingExcel() {
    const h = selected.header;
    const aoa = [
      ["<원 가 계 산 서>"],
      [],
      ["Division", h.division, "Garment", h.itemName, "생산구분", h.productionType, "", "", "", "", "날짜", h.date, "", "", "사이즈 전개", h.sizeRange],
      ["브랜드", h.brand, "", "", "품명", h.itemName, "", "", "", "", "생산처", h.factory, "", "", "환율", h.exchangeRate],
      ["품번", h.styleNo, "", "", "수량", h.quantity, "", "", "", "", "소비자가", h.retailPrice, "", "", "제조국", h.country],
      [],
      ["구분","품목","상세품명","업체명","규격","단위",
       "최초견적 요척","최초견적 단가","화폐","금액(USD)","금액(KRW)",
       "1차네고 요척","1차네고 단가","화폐","금액(USD)","금액(KRW)",
       "최종원가 요척","최종원가 단가","화폐","금액(USD)","금액(KRW)","비고"],
    ];
    selected.rows.forEach((row) => {
      aoa.push([
        row.category,row.item,row.description,row.supplier,row.spec,row.unit,
        row.quote.usage,row.quote.unitPrice,row.quote.currency,stageAmount(row,"quote",h.exchangeRate).usd,stageAmount(row,"quote",h.exchangeRate).krw,
        row.nego1.usage,row.nego1.unitPrice,row.nego1.currency,stageAmount(row,"nego1",h.exchangeRate).usd,stageAmount(row,"nego1",h.exchangeRate).krw,
        row.final.usage,row.final.unitPrice,row.final.currency,stageAmount(row,"final",h.exchangeRate).usd,stageAmount(row,"final",h.exchangeRate).krw,
        row.remark,
      ]);
    });
    aoa.push([]);
    STAGES.forEach((stage) => {
      const c = calculateStage(selected, stage.key);
      const sc = selected.stageCosts[stage.key];
      aoa.push([stage.label, "원자재 소계", c.material, "부자재 소계", c.trim, "원부자재계", c.materialTotal, "공정비", c.process, "기타비용", c.other]);
      aoa.push(["", "로컬 차지", sc.localCharge, "순수봉제비", sc.pureSewing, "업체 마진(%)", sc.marginRate, "마진금액", c.marginAmount]);
      aoa.push(["", "임가공료(CM)", c.cm, "소계(FOB)", c.fobValue, "관세 및 운송비", c.dutyFreight, "사입원가(VAT-)", c.purchaseVatMinus]);
      aoa.push(["", "공급분자재", sc.suppliedMaterial, "제조원가(VAT-)", c.manufacturingVatMinus, "제조원가(VAT+)", c.manufacturingVatPlus]);
      aoa.push([]);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{wch:10},{wch:18},{wch:34},{wch:18},{wch:12},{wch:10},...Array(15).fill({wch:13}),{wch:28}];
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, ws, (h.styleNo || "Costing").slice(0,31));
    XLSX.writeFile(book, `${h.styleNo || "Costing"}_원가견적서.xlsx`);
  }

  if (!selected) return <div className="panel"><button className="primary-button" onClick={newSheet}>+ 원가견적서 생성</button></div>;

  const calculations = Object.fromEntries(STAGES.map((s) => [s.key, calculateStage(selected, s.key)]));

  return (
    <div className="costing-page">
      <div className="costing-toolbar">
        <select value={selected.id} onChange={(e) => setSelectedId(e.target.value)}>
          {data.costSheets.map((s) => <option key={s.id} value={s.id}>{s.header.styleNo || "신규"} · V{s.version || 1} · {s.header.itemName || "품명 미정"}</option>)}
        </select>
        <div className="toolbar-actions">
          {canEdit && <button className="primary-button" onClick={newSheet}>+ 새 원가견적서</button>}
          {canEdit && <label className="excel-import">📤 Buyer Excel Import<input type="file" accept=".xlsx,.xls" onChange={importBuyerExcel} /></label>}
          <button className="excel-export" onClick={exportCostingExcel}>📥 Excel 추출</button>
          {canEdit && <button className="danger-button" onClick={deleteSheet}>삭제</button>}
        </div>
      </div>

      <div className="panel costing-header">
        <div className="section-title"><h2>원가계산서 기본정보</h2><span>실제 견적서 Header 항목</span></div>
        <div className="header-grid">
          {[
            ["division","Division"],["brand","브랜드"],["productionType","생산구분"],["styleNo","품번"],
            ["itemName","품명"],["quantity","수량"],["factory","생산처"],["date","날짜"],
            ["sizeRange","사이즈 전개"],["exchangeRate","환율"],["retailPrice","소비자가"],["country","제조국"],
          ].map(([key,label]) => <label key={key}>{label}<TextInput type={["quantity","exchangeRate","retailPrice"].includes(key)?"number":key==="date"?"date":"text"} value={selected.header[key]} onChange={(v)=>changeHeader(key,v)} /></label>)}
        </div>
      </div>

      <CostSection
        title="원자재"
        subtitle="Shell, Lining, Rib, Mesh 등 원하는 원단 항목을 직접 추가할 수 있습니다."
        category="원자재"
        rows={selected.rows.filter((r) => r.category === "원자재")}
        canEdit={canEdit}
        exchangeRate={selected.header.exchangeRate}
        changeRow={changeRow}
        changeStage={changeStage}
        addRow={addRow}
        removeRow={removeRow}
      />
      <CostSection
        title="부자재"
        subtitle="Zipper, Button, Label, Polybag 등 품목명을 자유롭게 입력할 수 있습니다."
        category="부자재"
        rows={selected.rows.filter((r) => r.category === "부자재")}
        canEdit={canEdit}
        exchangeRate={selected.header.exchangeRate}
        changeRow={changeRow}
        changeStage={changeStage}
        addRow={addRow}
        removeRow={removeRow}
      />
      <CostSection
        title="공정(CM)"
        subtitle="Cutting, Sewing, Iron, Packing, Wash, Bonding 등 공정을 직접 추가할 수 있습니다."
        category="공정"
        rows={selected.rows.filter((r) => r.category === "공정")}
        canEdit={canEdit}
        exchangeRate={selected.header.exchangeRate}
        changeRow={changeRow}
        changeStage={changeStage}
        addRow={addRow}
        removeRow={removeRow}
      />
      <CostSection
        title="기타비용"
        subtitle="Testing, Freight, Commission, Sample, Development 등 기타 비용을 직접 추가할 수 있습니다."
        category="기타비용"
        rows={selected.rows.filter((r) => r.category === "기타비용")}
        canEdit={canEdit}
        exchangeRate={selected.header.exchangeRate}
        changeRow={changeRow}
        changeStage={changeStage}
        addRow={addRow}
        removeRow={removeRow}
      />

      <div className="summary-grid">
        {STAGES.map((stage)=>{
          const c=calculations[stage.key], sc=selected.stageCosts[stage.key];
          return <div className="panel stage-summary" key={stage.key}>
            <h3>{stage.label}</h3>
            <div className="summary-line"><span>원자재 소계</span><strong>{usd(c.material)}</strong></div>
            <div className="summary-line"><span>부자재 소계</span><strong>{usd(c.trim)}</strong></div>
            <div className="summary-line total"><span>원부자재계</span><strong>{usd(c.materialTotal)}</strong></div>
            <div className="summary-line"><span>추가 공정비</span><strong>{usd(c.process)}</strong></div>
            <div className="summary-line"><span>기타비용</span><strong>{usd(c.other)}</strong></div>
            <label>로컬 차지<TextInput type="number" value={sc.localCharge} onChange={(v)=>changeStageCost(stage.key,"localCharge",v)} /></label>
            <label>순수봉제비(SMV×Rate)<TextInput type="number" value={sc.pureSewing} onChange={(v)=>changeStageCost(stage.key,"pureSewing",v)} /></label>
            <label>업체 마진(%)<TextInput type="number" value={sc.marginRate} onChange={(v)=>changeStageCost(stage.key,"marginRate",v)} /></label>
            <div className="summary-line"><span>업체 마진금액</span><strong>{usd(c.marginAmount)}</strong></div>
            <div className="summary-line"><span>임가공료(CM)</span><strong>{usd(c.cm)}</strong></div>
            <div className="summary-line highlight"><span>소계(FOB)</span><strong>{usd(c.fobValue)}</strong></div>
            <label>관세 및 운송비(%)<TextInput type="number" value={sc.dutyFreightRate} onChange={(v)=>changeStageCost(stage.key,"dutyFreightRate",v)} /></label>
            <div className="summary-line"><span>관세 및 운송비</span><strong>{usd(c.dutyFreight)}</strong></div>
            <div className="summary-line"><span>사입원가(VAT-)</span><strong>{usd(c.purchaseVatMinus)}</strong></div>
            <label>공급분자재<TextInput type="number" value={sc.suppliedMaterial} onChange={(v)=>changeStageCost(stage.key,"suppliedMaterial",v)} /></label>
            <div className="summary-line"><span>제조원가(VAT-)</span><strong>{usd(c.manufacturingVatMinus)}</strong></div>
            <label>VAT(%)<TextInput type="number" value={sc.vatRate} onChange={(v)=>changeStageCost(stage.key,"vatRate",v)} /></label>
            <div className="summary-line final"><span>제조원가(VAT+)</span><strong>{usd(c.manufacturingVatPlus)}</strong></div>
          </div>;
        })}
      </div>
    </div>
  );
}


function SalesProfitPage({ data, updateData, canEdit }) {
  const salesData = data.salesData || {};
  const [filter, setFilter] = useState("");

  function updateSelling(styleNo, value) {
    updateData({
      ...data,
      salesData: {
        ...salesData,
        [styleNo]: { ...(salesData[styleNo] || {}), sellingFob: num(value) },
      },
    });
  }

  const allRows = calculateSalesRows(data);
  const rows = allRows.filter(({ sheet, styleNo }) =>
    [styleNo, sheet.header.buyer, sheet.header.brand, sheet.header.itemName, sheet.header.factory]
      .join(" ").toLowerCase().includes(filter.toLowerCase())
  );

  const totals = rows.reduce((sum, row) => ({
    sales: sum.sales + row.metrics.totalSales,
    cost: sum.cost + row.metrics.totalCost,
    profit: sum.profit + row.metrics.totalProfit,
  }), { sales: 0, cost: 0, profit: 0 });
  const totalMargin = totals.sales > 0 ? totals.profit / totals.sales * 100 : 0;

  function exportSalesExcel() {
    const aoa = [
      ["Style No.","Version","Buyer/Brand","Item","Factory","Qty","최종 FOB","제조원가 VAT-","제조원가 VAT+","Selling FOB","Unit Profit","최종 Margin %","Markup %","Status","총매출","총제조원가","총이익"],
      ...rows.map(({ sheet, metrics, status }) => [
        sheet.header.styleNo,
        sheet.version || 1,
        sheet.header.buyer || sheet.header.brand,
        sheet.header.itemName,
        sheet.header.factory,
        metrics.quantity,
        metrics.finalFob,
        metrics.manufacturingVatMinus,
        metrics.manufacturingCost,
        metrics.selling,
        metrics.unitProfit,
        metrics.marginRate,
        metrics.markupRate,
        status.label,
        metrics.totalSales,
        metrics.totalCost,
        metrics.totalProfit,
      ]),
      [],
      ["TOTAL","","","","","","","","","","",totalMargin,"","",totals.sales,totals.cost,totals.profit],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{wch:16},{wch:10},{wch:18},{wch:30},{wch:18},{wch:12},...Array(11).fill({wch:16})];
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, ws, "Sales Profit");
    XLSX.writeFile(book, "매출수익_최종마진.xlsx");
  }

  return (
    <div className="sales-page">
      <div className="sales-toolbar">
        <div>
          <h2>매출 · 수익</h2>
          <p>각 스타일의 최신 원가 버전을 자동 사용합니다. Selling FOB만 입력하면 전체 수익성이 계산됩니다.</p>
        </div>
        <div className="sales-actions">
          <input className="sales-search" value={filter} onChange={(e) => setFilter(e.target.value)} placeholder="Style / Buyer / Item 검색" />
          <button className="excel-export" onClick={exportSalesExcel}>📥 매출·마진 Excel</button>
        </div>
      </div>

      <div className="stats-grid sales-stats">
        <div className="stat-card"><span>💵</span><div><p>총 매출</p><h3>{usd(totals.sales)}</h3><small>Selling FOB × Qty</small></div></div>
        <div className="stat-card"><span>🏭</span><div><p>총 제조원가</p><h3>{usd(totals.cost)}</h3><small>VAT+ 제조원가 × Qty</small></div></div>
        <div className="stat-card"><span>📈</span><div><p>총 이익</p><h3 className={totals.profit < 0 ? "negative" : "positive"}>{usd(totals.profit)}</h3><small>총매출 - 총제조원가</small></div></div>
        <div className="stat-card"><span>％</span><div><p>최종 Margin</p><h3 className={totalMargin < 0 ? "negative" : "positive"}>{totalMargin.toFixed(1)}%</h3><small>이익 ÷ 매출</small></div></div>
      </div>

      <div className="sales-card-list">
        {rows.length ? rows.map(({ sheet, styleNo, sellingFob, metrics, status }) => (
          <div className="panel sales-style-card" key={sheet.id}>
            <div className="sales-card-header">
              <div>
                <div className="sales-style-line"><strong>{styleNo}</strong><Badge>V{sheet.version || 1}</Badge><em className={`margin-status ${status.className}`}>{status.label}</em></div>
                <p>{sheet.header.buyer || sheet.header.brand} · {sheet.header.itemName} · {sheet.header.factory}</p>
              </div>
              <div className="sales-qty">Qty <strong>{metrics.quantity.toLocaleString()}</strong></div>
            </div>

            <div className="sales-card-grid">
              <div className="sales-metric"><span>최종 FOB</span><strong>{usd(metrics.finalFob)}</strong></div>
              <div className="sales-metric"><span>제조원가 VAT-</span><strong>{usd(metrics.manufacturingVatMinus)}</strong></div>
              <div className="sales-metric important"><span>제조원가 VAT+</span><strong>{usd(metrics.manufacturingCost)}</strong></div>
              <label className="selling-input">Selling FOB
                <TextInput type="number" value={sellingFob} onChange={(v) => updateSelling(styleNo, v)} />
              </label>
              <div className={`sales-metric ${metrics.unitProfit < 0 ? "loss-box" : "profit-box"}`}><span>Unit Profit</span><strong>{usd(metrics.unitProfit)}</strong></div>
              <div className="sales-metric"><span>Margin</span><strong>{metrics.marginRate.toFixed(1)}%</strong></div>
              <div className="sales-metric"><span>Markup</span><strong>{metrics.markupRate.toFixed(1)}%</strong></div>
            </div>

            <div className="sales-total-grid">
              <div><span>총매출</span><strong>{usd(metrics.totalSales)}</strong></div>
              <div><span>총제조원가</span><strong>{usd(metrics.totalCost)}</strong></div>
              <div className={metrics.totalProfit < 0 ? "negative" : "positive"}><span>총이익</span><strong>{usd(metrics.totalProfit)}</strong></div>
            </div>
          </div>
        )) : <div className="panel empty">원가견적서를 생성하거나 Buyer Excel을 Import해 주세요.</div>}
      </div>

      <div className="margin-note">
        <strong>계산 기준:</strong> Margin = (Selling FOB − 제조원가 VAT+) ÷ Selling FOB. 
        Markup = (Selling FOB − 제조원가 VAT+) ÷ 제조원가 VAT+.
      </div>
    </div>
  );
}

function GenericPage({ title, rows, columns }) {
  return <div className="panel"><h2>{title}</h2><SimpleTable rows={rows} columns={columns}/></div>;
}

export default function App() {
  const [data, setData] = useState(loadData);
  const [user, setUser] = useState(null);
  const [active, setActive] = useState("dashboard");
  const [search, setSearch] = useState("");

  function updateData(next) { setData(next); persist(next); }
  if (!user) return <Login data={data} onLogin={setUser}/>;

  const menus = MENU_GROUPS.map((g)=>({...g,items:g.items.filter((i)=>canAccess(user.role,i))})).filter((g)=>g.items.length);
  const current = menus.flatMap((g)=>g.items).find((i)=>i.id===active) || menus[0]?.items[0];
  const canEdit = user.role !== "Viewer";

  function content() {
    if (active==="dashboard") return <Dashboard data={data} search={search}/>;
    if (active==="costing") return <CostingPage data={data} updateData={updateData} canEdit={canEdit}/>;
    if (active==="styles") return <GenericPage title="상품 관리" rows={data.styles} columns={[{key:"styleNo",label:"Style No."},{key:"buyer",label:"Buyer"},{key:"item",label:"Item"},{key:"season",label:"Season"},{key:"factory",label:"Factory"},{key:"status",label:"상태",render:r=><Badge>{r.status}</Badge>}]}/>;
    if (active==="orders") return <GenericPage title="주문 관리" rows={data.orders} columns={[{key:"po",label:"PO No."},{key:"buyer",label:"Buyer"},{key:"styleNo",label:"Style No."},{key:"qty",label:"수량"},{key:"delivery",label:"납기"},{key:"status",label:"상태",render:r=><Badge>{r.status}</Badge>}]}/>;
    if (active==="production") return <GenericPage title="생산 관리" rows={data.production} columns={[{key:"styleNo",label:"Style No."},{key:"factory",label:"Factory"},{key:"cutting",label:"재단 %"},{key:"sewing",label:"봉제 %"},{key:"packing",label:"포장 %"},{key:"inspection",label:"검사"},{key:"issue",label:"이슈"}]}/>;
    if (active==="inventory") return <GenericPage title="재고 관리" rows={data.inventory} columns={[{key:"material",label:"자재명"},{key:"type",label:"구분"},{key:"stock",label:"재고"},{key:"unit",label:"단위"},{key:"status",label:"상태"}]}/>;
    if (active==="customers") return <GenericPage title="거래처 관리" rows={data.customers} columns={[{key:"buyer",label:"Buyer"},{key:"brand",label:"Brand"},{key:"country",label:"국가"},{key:"status",label:"상태"}]}/>;
    if (active==="suppliers") return <GenericPage title="공급업체" rows={data.suppliers} columns={[{key:"name",label:"업체명"},{key:"type",label:"구분"},{key:"country",label:"국가"},{key:"status",label:"상태"}]}/>;
    if (active==="purchase") return <GenericPage title="발주 관리" rows={data.purchase} columns={[{key:"po",label:"발주 No."},{key:"supplier",label:"업체"},{key:"material",label:"자재"},{key:"qty",label:"수량"},{key:"status",label:"상태"}]}/>;
    if (active==="users") return <GenericPage title="사용자 관리" rows={data.users} columns={[{key:"id",label:"ID"},{key:"name",label:"직원명"},{key:"role",label:"권한",render:r=><Badge>{ROLE_LABELS[r.role]}</Badge>},{key:"department",label:"부서"},{key:"status",label:"상태"}]}/>;
    if (active==="sales") return <SalesProfitPage data={data} updateData={updateData} canEdit={canEdit}/>;
    return <Dashboard data={data} search={search}/>;
  }

  return (
    <div className="app-shell">
      <style>{CSS}</style>
      <aside className="sidebar">
        <div className="brand"><div className="brand-icon">👗</div><div><h1>Fashion ERP</h1><p>의류 통합 경영</p></div></div>
        <nav>{menus.map((group)=><div className="nav-group" key={group.title}><div className="nav-title">{group.title}</div>{group.items.map((item)=><button className={`nav-item ${active===item.id?"active":""}`} onClick={()=>setActive(item.id)} key={item.id}><span>{item.icon}</span>{item.label}</button>)}</div>)}</nav>
        <div className="user-box"><div className="avatar">{user.name?.[0]||"U"}</div><div><strong>{user.name}</strong><span>{ROLE_LABELS[user.role]}</span></div></div>
      </aside>
      <main className="main">
        <header className="topbar">
          <div className="page-title"><h2>{current?.label}</h2><span>{ROLE_LABELS[user.role]} 권한</span></div>
          <div className="top-actions"><div className="top-search">🔍<input value={search} onChange={(e)=>setSearch(e.target.value)} placeholder="Style No / Buyer / Item / Factory 검색"/></div><button className="ghost-button" onClick={()=>setUser(null)}>Logout</button></div>
        </header>
        <section className="content">{content()}</section>
      </main>
    </div>
  );
}

const CSS = `
*{box-sizing:border-box}body{margin:0;font-family:Inter,Arial,sans-serif;background:#0d1020;color:#f8fafc}button,input,select{font:inherit}
.login-page{min-height:100vh;display:flex;align-items:center;justify-content:center}.login-card{width:380px;background:#14182a;border:1px solid #ffffff14;border-radius:24px;padding:32px}.login-logo{font-size:42px}.login-card h1{margin:8px 0}.login-card p{color:#8b95ad}.login-card label{display:block;margin-top:14px;color:#cbd5e1;font-size:13px;font-weight:700}.login-card input{width:100%;height:44px;margin-top:6px;border:1px solid #ffffff18;border-radius:12px;background:#0f1324;color:#fff;padding:0 12px}.login-card button{width:100%;height:44px;margin-top:20px;border:0;border-radius:12px;background:#6366f1;color:#fff;font-weight:900}.login-card small{display:block;margin-top:12px;color:#778198}.login-error{margin-top:12px;color:#fca5a5}
.app-shell{min-height:100vh;display:grid;grid-template-columns:240px 1fr}.sidebar{background:#121526;border-right:1px solid #ffffff12;display:flex;flex-direction:column}.brand{height:72px;padding:0 18px;display:flex;align-items:center;gap:12px;border-bottom:1px solid #ffffff12}.brand-icon{font-size:28px}.brand h1{font-size:17px;margin:0}.brand p{font-size:12px;color:#778198;margin:4px 0 0}.sidebar nav{flex:1;padding:16px 10px;overflow:auto}.nav-group{margin-bottom:18px}.nav-title{font-size:11px;color:#667089;padding:0 8px 7px}.nav-item{width:100%;border:0;background:transparent;color:#cbd5e1;padding:12px;border-radius:10px;text-align:left;display:flex;gap:11px;font-weight:700;cursor:pointer}.nav-item:hover{background:#ffffff0c}.nav-item.active{background:#292350;color:#fff}.user-box{height:74px;border-top:1px solid #ffffff12;display:flex;align-items:center;gap:10px;padding:0 16px}.avatar{width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#7c3aed,#ec4899);font-weight:900}.user-box strong,.user-box span{display:block}.user-box span{font-size:12px;color:#778198}
.main{min-width:0}.topbar{height:72px;border-bottom:1px solid #ffffff12;display:flex;align-items:center;justify-content:space-between;padding:0 26px}.page-title{display:flex;align-items:baseline;gap:12px}.page-title h2{margin:0}.page-title span{color:#778198;font-size:13px}.top-actions{display:flex;gap:10px}.top-search{height:42px;min-width:340px;border:1px solid #ffffff18;border-radius:13px;background:#181c2f;display:flex;align-items:center;gap:8px;padding:0 12px}.top-search input{width:100%;border:0;outline:0;background:transparent;color:#fff}.ghost-button,.excel-export,.danger-button,.primary-button{height:40px;border-radius:11px;padding:0 14px;border:1px solid #ffffff18;color:#fff;font-weight:800;cursor:pointer}.ghost-button{background:#181c2f}.primary-button{background:#6366f1;border:0}.excel-export{background:#0f766e}.danger-button{background:#7f1d1d}.content{padding:16px;overflow:auto}.content-grid{display:grid;gap:20px}.panel{background:#14182a;border:1px solid #ffffff12;border-radius:18px;padding:20px}.panel h2{margin-top:0}.stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}.stat-card{background:#14182a;border:1px solid #ffffff12;border-radius:18px;padding:18px;display:flex;gap:14px;align-items:center}.stat-card>span{font-size:25px}.stat-card p,.stat-card small{color:#8b95ad;margin:0}.stat-card h3{margin:4px 0;font-size:24px}
.table-wrap{overflow:auto}table{border-collapse:collapse;width:100%}th{font-size:12px;color:#8b95ad;text-align:left;padding:11px;border-bottom:1px solid #ffffff18;white-space:nowrap}td{padding:12px;border-bottom:1px solid #ffffff0d;white-space:nowrap}.empty{text-align:center;color:#8b95ad}.badge{display:inline-block;padding:5px 9px;border-radius:999px;background:#373070;color:#ddd6fe;font-size:12px;font-weight:800}
.costing-page{display:grid;gap:12px}.costing-toolbar{display:flex;align-items:center;justify-content:space-between}.costing-toolbar>select{min-width:380px;height:42px;border:1px solid #ffffff18;border-radius:12px;background:#181c2f;color:#fff;padding:0 12px}.toolbar-actions{display:flex;gap:8px}.section-title{margin-bottom:16px}.section-title h2{margin:0}.section-title span{color:#8b95ad;font-size:13px}.row-between{display:flex;align-items:center;justify-content:space-between}.clickable-title{cursor:pointer;user-select:none}.empty-section{text-align:center!important;color:#8b95ad!important;padding:22px!important}.inline-actions{display:flex;gap:8px}.inline-actions button{border:0;border-radius:10px;background:#23294a;color:#fff;padding:9px 12px;cursor:pointer}.header-grid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:8px}.header-grid label,.stage-summary label{color:#aab2c4;font-size:12px;font-weight:700}.header-grid input,.stage-summary input{width:100%;height:38px;margin-top:6px;border:1px solid #ffffff18;border-radius:9px;background:#0f1324;color:#fff;padding:0 9px}
.costing-table-panel{padding:18px}
.compact-cost-list{display:grid;gap:12px}
.compact-cost-row{border:1px solid #ffffff12;border-radius:14px;background:#0f1324;padding:12px}
.compact-row-top{display:grid;grid-template-columns:34px minmax(110px,.8fr) minmax(180px,1.5fr) minmax(130px,1fr) minmax(80px,.6fr) 72px 34px;gap:8px;align-items:end}
.compact-row-top label,.remark-field,.stage-input-grid label{font-size:11px;color:#9aa4ba;font-weight:700}
.compact-row-top input,.remark-field input,.stage-input-grid input,.stage-input-grid select{width:100%;height:32px;margin-top:5px;border:1px solid #ffffff14;border-radius:7px;background:#14182a;color:#fff;padding:0 7px}
.row-number{width:30px;height:30px;border-radius:8px;background:#23294a;display:flex;align-items:center;justify-content:center;color:#c7d2fe;font-weight:900}
.stage-cards{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-top:10px}
.stage-card{border-radius:11px;padding:10px;border:1px solid #ffffff12}
.stage-card-quote{background:#13233a}.stage-card-nego1{background:#302817}.stage-card-final{background:#123426}
.stage-card-title{font-weight:900;font-size:12px;margin-bottom:7px}
.stage-input-grid{display:grid;grid-template-columns:1fr 1fr 82px;gap:7px}
.stage-result{display:flex;justify-content:space-between;align-items:center;margin-top:8px;padding-top:8px;border-top:1px solid #ffffff12;font-size:12px}
.stage-result span{color:#cbd5e1}.stage-result strong{color:#86efac}
.remark-field{display:block;margin-top:9px}
.row-delete{width:30px;height:30px;border:0;border-radius:7px;background:#7f1d1d;color:#fff;cursor:pointer}
.empty-section{text-align:center!important;color:#8b95ad!important;padding:22px!important;border:1px dashed #ffffff18;border-radius:12px}

.summary-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}.stage-summary h3{margin-top:0}.summary-line{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid #ffffff0d}.summary-line.total{font-weight:900}.summary-line.highlight{margin-top:8px;padding:11px;border-radius:10px;background:#312e81}.summary-line.final{margin-top:8px;padding:13px;border-radius:10px;background:#14532d;font-size:17px}.stage-summary label{display:block;margin-top:10px}

.excel-import{position:relative;overflow:hidden;height:40px;border-radius:11px;padding:0 14px;background:#7c3aed;border:0;color:#fff;font-weight:800;cursor:pointer;display:inline-flex;align-items:center}.excel-import input{position:absolute;inset:0;opacity:0;cursor:pointer}
.sales-page{display:grid;gap:18px}.sales-toolbar{display:flex;justify-content:space-between;align-items:center}.sales-toolbar h2{margin:0}.sales-toolbar p{margin:6px 0 0;color:#8b95ad}.sales-table{min-width:1900px}.sales-table input{width:110px;height:34px;border:1px solid #ffffff18;border-radius:8px;background:#0f1324;color:#fff;padding:0 8px}.positive{color:#86efac!important}.negative{color:#fca5a5!important}.margin-note{padding:14px 16px;border-radius:12px;background:#181c2f;color:#aab2c4;font-size:13px}.sales-stats .stat-card h3{font-size:20px}



.dashboard-two-column{display:grid;grid-template-columns:1.35fr 1fr;gap:16px}
.dashboard-profit-list,.risk-list{display:grid;gap:8px}
.dashboard-profit-row,.risk-row{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:11px;border:1px solid #ffffff0d;border-radius:11px;background:#0f1324}
.dashboard-profit-row div:first-child,.risk-row div:first-child{min-width:0}
.dashboard-profit-row strong,.dashboard-profit-row span,.risk-row strong,.risk-row span{display:block}
.dashboard-profit-row span,.risk-row span{margin-top:4px;color:#8b95ad;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dashboard-profit-values{display:flex;align-items:center;gap:9px;white-space:nowrap}
.risk-row>div:last-child{display:flex;align-items:center;gap:8px}
.margin-status{display:inline-flex;align-items:center;justify-content:center;border-radius:999px;padding:5px 8px;font-size:10px;font-style:normal;font-weight:900;letter-spacing:.3px}
.status-loss,.status-critical{background:#7f1d1d;color:#fecaca}
.status-review{background:#78350f;color:#fde68a}
.status-good{background:#164e63;color:#a5f3fc}
.status-excellent{background:#14532d;color:#bbf7d0}
.sales-actions{display:flex;align-items:center;gap:8px}
.sales-search{height:40px;min-width:250px;border:1px solid #ffffff18;border-radius:11px;background:#181c2f;color:#fff;padding:0 12px}
.sales-card-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
.sales-style-card{display:grid;gap:12px}
.sales-card-header{display:flex;justify-content:space-between;align-items:flex-start;gap:12px}
.sales-card-header p{margin:6px 0 0;color:#8b95ad;font-size:12px}
.sales-style-line{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.sales-style-line>strong{font-size:18px}
.sales-qty{font-size:11px;color:#8b95ad;white-space:nowrap}.sales-qty strong{display:block;color:#fff;font-size:18px;margin-top:3px;text-align:right}
.sales-card-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}
.sales-metric,.selling-input{min-height:70px;border:1px solid #ffffff0d;border-radius:11px;background:#0f1324;padding:10px}
.sales-metric span,.selling-input{color:#8b95ad;font-size:11px;font-weight:700}
.sales-metric strong{display:block;margin-top:8px;color:#fff;font-size:16px}
.sales-metric.important{background:#24213f}
.selling-input input{height:34px;margin-top:7px}
.profit-box{background:#103427}.profit-box strong{color:#86efac}
.loss-box{background:#421818}.loss-box strong{color:#fca5a5}
.sales-total-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;padding-top:11px;border-top:1px solid #ffffff0d}
.sales-total-grid div{padding:8px;border-radius:9px;background:#181c2f}
.sales-total-grid span,.sales-total-grid strong{display:block}.sales-total-grid span{font-size:10px;color:#8b95ad}.sales-total-grid strong{margin-top:5px}

@media(max-width:1450px){
  .compact-row-top{grid-template-columns:30px minmax(100px,.8fr) minmax(150px,1.4fr) minmax(110px,1fr) 80px 64px 30px}
  .compact-row-top input,.stage-input-grid input,.stage-input-grid select{font-size:11px}
  .header-grid{grid-template-columns:repeat(4,minmax(0,1fr))}
}
@media(max-width:1180px){
  .stage-cards{grid-template-columns:1fr}
  .compact-row-top{grid-template-columns:30px repeat(3,minmax(0,1fr))}
  .compact-row-top .wide-field{grid-column:span 2}
  .compact-row-top .unit-field{grid-column:auto}
}


@media(max-width:1350px){
  .sales-card-list{grid-template-columns:1fr}
}
@media(max-width:1150px){
  .dashboard-two-column{grid-template-columns:1fr}
  .sales-card-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
}
@media(max-width:700px){
  .sales-toolbar,.sales-actions,.sales-card-header{align-items:stretch;flex-direction:column}
  .sales-search{width:100%;min-width:0}
  .sales-card-grid,.sales-total-grid{grid-template-columns:1fr}
}

@media(max-width:1100px){.app-shell{grid-template-columns:1fr}.sidebar{min-height:auto}.stats-grid,.summary-grid,.header-grid{grid-template-columns:1fr}.topbar{height:auto;min-height:72px;flex-direction:column;align-items:flex-start;padding:14px}.top-actions,.top-search{width:100%}.content{padding:14px}.costing-toolbar{align-items:stretch;gap:10px;flex-direction:column}.costing-toolbar>select{min-width:0;width:100%}}
`;
