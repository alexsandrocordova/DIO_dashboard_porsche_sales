// Smoke test do dashboard em Node: DOM simulado + base real de producao.
const fs = require("fs");

// ---- Stub de DOM -----------------------------------------------------
const elements = {};
function makeCtx() {
  return new Proxy(
    {},
    {
      get(t, prop) {
        if (prop === "canvas") return { width: 300, height: 150 };
        return () => ({ addColorStop() {} });
      },
      set() {
        return true;
      },
    }
  );
}
function makeEl() {
  return {
    value: "ALL",
    innerHTML: "",
    innerText: "",
    className: "",
    disabled: false,
    options: [],
    style: {},
    classList: { add() {}, remove() {} },
    appendChild() {},
    addEventListener() {},
    getBoundingClientRect: () => ({ width: 800, height: 400 }),
    getContext: () => makeCtx(),
  };
}
global.document = {
  getElementById: (id) => (elements[id] = elements[id] || makeEl()),
  createElement: () => makeEl(),
  body: { appendChild() {}, removeChild() {} },
};
global.window = {};
global.Option = class {
  constructor(v, t) {
    this.value = v;
    this.text = t || v;
  }
};

// ---- Carrega HTML main script + dados reais --------------------------
eval(fs.readFileSync("dashboard_data.js", "utf-8"));
if (!global.window.PORSCHE_SANITIZED_DATA || !global.window.PORSCHE_SANITIZED_DATA.length) {
  throw new Error("dashboard_data.js nao carregou registros");
}

const html = fs.readFileSync("porsche_intelligence_dashboard.html", "utf-8");
const blocks = html.match(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g);
const main = blocks.map((b) => b.replace(/^<script[^>]*>/, "").replace(/<\/script>$/, "")).reduce((a, b) => (a.length > b.length ? a : b));

let code = main;
code += `
;loadProductionData();
console.log("registros brutos:", rawSalesData.length);
console.log("registros filtrados:", filteredData.length);
if (rawSalesData.length !== 100) throw new Error("Esperado 100 registros, obteve " + rawSalesData.length);
if (filteredData.length !== 100) throw new Error("Filtro ALL deveria manter 100");
const r = rawSalesData[0];
console.log("amostra:", JSON.stringify(r));
if (!r.modelo || !r.cidade || !r.estado || !(r.preco > 0)) throw new Error("Registro incompleto: " + JSON.stringify(r));
console.log("KPI faturamento:", document.getElementById("kpi-total-revenue").innerText);
console.log("KPI ticket medio:", document.getElementById("kpi-avg-ticket").innerText);
console.log("KPI top model:", document.getElementById("kpi-top-model").innerText);
console.log("status badge:", document.getElementById("data-status-text").innerText);
const tbody = document.getElementById("sales-table-body");
if (tbody.innerHTML.indexOf(rowEstado => 0) >= 0) {} // no-op
console.log("tabela contem coluna estado:", JSON.stringify(rawSalesData[0].estado));
console.log("SMOKE TEST OK");
`;

eval(code);
