import { getSavedConfig, saveConfig, clearHistorianCache, getLastDebugMessage, testConnection } from "../shared/api.js";
function $(id) { return document.getElementById(id); }
function collectConfig() {
  return {
    baseUrl: $("baseUrl").value.trim(),
    historyEndpoint: $("historyEndpoint").value.trim(),
    proxyUrl: $("proxyUrl").value.trim(),
    requestMode: $("requestMode").value,
    dataSource: $("dataSource").value.trim(),
    field: $("field").value.trim(),
    hf: Number($("hf").value),
    rt: Number($("rt").value),
    stepped: Number($("stepped").value),
    useCredentials: $("useCredentials").checked,
    useProxy: $("proxyUrl").value.trim().length > 0
  };
}
function loadConfig() {
  const cfg = getSavedConfig();
  $("baseUrl").value = cfg.baseUrl; $("historyEndpoint").value = cfg.historyEndpoint; $("proxyUrl").value = cfg.proxyUrl || ""; $("requestMode").value = cfg.requestMode; $("dataSource").value = cfg.dataSource; $("field").value = cfg.field; $("hf").value = cfg.hf; $("rt").value = cfg.rt; $("stepped").value = cfg.stepped; $("useCredentials").checked = !!cfg.useCredentials;
}
function buildFormula(functionName) {
  const cfg = JSON.stringify(collectConfig()).replace(/"/g, '""');
  return `=${functionName}(${$("tagsAddress").value.trim()},${$("startAddress").value.trim()},${$("endAddress").value.trim()},${$("period").value.trim() || "1"},${$("pu").value},"${cfg}")`;
}
async function insertFormula(functionName) {
  const output = $("outputAddress").value.trim();
  const formula = buildFormula(functionName);
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(output);
    range.formulas = [[formula]];
    await context.sync();
  });
}
async function useSelectionFor(inputId) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    $(inputId).value = range.address;
  });
}
function refreshDebug() { $("debugText").textContent = getLastDebugMessage() || "(no debug messages yet)"; }
Office.onReady(() => {
  loadConfig(); refreshDebug();
  $("saveConfig").addEventListener("click", () => { const cfg = saveConfig(collectConfig()); $("connectionStatus").textContent = `Saved config for ${cfg.baseUrl}${cfg.historyEndpoint}`; refreshDebug(); });
  $("clearCache").addEventListener("click", () => { clearHistorianCache(); $("connectionStatus").textContent = "Historian cache cleared."; refreshDebug(); });
  $("testConnection").addEventListener("click", async () => { $("connectionStatus").textContent = "Testing..."; $("connectionStatus").textContent = await testConnection(JSON.stringify(collectConfig())); refreshDebug(); });
  $("insertHistoryData").addEventListener("click", async () => insertFormula("QKHISTORYDATA"));
  $("insertHistory").addEventListener("click", async () => insertFormula("QKHISTORY"));
  document.querySelectorAll("[data-pick]").forEach((button) => button.addEventListener("click", async () => useSelectionFor(button.getAttribute("data-pick"))));
  $("refreshDebug").addEventListener("click", refreshDebug);
});
