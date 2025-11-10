/* global Excel, Office */
console.log("Xerenity TaskPane loaded ✅");

const RPC_POST_URL = "https://tvpehjbqxpiswkqszwwv.supabase.co/rest/v1/rpc/search";
const PARAM_NAME = "ticker";
const PROFILE_NAME = "public";
const API_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR2cGVoamJxeHBpc3drcXN6d3d2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE2OTY0NTEzODksImV4cCI6MjAxMjAyNzM4OX0.LZW0i9HU81lCdyjAdqjwwF4hkuSVtsJsSDQh7blzozw";

// Convertir fecha YYYY-MM-DD a número de serie Excel
function ymdToExcelSerial(isoYmd) {
  const parts = String(isoYmd).split("-");
  const y = Number(parts[0]), m = Number(parts[1]) || 1, d = Number(parts[2]) || 1;
  const excelEpochUTC = Date.UTC(1899, 11, 30);
  const thisUTC = Date.UTC(y, m - 1, d);
  return (thisUTC - excelEpochUTC) / 86400000;
}

// 🧭 Botón PING
async function runPING() {
  await Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.values = [[1]];
    cell.format.fill.color = "lightgreen";
    await context.sync();
  });
}

// 📈 Botón XTY (consulta Supabase)
async function runXTY() {
  await Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.values = [["⏳ Cargando..."]];
    await context.sync();

    try {
      const ticker = prompt("Ticker a consultar:", "ibr_1yr");
      if (!ticker) {
        cell.values = [["❌ Cancelado"]];
        return;
      }

      const headers = {
        "Content-Type": "application/json",
        "content-profile": PROFILE_NAME,
        "Accept-Profile": PROFILE_NAME,
        "apikey": API_KEY,
        "Authorization": "Bearer " + API_KEY,
      };

      const bodyObj = {}; bodyObj[PARAM_NAME] = ticker.trim();

      const res = await fetch(RPC_POST_URL, {
        method: "POST",
        headers,
        body: JSON.stringify(bodyObj),
      });

      if (!res.ok) {
        const t = await res.text();
        throw new Error("HTTP " + res.status + (t ? " - " + t : ""));
      }

      const payload = await res.json();
      const rows = Array.isArray(payload)
        ? payload
        : payload && Array.isArray(payload.data)
        ? payload.data
        : [];

      if (!rows.length) {
        cell.values = [["⚠️ Sin datos"]];
        return;
      }

      const out = [["time", "value"], ...rows.map(r => [ymdToExcelSerial(r.time), Number(r.value)])];
      const range = context.workbook.getActiveWorksheet().getRange("A1").getResizedRange(out.length - 1, 1);
      range.values = out;
      range.format.autofitColumns();
      await context.sync();
    } catch (err) {
      cell.values = [["❌ Error: " + err.message]];
      await context.sync();
    }
  });
}

Office.onReady(() => console.log("✅ Office ready - Xerenity loaded"));

if (typeof window !== "undefined") {
  window.runPING = runPING;
  window.runXTY = runXTY;
}