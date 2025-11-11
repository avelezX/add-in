/* global Excel, Office */

console.log("Xerenity TaskPane script loaded ✅");

Office.onReady(() => {
  console.log("✅ Office ready - Xerenity panel active");
  document.getElementById("status").textContent = "Panel cargado correctamente ✅";
});

// 🔹 API Key y URL de Supabase (igual que en functions.js)
const RPC_POST_URL = "https://tvpehjbqxpiswkqszwwv.supabase.co/rest/v1/rpc/search";
const API_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR2cGVoamJxeHBpc3drcXN6d3d2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE2OTY0NTEzODksImV4cCI6MjAxMjAyNzM4OX0.LZW0i9HU81lCdyjAdqjwwF4hkuSVtsJsSDQh7blzozw";

// 🔹 Conversión de fechas a serial Excel
function ymdToExcelSerial(isoYmd) {
  const parts = String(isoYmd).split("-");
  const y = Number(parts[0]),
        m = Number(parts[1]) || 1,
        d = Number(parts[2]) || 1;
  const excelEpochUTC = Date.UTC(1899, 11, 30);
  const thisUTC = Date.UTC(y, m - 1, d);
  return (thisUTC - excelEpochUTC) / 86400000;
}

// 🔹 Botón de prueba PING
async function runPing() {
  await Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.values = [[1]];
    await context.sync();
  });
  console.log("✅ PING ejecutado correctamente");
  alert("PING ejecutado correctamente ✅");
}

// 🔹 Botón principal: Run XTY
async function runXTY() {
  try {
    const ticket = prompt("Introduce el ticket (por ejemplo: ibr_1yr):");
    if (!ticket) {
      alert("❗ Debes ingresar un ticket válido");
      return;
    }

    console.log("🚀 Consultando Supabase para ticket:", ticket);

    const res = await fetch(RPC_POST_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": API_KEY,
        "Authorization": "Bearer " + API_KEY
      },
      body: JSON.stringify({ ticket: ticket.trim() })
    });

    const text = await res.text();
    if (!res.ok) throw new Error(`HTTP ${res.status} - ${text}`);

    const payload = JSON.parse(text);
    console.log("📦 Respuesta Supabase:", payload);

    const rows = Array.isArray(payload)
      ? payload
      : payload?.data || [];

    if (!rows.length) {
      alert("⚠️ No se encontraron datos para ese ticket.");
      return;
    }

    // 🔸 Escribir los datos directamente en Excel
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = context.workbook.getActiveCell();

      const output = [["time", "value"]];
      for (const r of rows) {
        if (r.time && r.value != null) {
          output.push([ymdToExcelSerial(r.time), Number(r.value)]);
        }
      }

      const range = cell.getResizedRange(output.length - 1, 1);
      range.values = output;
      await context.sync();
    });

    alert("✅ Datos insertados en Excel con éxito.");
  } catch (err) {
    console.error("❌ Error en XTY:", err);
    alert(`Error en XTY ❌: ${err.message}`);
  }
}

// 🔹 Exponer funciones globales al HTML
window.runPing = runPing;
window.runXTY = runXTY;