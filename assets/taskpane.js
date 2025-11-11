console.log("Xerenity Taskpane script loaded ✅");

Office.onReady(() => {
  console.log("Office ready - Xerenity panel active");
  document.getElementById("status").textContent = "Conectado a Excel ✔️";

  // 🧮 SAMÁN TOOLS - Promedio
  document.getElementById("btnPromedio").addEventListener("click", async () => {
    document.getElementById("status").textContent = "Calculando promedio...";
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "address", "rowCount", "columnCount"]);
      await ctx.sync(); // Necesario antes de usar rowCount/columnCount

      const valores = [];
      for (const fila of range.values) {
        for (const v of fila) {
          if (typeof v === "number" && isFinite(v)) valores.push(v);
        }
      }
      if (!valores.length) throw new Error("No hay números en el rango seleccionado.");

      const promedio = valores.reduce((a, b) => a + b, 0) / valores.length;

      // Escribir una celda debajo del rango seleccionado (misma columna inicial)
      const target = range.getOffsetRange(range.rowCount, 0).getCell(0, 0);
      target.values = [[`Promedio: ${promedio.toFixed(4)}`]];

      await ctx.sync();
      document.getElementById("status").textContent =
        `Promedio calculado para ${range.address} ✅`;
    }).catch(err => {
      console.error(err);
      document.getElementById("status").textContent = `Error ❌: ${err.message}`;
    });
  });

  // ===========================
  // ⚡ XERENITY TOOLS
  // ===========================

  // PING
  document.getElementById("btnPing").addEventListener("click", async () => {
    try {
      const result = await PING();
      await Excel.run(async (ctx) => {
        const cell = ctx.workbook.getActiveCell();
        cell.values = [[result]];
        await ctx.sync();
      });
      document.getElementById("status").textContent = "PING ejecutado correctamente ✅";
    } catch (err) {
      document.getElementById("status").textContent = "Error en PING ❌: " + err.message;
    }
  });

  // SERIES (antes XTY)
  document.getElementById("btnSeries").addEventListener("click", async () => {
    const ticket = document.getElementById("tickerInput").value.trim();
    if (!ticket) {
      document.getElementById("status").textContent = "⚠️ Debes ingresar un ticker.";
      return;
    }

    document.getElementById("status").textContent = `Consultando ${ticket}...`;

    try {
      const data = await XTY(ticket);
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const startCell = ctx.workbook.getActiveCell();
        const range = startCell.getResizedRange(data.length - 1, data[0].length - 1);
        range.values = data;
        await ctx.sync();
      });
      document.getElementById("status").textContent = `Datos de ${ticket} cargados ✅`;
    } catch (err) {
      console.error(err);
      document.getElementById("status").textContent = `Error en SERIES ❌: ${err.message}`;
    }
  });
});