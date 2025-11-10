console.log("Xerenity Taskpane script loaded ✅");

Office.onReady(() => {
  console.log("Office ready - Xerenity panel active");
  document.getElementById("status").textContent = "Conectado a Excel ✔️";

  // Botón PING
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

  // Botón XTY
  document.getElementById("btnXty").addEventListener("click", async () => {
    const ticker = prompt("Introduce el ticker (ej. ibr_1yr):");
    if (!ticker) return;

    document.getElementById("status").textContent = "Consultando " + ticker + "...";

    try {
      const data = await XTY(ticker);
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getActiveCell().getResizedRange(data.length - 1, data[0].length - 1);
        range.values = data;
        await ctx.sync();
      });
      document.getElementById("status").textContent = "Datos cargados en Excel ✅";
    } catch (err) {
      console.error(err);
      document.getElementById("status").textContent = "Error al ejecutar XTY ❌: " + err.message;
    }
  });
});