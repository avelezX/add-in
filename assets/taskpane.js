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
    const ticker = document.getElementById("tickerInput").value.trim();
    if (!ticker) {
      document.getElementById("status").textContent = "⚠️ Debes ingresar un ticker.";
      return;
    }

    document.getElementById("status").textContent = `Consultando ${ticker}...`;

    try {
      const data = await XTY(ticker);
      // Reemplaza TODO el bloque Excel.run del botón XTY por este:
      await Excel.run(async (ctx) => {
        // 1) Celda activa correcta
        const start = ctx.workbook.getActiveCell();

        // 2) Dimensiones seguras
        const rows = Array.isArray(data) ? data.length : 0;
        const cols = rows > 0 && Array.isArray(data[0]) ? data[0].length : 0;
        if (rows === 0 || cols === 0) throw new Error("La consulta no devolvió filas.");

        // 3) Escribe la matriz empezando en la celda activa
        const range = start.getResizedRange(rows - 1, cols - 1);
        range.values = data;

        await ctx.sync();
      });
      document.getElementById("status").textContent = `Datos de ${ticker} cargados ✅`;
    } catch (err) {
      console.error(err);
      document.getElementById("status").textContent = `Error en XTY ❌: ${err.message}`;
    }
  });
});