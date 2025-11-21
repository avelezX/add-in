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
      document.getElementById("status").textContent = `Promedio calculado para ${range.address} ✅`;
    }).catch((err) => {
      console.error(err);
      document.getElementById("status").textContent = `Error ❌: ${err.message}`;
    });
  });
  // ============================
  //       INTERPOLACIÓN
  // ============================
  let interp_X = null;
  let interp_Y = null;

  // OK Años
  document.getElementById("btnPickX").addEventListener("click", async () => {
    try {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("values");
        await ctx.sync();

        interp_X = range.values.map((r) => r[0]); // extraer columna
        document.getElementById("status").textContent =
          "✓ Rango de AÑOS (X) guardado correctamente.";
      });
    } catch (err) {
      document.getElementById("status").textContent = "❌ Error guardando AÑOS (X): " + err.message;
    }
  });

  // OK Tasas
  document.getElementById("btnPickY").addEventListener("click", async () => {
    try {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("values");
        await ctx.sync();

        interp_Y = range.values.map((r) => r[0]);
        document.getElementById("status").textContent =
          "✓ Rango de TASAS (Y) guardado correctamente.";
      });
    } catch (err) {
      document.getElementById("status").textContent =
        "❌ Error guardando TASAS (Y): " + err.message;
    }
  });

  // Botón Interpolar
  document.getElementById("btnInterp").addEventListener("click", async () => {
    try {
      const X = interp_X;
      const Y = interp_Y;

      if (!X || !Y) {
        document.getElementById("status").textContent =
          "❌ Primero debes guardar los rangos X e Y.";
        return;
      }

      if (X.length !== Y.length) {
        document.getElementById("status").textContent =
          "❌ Los rangos X e Y deben tener la misma cantidad de datos.";
        return;
      }

      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load(["values", "rowCount", "columnCount", "address"]);
        await ctx.sync();

        const rows = range.rowCount;
        const cols = range.columnCount;

        // ❌ No permitir rangos 2D
        if (rows > 1 && cols > 1) {
          document.getElementById("status").textContent =
            "❌ Selecciona solo una fila o una columna (no rangos 2D).";
          return;
        }

        // Determinar orientación
        const isColumn = rows > 1;
        const isRow = cols > 1;

        // Construir pares ordenados para interpolación
        const pairs = X.map((x, i) => ({ x, y: Y[i] })).sort((a, b) => a.x - b.x);

        // Función interna para interpolar un valor
        function interpolarValor(year) {
          if (isNaN(year)) return null;

          // EXTRAPOLACIÓN HACIA ATRÁS
          if (year < pairs[0].x) {
            const x1 = pairs[0].x;
            const y1 = pairs[0].y;
            const x2 = pairs[1].x;
            const y2 = pairs[1].y;
            const slope = (y2 - y1) / (x2 - x1);
            return y1 + (year - x1) * slope;
          }

          // EXTRAPOLACIÓN HACIA ADELANTE
          if (year > pairs[pairs.length - 1].x) {
            const x1 = pairs[pairs.length - 2].x;
            const y1 = pairs[pairs.length - 2].y;
            const x2 = pairs[pairs.length - 1].x;
            const y2 = pairs[pairs.length - 1].y;
            const slope = (y2 - y1) / (x2 - x1);
            return y2 + (year - x2) * slope;
          }

          for (let i = 0; i < pairs.length - 1; i++) {
            if (year >= pairs[i].x && year <= pairs[i + 1].x) {
              const { x: x1, y: y1 } = pairs[i];
              const { x: x2, y: y2 } = pairs[i + 1];
              return y1 + ((year - x1) * (y2 - y1)) / (x2 - x1);
            }
          }
          return null;
        }

        // --- ESCRITURA FINAL EN EXCEL (SIN ERRORES DE DIMENSIONES) ---
        if (rows === 1 && cols === 1) {
          // Celda única → escribir a la derecha
          const year = range.values[0][0];
          const result = interpolarValor(year);

          const target = range.getOffsetRange(0, 1);
          target.values = [[result]];
        } else if (rows === 1 && cols > 1) {
          // FILA → escribir debajo
          const row = range.values[0];
          const interpolados = row.map((year) => interpolarValor(year));

          // Crear un rango exacto 1 x N debajo de la selección
          const target = range
            .getOffsetRange(1, 0)
            .getCell(0, 0)
            .getResizedRange(0, cols - 1);

          target.values = [interpolados]; // matriz horizontal 1 × N
        } else if (rows > 1 && cols === 1) {
          // COLUMNA → escribir derecha
          const interpolados = [];
          for (let r = 0; r < rows; r++) {
            const year = range.values[r][0];
            interpolados.push([interpolarValor(year)]); // matriz vertical N × 1
          }

          // Crear un rango exacto N x 1 a la derecha de la selección
          const target = range
            .getOffsetRange(0, 1)
            .getCell(0, 0)
            .getResizedRange(rows - 1, 0);

          target.values = interpolados;
        } else {
          document.getElementById("status").textContent =
            "❌ Selecciona solo una fila o una columna.";
          return;
        }

        await ctx.sync();

        document.getElementById("status").textContent =
          `✓ Interpolación completada (${rows * cols} valores).`;
      });
    } catch (err) {
      console.error(err);
      document.getElementById("status").textContent = "❌ Error interpolando: " + err.message;
    }
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

  // ========== SERIES: ALL ==========
  document.getElementById("btnSeriesAll").addEventListener("click", async () => {
    const ticket = document.getElementById("tickerInput").value.trim();
    if (!ticket) {
      document.getElementById("status").textContent = "⚠️ Debes ingresar un ticker.";
      return;
    }

    document.getElementById("status").textContent = `Consultando serie completa de ${ticket}...`;

    try {
      const data = await XTY(ticket);
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const startCell = ctx.workbook.getActiveCell();
        const range = startCell.getResizedRange(data.length - 1, data[0].length - 1);
        range.values = data;
        await ctx.sync();
      });
      document.getElementById("status").textContent = `Serie completa de ${ticket} cargada ✅`;
    } catch (err) {
      console.error(err);
      document.getElementById("status").textContent = `Error en SERIES (All) ❌: ${err.message}`;
    }
  });

  // ========== SERIES: LAST ==========
  document.getElementById("btnSeriesLast").addEventListener("click", async () => {
    const ticket = document.getElementById("tickerInput").value.trim();
    if (!ticket) {
      document.getElementById("status").textContent = "⚠️ Debes ingresar un ticker.";
      return;
    }

    document.getElementById("status").textContent = `Consultando último valor de ${ticket}...`;

    try {
      const data = await XTY(ticket);
      if (!data || data.length < 2) throw new Error("No se recibieron datos.");

      const last = data[data.length - 1]; // último registro [fecha, valor]
      const headers = data[0]; // ["time", "value"]
      const output = [headers, last];

      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const startCell = ctx.workbook.getActiveCell();
        const range = startCell.getResizedRange(output.length - 1, output[0].length - 1);
        range.values = output;
        await ctx.sync();
      });

      document.getElementById("status").textContent = `Último dato de ${ticket} cargado ✅`;
    } catch (err) {
      console.error(err);
      document.getElementById("status").textContent = `Error en SERIES (Last) ❌: ${err.message}`;
    }
  });

  // ========== COLTES LAST ==========
  document.getElementById("btnColtesLast").addEventListener("click", async () => {
    document.getElementById("status").textContent = "Consultando últimas series COLTES...";

    // Lista fija de series
    const coltesSeries = [
      ["COLTES 4.75 2035-04-04", "04b0a5375511340152e500a991fb6202"],
      ["COLTES 7 2031-03-26", "14b8d27bc31dd101c1c028710d92008c"],
      ["COLTES 6.25 2025-11-26", "1be6ad6b20ce6d17a9febd3a16a84ccd"],
      ["COLTES 6 2028-04-28", "28b42938a3491be73e9e3edb27d0a341"],
      ["COLTES 7.5 2026-08-26", "2f290fdd2e7b697ffa5129376aeeb03d"],
      ["COLTES 7.25 2050-10-26", "37cff456093349e7a656132b73409a7d"],
      ["COLTES 3.3 2027-03-17", "388a20606f87cf27babe68605a21dbe8"],
      ["COLTES 7 2031-03-26", "44beef76466d04b71805f73ac77eaa6e"],
      ["COLTES 7.75 2030-09-18", "45b93ff81a7f9274ae2305c8fa8c7492"],
      ["COLTES 5.75 2027-11-03", "4779cf6fd8ffafa88f6acb3968dc16ec"],
      ["COLTES 3 2033-03-25", "4c60c3257b1ab2a21c84d716cd21471f"],
      ["COLTES 7 2032-06-30", "58b363fbb6a935e2e271a934b32ac170"],
      ["COLTES 2.25 2029-04-18", "5acb5c01180aeef3d7e5060ee0f5228f"],
      ["COLTES 11 2024-07-24", "5deb428485a802abcd481efa1ca0f5c7"],
      ["COLTES 3.5 2025-07-05", "6223be130b2b724adc9fe7cd7d27b204"],
      ["COLTES 7.25 2034-10-18", "6611fc76da930ab18fea79efa11f644e"],
      ["COLTES 9.25 2042-05-28", "a567521827d8d5c922b1bf7d6504b780"],
      ["COLTES 6.25 2036-09-07", "c41ef5179b3c1b325a4d9c2665f3deb3"],
      ["COLTES 3.75 2037-02-25", "cc760f9770f4f1f84ef0ee4d5c5aab9b"],
      ["COLTES 3.75 2049-06-16", "e4a8c4350b378f4f6cb764a4cc18d396"],
      ["COLTES 13.25 2033-09-02", "e7e6a55b233278f79c07e43fda0bde10"],
    ];

    try {
      // Ejecutar todas las consultas en paralelo
      const results = await Promise.allSettled(
        coltesSeries.map(async ([name, ticker]) => {
          const data = await XTY(ticker);
          if (!data || data.length < 2) throw new Error("Sin datos");
          const last = data[data.length - 1]; // [fecha, valor]
          return [name, last[0], last[1]];
        })
      );

      // Construir tabla: encabezado + resultados válidos
      const rows = [["Serie", "Fecha", "Valor"]];
      for (const r of results) {
        if (r.status === "fulfilled") rows.push(r.value);
        else rows.push([coltesSeries[results.indexOf(r)][0], "Error", "—"]);
      }

      // Escribir en Excel
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const startCell = ctx.workbook.getActiveCell();
        const range = startCell.getResizedRange(rows.length - 1, rows[0].length - 1);
        range.values = rows;
        await ctx.sync();
      });

      document.getElementById("status").textContent = "Últimos valores COLTES cargados ✅";
    } catch (err) {
      console.error(err);
      document.getElementById("status").textContent = `Error en COLTES Last ❌: ${err.message}`;
    }
  });
});
