function procesarDatosAlMaestro(datosCrudos, mesSuministro) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMaestro = ss.getSheetByName("Padron_Maestro");
    const sheetConfig = ss.getSheetByName("Config_Rutas");
    
    if (!sheetMaestro || !sheetConfig) {
      throw new Error("Faltan pestañas: Padron_Maestro o Config_Rutas.");
    }

    // 1. Cargar configuración de cuadrillas
    const configData = sheetConfig.getDataRange().getValues();
    const cuadrillaMap = {};
    for (let i = 1; i < configData.length; i++) {
      cuadrillaMap[configData[i][0].toString().trim()] = configData[i][1];
    }

    // 2. Extraer encabezados del Excel
    const headers = datosCrudos[0];
    const h = {};
    headers.forEach((hd, idx) => h[hd.toString().trim()] = idx);

    const records = {};

    // 3. Agrupar (Pivotear) 6 filas en 1 sola por suministro
    for (let i = 1; i < datosCrudos.length; i++) {
      const row = datosCrudos[i];
      const suministro = row[h["Suministro"]];
      if (!suministro) continue; // Saltar filas vacías

      const sumId = suministro.toString().trim();
      
      // Si es nuevo, lo creamos
      if (!records[sumId]) {
        const rutaStr = row[h["Ruta"]] ? row[h["Ruta"]].toString().trim() : "";
        records[sumId] = {
          suministro: sumId,
          nombre_cliente: row[h["Nombre"]],
          direccion: row[h["Direccion"]],
          marca: row[h["Marca"]],
          modelo: row[h["Modelo"]],
          serie_fab: row[h["SerieFab"]],
          ruta: rutaStr,
          nomb_ruta: row[h["NombRuta"]],
          tarifa: row[h["Tarifa"]],
          periodo: mesSuministro, // AQUÍ ASIGNAMOS EL MES (Ej. 202605)
          lec_ant: {},
          cuadrilla: cuadrillaMap[rutaStr] || "Sin Asignar",
          estado: "PENDIENTE"
        };
      }
      
      // Guardar el concepto (1 al 6)
      const concepto = row[h["Concepto"]].toString().trim();
      records[sumId].lec_ant[concepto] = row[h["LecAnt"]];
    }

    // 4. Preparar las filas para pegar en Google Sheets
    const outputRows = [];
    for (const key in records) {
      const rec = records[key];
      outputRows.push([
        rec.suministro, rec.nombre_cliente, rec.direccion, rec.marca, rec.modelo,
        rec.serie_fab, rec.ruta, rec.nomb_ruta, rec.tarifa, rec.periodo,
        rec.lec_ant["1"] || 0, rec.lec_ant["2"] || 0, rec.lec_ant["3"] || 0, 
        rec.lec_ant["4"] || 0, rec.lec_ant["5"] || 0, rec.lec_ant["6"] || 0, 
        rec.cuadrilla, rec.estado, "Excel ENOSA", "Múltiple", "", "", "", "", "", "", "", ""
      ]);
    }

    // 5. Pegar en la base de datos
    if (outputRows.length > 0) {
      // Ojo: Usamos append (getLastRow() + 1) para que los meses se vayan sumando hacia abajo
      const startRow = sheetMaestro.getLastRow() + 1;
      sheetMaestro.getRange(startRow, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
    }

    return { 
      success: true, 
      cantidad: outputRows.length, 
      mensaje: `¡Padrón procesado! ${outputRows.length} suministros de ${mesSuministro} fueron consolidados en el Padrón Maestro.` 
    };

  } catch (e) {
    return { success: false, mensaje: e.message };
  }
}

function obtenerMetricasDashboard(mesSuministro) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Padron_Maestro");
    
    if (!sheet) throw new Error("No se encontró la pestaña Padron_Maestro");

    const data = sheet.getDataRange().getValues();
    
    // Si la hoja solo tiene encabezados, retornamos ceros
    if (data.length <= 1) {
      return { success: true, metricas: { total: 0, pendientes: 0, telemedicion: 0, campo: 0, completos: 0 } };
    }

    const headers = data[0];
    
    // Buscar dinámicamente en qué columna están el Periodo y el Estado
    const colPeriodo = headers.findIndex(h => h.toString().toLowerCase() === 'periodo');
    const colEstado = headers.findIndex(h => h.toString().toLowerCase() === 'estado');

    if (colPeriodo === -1 || colEstado === -1) {
      throw new Error("Faltan las columnas 'periodo' o 'estado' en el Padrón Maestro.");
    }

    // Contadores en cero
    let metricas = { total: 0, pendientes: 0, telemedicion: 0, campo: 0, completos: 0 };

    // Iterar desde la fila 1 (saltando encabezados)
    for (let i = 1; i < data.length; i++) {
      let filaMes = data[i][colPeriodo].toString().trim();
      
      // Filtrar estrictamente por el mes que el usuario seleccionó en la web
      if (filaMes === mesSuministro) {
        metricas.total++;
        let estado = data[i][colEstado].toString().trim().toUpperCase();

        // Clasificar según los estados de tu proceso
        if (estado === "PENDIENTE") {
          metricas.pendientes++;
        } else if (estado === "TELEMEDICIÓN OK") {
          metricas.telemedicion++;
        } else if (estado === "ASIGNADO A CAMPO" || estado === "EN PROCESO") {
          // Agrupamos en campo a los que están asignados o pendientes de extraer de Drive
          metricas.campo++;
        } else if (estado === "MEDICIÓN COMPLETA" || estado === "CORTADO / RETIRADO") {
          metricas.completos++;
        }
      }
    }

    return { success: true, metricas: metricas };

  } catch (e) {
    return { success: false, mensaje: e.message };
  }
}