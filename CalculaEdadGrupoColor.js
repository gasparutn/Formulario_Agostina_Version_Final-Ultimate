
/**
 * Archivo: CalculaEdadGrupoColor.gs
 * Contiene la lógica para calcular grupos por FECHA DE CORTE
 * y asignar colores.
 *
 * (CORRECCIÓN): Se ha hecho la comparación de 'aplicarColorGrupo'
 * insensible a mayúsculas/minúsculas.
 */

/**
 * (LÓGICA PROVISTA POR EL USUARIO)
 * Determina el grupo basado en una fecha de corte fija (30 de julio).
 * @param {string} fechaNacStr - La fecha de nacimiento (ej. "2017-10-20").
 * @returns {string} - El texto del grupo (ej. "Grupo 8 años").
 */
function determinarGrupoPorFecha(fechaNacStr) {
  if (!fechaNacStr) return "Sin Fecha";

  try {
    const fechaNac = new Date(fechaNacStr + "T00:00:00Z");
    const mesCorte = 6; // Julio (0-11)
    const diaCorte = 30;

    if (fechaNac >= new Date(Date.UTC(2022, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2023, mesCorte, diaCorte))) return "Grupo 3 años";
    if (fechaNac >= new Date(Date.UTC(2021, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2022, mesCorte, diaCorte))) return "Grupo 4 años";
    if (fechaNac >= new Date(Date.UTC(2020, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2021, mesCorte, diaCorte))) return "Grupo 5 años";
    if (fechaNac >= new Date(Date.UTC(2019, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2020, mesCorte, diaCorte))) return "Grupo 6 años";
    if (fechaNac >= new Date(Date.UTC(2018, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2019, mesCorte, diaCorte))) return "Grupo 7 años";
    if (fechaNac >= new Date(Date.UTC(2017, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2018, mesCorte, diaCorte))) return "Grupo 8 años";
    if (fechaNac >= new Date(Date.UTC(2016, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2017, mesCorte, diaCorte))) return "Grupo 9 años";
    if (fechaNac >= new Date(Date.UTC(2015, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2016, mesCorte, diaCorte))) return "Grupo 10 años";
    if (fechaNac >= new Date(Date.UTC(2014, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2015, mesCorte, diaCorte))) return "Grupo 11 años";
    if (fechaNac >= new Date(Date.UTC(2013, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2014, mesCorte, diaCorte))) return "Grupo 12 años";
    if (fechaNac >= new Date(Date.UTC(2012, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2013, mesCorte, diaCorte))) return "Grupo 13 años";
    if (fechaNac >= new Date(Date.UTC(2011, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2012, mesCorte, diaCorte))) return "Grupo 14 años";
    if (fechaNac >= new Date(Date.UTC(2010, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2011, mesCorte, diaCorte))) return "Grupo 15 años";

    return "Fuera de rango";

  } catch (e) {
    Logger.log("Error al parsear fecha en determinarGrupoPorFecha: " + fechaNacStr + " | Error: " + e.message);
    return "Error Fecha";
  }
}
/*
// este bloque calcula igual la edad con fecha de corte sin la referencia, pero es atumàtico, no requiere volver a refecncias el año
function determinarGrupoPorFecha(fechaNacStr, anioReferencia) {
  if (!fechaNacStr) return "Sin Fecha";

  try {
    const fechaNac = new Date(fechaNacStr + "T00:00:00Z");
    const fechaCorte = new Date(Date.UTC(anioReferencia, 5, 30)); // 30 de junio (mes 5 porque enero=0)

    // Calcular edad al corte
    let edad = fechaCorte.getUTCFullYear() - fechaNac.getUTCFullYear();
    const cumpleAntesDelCorte = 
      (fechaNac.getUTCMonth() < fechaCorte.getUTCMonth()) ||
      (fechaNac.getUTCMonth() === fechaCorte.getUTCMonth() && fechaNac.getUTCDate() <= fechaCorte.getUTCDate());

    if (!cumpleAntesDelCorte) {
      edad--; // Si cumple después del corte, aún no alcanza esa edad
    }

    return Grupo ${edad} años;

  } catch (e) {
    return "Error Fecha";
  }
}
*/
/**
 * Aplica el color de fondo a la celda del grupo basado en la hoja de Configuración.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja - La hoja de "Registros".
 * @param {number} fila - El número de fila a colorear.
 * @param {string} textoGrupo - El texto del grupo (ej. "Grupo 5 años").
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaConfig - La hoja de "Config".
 */
function aplicarColorGrupo(hoja, fila, textoGrupo, hojaConfig) {
  try {
    const rangoGrupos = hojaConfig.getRange("A30:B41");
    const valoresGrupos = rangoGrupos.getValues();
    const coloresGrupos = rangoGrupos.getBackgrounds();

    for (let i = 0; i < valoresGrupos.length; i++) {
      
      // =========================================================
      // --- ¡¡AQUÍ ESTÁ LA CORRECCIÓN!! ---
      // Comparamos todo en mayúsculas y sin espacios extra.
      // =========================================================
      if (valoresGrupos[i][0].toString().trim().toUpperCase() == textoGrupo.toString().trim().toUpperCase()) {
        const color = coloresGrupos[i][1]; 
        
        hoja.getRange(fila, COL_GRUPOS).setBackground(color);
        return;
      }
    }
    // Si llega aquí, es porque no encontró una coincidencia de texto
    Logger.log(`No se encontró color para el grupo: "${textoGrupo}" en la hoja Config!A30:A41`);

  } catch (e) {
    Logger.log(`Error al aplicar color para el grupo ${textoGrupo} en fila ${fila}: ${e.message}`);
  }
}

function procesarFilaIndividual(ss, hoja, fila, fechaNac) {
  try {
    // 1. Determinar el grupo
    const fechaNacStr = Utilities.formatDate(fechaNac, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    const grupo = determinarGrupoPorFecha(fechaNacStr);

    // 2. Escribir el grupo en la celda correspondiente
    hoja.getRange(fila, COL_GRUPOS).setValue(grupo);

    // 3. Aplicar el color
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    if (hojaConfig) {
      aplicarColorGrupo(hoja, fila, grupo, hojaConfig);
    } else {
      Logger.log(`No se encontró la hoja de configuración "${NOMBRE_HOJA_CONFIG}" para aplicar color.`);
    }
  } catch (e) {
    Logger.log(`Error en procesarFilaIndividual para fila ${fila}: ${e.message}`);
  }
}
