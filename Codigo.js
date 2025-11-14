/**
 * (MODIFICADO v15-CORREGIDO)
 * - Se actualiza el número de columnas de 48 a 47 en `registrarDatos`.
 * - Todas las demás referencias (COL_COLEGIO_JARDIN, COL_DNI_INSCRIPTO, etc.)
 * se actualizan automáticamente desde el nuevo `Constantes.js`.
 */
function doGet(e) {
  try {
    const params = e.parameter;
    Logger.log("doGet INICIADO. Parámetros de URL: " + JSON.stringify(params));

    const appUrl = ScriptApp.getService().getUrl();
    const htmlTemplate = HtmlService.createTemplateFromFile("Index");
    htmlTemplate.appUrl = appUrl;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    htmlTemplate.pagoTotalMPVisible = hojaConfig.getRange("B24").getValue();

    htmlTemplate.dniHermano = params.dni || "";
    htmlTemplate.tipoHermano = params.tipo || "";
    htmlTemplate.nombreHermano = "";
    htmlTemplate.apellidoHermano = "";
    htmlTemplate.fechaNacHermano = "";

    const html = htmlTemplate
      .evaluate()
      .setTitle("Formulario de Registro")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
    return html;
  } catch (err) {
    Logger.log(
      "Error en la detección de parámetros de doGet: " + err.toString()
    );
    // (CORRECCIÓN v8)
    return HtmlService.createHtmlOutput(
      "<b>Ocurrió un error:</b> " + err.toString()
    );
  }
}

/**
 * 'doPost' ya no es necesario para el webhook de MP.
 */
function doPost(e) {
  Logger.log("doPost llamado pero Mercado Pago está deshabilitado.");
  return ContentService.createTextOutput("OK");
}

// =========================================================
// --- FUNCIÓN AÑADIDA (SOLICITUD DEL USUARIO) ---
// Esta es la función que proporcionaste, determina el grupo por fecha de corte.
// =========================================================
function obtenerGrupoPorFechaNacimiento(fechaNacStr) {
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
    Logger.log("Error al parsear fecha en obtenerGrupoPorFechaNacimiento: " + fechaNacStr + " | Error: " + e.message);
    return "Error Fecha";
  }
}
// =========================================================
// --- FIN FUNCIÓN AÑADIDA ---
// =========================================================

// =========================================================
// --- FUNCIÓN AÑADIDA Y MODIFICADA (Para pintar celdas) ---
// Se trae de 'CalculaEdadGrupoColor.js' y se modifica 
// para que NO convierta a mayúsculas.
// =========================================================
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
      // Comparamos el texto exacto (ej. "Grupo 3 años") 
      // sin forzar mayúsculas.
      // =========================================================
      if (valoresGrupos[i][0].toString().trim() == textoGrupo.toString().trim()) {
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
// =========================================================
// --- FIN FUNCIÓN MODIFICADA ---
// =========================================================


/**
* (MODIFICADO v15-CORREGIDO)
* - Se actualiza el número de columnas de 48 a 47 en `registrarDatos`.
* - Todas las constantes (COL_COLEGIO_JARDIN, COL_DNI_INSCRIPTO, etc.)
* se leen desde el nuevo Constantes.js y apuntan a las columnas correctas.
*/
function registrarDatos(datos, testSheetName) {
  Logger.log("REGISTRAR DATOS INICIADO. Datos: " + JSON.stringify(datos));
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);

    const fechaNacPrincipal = datos.fechaNacimiento;
    if (
      !fechaNacPrincipal ||
      fechaNacPrincipal < "2010-01-01" ||
      fechaNacPrincipal > "2023-12-31"
    ) {
      return {
        status: "ERROR",
        message:
          "La fecha de nacimiento del inscripto principal debe estar entre 01/01/2010 y 31/12/2023.",
      };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    let estadoActual = obtenerEstadoRegistro();

    // Solo aplicar cupos si NO es un hermano (los hermanos ya tienen cupo)
    if (datos.tipoInscripto !== "hermano/a") {
      if (estadoActual.cierreManual)
        return {
          status: "CERRADO",
          message: "El registro se encuentra cerrado.",
        };
      if (datos.tipoInscripto !== "preventa" && estadoActual.alcanzado)
        return {
          status: "LIMITE_ALCANZADO",
          message: "Se ha alcanzado el cupo máximo.",
        };
      if (
        datos.tipoInscripto !== "preventa" &&
        datos.jornada === "Jornada Normal extendida" &&
        estadoActual.jornadaExtendidaAlcanzada
      ) {
        return {
          status: "LIMITE_EXTENDIDA",
          message: "Se agotó el cupo para Jornada Extendida.",
        };
      }
    }

    const dniBuscado = limpiarDNI(datos.dni);
    const hojaRegistroName = testSheetName || NOMBRE_HOJA_REGISTRO;
    let hojaRegistro = ss.getSheetByName(hojaRegistroName);
    
    if (!hojaRegistro) {
      return { status: 'ERROR', message: "Hoja de registros '" + hojaRegistroName + "' no encontrada." };
    }

    // Normalizar DNI
    const dniLimpio = limpiarDNI(datos.dni);

    // Buscar si ya existe el DNI en registros
    let filaExistente = null;
    if (hojaRegistro.getLastRow() > 1) {
      // (MODIFICADO v15-CORREGIDO) COL_DNI_INSCRIPTO ahora apunta a J (10)
      const rangoDNI = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
      const celda = rangoDNI.createTextFinder(dniLimpio).matchEntireCell(true).findNext();
      if (celda) filaExistente = celda.getRow();
    }

    // Si existe, actualizar en su lugar (comportamiento seguro para evitar duplicados)
    if (filaExistente) {
      try {
        // Actualizar campos principales
        
        const telResp1 = datos.telAreaResp1 && datos.telNumResp1 ? '(' + datos.telAreaResp1 + ') ' + datos.telNumResp1 : datos.telResp1 || '';
        const telResp2 = datos.telAreaResp2 && datos.telNumResp2 ? '(' + datos.telAreaResp2 + ') ' + datos.telNumResp2 : datos.telResp2 || '';

        hojaRegistro.getRange(filaExistente, COL_EMAIL).setValue(datos.email || ''); // E
        hojaRegistro.getRange(filaExistente, COL_NOMBRE).setValue(datos.nombre || ''); // F
        hojaRegistro.getRange(filaExistente, COL_APELLIDO).setValue(datos.apellido || ''); // G

        let fechaNacObj = null;
        if (datos.fechaNacimiento) {
          fechaNacObj = new Date(datos.fechaNacimiento);
          fechaNacObj.setMinutes(fechaNacObj.getMinutes() + fechaNacObj.getTimezoneOffset());
        }
        hojaRegistro.getRange(filaExistente, COL_FECHA_NACIMIENTO_REGISTRO).setValue(fechaNacObj || ''); // H

        // (MODIFICADO v15-CORREGIDO) COL_OBRA_SOCIAL (K) y COL_COLEGIO_JARDIN (L)
        hojaRegistro.getRange(filaExistente, COL_OBRA_SOCIAL).setValue(datos.obraSocial || ''); // K
        hojaRegistro.getRange(filaExistente, COL_COLEGIO_JARDIN).setValue(datos.colegioJardin || ''); // L
        
        hojaRegistro.getRange(filaExistente, COL_ADULTO_RESPONSABLE_1).setValue(datos.adultoResponsable1 || ''); // M
        hojaRegistro.getRange(filaExistente, COL_DNI_RESPONSABLE_1).setValue(datos.dniResponsable1 || ''); // N
        hojaRegistro.getRange(filaExistente, COL_TEL_RESPONSABLE_1).setValue(telResp1); // O
        hojaRegistro.getRange(filaExistente, COL_ADULTO_RESPONSABLE_2).setValue(datos.adultoResponsable2 || ''); // P
        hojaRegistro.getRange(filaExistente, COL_TEL_RESPONSABLE_2).setValue(telResp2); // Q
        hojaRegistro.getRange(filaExistente, COL_PERSONAS_AUTORIZADAS).setValue(datos.personasAutorizadas || ''); // R
        
        // Campos de salud (S-X)
        hojaRegistro.getRange(filaExistente, COL_PRACTICA_DEPORTE).setValue(datos.practicaDeporte || ''); // S
        hojaRegistro.getRange(filaExistente, COL_ESPECIFIQUE_DEPORTE).setValue(datos.especifiqueDeporte || ''); // T
        hojaRegistro.getRange(filaExistente, COL_TIENE_ENFERMEDAD).setValue(datos.tieneEnfermedad || ''); // U
        hojaRegistro.getRange(filaExistente, COL_ESPECIFIQUE_ENFERMEDAD).setValue(datos.especifiqueEnfermedad || ''); // V
        hojaRegistro.getRange(filaExistente, COL_ES_ALERGICO).setValue(datos.esAlergico || ''); // W
        hojaRegistro.getRange(filaExistente, COL_ESPECIFIQUE_ALERGIA).setValue(datos.especifiqueAlergia || ''); // X

        // Jornada y Pago (AA en adelante)
        hojaRegistro.getRange(filaExistente, COL_JORNADA).setValue(datos.jornada || ''); // AA
        hojaRegistro.getRange(filaExistente, COL_SOCIO).setValue(datos.esSocio || ''); // AB
        
        // (MODIFICADO v15) La lógica de v14 sigue funcionando, pero las constantes apuntan a nuevos lugares
        hojaRegistro.getRange(filaExistente, COL_METODO_PAGO).setValue(datos.metodoPago || ''); // AC
        hojaRegistro.getRange(filaExistente, COL_MODO_PAGO_CUOTA).setValue(datos.subMetodoCuotas || ''); // AD
        hojaRegistro.getRange(filaExistente, COL_PRECIO).setValue(datos.precio || ''); // AE (Precio Total)
        hojaRegistro.getRange(filaExistente, COL_CANTIDAD_CUOTAS).setValue(datos.cantidadCuotas || ''); // AI
        hojaRegistro.getRange(filaExistente, COL_ESTADO_PAGO).setValue(datos.estadoPago || ''); // AJ
        hojaRegistro.getRange(filaExistente, COL_MONTO_A_PAGAR).setValue(datos.montoAPagar); // AK (Vacío o Total)

        // Escribir el valor de cuota individual en AF, AG, AH
        if (datos.cantidadCuotas === 3 && datos.valorCuota > 0) {
          hojaRegistro.getRange(filaExistente, COL_CUOTA_1).setValue(datos.valorCuota); // AF
          hojaRegistro.getRange(filaExistente, COL_CUOTA_2).setValue(datos.valorCuota); // AG
          hojaRegistro.getRange(filaExistente, COL_CUOTA_3).setValue(datos.valorCuota); // AH
        }

        // Pagador manual (AL, AM)
        if (datos.pagadorNombreManual) {
          hojaRegistro.getRange(filaExistente, COL_PAGADOR_NOMBRE_MANUAL).setValue(datos.pagadorNombreManual); // AL
        }
        if (datos.pagadorDniManual) {
          hojaRegistro.getRange(filaExistente, COL_PAGADOR_DNI_MANUAL).setValue(datos.pagadorDniManual); // AM
        }

        // Aptitud (Y) / Foto (Z)
        if (datos.urlCertificadoAptitud) {
          const val = String(datos.urlCertificadoAptitud).startsWith('=HYPERLINK') ? datos.urlCertificadoAptitud : `=HYPERLINK("${datos.urlCertificadoAptitud}"; "Aptitud_${dniLimpio}")`;
          hojaRegistro.getRange(filaExistente, COL_APTITUD_FISICA).setValue(val); // Y
        }
        if (datos.urlFotoCarnet) {
          const valf = String(datos.urlFotoCarnet).startsWith('=HYPERLINK') ? datos.urlFotoCarnet : `=HYPERLINK("${datos.urlFotoCarnet}"; "Foto_${dniLimpio}")`;
          hojaRegistro.getRange(filaExistente, COL_FOTO_CARNET).setValue(valf); // Z
        }
        
        // Vínculo familiar (AR)
        if (datos.vinculoPrincipal) {
           hojaRegistro.getRange(filaExistente, COL_VINCULO_PRINCIPAL).setValue(datos.vinculoPrincipal); // AR
        }

        // --- Cálculo de Grupo y Color (H, I) ---
        try {
          const fechaNacStr = datos.fechaNacimiento;
          if (fechaNacStr) {
            // =========================================================
            // --- ¡¡MODIFICACIÓN #1!! ---
            // Se llama a la nueva función del usuario.
            // =========================================================
            const grupo = obtenerGrupoPorFechaNacimiento(fechaNacStr);
            // =========================================================
            
            hojaRegistro.getRange(filaExistente, COL_GRUPOS).setValue(grupo); // I
            aplicarColorGrupo(hojaRegistro, filaExistente, grupo, hojaConfig);
            Logger.log(`Grupo [${grupo}] y color RE-aplicados para DNI ${dniLimpio} en fila ${filaExistente}.`);
          }
        } catch (e) {
          Logger.log(`Error al RE-aplicar grupo/color para ${dniLimpio} en fila ${filaExistente}: ${e.message}`);
        }
        // --- FIN DE LA CORRECCIÓN ---

        SpreadsheetApp.flush();
        const numeroDeTurno = hojaRegistro.getRange(filaExistente, COL_NUMERO_TURNO).getValue();
        Logger.log(`Registro actualizado para DNI ${dniLimpio} en fila ${filaExistente}. Turno: ${numeroDeTurno}`);
        return { status: 'OK_REGISTRO', message: 'Registro actualizado correctamente.', numeroDeTurno: numeroDeTurno, datos: datos };
      } catch (e) {
        Logger.log('Error actualizando registro existente: ' + e.toString());
        return { status: 'ERROR', message: 'Error al actualizar el registro: ' + e.message };
      }
    }

    // Nuevo registro: preparar datos e insertar
    const lastRow = hojaRegistro.getLastRow();
    const registrosActuales = lastRow > 1 ? hojaRegistro.getRange(2, COL_NUMERO_TURNO, lastRow - 1, 1).getValues().filter(r => r[0] !== '' && r[0] != null).length : 0;
    const numeroDeTurno = registrosActuales + 1;

    // --- MODIFICACIÓN v15-CORREGIDO ---
    // Usamos el número de columnas de la hoja (que ahora es 47)
    const totalCols = Math.max(47, hojaRegistro.getLastColumn()); 
    const valoresFila = new Array(totalCols).fill('');
    // --- FIN MODIFICACIÓN v15-CORREGIDO ---

    valoresFila[COL_NUMERO_TURNO - 1] = numeroDeTurno; // A
    valoresFila[COL_MARCA_TEMPORAL - 1] = new Date(); // B

    const esPreventaReg = datos.esPreventa === true || datos.tipoInscripto === 'preventa';
    let marcaNE = '';
    if (datos.jornada === 'Jornada Normal extendida') {
      marcaNE = esPreventaReg ? 'Extendida (Pre-venta)' : 'Extendida';
    } else {
      marcaNE = esPreventaReg ? 'Normal (Pre-Venta)' : 'Normal';
    }
    valoresFila[COL_MARCA_N_E_A - 1] = marcaNE; // C

    let estadoNuevoAnt = 'Nuevo';
    if (datos.tipoInscripto === 'hermano/a') {
      if (datos.tipoInscripcionOriginal === 'preventa') {
        estadoNuevoAnt = 'Pre-Venta Hermano/a';
      } else if (datos.tipoInscripcionOriginal === 'anterior') {
        estadoNuevoAnt = 'Anterior Hermano/a';
      } else {
        estadoNuevoAnt = 'Nuevo Hermano/a';
      }
    } else {
      if (esPreventaReg) estadoNuevoAnt = 'Pre-Venta';
      else if (datos.tipoInscripto === 'anterior') estadoNuevoAnt = 'Anterior';
    }
    valoresFila[COL_ESTADO_NUEVO_ANT - 1] = estadoNuevoAnt; // D

    valoresFila[COL_EMAIL - 1] = datos.email || ''; // E
    valoresFila[COL_NOMBRE - 1] = datos.nombre || ''; // F
    valoresFila[COL_APELLIDO - 1] = datos.apellido || ''; // G

    let fechaNacObjNueva = null;
    if (datos.fechaNacimiento) {
      fechaNacObjNueva = new Date(datos.fechaNacimiento);
      fechaNacObjNueva.setMinutes(fechaNacObjNueva.getMinutes() + fechaNacObjNueva.getTimezoneOffset());
    }
    valoresFila[COL_FECHA_NACIMIENTO_REGISTRO - 1] = fechaNacObjNueva || ''; // H

    // (MODIFICADO v15-CORREGIDO)
    valoresFila[COL_DNI_INSCRIPTO - 1] = dniLimpio || ''; // J
    valoresFila[COL_OBRA_SOCIAL - 1] = datos.obraSocial || ''; // K
    valoresFila[COL_COLEGIO_JARDIN - 1] = datos.colegioJardin || ''; // L
    
    valoresFila[COL_ADULTO_RESPONSABLE_1 - 1] = datos.adultoResponsable1 || ''; // M
    valoresFila[COL_DNI_RESPONSABLE_1 - 1] = datos.dniResponsable1 || ''; // N
    const telResp1 = (datos.telAreaResp1 && datos.telNumResp1) ? '(' + datos.telAreaResp1 + ') ' + datos.telNumResp1 : '';
    const telResp2 = (datos.telAreaResp2 && datos.telNumResp2) ? '(' + datos.telAreaResp2 + ') ' + datos.telNumResp2 : '';
    valoresFila[COL_TEL_RESPONSABLE_1 - 1] = telResp1; // O
    valoresFila[COL_ADULTO_RESPONSABLE_2 - 1] = datos.adultoResponsable2 || ''; // P
    valoresFila[COL_TEL_RESPONSABLE_2 - 1] = telResp2; // Q

    valoresFila[COL_PERSONAS_AUTORIZADAS - 1] = datos.personasAutorizadas || ''; // R

    // Campos de salud (S-X)
    valoresFila[COL_PRACTICA_DEPORTE - 1] = datos.practicaDeporte || ''; // S
    valoresFila[COL_ESPECIFIQUE_DEPORTE - 1] = datos.especifiqueDeporte || ''; // T
    valoresFila[COL_TIENE_ENFERMEDAD - 1] = datos.tieneEnfermedad || ''; // U
    valoresFila[COL_ESPECIFIQUE_ENFERMEDAD - 1] = datos.especifiqueEnfermedad || ''; // V
    valoresFila[COL_ES_ALERGICO - 1] = datos.esAlergico || ''; // W
    valoresFila[COL_ESPECIFIQUE_ALERGIA - 1] = datos.especifiqueAlergia || ''; // X

    // Aptitud (Y) y Foto (Z)
    if (datos.urlCertificadoAptitud) {
      valoresFila[COL_APTITUD_FISICA - 1] = String(datos.urlCertificadoAptitud).startsWith('=HYPERLINK') ? datos.urlCertificadoAptitud : `=HYPERLINK("${datos.urlCertificadoAptitud}"; "Aptitud_${dniLimpio}")`; // Y
    }
    if (datos.urlFotoCarnet) {
      valoresFila[COL_FOTO_CARNET - 1] = String(datos.urlFotoCarnet).startsWith('=HYPERLINK') ? datos.urlFotoCarnet : `=HYPERLINK("${datos.urlFotoCarnet}"; "Foto_${dniLimpio}")`; // Z
    }

    valoresFila[COL_JORNADA - 1] = datos.jornada || ''; // AA
    valoresFila[COL_SOCIO - 1] = datos.esSocio || ''; // AB

    // (MODIFICADO v15) La lógica de v14 sigue funcionando, pero las constantes apuntan a nuevos lugares
    valoresFila[COL_METODO_PAGO - 1] = datos.metodoPago || ''; // AC
    valoresFila[COL_MODO_PAGO_CUOTA - 1] = datos.subMetodoCuotas || ''; // AD
    valoresFila[COL_PRECIO - 1] = datos.precio || ''; // AE (Precio Total)
    valoresFila[COL_CANTIDAD_CUOTAS - 1] = datos.cantidadCuotas || ''; // AI
    valoresFila[COL_ESTADO_PAGO - 1] = datos.estadoPago || ''; // AJ
    valoresFila[COL_MONTO_A_PAGAR - 1] = datos.montoAPagar; // AK (Vacío o Total)

    // Escribir el valor de cuota individual en AF, AG, AH
    if (datos.cantidadCuotas === 3 && datos.valorCuota > 0) {
      valoresFila[COL_CUOTA_1 - 1] = datos.valorCuota; // AF
      valoresFila[COL_CUOTA_2 - 1] = datos.valorCuota; // AG
      valoresFila[COL_CUOTA_3 - 1] = datos.valorCuota; // AH
    }
    
    // Vinculo familiar (AR)
    if (datos.vinculoPrincipal) {
      valoresFila[COL_VINCULO_PRINCIPAL - 1] = datos.vinculoPrincipal;
    } else if (datos.hermanos && datos.hermanos.length > 0) {
      valoresFila[COL_VINCULO_PRINCIPAL - 1] = `FAM_${numeroDeTurno}`;
    }

    // Pagador manual inicial (AL, AM)
    if (datos.pagadorNombreManual) valoresFila[COL_PAGADOR_NOMBRE_MANUAL - 1] = datos.pagadorNombreManual; // AL
    if (datos.pagadorDniManual) valoresFila[COL_PAGADOR_DNI_MANUAL - 1] = datos.pagadorDniManual; // AM

    // Insertar la fila
    hojaRegistro.appendRow(valoresFila);
    const nuevaFila = hojaRegistro.getLastRow();
    
    // --- Cálculo de Grupo y Color (H, I) ---
    try {
      const fechaNacStr = datos.fechaNacimiento;
      if (fechaNacStr) {
        // =========================================================
        // --- ¡¡MODIFICACIÓN #2!! ---
        // Se llama a la nueva función del usuario.
        // =========================================================
        const grupo = obtenerGrupoPorFechaNacimiento(fechaNacStr);
        // =========================================================

        hojaRegistro.getRange(nuevaFila, COL_GRUPOS).setValue(grupo); // I
        aplicarColorGrupo(hojaRegistro, nuevaFila, grupo, hojaConfig);
      } else {
         Logger.log(`No se pudo calcular el grupo para ${dniLimpio}: sin fecha de nacimiento.`);
      }
    } catch (e) {
      Logger.log(`Error al aplicar grupo/color para ${dniLimpio} en fila ${nuevaFila}: ${e.message}`);
    }
    // --- FIN DE LA CORRECCIÓN ---
    SpreadsheetApp.flush();
    Logger.log(`Nuevo registro creado para DNI ${dniLimpio} en fila ${nuevaFila}. Turno: ${numeroDeTurno}`);

    return { status: 'OK_REGISTRO', message: 'Registro creado con éxito.', numeroDeTurno: numeroDeTurno, datos: datos };
  } catch (e) {
    Logger.log("Error en registrarDatos: " + e.toString() + " Stack: " + e.stack);
    return { status: "ERROR", message: "Error en el servidor: " + e.message };
  } finally {
    try {
      lock.releaseLock();
    } catch (er) {
      // noop
    }
  }
}

/**
 * Permite a un usuario ya registrado editar campos específicos.
 * (MODIFICADO v15) No requiere cambios, usa constantes.
 */
function actualizarDatosPersonales(dni, datosEditados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const dniLimpio = limpiarDNI(dni);
    if (!dniLimpio || !datosEditados)
      return {
        status: "ERROR",
        message: "Faltan datos (DNI o datos a editar).",
      };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hoja)
      throw new Error(`La hoja "${NOMBRE_HOJA_REGISTRO}" no fue encontrada.`);
    const rangoDni = hoja.getRange(
      2,
      COL_DNI_INSCRIPTO,
      hoja.getLastRow() - 1,
      1
    );
    const celdaEncontrada = rangoDni
      .createTextFinder(dniLimpio)
      .matchEntireCell(true)
      .findNext();

    if (celdaEncontrada) {
      const fila = celdaEncontrada.getRow();
      if (datosEditados.adultoResponsable1 !== undefined)
        hoja
          .getRange(fila, COL_ADULTO_RESPONSABLE_1)
          .setValue(datosEditados.adultoResponsable1);
      if (datosEditados.dniResponsable1 !== undefined)
        hoja
          .getRange(fila, COL_DNI_RESPONSABLE_1)
          .setValue(datosEditados.dniResponsable1);
      if (datosEditados.telResp1 !== undefined)
        hoja
          .getRange(fila, COL_TEL_RESPONSABLE_1)
          .setValue(datosEditados.telResp1);
      if (datosEditados.adultoResponsable2 !== undefined)
        hoja
          .getRange(fila, COL_ADULTO_RESPONSABLE_2)
          .setValue(datosEditados.adultoResponsable2);
      if (datosEditados.telResp2 !== undefined)
        hoja
          .getRange(fila, COL_TEL_RESPONSABLE_2)
          .setValue(datosEditados.telResp2);
      if (datosEditados.personasAutorizadas !== undefined)
        hoja
          .getRange(fila, COL_PERSONAS_AUTORIZADAS)
          .setValue(datosEditados.personasAutorizadas);

      if (
        datosEditados.urlCertificadoAptitud !== undefined &&
        datosEditados.urlCertificadoAptitud.startsWith("http")
      ) {
        const extension = datosEditados.urlCertificadoAptitud.includes(".")
          ? datosEditados.urlCertificadoAptitud.split(".").pop()
          : "pdf";
        const nuevoNombreAptitud = `AptitudFisica_${dniLimpio}.${extension}`;
        const formulaLink = `=HYPERLINK("${datosEditados.urlCertificadoAptitud}"; "${nuevoNombreAptitud}")`;
        hoja.getRange(fila, COL_APTITUD_FISICA).setValue(formulaLink);
      }

      Logger.log(
        `Datos personales actualizados para DNI ${dniLimpio} en fila ${fila}.`
      );
      return { status: "OK", message: "¡Datos actualizados con éxito!" };
    } else {
      Logger.log(
        `No se encontró DNI ${dniLimpio} para actualizar datos personales.`
      );
      return {
        status: "ERROR",
        message: `No se encontró el registro para el DNI ${dniLimpio}.`,
      };
    }
  } catch (e) {
    Logger.log("Error en actualizarDatosPersonales: " + e.toString());
    return { status: "ERROR", message: "Error en el servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Calcula la edad en años, meses y días a partir de una fecha de nacimiento.
 */
function calcularEdadDetallada(fechaNacimiento) {
  if (!fechaNacimiento || !(fechaNacimiento instanceof Date)) {
    return { anos: 0, meses: 0, dias: 0 };
  }
  const hoy = new Date();
  const fechaNac = new Date(fechaNacimiento);
  fechaNac.setMinutes(fechaNac.getMinutes() + fechaNac.getTimezoneOffset());

  let anos = hoy.getFullYear() - fechaNac.getFullYear();
  let meses = hoy.getMonth() - fechaNac.getMonth();
  let dias = hoy.getDate() - fechaNac.getDate();

  if (dias < 0) {
    meses--;
    dias += new Date(hoy.getFullYear(), hoy.getMonth(), 0).getDate();
  }
  if (meses < 0) {
    anos--;
    meses += 12;
  }
  return { anos, meses, dias };
}

function limpiarDNI(dni) {
  if (!dni) return "";
  return String(dni)
    .replace(/[.\s-]/g, "")
    .trim();
}

function obtenerEstadoRegistro() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hojaConfig)
      throw new Error(`Hoja "${NOMBRE_HOJA_CONFIG}" no encontrada.`);
    const limiteCupos = parseInt(hojaConfig.getRange("B1").getValue()) || 100;
    const limiteJornadaExtendida = parseInt(
      hojaConfig.getRange("B4").getValue()
    );
    const formularioAbierto = hojaConfig.getRange("B11").getValue() === true;
    let registrosActuales = 0;
    let registrosJornadaExtendida = 0;
    if (hojaRegistro && hojaRegistro.getLastRow() > 1) {
      const lastRow = hojaRegistro.getLastRow();
      registrosActuales = hojaRegistro
        .getRange(2, COL_NUMERO_TURNO, lastRow - 1, 1)
        .getValues()
        .filter((fila) => fila[0] != null && fila[0] != "").length;
      registrosJornadaExtendida = hojaRegistro
        .getRange(2, COL_MARCA_N_E_A, lastRow - 1, 1)
        .getValues()
        .filter((row) => String(row[0]).startsWith("Extendida")).length;
    }
    hojaConfig.getRange("B2").setValue(registrosActuales);
    hojaConfig.getRange("B5").setValue(registrosJornadaExtendida);
    return {
      alcanzado: registrosActuales >= limiteCupos,
      jornadaExtendidaAlcanzada:
        registrosJornadaExtendida >= limiteJornadaExtendida,
      cierreManual: !formularioAbierto,
    };
  } catch (e) {
    Logger.log("Error en obtenerEstadoRegistro: " + e.message);
    return {
      cierreManual: true,
      message: "Error al leer config: " + e.message,
    };
  }
}

/**
 * (MODIFICADO v15-CORREGIDO)
 * - Esta función no requiere cambios internos, ya que todas las
 * lecturas de columnas (COL_DNI_INSCRIPTO, COL_ESTADO_PAGO, etc.)
 * se actualizan automáticamente desde `Constantes.js`.
 */
function validarAcceso(dni, tipoInscripto) {
  try {
    if (!dni || !/^[0-9]{8}$/.test(dni.trim()))
      return {
        status: "ERROR",
        message: "El DNI debe tener exactamente 8 dígitos numéricos.",
      };
    const dniLimpio = limpiarDNI(dni);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    if (!hojaConfig)
      return {
        status: "ERROR",
        message: `La hoja de configuración "${NOMBRE_HOJA_CONFIG}" no fue encontrada.`,
      };
    const pagoTotalMPVisible = hojaConfig.getRange("B24").getValue() === true;
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);

    if (hojaRegistro && hojaRegistro.getLastRow() > 1) {
      const celdaRegistro = hojaRegistro
        .getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1) // Ahora J (10)
        .createTextFinder(dniLimpio)
        .matchEntireCell(true)
        .findNext();
      if (celdaRegistro) {
        const estado = obtenerEstadoRegistro();
        if (estado.cierreManual)
          return {
            status: "CERRADO",
            message: "El formulario se encuentra cerrado por mantenimiento.",
          };
        return gestionarUsuarioYaRegistrado(
          ss,
          hojaRegistro,
          celdaRegistro.getRow(),
          dniLimpio,
          estado,
          tipoInscripto,
          pagoTotalMPVisible
        );
      }
    }

    const estado = obtenerEstadoRegistro();
    if (estado.cierreManual)
      return {
        status: "CERRADO",
        message: "El formulario se encuentra cerrado por mantenimiento.",
      };
    if (estado.alcanzado && tipoInscripto !== "preventa")
      return {
        status: "LIMITE_ALCANZADO",
        message: "Se ha alcanzado el cupo máximo para nuevos registros.",
      };

    const hojaPreventa = ss.getSheetByName(NOMBRE_HOJA_PREVENTA);
    if (tipoInscripto === "preventa") {
      if (!hojaPreventa)
        return {
          status: "ERROR",
          message: `La hoja de configuración "${NOMBRE_HOJA_PREVENTA}" no fue encontrada.`,
        };
      const celdaEncontrada = hojaPreventa
        .getRange(2, COL_PREVENTA_DNI, hojaPreventa.getLastRow() - 1, 1)
        .createTextFinder(dniLimpio)
        .matchEntireCell(true)
        .findNext();
      if (!celdaEncontrada)
        return {
          status: "ERROR_TIPO_ANT",
          message: `El DNI ${dniLimpio} no se encuentra en la base de datos de Pre-Venta.`,
        };

      const fila = hojaPreventa
        .getRange(celdaEncontrada.getRow(), 1, 1, hojaPreventa.getLastColumn())
        .getValues()[0];
      const jornadaGuarda = String(fila[COL_PREVENTA_GUARDA - 1])
        .trim()
        .toLowerCase();
      const jornadaPredefinida =
        jornadaGuarda.includes("si") || jornadaGuarda.includes("extendida")
          ? "Jornada Normal extendida"
          : "Jornada Normal";
      if (
        jornadaPredefinida === "Jornada Normal extendida" &&
        estado.jornadaExtendidaAlcanzada
      )
        return {
          status: "LIMITE_EXTENDIDA",
          message:
            "Su DNI de Pre-Venta corresponde a Jornada Extendida, pero el cupo ya se ha agotado.",
        };

      const fechaNacimientoRaw = fila[COL_PREVENTA_FECHA_NAC - 1];
      const fechaNacimientoStr =
        fechaNacimientoRaw instanceof Date
          ? Utilities.formatDate(
              fechaNacimientoRaw,
              ss.getSpreadsheetTimeZone(),
              "yyyy-MM-dd"
            )
          : "";
      return {
        status: "OK_PREVENTA",
        message: "✅ DNI de Pre-Venta validado. Se autocompletarán sus datos.",
        datos: {
          email: fila[COL_PREVENTA_EMAIL - 1],
          nombre: fila[COL_PREVENTA_NOMBRE - 1],
          apellido: fila[COL_PREVENTA_APELLIDO - 1],
          dni: dniLimpio,
          fechaNacimiento: fechaNacimientoStr,
          jornada: jornadaPredefinida,
          esPreventa: true,
        },
        jornadaExtendidaAlcanzada: estado.jornadaExtendidaAlcanzada,
        tipoInscripto: tipoInscripto,
        pagoTotalMPVisible: pagoTotalMPVisible,
      };
    }

    const hojaBusqueda = ss.getSheetByName(NOMBRE_HOJA_BUSQUEDA);
    if (!hojaBusqueda)
      return {
        status: "ERROR",
        message: `La hoja "${NOMBRE_HOJA_BUSQUEDA}" no fue encontrada.`,
      };
    const celdaEncontrada = hojaBusqueda
      .getRange(2, COL_DNI_BUSQUEDA, hojaBusqueda.getLastRow() - 1, 1)
      .createTextFinder(dniLimpio)
      .matchEntireCell(true)
      .findNext();

    if (celdaEncontrada) {
      if (
        hojaPreventa &&
        hojaPreventa.getLastRow() > 1 &&
        hojaPreventa
          .getRange(2, COL_PREVENTA_DNI, hojaPreventa.getLastRow() - 1, 1)
          .createTextFinder(dniLimpio)
          .matchEntireCell(true)
          .findNext()
      ) {
        return {
          status: "ERROR_TIPO_ANT",
          message:
            'Usted tiene un cupo Pre-Venta. Por favor, elija la opción "Inscripto PRE-VENTA" para validar.',
        };
      }
      if (tipoInscripto === "nuevo")
        return {
          status: "ERROR_TIPO_NUEVO",
          message:
            "El DNI se encuentra en nuestra base de datos. Por favor, seleccione 'Soy Inscripto Anterior' y valide nuevamente.",
        };

      const fila = hojaBusqueda
        .getRange(celdaEncontrada.getRow(), COL_HABILITADO_BUSQUEDA, 1, 10)
        .getValues()[0];
      if (fila[0] !== true)
        return {
          status: "NO_HABILITADO",
          message:
            "El DNI se encuentra en la base de datos, pero no está habilitado para la inscripción.",
        };

      const fechaNacimientoRaw = fila[COL_FECHA_NACIMIENTO_BUSQUEDA - COL_HABILITADO_BUSQUEDA]; // (Corregido índice)
      const fechaNacimientoStr =
        fechaNacimientoRaw instanceof Date
          ? Utilities.formatDate(
              fechaNacimientoRaw,
              ss.getSpreadsheetTimeZone(),
              "yyyy-MM-dd"
            )
          : "";
      return {
        status: "OK",
        datos: {
          nombre: fila[COL_NOMBRE_BUSQUEDA - COL_HABILITADO_BUSQUEDA], // (Corregido índice)
          apellido: fila[COL_APELLIDO_BUSQUEDA - COL_HABILITADO_BUSQUEDA], // (Corregido índice)
          dni: dniLimpio,
          fechaNacimiento: fechaNacimientoStr,
          obraSocial: String(fila[COL_OBRASOCIAL_BUSQUEDA - COL_HABILITADO_BUSQUEDA] || "").trim(), // (Corregido índice)
          colegioJardin: String(fila[COL_COLEGIO_BUSQUEDA - COL_HABILITADO_BUSQUEDA] || "").trim(), // (Corregido índice)
          adultoResponsable1: String(fila[COL_RESPONSABLE_BUSQUEDA - COL_HABILITADO_BUSQUEDA] || "").trim(), // (Corregido índice)
          esPreventa: false,
        },
        edad: calcularEdadDetallada(fechaNacimientoStr),
        jornadaExtendidaAlcanzada: estado.jornadaExtendidaAlcanzada,
        tipoInscripto: tipoInscripto,
        pagoTotalMPVisible: pagoTotalMPVisible,
      };
    } else {
      if (
        hojaPreventa &&
        hojaPreventa.getLastRow() > 1 &&
        hojaPreventa
          .getRange(2, COL_PREVENTA_DNI, hojaPreventa.getLastRow() - 1, 1)
          .createTextFinder(dniLimpio)
          .matchEntireCell(true)
          .findNext()
      ) {
        return {
          status: "ERROR_TIPO_ANT",
          message:
            'Usted tiene un cupo Pre-Venta. Por favor, elija la opción "Inscripto PRE-VENTA" para validar.',
        };
      }
      if (tipoInscripto === "anterior")
        return {
          status: "ERROR_TIPO_ANT",
          message:
            "No se encuentra en la base de datos de años anteriores. Por favor, seleccione 'Soy Nuevo Inscripto'.",
        };
      return {
        status: "OK_NUEVO",
        message: "✅ DNI validado. Proceda al registro.",
        jornadaExtendidaAlcanzada: estado.jornadaExtendidaAlcanzada,
        tipoInscripto: tipoInscripto,
        datos: { dni: dniLimpio, esPreventa: false },
        pagoTotalMPVisible: pagoTotalMPVisible,
      };
    }
  } catch (e) {
    Logger.log("Error en validarAcceso: " + e.message + " Stack: " + e.stack);
    return {
      status: "ERROR",
      message: "Ocurrió un error al validar el DNI. " + e.message,
    };
  }
}

/**
 * (MODIFICADO v15-CORREGIDO)
 * - Esta función no requiere cambios internos, ya que todas las
 * lecturas de columnas (COL_ESTADO_PAGO, COL_CANTIDAD_CUOTAS, etc.)
 * se actualizan automáticamente desde `Constantes.js`.
 */
function configurarColumnaEnviarEmailComoCheckboxDeshabilitada() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hoja) throw new Error(`Hoja '${NOMBRE_HOJA_REGISTRO}' no encontrada.`);

    const lastRow = Math.max(hoja.getLastRow(), 2);
    const numRows = Math.max(1, lastRow - 1);
    
    // (MODIFICADO v15) Esta constante (COL_ENVIAR_EMAIL_MANUAL) ahora apunta a la columna 45 (AS)
    const rango = hoja.getRange(2, COL_ENVIAR_EMAIL_MANUAL, numRows, 1);

    rango.clearContent();

    const regla = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
    rango.setDataValidation(regla);

    const valores = [];
    for (let i = 0; i < numRows; i++) valores.push([false]);
    rango.setValues(valores);

    try {
      const proteccion = rango.protect().setDescription('Columna AS: checkbox deshabilitada por configuración.');
      const me = Session.getEffectiveUser();
      proteccion.addEditor(me);
      const editors = proteccion.getEditors();
      editors.forEach(function (ed) {
        try {
          if (ed.getEmail && ed.getEmail() !== me.getEmail()) proteccion.removeEditor(ed);
        } catch (e) {
          // noop
        }
      });
      if (proteccion.canDomainEdit && proteccion.canDomainEdit()) proteccion.setDomainEdit(false);
    } catch (e) {
      Logger.log('Advertencia: no se pudo aplicar protección a la columna AS: ' + e.message);
    }

    return { status: 'OK', message: 'Columna AS configurada como casillas y deshabilitada (protección aplicada cuando fue posible).' };
  } catch (e) {
    Logger.log('Error en configurarColumnaEnviarEmailComoCheckboxDeshabilitada: ' + e.message);
    return { status: 'ERROR', message: 'Error al configurar la columna AS: ' + e.message };
  }
}


/**
 * (MODIFICADO v15-CORREGIDO)
 * - Esta función no requiere cambios internos, ya que todas las
 * lecturas de columnas (COL_ESTADO_PAGO, COL_CANTIDAD_CUOTAS, etc.)
 * se actualizan automáticamente desde `Constantes.js`.
 */
function gestionarUsuarioYaRegistrado(
  ss,
  hojaRegistro,
  filaRegistro,
  dniLimpio,
  estado,
  tipoInscripto,
  pagoTotalMPVisible
) {
  let rangoFila = hojaRegistro
    .getRange(filaRegistro, 1, 1, hojaRegistro.getLastColumn())
    .getValues()[0];

  let estadoPago = rangoFila[COL_ESTADO_PAGO - 1]; // Nueva Col AJ (36)
  const metodoPago = rangoFila[COL_METODO_PAGO - 1]; // Sin cambio (AC)
  
  const nombreRegistrado = rangoFila[COL_NOMBRE - 1]; // Sin cambio (F)
  const apellidoRegistrado = rangoFila[COL_APELLIDO - 1]; // Sin cambio (G)
  const nombreCompleto = `${nombreRegistrado} ${apellidoRegistrado}`;

  const estadoInscripto = rangoFila[COL_ESTADO_NUEVO_ANT - 1]; // Sin cambio (D)
  const estadoInscriptoTrim = estadoInscripto
    ? String(estadoInscripto).trim().toLowerCase()
    : "";

  // --- (Validación de Tipo) ---
  if (estadoInscriptoTrim.includes('anterior') && tipoInscripto !== 'anterior') {
    return { status: 'ERROR', message: 'Este DNI ya está registrado como "Inscripto Anterior". Por favor, seleccione esa opción y valide de nuevo.' };
  }
  // ... (resto de validaciones de tipo) ...
  if (estadoInscriptoTrim.includes('nuevo') && tipoInscripto !== 'nuevo' && tipoInscripto !== 'hermano/a') {
    return { status: 'ERROR', message: 'Este DNI ya está registrado como "Nuevo Inscripto". Por favor, seleccione esa opción y valide de nuevo.' };
  }
  if (estadoInscriptoTrim.includes('pre-venta') && tipoInscripto !== 'preventa' && tipoInscripto !== 'hermano/a') {
    return { status: 'ERROR', message: 'Este DNI está registrado como "Pre-Venta". Por favor, seleccione esa opción y valide de nuevo.' };
  }


  // (Preparar datos para la función de Editar)
  const datosParaEdicion = {
    dni: dniLimpio,
    nombre: nombreRegistrado,
    apellido: apellidoRegistrado,
    email: rangoFila[COL_EMAIL - 1] || '',
    adultoResponsable1: rangoFila[COL_ADULTO_RESPONSABLE_1 - 1] || '',
    dniResponsable1: rangoFila[COL_DNI_RESPONSABLE_1 - 1] || '',
    telResponsable1: rangoFila[COL_TEL_RESPONSABLE_1 - 1] || '',
    adultoResponsable2: rangoFila[COL_ADULTO_RESPONSABLE_2 - 1] || '',
    telResp2: rangoFila[COL_TEL_RESPONSABLE_2 - 1] || '',
    personasAutorizadas: rangoFila[COL_PERSONAS_AUTORIZADAS - 1] || '',
    urlCertificadoAptitud: rangoFila[COL_APTITUD_FISICA - 1] || '' // Y
  };


  // =========================================================
  // --- BUG REDIRECCIÓN HERMANOS ---
  // =========================================================
  const campoSaludVacio = !rangoFila[COL_PRACTICA_DEPORTE - 1] || rangoFila[COL_PRACTICA_DEPORTE - 1] === ""; // S

  if (estadoInscriptoTrim.includes('hermano/a') && campoSaludVacio) { 
  // =========================================================
    
    let tipoOriginal = "nuevo"; // Default
    if (estadoInscriptoTrim.includes("anterior")) tipoOriginal = "anterior";
    if (estadoInscriptoTrim.includes("pre-venta")) tipoOriginal = "preventa";

    // (Validación de tipo específica para hermanos incompletos)
    if (tipoInscripto !== tipoOriginal) {
       let tipoCorrecto = "Soy Nuevo Inscripto"; // Default
       if (tipoOriginal === "anterior") {
         tipoCorrecto = "Soy Inscripto Anterior";
       } else if (tipoOriginal === "preventa") {
         tipoCorrecto = "Soy Inscripto PRE-VENTA";
       }
       return {
         status: "ERROR",
         message: `Este DNI ya está registrado en la categoría "${estadoInscripto}". Por favor, seleccione la opción "<strong>${tipoCorrecto}</strong>" y valide de nuevo.`,
       };
    }
    
    // Si el tipo es correcto, proceder a completar datos
     const datosCompletos = {
      ...datosParaEdicion,
      fechaNacimiento: rangoFila[COL_FECHA_NACIMIENTO_REGISTRO - 1] ? Utilities.formatDate(new Date(rangoFila[COL_FECHA_NACIMIENTO_REGISTRO - 1]), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : '',
      obraSocial: rangoFila[COL_OBRA_SOCIAL - 1] || '', // K
      colegioJardin: rangoFila[COL_COLEGIO_JARDIN - 1] || '', // L
      esHermanoCompletando: true, // Flag para el cliente
      esPreventa: estadoInscriptoTrim.includes('pre-venta') // Pasar si es preventa
    };

    return {
      status: 'HERMANO_COMPLETAR',
      message: `⚠️ ¡Hola ${datosCompletos.nombre}! Eres un hermano/a pre-registrado.\n` +
      `Por favor, complete/verifique TODOS los campos del formulario para obtener el cupo definitivo.`,
      datos: datosCompletos,
      jornadaExtendidaAlcanzada: estado.jornadaExtendidaAlcanzada,
      tipoInscripto: tipoInscripto, 
      pagoTotalMPVisible: pagoTotalMPVisible
    };
  }
  // --- (FIN CORRECCIÓN Error 2) ---


  // --- (Inicio de Lógica de Pagos para usuarios COMPLETOS) ---

  const idFamiliar = rangoFila[COL_VINCULO_PRINCIPAL - 1]; // Nueva Col AR (44)
  
  let tieneHermanos = false;
  
  if (idFamiliar) {
    const finder = hojaRegistro
      .getRange(2, COL_VINCULO_PRINCIPAL, hojaRegistro.getLastRow() - 1, 1)
      .createTextFinder(idFamiliar)
      .matchEntireCell(true);
    
    if (finder.findAll().length > 1) {
       tieneHermanos = true;
    }
  }

  let cantidadCuotasRegistrada = parseInt(rangoFila[COL_CANTIDAD_CUOTAS - 1]); // Nueva Col AI (35)
  if (
    metodoPago === "Pago en Cuotas" &&
    (isNaN(cantidadCuotasRegistrada) || cantidadCuotasRegistrada < 1)
  ) {
    cantidadCuotasRegistrada = 3; 
    hojaRegistro.getRange(filaRegistro, COL_CANTIDAD_CUOTAS).setValue(3); 
    rangoFila[COL_CANTIDAD_CUOTAS - 1] = 3; 
  } else if (isNaN(cantidadCuotasRegistrada)) {
    cantidadCuotasRegistrada = 0; 
  }

  let estadoPagoActual = estadoPago; // Col AJ (36)
  const c_total = rangoFila[COL_COMPROBANTE_MANUAL_TOTAL_EXT - 1], // Nueva Col AN (40)
    c_c1 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA1 - 1], // Nueva Col AO (41)
    c_c2 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA2 - 1], // Nueva Col AP (42)
    c_c3 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA3 - 1]; // Nueva Col AQ (43)
  const tieneComprobantes = c_total || c_c1 || c_c2 || c_c3;

  if (
    !tieneComprobantes &&
    (String(estadoPagoActual).includes("En revisión") ||
      String(estadoPagoActual).includes("Pagado") ||
      String(estadoPagoActual).includes("Total") ||
      String(estadoPagoActual).includes("Pago Cuota"))
  ) {
    Logger.log(
      `Corrigiendo estado para DNI ${dniLimpio}: El estado era '${estadoPagoActual}' pero no hay comprobantes. Reseteando.`
    );
    estadoPagoActual =
      metodoPago === "Pago en Cuotas"
        ? `Pendiente (${cantidadCuotasRegistrada} Cuotas)`
        : "Pendiente (Transferencia)";
    hojaRegistro
      .getRange(filaRegistro, COL_ESTADO_PAGO)
      .setValue(estadoPagoActual);
    hojaRegistro.getRange(filaRegistro, COL_CUOTA_1, 1, 3).clearContent(); // Limpia AF, AG, AH
    rangoFila = hojaRegistro
      .getRange(filaRegistro, 1, 1, hojaRegistro.getLastColumn())
      .getValues()[0];
    estadoPago = estadoPagoActual; 
  }

  if (
    String(estadoPagoActual).startsWith("Pendiente") &&
    (String(rangoFila[COL_CUOTA_1 - 1]).startsWith("Pagada") || // Col AF
      String(rangoFila[COL_CUOTA_2 - 1]).startsWith("Pagada") || // Col AG
      String(rangoFila[COL_CUOTA_3 - 1]).startsWith("Pagada")) // Col AH
  ) {
    Logger.log(
      `Corrigiendo datos inconsistentes para DNI ${dniLimpio}: El estado era ${estadoPagoActual} pero las cuotas estaban pagadas. Reseteando cuotas.`
    );
    hojaRegistro.getRange(filaRegistro, COL_CUOTA_1, 1, 3).clearContent(); // Limpia AF, AG, AH
    rangoFila = hojaRegistro
      .getRange(filaRegistro, 1, 1, hojaRegistro.getLastColumn())
      .getValues()[0];
  }

  let proximaCuotaPendiente = null;
  let cuotasPendientes = 0;
  let cuotasPagadas = []; 
  let estadoPagoRecalculado = estadoPago; 

  if (metodoPago === "Pago en Cuotas") {
    const cuotas = [
      rangoFila[COL_CUOTA_1 - 1], // Col AF
      rangoFila[COL_CUOTA_2 - 1], // Col AG
      rangoFila[COL_CUOTA_3 - 1], // Col AH
    ];

    const tieneComp1 = c_c1 && String(c_c1).trim() !== "";
    const tieneComp2 = c_c2 && String(c_c2).trim() !== "";
    const tieneComp3 = c_c3 && String(c_c3).trim() !== "";

    const cuotaText0 = String(cuotas[0] || "");
    const cuotaText1 = String(cuotas[1] || "");
    const cuotaText2 = String(cuotas[2] || "");
    const aiText = String(estadoPago || ""); // Col AJ

    let estadosTexto = [];
    let pagadasCount = 0;

    const cuotaEstaPagada = (index, tieneComp, cuotaText, aiContains) => {
      // (MODIFICADO v14) Si la celda de la cuota (AF, AG, AH) es un número, NO está pagada.
      // Solo contamos "Pagada" si tiene texto o si tiene comprobante.
      if (typeof cuotaText === 'number') return false; 
      if (tieneComp) return true;
      if (String(cuotaText).toLowerCase().startsWith("pagada")) return true;
      if (aiContains) return true;
      return false;
    };

    const construirTextoCuota = (
      index,
      tieneComp,
      cuotaText,
      aiContainsFamiliar
    ) => {
      const estaPagada = cuotaEstaPagada(
        index,
        tieneComp,
        cuotaText,
        aiText.includes(`C${index} Pagada`)
      );
      if (estaPagada) {
        pagadasCount++;
        cuotasPagadas.push(`mp_cuota_${index}`);
        const tieneFamiliar =
          aiText.includes(`C${index} Familiar`) ||
          (cuotaText && cuotaText.toLowerCase().includes("familiar"));
        return `C${index} ${tieneFamiliar ? "Familiar Pagada" : "Pagada"}`;
      } else {
        if (!proximaCuotaPendiente) proximaCuotaPendiente = `C${index}`;
        return `C${index} Pendiente`;
      }
    };

    estadosTexto.push(
      construirTextoCuota(
        1,
        tieneComp1,
        cuotaText0,
        aiText.includes("C1 Familiar")
      )
    );
    estadosTexto.push(
      construirTextoCuota(
        2,
        tieneComp2,
        cuotaText1,
        aiText.includes("C2 Familiar")
      )
    );
    estadosTexto.push(
      construirTextoCuota(
        3,
        tieneComp3,
        cuotaText2,
        aiText.includes("C3 Familiar")
      )
    );

    cuotasPendientes = cantidadCuotasRegistrada - pagadasCount;

    const comprobantesPresentes =
      (tieneComp1 ? 1 : 0) + (tieneComp2 ? 1 : 0) + (tieneComp3 ? 1 : 0);

    if (
      cantidadCuotasRegistrada > 0 &&
      pagadasCount >= cantidadCuotasRegistrada &&
      (comprobantesPresentes >= cantidadCuotasRegistrada || c_total)
    ) {
      if (cantidadCuotasRegistrada === 2) {
        estadoPagoRecalculado = [estadosTexto[0], estadosTexto[1]].join(", ");
      } else if (cantidadCuotasRegistrada === 1) {
        estadoPagoRecalculado = estadosTexto[0];
      } else {
        estadoPagoRecalculado = estadosTexto.join(", ");
      }
    } else if (pagadasCount === 0) {
      estadoPagoRecalculado = `Pendiente (${cantidadCuotasRegistrada} Cuotas)`;
    } else {
      if (cantidadCuotasRegistrada === 2)
        estadosTexto = [estadosTexto[0], estadosTexto[1]];
      if (cantidadCuotasRegistrada === 1) estadosTexto = [estadosTexto[0]];
      estadoPagoRecalculado = estadosTexto.join(", ");
    }

    if (estadoPagoRecalculado !== estadoPago) {
      Logger.log(
        `Corrigiendo estado para DNI ${dniLimpio}: De '${estadoPago}' a '${estadoPagoRecalculado}'`
      );
      hojaRegistro
        .getRange(filaRegistro, COL_ESTADO_PAGO)
        .setValue(estadoPagoRecalculado);
      estadoPago = estadoPagoRecalculado; 
    }
  }

  const cuotasPagadasPorComprobante = [];
  if (c_c1 && String(c_c1).trim() !== "")
    cuotasPagadasPorComprobante.push("mp_cuota_1");
  if (c_c2 && String(c_c2).trim() !== "")
    cuotasPagadasPorComprobante.push("mp_cuota_2");
  if (c_c3 && String(c_c3).trim() !== "")
    cuotasPagadasPorComprobante.push("mp_cuota_3");

  let cuotasPagadasFinal = cuotasPagadasPorComprobante.slice();
  if (cantidadCuotasRegistrada === 2)
    cuotasPagadasFinal = cuotasPagadasFinal.filter((c) => c !== "mp_cuota_3");
  if (cantidadCuotasRegistrada === 1)
    cuotasPagadasFinal = cuotasPagadasFinal.filter((c) => c === "mp_cuota_1");

  const pagadasCountByComp = cuotasPagadasFinal.length;
  const cuotasPendientesByComp = Math.max(
    0,
    cantidadCuotasRegistrada - pagadasCountByComp
  );

  let algunoHermanoCompletos = false; 
  if (tieneHermanos && idFamiliar) {
    try {
      const finder = hojaRegistro
        .getRange(2, COL_VINCULO_PRINCIPAL, hojaRegistro.getLastRow() - 1, 1) // AR
        .createTextFinder(idFamiliar)
        .matchEntireCell(true);
      
      const filasHermanos = finder.findAll();

      for (const celda of filasHermanos) {
        const r = celda.getRow();
        if (r === filaRegistro) continue; 

        const filaVals = hojaRegistro
            .getRange(r, 1, 1, hojaRegistro.getLastColumn())
            .getValues()[0];
        
        const ctot = filaVals[COL_COMPROBANTE_MANUAL_TOTAL_EXT - 1]; // AN
        const cc1 = filaVals[COL_COMPROBANTE_MANUAL_CUOTA1 - 1]; // AO
        const cc2 = filaVals[COL_COMPROBANTE_MANUAL_CUOTA2 - 1]; // AP
        const cc3 = filaVals[COL_COMPROBANTE_MANUAL_CUOTA3 - 1]; // AQ
        let cnt = 0;
        if (cc1 && String(cc1).trim() !== "") cnt++;
        if (cc2 && String(cc2).trim() !== "") cnt++;
        if (cc3 && String(cc3).trim() !== "") cnt++;
        const cantidadCuotasFila = parseInt(
          filaVals[COL_CANTIDAD_CUOTAS - 1] // AI
        );
        const cantidadReal =
          isNaN(cantidadCuotasFila) || cantidadCuotasFila < 1
            ? filaVals[COL_METODO_PAGO - 1] === "Pago en Cuotas"
              ? 3
              : 0
            : cantidadCuotasFila;
        const completos =
          (cantidadReal > 0 && cnt >= cantidadReal) || Boolean(ctot);
        if (completos) {
          algunoHermanoCompletos = true;
          break; 
        }
      }
    } catch (e) {
      Logger.log("Error calculando hermanos comprobantes completos: " + e.toString());
      algunoHermanoCompletos = false;
    }
  }


  const baseResponse = {
    status: "REGISTRO_ENCONTRADO",
    adeudaAptitud: !rangoFila[COL_APTITUD_FISICA - 1], // Y
    metodoPago: metodoPago,
    pagoTotalMPVisible: pagoTotalMPVisible,
    datos: datosParaEdicion,
    tieneHermanos: tieneHermanos,
    algunoHermanoConComprobantesCompletos: algunoHermanoCompletos,
    cantidadCuotasRegistrada: cantidadCuotasRegistrada,
    cuotasPagadas: cuotasPagadasFinal,
    cuotasPendientes: cuotasPendientesByComp,
    comprobantesCompletos:
      (cantidadCuotasRegistrada > 0 &&
        pagadasCountByComp >= cantidadCuotasRegistrada) ||
      Boolean(c_total),
  };

  if (
    String(estadoPago).startsWith("Pago total") ||
    String(estadoPago).startsWith("Pagado Total") ||
    String(estadoPago).startsWith("Pagado") ||
    String(estadoPago).startsWith("Pago Total Familiar")

  ) {
    return {
      ...baseResponse,
      message: `✅ El DNI ${dniLimpio} (${nombreCompleto}) ya se encuentra REGISTRADO y la inscripción está PAGADA (${estadoPago}).`,
      proximaCuotaPendiente: null,
      cuotasPendientes: 0,
    };
  }

  if (
    String(estadoPago).includes("Pendiente") &&
    String(estadoPago).includes("Pagada")
  ) {
    return {
      ...baseResponse,
      message: `⚠️ El DNI ${dniLimpio} (${nombreCompleto}) ya se encuentra REGISTRADO. Estado: ${estadoPago}.`,
      proximaCuotaPendiente: proximaCuotaPendiente,
    };
  }

  if (
    String(estadoPago).includes("En revisión") ||
    String(estadoPago).includes("Pago Cuota")
  ) {
    let mensajeRevision = `⚠️ El DNI ${dniLimpio} (${nombreCompleto}) ya se encuentra REGISTRADO. Estado: ${estadoPago}.`;
    if (cuotasPendientes > 0)
      mensajeRevision += ` Le quedan ${cuotasPendientes} cuota${
        cuotasPendientes > 1 ? "s" : ""
      } pendiente${cuotasPendientes > 1 ? "s" : ""}.`;
    return {
      ...baseResponse,
      message: mensajeRevision,
      proximaCuotaPendiente: proximaCuotaPendiente,
    };
  }

  let mensajePendiente = `⚠️ El DNI ${dniLimpio} (${nombreCompleto}) ya se encuentra REGISTRADO. El pago (${metodoPago}) está PENDIENTE.`;
  if (metodoPago === "Pago en Cuotas") {
    mensajePendiente = `⚠️ El DNI ${dniLimpio} (${nombreCompleto}) ya se encuentra REGISTRADO. El pago (${metodoPago}) está PENDIENTE. Tiene ${cuotasPendientes} cuota${
      cuotasPendientes !== 1 ? "s" : ""
    } pendiente${cuotasPendientes !== 1 ? "s" : ""}.`;
  }

  return {
    ...baseResponse,
    message: mensajePendiente,
    proximaCuotaPendiente: proximaCuotaPendiente,
    estadoPago: estadoPago,
  };
}


/**
 * Sube un archivo a Drive. No requiere cambios.
 */
function subirArchivoIndividual(fileData, dni, tipoArchivo) {
  try {
    if (!fileData || !dni || !tipoArchivo) {
      return {
        status: "ERROR",
        message: "Faltan datos para la subida (DNI, archivo o tipo).",
      };
    }
    const dniLimpio = limpiarDNI(dni);

    // --- (INICIO MODIFICACIÓN NOMBRE DE ARCHIVO) ---
    let nuevoNombre = fileData.fileName; // Por defecto usa el nombre ya construido
    const extension = fileData.fileName.includes(".")
      ? fileData.fileName.split(".").pop()
      : "jpg";

    if (tipoArchivo === "foto") {
      nuevoNombre = `FotoCarnet_${dniLimpio}.${extension}`;
    } else if (tipoArchivo === "ficha") {
      nuevoNombre = `AptitudFisica_${dniLimpio}.${extension}`;
    }
    // Si es 'comprobante', el nombre ya viene pre-formateado desde 'subirComprobanteManual'

    const fileUrl = uploadFileToDrive(
      fileData.data,
      fileData.mimeType,
      nuevoNombre, // Usamos el nombre nuevo
      dniLimpio,
      tipoArchivo
    );
    // --- (FIN MODIFICACIÓN) ---

    if (typeof fileUrl === "object" && fileUrl.status === "ERROR") {
      return fileUrl;
    }

    // Devuelve la URL simple, el =HYPERLINK se genera en la función que lo llama
    return {
      status: "OK",
      url: fileUrl.replace('=HYPERLINK("', "").split('";')[0],
    };
  } catch (e) {
    Logger.log("Error en subirArchivoIndividual: " + e.toString());
    return {
      status: "ERROR",
      message: "Error del servidor al subir: " + e.message,
    };
  }
}

/**
 * Sube aptitud física. No requiere cambios.
 */
function subirAptitudManual(dni, fileData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const dniLimpio = limpiarDNI(dni);
    if (!dniLimpio || !fileData) {
      return { status: "ERROR", message: "Faltan datos (DNI o archivo)." };
    }

    const extension = fileData.fileName.includes(".")
      ? fileData.fileName.split(".").pop()
      : "pdf";
    const nuevoNombreAptitud = `AptitudFisica_${dniLimpio}.${extension}`;

    const fileUrlFormula = uploadFileToDrive(
      fileData.data,
      fileData.mimeType,
      nuevoNombreAptitud,
      dniLimpio,
      "ficha"
    );

    if (
      typeof fileUrlFormula !== "string" ||
      !fileUrlFormula.startsWith("=HYPERLINK")
    ) {
      throw new Error("Error al subir el archivo a Drive.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hoja)
      throw new Error(`La hoja "${NOMBRE_HOJA_REGISTRO}" no fue encontrada.`);

    const rangoDni = hoja.getRange(
      2,
      COL_DNI_INSCRIPTO,
      hoja.getLastRow() - 1,
      1
    );
    const celdaEncontrada = rangoDni
      .createTextFinder(dniLimpio)
      .matchEntireCell(true)
      .findNext();

    if (celdaEncontrada) {
      const fila = celdaEncontrada.getRow();
      hoja.getRange(fila, COL_APTITUD_FISICA).setValue(fileUrlFormula); // Guardar la fórmula

      Logger.log(
        `Aptitud Física subida para DNI ${dniLimpio} en fila ${fila}.`
      );
      return {
        status: "OK",
        message: "¡Certificado de Aptitud subido con éxito!",
      };
    } else {
      Logger.log(`No se encontró DNI ${dniLimpio} para subir aptitud física.`);
      return {
        status: "ERROR",
        message: `No se encontró el registro para el DNI ${dniLimpio}.`,
      };
    }
  } catch (e) {
    Logger.log("Error en subirAptitudManual: " + e.toString());
    return { status: "ERROR", message: "Error en el servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Valida DNI de hermano. No requiere cambios.
 */
function validarDNIHermano(dniHermano, dniPrincipal) {
  try {
    const dniLimpio = limpiarDNI(dniHermano);
    const dniPrincipalLimpio = limpiarDNI(dniPrincipal);

    if (!/^[0-9]{8}$/.test(dniLimpio)) {
      return { status: 'ERROR', message: 'El DNI del hermano/a debe tener 8 dígitos.' };
    }
    if (dniLimpio === dniPrincipalLimpio) {
      return { status: 'ERROR', message: 'El DNI del hermano/a no puede ser igual al del inscripto principal.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Chequear duplicados en Registros
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (hojaRegistro && hojaRegistro.getLastRow() > 1) {
      const rangoDniRegistro = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
      const celdaRegistro = rangoDniRegistro.createTextFinder(dniLimpio).matchEntireCell(true).findNext();
      if (celdaRegistro) {
        return { status: 'ERROR', message: `El DNI ${dniLimpio} ya se encuentra registrado en la base de datos (Fila ${celdaRegistro.getRow()}). No se puede agregar como hermano.` };
      }
    }

    // 2. Chequear en PRE-VENTA
    const hojaPreventa = ss.getSheetByName(NOMBRE_HOJA_PREVENTA);
    if (hojaPreventa && hojaPreventa.getLastRow() > 1) {
      const rangoDniPreventa = hojaPreventa.getRange(2, COL_PREVENTA_DNI, hojaPreventa.getLastRow() - 1, 1);
      const celdaEncontradaPreventa = rangoDniPreventa.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

      if (celdaEncontradaPreventa) {
        const fila = hojaPreventa.getRange(celdaEncontradaPreventa.getRow(), 1, 1, hojaPreventa.getLastColumn()).getValues()[0];
        const fechaNacimientoRaw = fila[COL_PREVENTA_FECHA_NAC - 1];
        const fechaNacimientoStr = (fechaNacimientoRaw instanceof Date) ? Utilities.formatDate(fechaNacimientoRaw, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : '';

        return {
          status: 'OK_PREVENTA',
          message: '¡DNI de Pre-Venta encontrado! Se autocompletarán los datos del hermano/a.',
          datos: {
            dni: dniLimpio,
            nombre: fila[COL_PREVENTA_NOMBRE - 1],
            apellido: fila[COL_PREVENTA_APELLIDO - 1],
            fechaNacimiento: fechaNacimientoStr,
            obraSocial: '',
            colegio: '',
            tipo: 'preventa'
          },
        };
      }
    }

    // 3. Chequear en Base de Datos (Anteriores)
    const hojaBusqueda = ss.getSheetByName(NOMBRE_HOJA_BUSQUEDA);
    if (hojaBusqueda && hojaBusqueda.getLastRow() > 1) {
      const rangoDNI = hojaBusqueda.getRange(2, COL_DNI_BUSQUEDA, hojaBusqueda.getLastRow() - 1, 1);
      const celdaEncontrada_BD = rangoDNI.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

      if (celdaEncontrada_BD) {
        const fila = hojaBusqueda.getRange(celdaEncontrada_BD.getRow(), COL_HABILITADO_BUSQUEDA, 1, 10).getValues()[0];
        const fechaNacimientoRaw = fila[COL_FECHA_NACIMIENTO_BUSQUEDA - COL_HABILITADO_BUSQUEDA];
        const fechaNacimientoStr = (fechaNacimientoRaw instanceof Date) ? Utilities.formatDate(fechaNacimientoRaw, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : '';

        return {
          status: 'OK_ANTERIOR',
          message: '¡DNI de Inscripto Anterior encontrado! Se autocompletarán los datos del hermano/a.',
          datos: {
            dni: dniLimpio,
            nombre: fila[COL_NOMBRE_BUSQUEDA - COL_HABILITADO_BUSQUEDA],
            apellido: fila[COL_APELLIDO_BUSQUEDA - COL_HABILITADO_BUSQUEDA],
            fechaNacimiento: fechaNacimientoStr,
            obraSocial: String(fila[COL_OBRASOCIAL_BUSQUEDA - COL_HABILITADO_BUSQUEDA] || '').trim(),
            colegio: String(fila[COL_COLEGIO_BUSQUEDA - COL_HABILITADO_BUSQUEDA] || '').trim(),
            tipo: 'anterior'
          },
        };
      }
    }

    // 4. No encontrado (Nuevo)
    return {
      status: 'OK_NUEVO',
      message: 'DNI no encontrado en Pre-Venta ni en registros Anteriores. Por favor, complete todos los datos del hermano/a.',
      datos: {
        dni: dniLimpio,
        nombre: '',
        apellido: '',
        fechaNacimiento: '',
        obraSocial: '',
        colegio: '',
        tipo: 'nuevo'
      },
    };
  } catch (e) {
    Logger.log("Error en validarDNIHermano: " + e.message);
    return { status: 'ERROR', message: 'Error del servidor: ' + e.message };
  }
}

/**
 * Sube archivo a Drive. No requiere cambios.
 */
function uploadFileToDrive(data, mimeType, newFilename, dni, tipoArchivo) {
  try {
    if (!dni) return { status: "ERROR", message: "No se recibió DNI." };
    let parentFolderId;
    switch (tipoArchivo) {
      case "foto":
        parentFolderId = FOLDER_ID_FOTOS;
        break;
      case "ficha":
        parentFolderId = FOLDER_ID_FICHAS;
        break;
      case "comprobante":
        parentFolderId = FOLDER_ID_COMPROBANTES;
        break;
      default:
        return { status: "ERROR", message: "Tipo de archivo no reconocido." };
    }
    if (!parentFolderId || parentFolderId.includes("AQUI_VA_EL_ID")) {
      return { status: "ERROR", message: "IDs de carpetas no configurados." };
    }

    const parentFolder = DriveApp.getFolderById(parentFolderId);
    let subFolder;
    const folders = parentFolder.getFoldersByName(dni);
    subFolder = folders.hasNext()
      ? folders.next()
      : parentFolder.createFolder(dni);

    const decodedData = Utilities.base64Decode(data.split(",")[1]);
    const blob = Utilities.newBlob(decodedData, mimeType, newFilename);
    const file = subFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // --- (MODIFICACIÓN) ---
    // Devolver la URL con el nombre de archivo como hipervínculo para la hoja
    return `=HYPERLINK("${file.getUrl()}"; "${newFilename}")`;
    // --- (FIN MODIFICACIÓN) ---
  } catch (e) {
    Logger.log("Error en uploadFileToDrive: " + e.toString());
    return { status: "ERROR", message: "Error al subir archivo: " + e.message };
  }
}