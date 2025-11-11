// =========================================================
// Las constantes (COL_EMAIL, COL_ESTADO_PAGO, etc.) se leen
// automáticamente desde el archivo 'Constantes.gs'.
//
// (MODIFICADO) TODA LA LÓGICA DE MERCADO PAGO HA SIDO ELIMINADA.
// =========================================================

/**
* (PASO 1 - CORREGIDO)
* (Punto 10) Añadida lógica para "Transferencia"
* (Punto 28) Lógica de "Pago en Cuotas" ajustada para "Pago en 3 Cuotas"
* (FIX #3) Añadida la lógica de registro de hermanos que se había perdido.
* (MODIFICADO) Se pasa el 'tipo' de hermano a 'registrarDatos'.
*/
function paso1_registrarRegistro(datos) {
  Logger.log("PASO 1 INICIADO. Datos recibidos: " + JSON.stringify(datos));
  try {
    if (!datos.urlFotoCarnet && !datos.esHermanoCompletando) { // (Punto 6) Los hermanos no suben foto en el registro inicial
      Logger.log("Error: El formulario se envió sin la URL de la Foto Carnet.");
      return { status: 'ERROR', message: 'Falta la Foto Carnet. Por favor, asegúrese de que el archivo se haya subido correctamente.' };
    }

    // (Punto 10) Nuevos estados de pago
    if (datos.metodoPago === 'Pago Efectivo (Adm del Club)') {
      datos.estadoPago = "Pendiente (Efectivo)";
    } else if (datos.metodoPago === 'Transferencia') {
      datos.estadoPago = "Pendiente (Transferencia)"; // NUEVO
    } else if (datos.metodoPago === 'Pago en Cuotas') {
      datos.estadoPago = `Pendiente (${datos.cantidadCuotas} Cuotas)`; // (datos.cantidadCuotas será 3)
    } else { 
      // Fallback por si algún método de MP quedó cacheado (ya no debería pasar)
      datos.estadoPago = "Pendiente (Transferencia)";
    }

    // (Punto 12) Si es un hermano completando, llamamos a una función diferente
    if (datos.esHermanoCompletando === true) {
      // (MODIFICADO) Se pasa 'datos' directamente
      const respuestaUpdate = actualizarDatosHermano(datos);
      // Asignar datos de nombre/apellido a la respuesta para el 'paso2'
      respuestaUpdate.datos = datos; 
      return respuestaUpdate;
    } else {
      
      // --- INICIO DE LA CORRECCIÓN (FIX #3) ---
      // 1. Registrar al inscripto principal
      // (registrarDatos() vive en codigo.gs y ya calcula el grupo)
      const respuestaRegistro = registrarDatos(datos); 
      
      if (respuestaRegistro.status !== 'OK_REGISTRO') {
        Logger.log("Fallo el registro principal: " + respuestaRegistro.message);
        return respuestaRegistro; // Si falló el principal, detenerse
      }

      // 2. Si hay hermanos, registrarlos
      const hermanosRegistrados = [];
      if (datos.hermanos && datos.hermanos.length > 0) {
        const idVinculo = `FAM_${respuestaRegistro.numeroDeTurno}`;
        respuestaRegistro.datos.vinculoPrincipal = idVinculo; // Para el 'paso2'
        
        try {
          // Actualizar la fila del principal con el ID de vínculo
          const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
          const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
          const dniPrincipalLimpio = limpiarDNI(datos.dni); // (Usamos limpiarDNI de Código.js)
          
          if (hojaRegistro.getLastRow() > 1) {
            const rangoDni = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
            const celdaEncontrada = rangoDni.createTextFinder(dniPrincipalLimpio).matchEntireCell(true).findNext();
            
            if (celdaEncontrada) {
              hojaRegistro.getRange(celdaEncontrada.getRow(), COL_VINCULO_PRINCIPAL).setValue(idVinculo);
              Logger.log(`ID Vínculo ${idVinculo} seteado en fila ${celdaEncontrada.getRow()} para DNI ${dniPrincipalLimpio}`);
            } else {
               Logger.log(`No se encontró la fila para el DNI ${dniPrincipalLimpio} para setear el ID de vínculo.`);
            }
          }
        } catch (e) {
          Logger.log("Error al setear el ID de vínculo familiar: " + e.message);
        }

        // Registrar a cada hermano
        datos.hermanos.forEach(hermano => {
          try {
            // (CORRECCIÓN) Determinar el tipo de inscripción REAL del hermano.
            const tipoInscripcionHermano = hermano.tipo || 'nuevo'; // 'nuevo', 'anterior', 'preventa'

            const datosHermano = {
              nombre: hermano.nombre,
              apellido: hermano.apellido,
              dni: hermano.dni,
              fechaNacimiento: hermano.fechaNac,
              obraSocial: hermano.obraSocial,
              colegioJardin: hermano.colegio,
              tipoInscripto: 'hermano/a', // Identificador genérico para la lógica de 'completar datos'
              tipoInscripcionOriginal: tipoInscripcionHermano, // (NUEVO) Guardamos el tipo real
              esPreventa: tipoInscripcionHermano === 'preventa', 
              email: datos.email, 
              adultoResponsable1: datos.adultoResponsable1,
              dniResponsable1: datos.dniResponsable1,
              telAreaResp1: datos.telAreaResp1, 
              telNumResp1: datos.telNumResp1,
              estadoPago: "Pendiente (Hermano/a)",
              metodoPago: "", 
              jornada: "", 
              esSocio: "", 
              vinculoPrincipal: idVinculo 
            };

            // Llamar a registrarDatos para el hermano (esto también calculará el grupo)
            const respHermano = registrarDatos(datosHermano);
            if (respHermano.status === 'OK_REGISTRO') {
              
              // =========================================================
              // --- ¡¡AQUÍ ESTÁ LA CORRECCIÓN (Error 2)!! ---
              // (Se pasa el tipo real para la redirección)
              // =========================================================
              hermanosRegistrados.push({
                nombre: hermano.nombre,
                apellido: hermano.apellido,
                dni: hermano.dni,
                tipo: tipoInscripcionHermano // <-- CORREGIDO (antes era "hermano/a")
              });
              // =========================================================

            } else {
               Logger.log(`Fallo al pre-registrar hermano ${hermano.dni}: ${respHermano.message}`);
            }
          } catch (e) {
             Logger.log(`Error crítico pre-registrando hermano ${hermano.dni}: ${e.message}`);
          }
        });
      }
      
      respuestaRegistro.hermanosRegistrados = hermanosRegistrados;
      Logger.log("PASO 1 FINALIZADO. Respuesta: " + JSON.stringify(respuestaRegistro));
      return respuestaRegistro;
      // --- FIN DE LA CORRECCIÓN (FIX #3) ---
    }

  } catch (e) {
    Logger.log("Error en paso1_registrarRegistro: " + e.message + " Stack: " + e.stack);
    return { status: 'ERROR', message: 'Error general en el servidor (Paso 1): ' + e.message };
  }
}

// =========================================================
// (NUEVA FUNCIÓN HELPER para solucionar error de 'hermano')
// =========================================================
/**
 * Obtiene el precio y el monto a pagar desde la hoja de Config.
 */
function obtenerPrecioDesdeConfig(metodoPago, cantidadCuotasStr, hojaConfig) {
  let precio = 0;
  let montoAPagar = 0;
  try {
    const precioCuota = hojaConfig.getRange("B20").getValue();
    const precioTotal = hojaConfig.getRange("B14").getValue();

    if (metodoPago === 'Pago en Cuotas') {
      const numCuotas = parseInt(cantidadCuotasStr) || 3;
      precio = precioCuota * numCuotas;
      montoAPagar = precio;
    } else if (metodoPago === 'Pago Efectivo (Adm del Club)' || metodoPago === 'Transferencia') {
      precio = precioTotal;
      montoAPagar = precio;
    }

    if (precio === 0 && precioTotal > 0) {
      precio = precioTotal;
    }
    if (montoAPagar === 0 && precio > 0) {
       montoAPagar = precio;
    }

    return { precio, montoAPagar };

  } catch (e) {
    Logger.log("Error en obtenerPrecioDesdeConfig: " + e.message);
    return { precio: 0, montoAPagar: 0 };
  }
}


/**
* (MODIFICADO)
* - (FIX #1) Corregida la escritura de HYPERLINK para Foto Carnet y Aptitud.
* - (NUEVO) Añadida la lógica de cálculo de grupo y color al actualizar.
*/
function actualizarDatosHermano(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    const dniBuscado = limpiarDNI(datos.dni); 

    if (!hojaRegistro) throw new Error("Hoja de Registros no encontrada");

    const rangoDni = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
    const celdaEncontrada = rangoDni.createTextFinder(dniBuscado).matchEntireCell(true).findNext();

    if (!celdaEncontrada) {
      return { status: 'ERROR', message: 'No se encontró el registro del hermano para actualizar.' };
    }

    const fila = celdaEncontrada.getRow();

    // --- CÁLCULO DE PRECIOS ---
    const { precio, montoAPagar } = obtenerPrecioDesdeConfig(datos.metodoPago, datos.cantidadCuotas, hojaConfig);

    const telResp1 = `(${datos.telAreaResp1}) ${datos.telNumResp1}`;
    const telResp2 = (datos.telAreaResp2 && datos.telNumResp2) ? `(${datos.telAreaResp2}) ${datos.telNumResp2}` : '';

    // --- (MODIFICACIÓN) ---
    const esPreventa = (datos.esPreventa === true); 
    let marcaNE = "";
    if (datos.jornada === 'Jornada Normal extendida') {
      marcaNE = esPreventa ? "Extendida (Pre-venta)" : "Extendida";
    } else { 
      marcaNE = esPreventa ? "Normal (Pre-Venta)" : "Normal";
    }
    // --- (FIN MODIFICACIÓN) ---


    // (Punto 6, 27) Actualizar la fila del hermano con los datos completos
    hojaRegistro.getRange(fila, COL_MARCA_N_E_A).setValue(marcaNE);
    hojaRegistro.getRange(fila, COL_EMAIL).setValue(datos.email);
    hojaRegistro.getRange(fila, COL_OBRA_SOCIAL).setValue(datos.obraSocial);
    hojaRegistro.getRange(fila, COL_COLEGIO_JARDIN).setValue(datos.colegioJardin);
    hojaRegistro.getRange(fila, COL_ADULTO_RESPONSABLE_1).setValue(datos.adultoResponsable1);
    hojaRegistro.getRange(fila, COL_DNI_RESPONSABLE_1).setValue(datos.dniResponsable1);
    hojaRegistro.getRange(fila, COL_TEL_RESPONSABLE_1).setValue(telResp1);
    hojaRegistro.getRange(fila, COL_ADULTO_RESPONSABLE_2).setValue(datos.adultoResponsable2);
    hojaRegistro.getRange(fila, COL_TEL_RESPONSABLE_2).setValue(telResp2);
    hojaRegistro.getRange(fila, COL_PERSONAS_AUTORIZADAS).setValue(datos.personasAutorizadas);
    hojaRegistro.getRange(fila, COL_PRACTICA_DEPORTE).setValue(datos.practicaDeporte);
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_DEPORTE).setValue(datos.especifiqueDeporte);
    hojaRegistro.getRange(fila, COL_TIENE_ENFERMEDAD).setValue(datos.tieneEnfermedad);
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_ENFERMEDAD).setValue(datos.especifiqueEnfermedad);
    hojaRegistro.getRange(fila, COL_ES_ALERGICO).setValue(datos.esAlergico);
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_ALERGIA).setValue(datos.especifiqueAlergia);

    // --- INICIO DE LA CORRECCIÓN (FIX #1) ---
    const urlAptitud = datos.urlCertificadoAptitud || '';
    if (urlAptitud) {
      const valAptitud = String(urlAptitud).startsWith('=HYPERLINK') ? urlAptitud : `=HYPERLINK("${urlAptitud}"; "Aptitud_${dniBuscado}")`;
      hojaRegistro.getRange(fila, COL_APTITUD_FISICA).setValue(valAptitud);
    } else {
      hojaRegistro.getRange(fila, COL_APTITUD_FISICA).setValue('');
    }
    
    const urlFoto = datos.urlFotoCarnet || '';
    if (urlFoto) {
      const valFoto = String(urlFoto).startsWith('=HYPERLINK') ? urlFoto : `=HYPERLINK("${urlFoto}"; "Foto_${dniBuscado}")`;
      hojaRegistro.getRange(fila, COL_FOTO_CARNET).setValue(valFoto);
    } else {
      hojaRegistro.getRange(fila, COL_FOTO_CARNET).setValue('');
    }
    // --- FIN DE LA CORRECCIÓN (FIX #1) ---

    hojaRegistro.getRange(fila, COL_JORNADA).setValue(datos.jornada);
    hojaRegistro.getRange(fila, COL_SOCIO).setValue(datos.esSocio); 
    hojaRegistro.getRange(fila, COL_METODO_PAGO).setValue(datos.metodoPago);
    hojaRegistro.getRange(fila, COL_PRECIO).setValue(precio);
    hojaRegistro.getRange(fila, COL_CANTIDAD_CUOTAS).setValue(parseInt(datos.cantidadCuotas) || 0); 
    hojaRegistro.getRange(fila, COL_ESTADO_PAGO).setValue(datos.estadoPago);
    hojaRegistro.getRange(fila, COL_MONTO_A_PAGAR).setValue(montoAPagar);
    
    // --- (INICIO DE MODIFICACIÓN - CÁLCULO DE GRUPO) ---
    // (determinarGrupoPorFecha y aplicarColorGrupo están en Código.js)
    const fechaNacHermano = hojaRegistro.getRange(fila, COL_FECHA_NACIMIENTO_REGISTRO).getValue();
    
    // (Corrección) Asegurarse de que la fecha sea un string 'yyyy-MM-dd' para la función
    let fechaNacHermanoStr = "";
    if (fechaNacHermano instanceof Date) {
        fechaNacHermanoStr = Utilities.formatDate(fechaNacHermano, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    } else if (fechaNacHermano) {
        // Intentar parsear si es un string (aunque getValue() suele devolver Date)
        try {
            fechaNacHermanoStr = Utilities.formatDate(new Date(fechaNacHermano), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        } catch(e) {
            Logger.log("No se pudo parsear la fecha del hermano: " + fechaNacHermano);
        }
    }

    if (fechaNacHermanoStr) {
        const grupoAsignado = determinarGrupoPorFecha(fechaNacHermanoStr);
        hojaRegistro.getRange(fila, COL_GRUPOS).setValue(grupoAsignado);
        aplicarColorGrupo(hojaRegistro, fila, grupoAsignado, hojaConfig);
    }
    // --- (FIN DE MODIFICACIÓN) ---

    SpreadsheetApp.flush();

    datos.nombre = hojaRegistro.getRange(fila, COL_NOMBRE).getValue();
    datos.apellido = hojaRegistro.getRange(fila, COL_APELLIDO).getValue();

    return { status: 'OK_REGISTRO', message: '¡Registro de Hermano Actualizado!', numeroDeTurno: hojaRegistro.getRange(fila, COL_NUMERO_TURNO).getValue() };

  } catch (e) {
    Logger.log("Error en actualizarDatosHermano: " + e.message + " Stack: " + e.stack);
    return { status: 'ERROR', message: 'Error general en el servidor (Actualizar Hermano): ' + e.message };
  } finally {
    lock.releaseLock();
  }
}


/**
* (PASO 2 - MODIFICADO)
*/
function paso2_procesarPostRegistro(datos, numeroDeTurno, hermanosRegistrados = null) {
  try {
    const hermanos = hermanosRegistrados || [];
    const dniRegistrado = datos.dni;
    let message = "";

    if (datos.metodoPago === 'Pago Efectivo (Adm del Club)') {
      message = `¡Registro guardado con éxito!!.<br>Su método de pago es: <strong>${datos.metodoPago}</strong>. acérquese a la Secretaría del Club de Martes a Sábados de 11hs a 18hs.`;
    } else if (datos.metodoPago === 'Transferencia') {
      message = `¡Registro guardado con éxito!!.<br>Su método de pago es: <strong>${datos.metodoPago}</strong>. Realice la transferencia y vuelva a ingresar con su DNI para subir el comprobante.`;
    } else if (datos.metodoPago === 'Pago en Cuotas') {
      message = `¡Registro guardado con éxito!!.<br>Su método de pago es: <strong>Pago en 3 Cuotas</strong>. Realice la transferencia de la primer cuota y vuelva a ingresar con su DNI para subir el comprobante.`;
    } else {
      message = `¡Registro guardado con éxito!!. Contacte a la administración para coordinar el pago.`;
    }

    Logger.log(`(Paso 2) Registro exitoso para DNI ${dniRegistrado}. Método: ${datos.metodoPago}. Email desactivado.`);

    return { 
      status: 'OK_EFECTIVO', 
      message: message, 
      hermanos: hermanos,
      dniRegistrado: dniRegistrado,
      datos: datos 
    };

  } catch (e) {
    Logger.log("Error en paso2_procesarPostRegistro: " + e.message);
    return { 
      status: 'ERROR', 
      message: 'Error general en el servidor (Paso 2): ' + e.message, 
      hermanos: [],
      dniRegistrado: datos.dni,
      datos: datos 
    };
  }
}