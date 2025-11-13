// =========================================================
// (MODIFICADO v15-CORREGIDO)
// Este archivo no requiere cambios internos, ya que todas las
// constantes (COL_DNI_INSCRIPTO, COL_MODO_PAGO_CUOTA, etc.)
// se leen desde el nuevo `Constantes.js` y apuntan a las
// columnas correctas (J, K, L, AD, AE, etc.).
// =========================================================

/**
* (PASO 1)
*/
function paso1_registrarRegistro(datos) {
  Logger.log("PASO 1 INICIADO. Datos recibidos: " + JSON.stringify(datos));
  try {
    if (!datos.urlFotoCarnet && !datos.esHermanoCompletando) { 
      Logger.log("Error: El formulario se envió sin la URL de la Foto Carnet.");
      return { status: 'ERROR', message: 'Falta la Foto Carnet. Por favor, asegúrese de que el archivo se haya subido correctamente.' };
    }

    // =========================================================
    // --- LÓGICA DE PRECIO HÍBRIDA (PRINCIPAL) ---
    // =========================================================
    const hojaConfig = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NOMBRE_HOJA_CONFIG);
    
    const totalHijos = 1 + (datos.hermanos ? datos.hermanos.length : 0);
    const indicePrecioAplicar = Math.min(totalHijos - 1, 2);
    
    const infoPrecioPrincipal = obtenerPrecioYConfiguracion(datos, hojaConfig, indicePrecioAplicar);
    
    Logger.log(`Total Hijos: ${totalHijos}. Índice de precio aplicado: ${indicePrecioAplicar}. Precio H1: ${infoPrecioPrincipal.precio}`);

    // 4. Aplicar ese precio al Inscripto Principal (datos)
    datos.precio = infoPrecioPrincipal.precio; // Col AE (Precio Total)
    datos.montoAPagar = infoPrecioPrincipal.montoAPagar; // Col AK ('' si es cuotas, $Total si es único)
    datos.cantidadCuotas = infoPrecioPrincipal.cantidadCuotas; // Col AI
    datos.valorCuota = infoPrecioPrincipal.valorCuota; // (NUEVO) Valor para AF, AG, AH
    // (datos.subMetodoCuotas ya viene en el objeto 'datos')
    // =========================================================

    if (datos.metodoPago === 'Pago Efectivo (Adm del Club)') {
      datos.estadoPago = "Pendiente (Efectivo)";
    } else if (datos.metodoPago === 'Transferencia') {
      datos.estadoPago = "Pendiente (Transferencia)";
    } else if (datos.metodoPago === 'Pago en Cuotas') {
      datos.estadoPago = `Pendiente (${datos.cantidadCuotas} Cuotas)`;
    } else { 
      datos.estadoPago = "Pendiente (Transferencia)";
    }

    if (datos.esHermanoCompletando === true) {
      const respuestaUpdate = actualizarDatosHermano(datos); 
      respuestaUpdate.datos = datos; 
      return respuestaUpdate;
    } else {
      
      const respuestaRegistro = registrarDatos(datos); 
      
      if (respuestaRegistro.status !== 'OK_REGISTRO') {
        Logger.log("Fallo el registro principal: " + respuestaRegistro.message);
        return respuestaRegistro;
      }

      const hermanosRegistrados = [];
      if (datos.hermanos && datos.hermanos.length > 0) {
        const idVinculo = `FAM_${respuestaRegistro.numeroDeTurno}`;
        respuestaRegistro.datos.vinculoPrincipal = idVinculo;
        
        try {
          const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
          const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
          const dniPrincipalLimpio = limpiarDNI(datos.dni); 
          
          if (hojaRegistro.getLastRow() > 1) {
            const rangoDni = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
            const celdaEncontrada = rangoDni.createTextFinder(dniPrincipalLimpio).matchEntireCell(true).findNext();
            
            if (celdaEncontrada) {
              hojaRegistro.getRange(celdaEncontrada.getRow(), COL_VINCULO_PRINCIPAL).setValue(idVinculo);
            }
          }
        } catch (e) {
          Logger.log("Error al setear el ID de vínculo familiar: " + e.message);
        }

        datos.hermanos.forEach((hermano, i) => {
          try {
            const tipoInscripcionHermano = hermano.tipo || 'nuevo';

            const datosHermano = {
              nombre: hermano.nombre,
              apellido: hermano.apellido,
              dni: hermano.dni,
              fechaNacimiento: hermano.fechaNac,
              obraSocial: hermano.obraSocial,
              colegioJardin: hermano.colegio,
              tipoInscripto: 'hermano/a',
              tipoInscripcionOriginal: tipoInscripcionHermano,
              esPreventa: tipoInscripcionHermano === 'preventa', 
              email: datos.email, 
              adultoResponsable1: datos.adultoResponsable1,
              dniResponsable1: datos.dniResponsable1,
              telAreaResp1: datos.telAreaResp1, 
              telNumResp1: datos.telNumResp1,
              metodoPago: "",
              subMetodoCuotas: "", // (NUEVO) Dejar vacío
              jornada: "",
              esSocio: "",
              vinculoPrincipal: idVinculo,
              precio: 0,
              montoAPagar: 0,
              cantidadCuotas: 0,
              valorCuota: 0, // (NUEVO)
              estadoPago: "Pendiente (Hermano/a)", // Estado clave
            };

            const respHermano = registrarDatos(datosHermano);
            if (respHermano.status === 'OK_REGISTRO') {
              hermanosRegistrados.push({
                nombre: hermano.nombre,
                apellido: hermano.apellido,
                dni: hermano.dni,
                tipo: tipoInscripcionHermano 
              });
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
    }

  } catch (e) {
    Logger.log("Error en paso1_registrarRegistro: " + e.message + " Stack: " + e.stack);
    return { status: 'ERROR', message: 'Error general en el servidor (Paso 1): ' + e.message };
  }
}


/**
* (MODIFICADO v15-CORREGIDO)
* - No requiere cambios internos, ya que todas las
* constantes (COL_DNI_INSCRIPTO, COL_MODO_PAGO_CUOTA, etc.)
* se leen desde el nuevo `Constantes.js`.
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

    const fila = celdaEncontrada.getRow(); // <- Esta es la fila de ESTE hermano

    // =========================================================
    // --- LÓGICA DE PRECIO ESCALONADO (HERMANO COMPLETANDO) ---
    // =========================================================
    
    const indiceHijo = _obtenerIndiceHijo(hojaRegistro, fila);
    const infoPrecio = obtenerPrecioYConfiguracion(datos, hojaConfig, indiceHijo);
    
    const precio = infoPrecio.precio; // Col AE (Precio Total)
    const montoAPagar = infoPrecio.montoAPagar; // Col AK ('' si es cuotas, $Total si es único)
    const valorCuota = infoPrecio.valorCuota; // Valor para AF, AG, AH
    
    datos.cantidadCuotas = infoPrecio.cantidadCuotas;
    datos.valorCuota = valorCuota;
    
    if (datos.metodoPago === 'Pago en Cuotas') {
      datos.estadoPago = `Pendiente (${datos.cantidadCuotas} Cuotas)`;
    } else if (datos.metodoPago === 'Pago Efectivo (Adm del Club)') {
      datos.estadoPago = "Pendiente (Efectivo)";
    } else if (datos.metodoPago === 'Transferencia') {
      datos.estadoPago = "Pendiente (Transferencia)";
    }
    // =========================================================


    const telResp1 = `(${datos.telAreaResp1}) ${datos.telNumResp1}`;
    const telResp2 = (datos.telAreaResp2 && datos.telNumResp2) ? `(${datos.telAreaResp2}) ${datos.telNumResp2}` : '';

    const esPreventa = (datos.esPreventa === true); 
    let marcaNE = "";
    if (datos.jornada === 'Jornada Normal extendida') {
      marcaNE = esPreventa ? "Extendida (Pre-venta)" : "Extendida";
    } else { 
      marcaNE = esPreventa ? "Normal (Pre-Venta)" : "Normal";
    }

    // (MODIFICADO v15-CORREGIDO) Todas las constantes apuntan a las nuevas columnas
    hojaRegistro.getRange(fila, COL_MARCA_N_E_A).setValue(marcaNE); // C
    hojaRegistro.getRange(fila, COL_EMAIL).setValue(datos.email); // E
    hojaRegistro.getRange(fila, COL_OBRA_SOCIAL).setValue(datos.obraSocial); // K
    hojaRegistro.getRange(fila, COL_COLEGIO_JARDIN).setValue(datos.colegioJardin); // L
    hojaRegistro.getRange(fila, COL_ADULTO_RESPONSABLE_1).setValue(datos.adultoResponsable1); // M
    hojaRegistro.getRange(fila, COL_DNI_RESPONSABLE_1).setValue(datos.dniResponsable1); // N
    hojaRegistro.getRange(fila, COL_TEL_RESPONSABLE_1).setValue(telResp1); // O
    hojaRegistro.getRange(fila, COL_ADULTO_RESPONSABLE_2).setValue(datos.adultoResponsable2); // P
    hojaRegistro.getRange(fila, COL_TEL_RESPONSABLE_2).setValue(telResp2); // Q
    hojaRegistro.getRange(fila, COL_PERSONAS_AUTORIZADAS).setValue(datos.personasAutorizadas); // R
    hojaRegistro.getRange(fila, COL_PRACTICA_DEPORTE).setValue(datos.practicaDeporte); // S
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_DEPORTE).setValue(datos.especifiqueDeporte); // T
    hojaRegistro.getRange(fila, COL_TIENE_ENFERMEDAD).setValue(datos.tieneEnfermedad); // U
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_ENFERMEDAD).setValue(datos.especifiqueEnfermedad); // V
    hojaRegistro.getRange(fila, COL_ES_ALERGICO).setValue(datos.esAlergico); // W
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_ALERGIA).setValue(datos.especifiqueAlergia); // X

    // Fotos (Y, Z)
    const urlAptitud = datos.urlCertificadoAptitud || '';
    if (urlAptitud) {
      const valAptitud = String(urlAptitud).startsWith('=HYPERLINK') ? urlAptitud : `=HYPERLINK("${urlAptitud}"; "Aptitud_${dniBuscado}")`;
      hojaRegistro.getRange(fila, COL_APTITUD_FISICA).setValue(valAptitud); // Y
    } else {
      hojaRegistro.getRange(fila, COL_APTITUD_FISICA).setValue('');
    }
    const urlFoto = datos.urlFotoCarnet || '';
    if (urlFoto) {
      const valFoto = String(urlFoto).startsWith('=HYPERLINK') ? urlFoto : `=HYPERLINK("${urlFoto}"; "Foto_${dniBuscado}")`;
      hojaRegistro.getRange(fila, COL_FOTO_CARNET).setValue(valFoto); // Z
    } else {
      hojaRegistro.getRange(fila, COL_FOTO_CARNET).setValue('');
    }

    // Actualizar datos de PAGO (AA en adelante)
    hojaRegistro.getRange(fila, COL_JORNADA).setValue(datos.jornada); // AA
    hojaRegistro.getRange(fila, COL_SOCIO).setValue(datos.esSocio); // AB
    hojaRegistro.getRange(fila, COL_METODO_PAGO).setValue(datos.metodoPago); // AC
    hojaRegistro.getRange(fila, COL_MODO_PAGO_CUOTA).setValue(datos.subMetodoCuotas || ''); // AD
    hojaRegistro.getRange(fila, COL_PRECIO).setValue(precio); // AE (Precio Total)
    hojaRegistro.getRange(fila, COL_CANTIDAD_CUOTAS).setValue(parseInt(datos.cantidadCuotas) || 0); // AI
    hojaRegistro.getRange(fila, COL_ESTADO_PAGO).setValue(datos.estadoPago); // AJ
    hojaRegistro.getRange(fila, COL_MONTO_A_PAGAR).setValue(montoAPagar); // AK (Vacío o Total)
    
    // Escribir el valor de cuota individual en AF, AG, AH
    if (datos.cantidadCuotas === 3 && valorCuota > 0) {
      hojaRegistro.getRange(fila, COL_CUOTA_1).setValue(valorCuota); // AF
      hojaRegistro.getRange(fila, COL_CUOTA_2).setValue(valorCuota); // AG
      hojaRegistro.getRange(fila, COL_CUOTA_3).setValue(valorCuota); // AH
    }
    
    // --- Cálculo de Grupo y Color (H, I) ---
    const fechaNacHermano = hojaRegistro.getRange(fila, COL_FECHA_NACIMIENTO_REGISTRO).getValue();
    
    let fechaNacHermanoStr = "";
    if (fechaNacHermano instanceof Date) {
        fechaNacHermanoStr = Utilities.formatDate(fechaNacHermano, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    } else if (fechaNacHermano) {
        try {
            fechaNacHermanoStr = Utilities.formatDate(new Date(fechaNacHermano), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        } catch(e) {
            Logger.log("No se pudo parsear la fecha del hermano: " + fechaNacHermano);
        }
    }

    if (fechaNacHermanoStr) {
        const grupoAsignado = determinarGrupoPorFecha(fechaNacHermanoStr);
        hojaRegistro.getRange(fila, COL_GRUPOS).setValue(grupoAsignado); // I
        aplicarColorGrupo(hojaRegistro, fila, grupoAsignado, hojaConfig);
    }

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
* (PASO 2 - Sin cambios)
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
      const subMetodo = datos.subMetodoCuotas === "Efectivo" ? "Efectivo (Adm del Club)" : "Transferencia";
      message = `¡Registro guardado con éxito!!.<br>Su método de pago es: <strong>Pago en 3 Cuotas (${subMetodo})</strong>. Realice el pago de la primer cuota y vuelva a ingresar con su DNI para subir el comprobante.`;
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