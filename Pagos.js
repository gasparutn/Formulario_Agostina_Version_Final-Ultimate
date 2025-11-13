// =========================================================
// Las constantes (COL_EMAIL, COL_ESTADO_PAGO, etc.) se leen
// automáticamente desde el archivo 'Constantes.gs'.
//
// (MODIFICADO) Lógica de precios "Híbrida" (Definitiva)
// 1. El principal obtiene el precio del último hermano (ej: si son 2, obtiene el precio de "Registro 2").
// 2. Los hermanos se pre-registran sin precio.
// 3. Los hermanos, al completar, obtienen su precio escalonado (ej: "Registro 2", "Registro 3")
//    basado en SUS PROPIAS opciones de jornada/socio.
// =========================================================

/**
* (PASO 1 - MODIFICADO)
* Lógica de precio "Híbrida" (Definitiva)
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
    
    // 1. Contar el total de hijos en esta transacción (Principal + lista de hermanos)
    const totalHijos = 1 + (datos.hermanos ? datos.hermanos.length : 0);
    
    // 2. Determinar el índice de precio para el grupo (Si son 2 hijos, usamos índice 1. Si son 3+, usamos 2)
    const indicePrecioAplicar = Math.min(totalHijos - 1, 2);
    
    // 3. Obtener esa configuración de precio (ej: "Registro 2") usando las opciones de H1
    const infoPrecioPrincipal = obtenerPrecioYConfiguracion(datos, hojaConfig, indicePrecioAplicar);
    
    Logger.log(`Total Hijos: ${totalHijos}. Índice de precio aplicado: ${indicePrecioAplicar}. Precio H1: ${infoPrecioPrincipal.precio}`);

    // 4. Aplicar ese precio al Inscripto Principal (datos)
    datos.precio = infoPrecioPrincipal.precio;
    datos.montoAPagar = infoPrecioPrincipal.montoAPagar;
    datos.cantidadCuotas = infoPrecioPrincipal.cantidadCuotas;
    // =========================================================

    // (Punto 10) Nuevos estados de pago (solo para el principal)
    if (datos.metodoPago === 'Pago Efectivo (Adm del Club)') {
      datos.estadoPago = "Pendiente (Efectivo)";
    } else if (datos.metodoPago === 'Transferencia') {
      datos.estadoPago = "Pendiente (Transferencia)";
    } else if (datos.metodoPago === 'Pago en Cuotas') {
      datos.estadoPago = `Pendiente (${datos.cantidadCuotas} Cuotas)`;
    } else { 
      datos.estadoPago = "Pendiente (Transferencia)";
    }

    // (Punto 12) Si es un hermano completando, llamamos a una función diferente
    if (datos.esHermanoCompletando === true) {
      // Esta función (actualizarDatosHermano) tiene la lógica
      // para calcular el precio escalonado (ej: índice 1 o 2).
      const respuestaUpdate = actualizarDatosHermano(datos); 
      respuestaUpdate.datos = datos; 
      return respuestaUpdate;
    } else {
      
      // 1. Registrar al inscripto principal
      const respuestaRegistro = registrarDatos(datos); 
      
      if (respuestaRegistro.status !== 'OK_REGISTRO') {
        Logger.log("Fallo el registro principal: " + respuestaRegistro.message);
        return respuestaRegistro;
      }

      // 2. Si hay hermanos, registrarlos
      const hermanosRegistrados = [];
      if (datos.hermanos && datos.hermanos.length > 0) {
        const idVinculo = `FAM_${respuestaRegistro.numeroDeTurno}`;
        respuestaRegistro.datos.vinculoPrincipal = idVinculo;
        
        try {
          // Actualizar la fila del principal con el ID de vínculo
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

        // Registrar a cada hermano
        datos.hermanos.forEach((hermano, i) => {
          try {
            // =========================================================
            // --- LÓGICA DE PRECIO (HERMANOS) ---
            // =========================================================
            // Los hermanos se pre-registran SIN precio, jornada o método de pago.
            // Estos se calcularán cuando ellos completen sus datos.
            
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
              // Los datos de pago y precio se dejan VACÍOS
              metodoPago: "",
              jornada: "",
              esSocio: "",
              vinculoPrincipal: idVinculo,
              precio: 0,
              montoAPagar: 0,
              cantidadCuotas: 0,
              estadoPago: "Pendiente (Hermano/a)", // Estado clave
            };
            // =========================================================

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
* (MODIFICADO)
* - Lógica de precio "escalonada" (revertida).
* - Calcula el precio para ESTE hermano basado en su índice (1, 2, etc.).
* - YA NO actualiza los precios de otros miembros de la familia.
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
    
    // 1. Encontrar el índice de ESTE hermano (0, 1, 2...)
    // (Esta función vive en precios.gs)
    const indiceHijo = _obtenerIndiceHijo(hojaRegistro, fila);
    
    // 2. Calcular el precio usando los DATOS DE ESTE HERMANO y su ÍNDICE
    const infoPrecio = obtenerPrecioYConfiguracion(datos, hojaConfig, indiceHijo);
    
    const precio = infoPrecio.precio;
    const montoAPagar = infoPrecio.montoAPagar;
    
    // 3. Sobrescribir cantidadCuotas en el objeto 'datos'
    datos.cantidadCuotas = infoPrecio.cantidadCuotas;
    
    // 4. Actualizar el estado de pago si es cuotas
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

    // (Punto 6, 27) Actualizar la fila del hermano con los datos completos
    // ESTO SOLO AFECTA A LA 'fila' DE ESTE HERMANO
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

    // Fotos
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

    // Actualizar datos de PAGO (Jornada, Socio, Precio) SÓLO EN ESTA FILA
    hojaRegistro.getRange(fila, COL_JORNADA).setValue(datos.jornada);
    hojaRegistro.getRange(fila, COL_SOCIO).setValue(datos.esSocio); 
    hojaRegistro.getRange(fila, COL_METODO_PAGO).setValue(datos.metodoPago);
    hojaRegistro.getRange(fila, COL_PRECIO).setValue(precio); // <- Asigna el precio escalonado
    hojaRegistro.getRange(fila, COL_CANTIDAD_CUOTAS).setValue(parseInt(datos.cantidadCuotas) || 0);
    hojaRegistro.getRange(fila, COL_ESTADO_PAGO).setValue(datos.estadoPago);
    hojaRegistro.getRange(fila, COL_MONTO_A_PAGAR).setValue(montoAPagar);
    
    // --- Cálculo de Grupo y Color ---
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
        hojaRegistro.getRange(fila, COL_GRUPOS).setValue(grupoAsignado);
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