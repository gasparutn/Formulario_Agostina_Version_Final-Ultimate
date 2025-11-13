/**
 * Archivo: precios.gs
 * Contiene la lógica centralizada para calcular los precios de inscripción
 * basado en la grilla de la hoja "Config".
 *
 * Lógica: "Precio Escalonado". Cada hijo paga un precio individual
 * basado en su orden (1ro, 2do, 3ro+) y sus propias elecciones
 * de Jornada, Socio y Método de Pago.
 */

/**
 * Función principal para obtener el precio basado en la grilla de Config.
 *
 * @param {object} datos - Objeto con los datos del inscripto (.jornada, .metodoPago, .esSocio).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaConfig - La hoja de "Config" ya abierta.
 * @param {number} indiceHijo - El índice del hijo (0 = Principal, 1 = 1er Hermano, 2 = 2do Hermano, etc.).
 * @returns {{precio: number, montoAPagar: number, cantidadCuotas: number}}
 */
function obtenerPrecioYConfiguracion(datos, hojaConfig, indiceHijo = 0) {
  Logger.log(`Calculando precio para indiceHijo: ${indiceHijo}. Datos: Jornada=${datos.jornada}, Metodo=${datos.metodoPago}, Socio=${datos.esSocio}`);

  const jornada = datos.jornada;
  const metodoPago = datos.metodoPago;
  const esSocio = (datos.esSocio === "SÍ");

  let precio = 0;
  let montoAPagar = 0;
  let cantidadCuotas = 0;
  let celdaPrecio = ""; // Para debug

  try {
    // --- 1. Lógica de Mapeo de Columnas (E, F, G, H) ---
    let colLetra = "";
    
    if (esSocio) {
      if (metodoPago === "Pago Efectivo (Adm del Club)") {
        colLetra = "E"; // Columna E: Socio Pago Efectivo
      } else {
        colLetra = "F"; // Columna F: Socio Pago Transferencia (y Cuotas)
      }
    } else { // No Socio
      if (metodoPago === "Pago Efectivo (Adm del Club)") {
        colLetra = "G"; // Columna G: No Socio Pago Efectivo
      } else {
        colLetra = "H"; // Columna H: No Socio Pago Transferencia (y Cuotas)
      }
    }

    // --- 2. Lógica de Mapeo de Filas (17, 27, 37, 47, etc.) ---
    let baseRow = 0;

    if (jornada === "Jornada Normal") {
      if (metodoPago === "Pago en Cuotas") {
        baseRow = 37;
        cantidadCuotas = 3;
      } else {
        baseRow = 17;
        cantidadCuotas = 1; 
      }
    } else { // Asumimos "Jornada Normal extendida"
      if (metodoPago === "Pago en Cuotas") {
        baseRow = 47;
        cantidadCuotas = 3;
      } else {
        baseRow = 27;
        cantidadCuotas = 1;
      }
    }

    // Determina la fila exacta según el índice del hijo
    // (0=Registro 1, 1=Registro 2, 2+=Registro 3)
    let rowNum = 0;
    // (Asegurarse de que el índice no sea mayor a 2)
    const indiceLimitado = Math.min(indiceHijo, 2);

    if (indiceLimitado === 0) {
      rowNum = baseRow; // e.g., 17, 27, 37, 47
    } else if (indiceLimitado === 1) {
      rowNum = baseRow + 1; // e.g., 18, 28, 38, 48
    } else { // indiceLimitado === 2
      rowNum = baseRow + 2; // e.g., 19, 29, 39, 49 (para 3ro y subsiguientes)
    }

    // --- 3. Obtener el Precio de la Celda ---
    if (colLetra && rowNum > 0) {
      celdaPrecio = colLetra + rowNum;
      Logger.log("Obteniendo precio de la celda de Config: " + celdaPrecio);
      
      const valorCelda = hojaConfig.getRange(celdaPrecio).getValue();
      
      if (typeof valorCelda === 'number') {
        precio = valorCelda;
      } else if (typeof valorCelda === 'string') {
        const precioLimpio = valorCelda.replace(/[$.]/g, "").split(" ")[0].replace(/,/g, ".");
        precio = parseFloat(precioLimpio) || 0;
      }
    } else {
      Logger.log(`No se pudo determinar la celda. colLetra=${colLetra}, rowNum=${rowNum}`);
    }

  } catch (e) {
    Logger.log(`Error en obtenerPrecioYConfiguracion: ${e.message}. Celda: ${celdaPrecio}. Stack: ${e.stack}`);
    return { precio: 0, montoAPagar: 0, cantidadCuotas: 0 };
  }
  
  // --- 4. Determinar Monto Total ---
  if (cantidadCuotas === 3) {
    montoAPagar = precio * 3;
    precio = precio * 3;
  } else {
    montoAPagar = precio;
  }
  
  Logger.log(`Precio final: ${precio}, Monto a Pagar: ${montoAPagar}, Cuotas: ${cantidadCuotas}`);
  return { precio: precio, montoAPagar: montoAPagar, cantidadCuotas: cantidadCuotas };
}

/**
 * Función helper para encontrar el índice de un hermano (0, 1, 2+)
 * cuando está actualizando sus datos.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaRegistro - La hoja de "Registros".
 * @param {number} fila - La fila del DNI que se está actualizando.
 * @returns {number} - El índice del hijo (0, 1, 2, etc.).
 */
function _obtenerIndiceHijo(hojaRegistro, fila) {
  try {
    const vinculoPrincipal = hojaRegistro.getRange(fila, COL_VINCULO_PRINCIPAL).getValue();

    // Si no tiene vínculo (o es el principal sin hermanos), es el índice 0.
    if (!vinculoPrincipal) {
      return 0;
    }

    // Si tiene vínculo, buscar a toda la familia
    const rangoVinculos = hojaRegistro.getRange(2, COL_VINCULO_PRINCIPAL, hojaRegistro.getLastRow() - 1, 1);
    const todasLasCeldas = rangoVinculos.createTextFinder(vinculoPrincipal).matchEntireCell(true).findAll();
    
    let familia = [];
    todasLasCeldas.forEach(celda => {
      const filaNum = celda.getRow();
      const turno = hojaRegistro.getRange(filaNum, COL_NUMERO_TURNO).getValue();
      familia.push({ fila: filaNum, turno: turno });
    });

    // Ordenar la familia por N° de Turno (ascendente)
    familia.sort((a, b) => a.turno - b.turno);

    // Encontrar la posición (índice) de nuestra fila actual en la familia ordenada
    const indice = familia.findIndex(miembro => miembro.fila === fila);

    if (indice === -1) {
      Logger.log(`Error en _obtenerIndiceHijo: No se encontró la fila ${fila} en la familia ${vinculoPrincipal}`);
      return 0; // Fallback seguro
    }

    Logger.log(`_obtenerIndiceHijo: Fila ${fila} es el índice ${indice} en la familia ${vinculoPrincipal}`);
    
    // Retornar el índice, limitado a 2 (para 3er hermano o más)
    return Math.min(indice, 2); 

  } catch (e) {
    Logger.log(`Error en _obtenerIndiceHijo: ${e.message}`);
    return 0; // Fallback seguro
  }
}