/**
 * (MODIFICADO v15-CORREGIDO)
 * - Lógica de 'aplicarCambiosHermano' modificada.
 * - Si el pago principal es "Total" (externo) y el hermano paga "en Cuotas",
 * el comprobante se aplica a la PRIMERA CUOTA DISPONIBLE del hermano (AO, AP, o AQ).
 * - Se actualiza el estado (AF, AG, AH) y el estado principal (AJ) del hermano correctamente.
 * - Se corrigen los comentarios de las constantes (ej. VINCULO ahora es AR).
 */
function subirComprobanteManual(
  dni,
  fileData,
  cuotasSeleccionadas,
  datosExtras,
  esPagoFamiliar = false
) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  // --- (INICIO REFACTORIZACIÓN) ---
  // Mover variables clave al scope de la función principal
  const cuotasPagadasAhora = new Set(cuotasSeleccionadas);
  const pagandoC1 = cuotasPagadasAhora.has("mp_cuota_1");
  const pagandoC2 = cuotasPagadasAhora.has("mp_cuota_2");
  const pagandoC3 = cuotasPagadasAhora.has("mp_cuota_3");

  const nombrePagador = datosExtras.nombrePagador;
  const dniPagador = datosExtras.dniPagador;
  const mensajeFinalCompleto = `¡Inscripción completa!!!<br>Estimada familia, puede validar nuevamente con el dni y acceder a modificar datos de inscrpición en caso de que lo requiera.`;
  let familiaPagos = {}; // (CORRECCIÓN) Inicializar aquí para que siempre exista.
  // --- (FIN REFACTORIZACIÓN) ---

  try {
    const dniLimpio = limpiarDNI(dni);
    if (
      !dniLimpio ||
      !fileData ||
      !cuotasSeleccionadas ||
      cuotasSeleccionadas.length === 0
    ) {
      return {
        status: "ERROR",
        message: "Faltan datos (DNI, archivo o tipo de comprobante).",
      };
    }
    if (!datosExtras || !datosExtras.nombrePagador || !datosExtras.dniPagador) {
      return {
        status: "ERROR",
        message: "Faltan los datos del adulto pagador (Nombre o DNI).",
      };
    }
    if (!/^[0-9]{8}$/.test(datosExtras.dniPagador)) {
      return {
        status: "ERROR",
        message: "El DNI del pagador debe tener 8 dígitos.",
      };
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
      let rangoFilaPrincipal = hoja
        .getRange(fila, 1, 1, hoja.getLastColumn())
        .getValues()[0];

      const dniHoja = rangoFilaPrincipal[COL_DNI_INSCRIPTO - 1];
      const nombreHoja = rangoFilaPrincipal[COL_NOMBRE - 1];
      const apellidoHoja = rangoFilaPrincipal[COL_APELLIDO - 1];
      const metodoPagoHoja = rangoFilaPrincipal[COL_METODO_PAGO - 1] || "Pago";

      /**
       * (NUEVO) Función para actualizar el monto acumulado en la columna AK.
       * @param {number} filaAfectada - La fila a actualizar.
       * @param {Array<any>} filaDatos - Los datos de la fila leída.
       */
      const _actualizarMontoAcumulado = (filaAfectada, filaDatos) => {
        const metodoPago = filaDatos[COL_METODO_PAGO - 1];
        let montoAcumulado = 0;

        if (metodoPago === 'Pago en Cuotas') {
          // (CORRECCIÓN) Leer el valor existente en AK para sumar sobre él.
          const valorActualAK = hoja.getRange(filaAfectada, COL_MONTO_A_PAGAR).getValue();
          if (typeof valorActualAK === 'number') {
            montoAcumulado = valorActualAK;
          }

          // Identificar qué cuotas se están pagando AHORA y sumar su valor al acumulado.
          const valorC1 = filaDatos[COL_CUOTA_1 - 1];
          const valorC2 = filaDatos[COL_CUOTA_2 - 1];
          const valorC3 = filaDatos[COL_CUOTA_3 - 1];

          // (CORRECCIÓN CLAVE) Sumar solo el valor de la cuota que se está pagando en esta transacción.
          // Se asume que esta función se llama justo después de poner el comprobante.
          if (pagandoC1 && typeof valorC1 === 'number' && !hoja.getRange(filaAfectada, COL_CUOTA_1).getValue().toString().toLowerCase().includes('pagada')) {
            montoAcumulado += valorC1;
          }
          if (pagandoC2 && typeof valorC2 === 'number' && !hoja.getRange(filaAfectada, COL_CUOTA_2).getValue().toString().toLowerCase().includes('pagada')) {
            montoAcumulado += valorC2;
          }
          if (pagandoC3 && typeof valorC3 === 'number' && !hoja.getRange(filaAfectada, COL_CUOTA_3).getValue().toString().toLowerCase().includes('pagada')) {
            montoAcumulado += valorC3;
          }

        } else { // Pago único (Transferencia, Efectivo)
          const compTotal = filaDatos[COL_COMPROBANTE_MANUAL_TOTAL_EXT - 1];
          if (compTotal) {
            const precioTotal = filaDatos[COL_PRECIO - 1];
            if (typeof precioTotal === 'number') {
              montoAcumulado = precioTotal;
            }
          }
        }
        hoja.getRange(filaAfectada, COL_MONTO_A_PAGAR).setValue(montoAcumulado);
      };

      /**
       * (Función Helper REFACTORIZADA para aplicar cambios v9)
       * @param {number} filaAfectada - El número de fila a modificar.
       * @param {string} metodoPago - El método de pago (ej: "Pago en Cuotas").
       * @param {string} fileUrl - El link al comprobante.
       * @returns {{esTotal: boolean, nuevoEstado: string, cuotasPagadasNombres: string[], pagadasCount: number, cantidadCuotasRegistrada: number}}
       */
      const aplicarCambios = (filaAfectada, metodoPago, fileUrl) => {
        // 1. OBTENER DATOS PROPIOS DE LA FILA (Corrección Bug 2)
        let rangoFila = hoja
          .getRange(filaAfectada, 1, 1, hoja.getLastColumn())
          .getValues()[0];
        const [c1, c2, c3] = [
          rangoFila[COL_CUOTA_1 - 1], // AF
          rangoFila[COL_CUOTA_2 - 1], // AG
          rangoFila[COL_CUOTA_3 - 1], // AH
        ];
        const estadoAIActual = String(rangoFila[COL_ESTADO_PAGO - 1] || ""); // AJ

        // (Corrección Bug 0>=0)
        let cantidadCuotasRegistrada = parseInt(
          rangoFila[COL_CANTIDAD_CUOTAS - 1] // AI
        );
        if (
          metodoPago === "Pago en Cuotas"
        ) {
            const comp1 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA1 - 1];
            const comp2 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA2 - 1];
            const comp3 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA3 - 1];

            if (pagandoC1 && comp1 && String(comp1).trim() !== "") {
                throw new Error("La Cuota 1 ya tiene un comprobante registrado. No se puede volver a pagar.");
            }
            if (pagandoC2 && comp2 && String(comp2).trim() !== "") {
                throw new Error("La Cuota 2 ya tiene un comprobante registrado. No se puede volver a pagar.");
            }
            if (pagandoC3 && comp3 && String(comp3).trim() !== "") {
                throw new Error("La Cuota 3 ya tiene un comprobante registrado. No se puede volver a pagar.");
            }
        }
        if (
          metodoPago === "Pago en Cuotas" &&
          (isNaN(cantidadCuotasRegistrada) || cantidadCuotasRegistrada < 1)
        ) {
          cantidadCuotasRegistrada = 3;
          hoja.getRange(filaAfectada, COL_CANTIDAD_CUOTAS).setValue(3); // AI
        } else if (isNaN(cantidadCuotasRegistrada)) {
          cantidadCuotasRegistrada = 0;
        }

        // 2. CALCULAR ESTADO FUTURO (Corrección Bug 2)
        Logger.log(
          `aplicarCambios INICIO fila:${filaAfectada} c1:'${c1}' c2:'${c2}' c3:'${c3}' cuotasAhora:${Array.from(
            cuotasPagadasAhora
          ).join(
            "|"
          )} pagandoC1:${pagandoC1} pagandoC2:${pagandoC2} pagandoC3:${pagandoC3}`
        );
        // Determinar si la cuota estaba pagada previamente (en la fila) y si se está pagando AHORA
        // Considerar comprobantes asociados: si existe comprobante en columna correspondiente, contar como pagada
        const comp1 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA1 - 1]; // AO
        const comp2 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA2 - 1]; // AP
        const comp3 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA3 - 1]; // AQ
        // Solo considerar una cuota como pagada previamente si existe un comprobante asociado.
        const prevPagada1 = comp1 && String(comp1).toString().trim() !== "";
        const prevPagada2 = comp2 && String(comp2).toString().trim() !== "";
        const prevPagada3 = comp3 && String(comp3).toString().trim() !== "";
        const pagandoAhora1 = pagandoC1;
        const pagandoAhora2 = pagandoC2;
        const pagandoAhora3 = pagandoC3;

        const estadoC1 = prevPagada1 || pagandoAhora1;
        const estadoC2 = prevPagada2 || pagandoAhora2;
        const estadoC3 = prevPagada3 || pagandoAhora3;

        let pagadasCount = 0;
        if (estadoC1) pagadasCount++;
        if (estadoC2) pagadasCount++;
        if (estadoC3) pagadasCount++;

        // (Corrección Bug 0>=0) 'esTotal' solo puede ser true si hay cuotas que pagar
        let esTotal =
          cantidadCuotasRegistrada > 0 &&
          pagadasCount >= cantidadCuotasRegistrada;
        if (
          metodoPago !== "Pago en Cuotas" &&
          (cuotasPagadasAhora.has("mp_total") || cuotasPagadasAhora.has("externo"))
        ) {
          esTotal = true; // Pago total para Transferencia/Efectivo
        }

        // 4. DETERMINAR ESTADO DE PAGO (Columna AJ) (Corrección Bug 2)
        let nuevoEstadoPago = "";
        if (esTotal) {
          if (metodoPago === "Pago en Cuotas") {
            // Mostrar detalle por cuota incluso si se completaron todas, para mantener trazabilidad
            let estadosTotales = [];
            // Cuota1
            if (estadoC1) {
              if (pagandoAhora1 && esPagoFamiliar)
                estadosTotales.push("C1 Familiar Pagada");
              else if (prevPagada1 && estadoAIActual.includes("C1 Familiar"))
                estadosTotales.push("C1 Familiar Pagada");
              else estadosTotales.push("C1 Pagada");
            } else {
              estadosTotales.push("C1 Pendiente");
            }
            // Cuota2
            if (estadoC2) {
              if (pagandoAhora2 && esPagoFamiliar)
                estadosTotales.push("C2 Familiar Pagada");
              else if (prevPagada2 && estadoAIActual.includes("C2 Familiar"))
                estadosTotales.push("C2 Familiar Pagada");
              else estadosTotales.push("C2 Pagada");
            } else {
              estadosTotales.push("C2 Pendiente");
            }
            // Cuota3
            if (estadoC3) {
              if (pagandoAhora3 && esPagoFamiliar)
                estadosTotales.push("C3 Familiar Pagada");
              else if (prevPagada3 && estadoAIActual.includes("C3 Familiar"))
                estadosTotales.push("C3 Familiar Pagada");
              else estadosTotales.push("C3 Pagada");
            } else {
              estadosTotales.push("C3 Pendiente");
            }
            if (cantidadCuotasRegistrada === 2)
              estadosTotales = [estadosTotales[0], estadosTotales[1]];
            if (cantidadCuotasRegistrada === 1)
              estadosTotales = [estadosTotales[0]];
            nuevoEstadoPago = estadosTotales.join(", ");
          } else {
            nuevoEstadoPago = esPagoFamiliar ? "Pago Total Familiar" : "Pagado";
          }
        } else {
          // Lógica de estado parcial
          if (metodoPago === "Pago en Cuotas") {
            let estados = [];
            // Construir textos por cuota para estado parcial
            if (estadoC1) {
              if (pagandoAhora1 && esPagoFamiliar)
                estados.push("C1 Familiar Pagada");
              else if (prevPagada1 && estadoAIActual.includes("C1 Familiar"))
                estados.push("C1 Familiar Pagada");
              else estados.push("C1 Pagada");
            } else {
              estados.push("C1 Pendiente");
            }
            if (estadoC2) {
              if (pagandoAhora2 && esPagoFamiliar)
                estados.push("C2 Familiar Pagada");
              else if (prevPagada2 && estadoAIActual.includes("C2 Familiar"))
                estados.push("C2 Familiar Pagada");
              else estados.push("C2 Pagada");
            } else {
              estados.push("C2 Pendiente");
            }
            if (estadoC3) {
              if (pagandoAhora3 && esPagoFamiliar)
                estados.push("C3 Familiar Pagada");
              else if (prevPagada3 && estadoAIActual.includes("C3 Familiar"))
                estados.push("C3 Familiar Pagada");
              else estados.push("C3 Pagada");
            } else {
              estados.push("C3 Pendiente");
            }

            if (cantidadCuotasRegistrada === 2)
              estados = [estados[0], estados[1]];
            if (cantidadCuotasRegistrada === 1) estados = [estados[0]];

            nuevoEstadoPago = estados.join(", "); // Ej: "C1 Pagada, C2 Pagada, C3 Pendiente"
          } else {
            nuevoEstadoPago = "Pago Parcial (En revisión)"; // Transferencia/Efectivo parcial
          }
        }

        // 5. ACUMULAR DATOS PAGADOR (Columnas AL/AM)
        // --- (INICIO CORRECCIÓN v9 - Formato "Nombre Apellido" y "DNI") ---
        const datosNuevosNombre = nombrePagador; // Formato: "Nombre Apellido"
        const datosNuevosDNI = dniPagador; // Columna AM solo DNI
        // --- (FIN CORRECCIÓN v9) ---

        const celdaNombre = hoja.getRange(
          filaAfectada,
          COL_PAGADOR_NOMBRE_MANUAL // Columna AL
        );
        const celdaDNI = hoja.getRange(
          filaAfectada,
          COL_PAGADOR_DNI_MANUAL // Columna AM
        );
        const valorActualNombre = celdaNombre.getValue().toString().trim();
        const valorActualDNI = celdaDNI.getValue().toString().trim();

        const valorFinalNombre = valorActualNombre
          ? `${valorActualNombre}, ${datosNuevosNombre}`
          : datosNuevosNombre;
        const valorFinalDNI = valorActualDNI
          ? `${valorActualDNI}, ${datosNuevosDNI}`
          : datosNuevosDNI;

        // Escribir solo si se está pagando (fileUrl no está vacío)
        if (fileUrl) {
          celdaNombre.setValue(valorFinalNombre);
          celdaDNI.setValue(valorFinalDNI);
        }

        // Post-procesado: si es Pago Familiar y se pagaron cuotas ahora, asegurar que el texto del principal
        // refleje 'Familiar Pagada' para las cuotas pagadas en esta operación.
        if (esPagoFamiliar && cuotasPagadasAhora.size > 0) {
          let estadoProcesado = nuevoEstadoPago;
          if (cuotasPagadasAhora.has("mp_cuota_1")) {
            estadoProcesado = estadoProcesado.replace(
              /C1 Pagada/g,
              "C1 Familiar Pagada"
            );
          }
          if (cuotasPagadasAhora.has("mp_cuota_2")) {
            estadoProcesado = estadoProcesado.replace(
              /C2 Pagada/g,
              "C2 Familiar Pagada"
            );
          }
          if (cuotasPagadasAhora.has("mp_cuota_3")) {
            estadoProcesado = estadoProcesado.replace(
              /C3 Pagada/g,
              "C3 Familiar Pagada"
            );
          }
          nuevoEstadoPago = estadoProcesado;
        }

        // 6. SETEAR ESTADO DE PAGO (Columna AJ)
        hoja.getRange(filaAfectada, COL_ESTADO_PAGO).setValue(nuevoEstadoPago); // AJ
        Logger.log(
          `aplicarCambios FIN fila:${filaAfectada} nuevoEstado:'${nuevoEstadoPago}' pagadasCount:${pagadasCount} cantidadCuotas:${cantidadCuotasRegistrada}`
        );

        // 7. SETEAR LINK COMPROBANTE (Columnas AN, AO, AP, AQ)
        // Escribir solo si se está pagando (fileUrl no está vacío)
        if (fileUrl) {
          if (
            metodoPago === "Transferencia" ||
            metodoPago === "Pago Efectivo (Adm del Club)" ||
            (cuotasPagadasAhora.has("externo") || cuotasPagadasAhora.has("mp_total"))
          ) {
            hoja
              .getRange(filaAfectada, COL_COMPROBANTE_MANUAL_TOTAL_EXT) // AN
              .setValue(fileUrl);
          } else {
            // 'Pago en Cuotas' -> escribir en todas las cuotas seleccionadas (no usar else-if)
            cuotasPagadasAhora.forEach((cuota) => {
              if (cuota === "mp_cuota_1")
                hoja
                  .getRange(filaAfectada, COL_COMPROBANTE_MANUAL_CUOTA1) // AO
                  .setValue(fileUrl);
              if (cuota === "mp_cuota_2")
                hoja
                  .getRange(filaAfectada, COL_COMPROBANTE_MANUAL_CUOTA2) // AP
                  .setValue(fileUrl);
              if (cuota === "mp_cuota_3")
                hoja
                  .getRange(filaAfectada, COL_COMPROBANTE_MANUAL_CUOTA3) // AQ
                  .setValue(fileUrl);
            });
          }
        }

        // (CORRECCIÓN URGENTE) Actualizar el monto acumulado en AK ANTES de sobreescribir las celdas de cuota con "Pagada".
        const filaParaAcumular = hoja.getRange(filaAfectada, 1, 1, hoja.getLastColumn()).getValues()[0];
        _actualizarMontoAcumulado(filaAfectada, filaParaAcumular);

        // 8. SETEAR ESTADO CUOTAS (Columnas AF, AG, AH)
        // (CORRECCIÓN) Solo modificar columnas de cuotas si el método de pago es "Pago en Cuotas"
        if (metodoPago === "Pago en Cuotas") {
          if (esTotal) {
            hoja
              .getRange(filaAfectada, COL_CUOTA_1, 1, 3) // AF, AG, AH
              .setValues([["Pagada", "Pagada", "Pagada"]]);
          } else {
            // Solo marcar las cuotas pagadas AHORA
            cuotasPagadasAhora.forEach((cuota) => {
              if (cuota === "mp_cuota_1")
                hoja
                  .getRange(filaAfectada, COL_CUOTA_1) // AF
                  .setValue("Pagada");
              if (cuota === "mp_cuota_2")
                hoja
                  .getRange(filaAfectada, COL_CUOTA_2) // AG
                  .setValue("Pagada");
              if (cuota === "mp_cuota_3")
                hoja
                  .getRange(filaAfectada, COL_CUOTA_3) // AH
                  .setValue("Pagada");
            });
          }
        }

        // 9. Devolver el estado calculado para el mensaje de éxito
        let cuotasPagadasNombres = [];
        if (pagandoC1) cuotasPagadasNombres.push("Cuota 1");
        if (pagandoC2) cuotasPagadasNombres.push("Cuota 2");
        if (pagandoC3) cuotasPagadasNombres.push("Cuota 3");

        return {
          esTotal: esTotal,
          nuevoEstado: nuevoEstadoPago,
          cuotasPagadasNombres: cuotasPagadasNombres,
          pagadasCount: pagadasCount,
          cantidadCuotasRegistrada: cantidadCuotasRegistrada,
        };
      };
      // --- (FIN FUNCIÓN HELPER 'aplicarCambios') ---

      // --- 4. Construir Nombre del Archivo ---
      const fechaActual = Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "yyyy-MM-dd_HH-mm-ss");

      let baseNombreArchivo = "";

      if (metodoPagoHoja === 'Pago en Cuotas') {
        const cuotasNombres = cuotasSeleccionadas.map(c => {
          if (c === 'mp_cuota_1') return 'C1';
          if (c === 'mp_cuota_2') return 'C2';
          if (c === 'mp_cuota_3') return 'C3';
          return null;
        }).filter(Boolean);

        if (cuotasNombres.length > 0) {
          baseNombreArchivo = `${cuotasNombres.join('_')}_${dniHoja}_${fechaActual}`;
        } else {
          baseNombreArchivo = `Comprobante_${dniHoja}_${fechaActual}`;
        }
      } else { // Efectivo, Transferencia u otro.
        baseNombreArchivo = `Pago único${dniHoja}_${fechaActual}`;
      }

      const nombreArchivoLimpio = baseNombreArchivo.replace(/[^\w.()-]/g, "_");
      const extension = fileData.fileName.includes(".")
        ? fileData.fileName.split(".").pop()
        : "jpg";
      const nuevoNombreArchivo = `${nombreArchivoLimpio}.${extension}`;

      Logger.log(`Nuevo nombre de archivo: ${nuevoNombreArchivo}`);

      // --- 5. Subir el Archivo ---
      const fileUrl = uploadFileToDrive(
        fileData.data,
        fileData.mimeType,
        nuevoNombreArchivo,
        dniLimpio,
        "comprobante"
      );
      if (typeof fileUrl !== "string" || !fileUrl.startsWith("=HYPERLINK")) {
        throw new Error(
          "Error al subir el archivo a Drive: " +
            (fileUrl.message || "Error desconocido")
        );
      }

      // --- 6. Aplicar Cambios (Real) ---
      let mensajeExito = "";
      let resultadoPrincipal;

      if (esPagoFamiliar) {
        const idFamiliar = rangoFilaPrincipal[COL_VINCULO_PRINCIPAL - 1]; // AR (44)
        if (!idFamiliar) {
          Logger.log(
            `Pago Familiar marcado, pero no se encontró ID Familiar en fila ${fila}. Aplicando solo al DNI ${dniLimpio}.`
          );
          resultadoPrincipal = aplicarCambios(fila, metodoPagoHoja, fileUrl);
        } else {
          const rangoVinculos = hoja.getRange(
            2,
            COL_VINCULO_PRINCIPAL, // AR (44)
            hoja.getLastRow() - 1,
            1
          );
          const todasLasFilas = rangoVinculos
            .createTextFinder(idFamiliar)
            .matchEntireCell(true)
            .findAll();
          let nombresActualizados = [];

          // =========================================================
          // --- ¡¡INICIO DE LA CORRECCIÓN (Error 1)!! ---
          // (Helper modificado para pago total -> cuota 1)
          // =========================================================
          const aplicarCambiosHermano = (filaHermano, fileUrlHermano) => {
            let filaDatos = hoja
              .getRange(filaHermano, 1, 1, hoja.getLastColumn())
              .getValues()[0];
            const metodoPagoHermano = filaDatos[COL_METODO_PAGO - 1]; // AC
            const cantidadCuotasHermano =
              parseInt(filaDatos[COL_CANTIDAD_CUOTAS - 1]) || 0; // AI
            
            // Copia local del Set de cuotas del principal
            let cuotasPagadasAhoraLocal = new Set(cuotasPagadasAhora); 
            const esPagoTotalPrincipal = cuotasPagadasAhora.has("externo") || cuotasPagadasAhora.has("mp_total");

            if (esPagoTotalPrincipal && metodoPagoHermano === "Pago en Cuotas") {
                // El principal hizo un pago total. Esto debe contar como C1 (o la prox) para el hermano.
                const comp1h = filaDatos[COL_COMPROBANTE_MANUAL_CUOTA1 - 1]; // AO
                const comp2h = filaDatos[COL_COMPROBANTE_MANUAL_CUOTA2 - 1]; // AP
                const comp3h = filaDatos[COL_COMPROBANTE_MANUAL_CUOTA3 - 1]; // AQ

                // Busca el primer slot de cuota vacío
                if (!comp1h || String(comp1h).trim() === "") {
                    cuotasPagadasAhoraLocal.add("mp_cuota_1");
                } else if (!comp2h || String(comp2h).trim() === "") {
                    cuotasPagadasAhoraLocal.add("mp_cuota_2");
                } else if (!comp3h || String(comp3h).trim() === "") {
                    cuotasPagadasAhoraLocal.add("mp_cuota_3");
                }
                // Si todos están llenos, no se añade nada extra, pero el comprobante irá al "Total" (AN)
            }
            
            // 1) Append pagador manual (AL/AM)
            if (fileUrlHermano) {
              const celdaNombreH = hoja.getRange(
                filaHermano,
                COL_PAGADOR_NOMBRE_MANUAL // AL
              );
              const celdaDNIH = hoja.getRange(
                filaHermano,
                COL_PAGADOR_DNI_MANUAL // AM
              );
              const valNomAct = celdaNombreH.getValue().toString().trim();
              const valDniAct = celdaDNIH.getValue().toString().trim();
              const nuevoNom = valNomAct
                ? `${valNomAct}, ${nombrePagador}`
                : nombrePagador;
              const nuevoDni = valDniAct
                ? `${valDniAct}, ${dniPagador}`
                : dniPagador;
              celdaNombreH.setValue(nuevoNom);
              celdaDNIH.setValue(nuevoDni);
            }

            // 2) Marcar las cuotas pagadas AHORA (usando el Set local)
            if (metodoPagoHermano === "Pago en Cuotas") {
              if (cuotasPagadasAhoraLocal.has("mp_cuota_1"))
                hoja
                  .getRange(filaHermano, COL_CUOTA_1) // AF
                  .setValue("Pagada");
              if (cuotasPagadasAhoraLocal.has("mp_cuota_2"))
                hoja
                  .getRange(filaHermano, COL_CUOTA_2) // AG
                  .setValue("Pagada");
              if (cuotasPagadasAhoraLocal.has("mp_cuota_3"))
                hoja
                  .getRange(filaHermano, COL_CUOTA_3) // AH
                  .setValue("Pagada"); // (CORRECCIÓN)
            }
            
            // 3) Setear comprobantes (usando el Set local)
            if (fileUrlHermano) {
                if ((esPagoTotalPrincipal || cuotasPagadasAhoraLocal.has("externo") || cuotasPagadasAhoraLocal.has("mp_total")) && metodoPagoHermano !== "Pago en Cuotas") {
                    // Es un pago total (Efectivo/Transf) y el hermano también es de pago total
                    hoja.getRange(filaHermano, COL_COMPROBANTE_MANUAL_TOTAL_EXT).setValue(fileUrlHermano); // AN
                } else {
                    // Es pago en cuotas (o fue forzado a serlo)
                    cuotasPagadasAhoraLocal.forEach((cuota) => {
                      if (cuota === "mp_cuota_1")
                        hoja.getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA1).setValue(fileUrlHermano); // AO
                      if (cuota === "mp_cuota_2")
                        hoja.getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA2).setValue(fileUrlHermano); // AP
                      if (cuota === "mp_cuota_3")
                        hoja.getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA3).setValue(fileUrlHermano); // AQ
                      
                      // Fallback para el pago total del principal si el hermano es de cuotas pero no se encontró slot
                      if ((cuota === "externo" || cuota === "mp_total") && metodoPagoHermano === "Pago en Cuotas") {
                         hoja.getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA1).setValue(fileUrlHermano); // Pone en C1 (AO) por defecto
                      }
                    });
                }
            }
            // =========================================================
            // --- ¡¡FIN DE LA CORRECCIÓN (Error 1)!! ---
            // =========================================================

            // 4) Releer la fila y recalcular el estado (usando el Set local)
            let filaActualizada = hoja
              .getRange(filaHermano, 1, 1, hoja.getLastColumn())
              .getValues()[0];
            const estadoAIHermano = String(
              filaActualizada[COL_ESTADO_PAGO - 1] || "" // AJ
            );
            
            const comp1h = filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA1 - 1]; // AO
            const comp2h = filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA2 - 1]; // AP
            const comp3h = filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA3 - 1]; // AQ
            const compTotalh = filaActualizada[COL_COMPROBANTE_MANUAL_TOTAL_EXT - 1]; // AN

            const prevPagada1 =
              comp1h && String(comp1h).toString().trim() !== "";
            const prevPagada2 =
              comp2h && String(comp2h).toString().trim() !== "";
            const prevPagada3 =
              comp3h && String(comp3h).toString().trim() !== "";
            
            // ¡IMPORTANTE! 'ahoraPagada' se basa en el Set Local
            const ahoraPagada1 = prevPagada1; // El comprobante ya se escribió
            const ahoraPagada2 = prevPagada2;
            const ahoraPagada3 = prevPagada3;

            let estados = [];
            let pagadasCountHermano = 0;

            // Cuota 1
            if (ahoraPagada1) {
              pagadasCountHermano++;
              if ((cuotasPagadasAhoraLocal.has("mp_cuota_1") || esPagoTotalPrincipal) && esPagoFamiliar)
                estados.push("C1 Familiar Pagada");
              else if (prevPagada1 && estadoAIHermano.includes("C1 Familiar"))
                estados.push("C1 Familiar Pagada");
              else estados.push("C1 Pagada");
            } else {
              estados.push("C1 Pendiente");
            }
            // Cuota 2
            if (ahoraPagada2) {
              pagadasCountHermano++;
              if ((cuotasPagadasAhoraLocal.has("mp_cuota_2")) && esPagoFamiliar)
                estados.push("C2 Familiar Pagada");
              else if (prevPagada2 && estadoAIHermano.includes("C2 Familiar"))
                estados.push("C2 Familiar Pagada");
              else estados.push("C2 Pagada");
            } else {
              estados.push("C2 Pendiente");
            }
            // Cuota 3
            if (ahoraPagada3) {
              pagadasCountHermano++;
              if ((cuotasPagadasAhoraLocal.has("mp_cuota_3")) && esPagoFamiliar)
                estados.push("C3 Familiar Pagada");
              else if (prevPagada3 && estadoAIHermano.includes("C3 Familiar"))
                estados.push("C3 Familiar Pagada");
              else estados.push("C3 Pagada");
            } else {
              estados.push("C3 Pendiente");
            }
            
            // (Lógica de forzar "Familiar" movida arriba)
            
            if (cantidadCuotasHermano === 2) estados = [estados[0], estados[1]];
            if (cantidadCuotasHermano === 1) estados = [estados[0]];

            pagadasCountHermano = estados.filter(
              (s) => !s.toLowerCase().includes("pendiente")
            ).length;

            if (cantidadCuotasHermano === 0 && fileUrlHermano) {
               hoja
                .getRange(filaHermano, COL_ESTADO_PAGO) // AJ
                .setValue(esPagoFamiliar ? "Pago Total Familiar" : "Pagado");
              _actualizarMontoAcumulado(filaHermano, filaActualizada);
              return { esTotal: true, nuevoEstado: esPagoFamiliar ? "Pago Total Familiar" : "Pagado" };
            }

            let nuevoEstadoH = "";
            const esTotalHermano = cantidadCuotasHermano > 0 && pagadasCountHermano >= cantidadCuotasHermano;
            
            if (esTotalHermano) {
                nuevoEstadoH = estados.join(", ");
            } else if (pagadasCountHermano === 0) {
              nuevoEstadoH = `Pendiente (${cantidadCuotasHermano || 3} Cuotas)`;
            } else {
              nuevoEstadoH = estados.join(", ");
            }

            // Si el hermano NO es de cuotas, pero el principal pagó total
            if (metodoPagoHermano !== "Pago en Cuotas" && esPagoTotalPrincipal) {
                nuevoEstadoH = esPagoFamiliar ? "Pago Total Familiar" : "Pagado";
            }

            hoja.getRange(filaHermano, COL_ESTADO_PAGO).setValue(nuevoEstadoH); // AJ
            Logger.log(
              `aplicarCambiosHermano FIN. fila:${filaHermano} nuevoEstado:${nuevoEstadoH} pagadasCount:${pagadasCountHermano} cantidadCuotas:${cantidadCuotasHermano}`
            );

            // (NUEVO - CORRECCIÓN) Actualizar el monto acumulado en AK para el hermano
            _actualizarMontoAcumulado(filaHermano, filaActualizada);

            return {
              esTotal: esTotalHermano || (compTotalh && String(compTotalh).trim() !== ""),
              nuevoEstado: nuevoEstadoH,
            };
          };

          // Aplicar para cada miembro: principal con la función completa, hermanos con la función ligera
          todasLasFilas.forEach((celda) => {
            const rowNum = celda.getRow();
            if (datosExtras.subMetodo) {
              hoja.getRange(rowNum, COL_MODO_PAGO_CUOTA).setValue(datosExtras.subMetodo);
            }
            if (rowNum === fila) {
              const resultadoFila = aplicarCambios(
                rowNum,
                metodoPagoHoja,
                fileUrl
              );
              resultadoPrincipal = resultadoFila; // Guardar resultado del principal
              nombresActualizados.push(
                hoja.getRange(rowNum, COL_NOMBRE).getValue()
              );
            } else {
              const resultadoHermano = aplicarCambiosHermano(rowNum, fileUrl);
              nombresActualizados.push(
                hoja.getRange(rowNum, COL_NOMBRE).getValue()
              );
            }
          });

          Logger.log(
            `Pago Familiar aplicado a ${
              nombresActualizados.length
            } miembros: ${nombresActualizados.join(", ")}`
          );

          if (resultadoPrincipal.esTotal) {
            mensajeExito = `¡Pago Familiar Total registrado con éxito para ${nombresActualizados.length} inscriptos!<br>${mensajeFinalCompleto}`;
          } else {
            mensajeExito = `Se registró el pago de ${resultadoPrincipal.cuotasPagadasNombres.join(
              " y "
            )} para ${nombresActualizados.length} inscriptos.`;
          }
        }
      } else {
        // Aplicación Individual
        if (datosExtras.subMetodo) {
          hoja.getRange(fila, COL_MODO_PAGO_CUOTA).setValue(datosExtras.subMetodo);
        }
        resultadoPrincipal = aplicarCambios(fila, metodoPagoHoja, fileUrl);
      }

      // --- 7. Formular Mensaje de Éxito ---
      if (!mensajeExito) {
        // Si el mensaje no se seteó en el bloque familiar (porque fue individual)
        if (resultadoPrincipal.esTotal) {
          mensajeExito = mensajeFinalCompleto;
        } else {
          mensajeExito = `Se registró el pago de: ${resultadoPrincipal.cuotasPagadasNombres.join(
            " y "
          )}.`;
          const pendientes =
            resultadoPrincipal.cantidadCuotasRegistrada -
            resultadoPrincipal.pagadasCount;

          if (pendientes > 0) {
            mensajeExito += ` Le quedan ${pendientes} cuota${
              pendientes > 1 ? "s" : ""
            } pendiente${pendientes > 1 ? "s" : ""}.`;
          } else {
            mensajeExito = `¡Felicidades! Ha completado todas las cuotas.<br>${mensajeFinalCompleto}`;
          }
        }
      }

      Logger.log(
        `Comprobante subido para DNI ${dniLimpio}. Estado final: ${resultadoPrincipal.nuevoEstado}. ¿Familiar?: ${esPagoFamiliar}`
      );

      // Leer la fila actualizada del principal para calcular comprobantes/pendientes que usa la UI
      const filaActualizadaPrincipal = hoja
        .getRange(fila, 1, 1, hoja.getLastColumn())
        .getValues()[0];
      const c_total_p =
        filaActualizadaPrincipal[COL_COMPROBANTE_MANUAL_TOTAL_EXT - 1]; // AN
      const c_c1_p =
        filaActualizadaPrincipal[COL_COMPROBANTE_MANUAL_CUOTA1 - 1]; // AO
      const c_c2_p =
        filaActualizadaPrincipal[COL_COMPROBANTE_MANUAL_CUOTA2 - 1]; // AP
      const c_c3_p =
        filaActualizadaPrincipal[COL_COMPROBANTE_MANUAL_CUOTA3 - 1]; // AQ
      
      // (CORRECCIÓN) Volver a calcular las cuotas pagadas con los datos frescos para devolver a la UI
      const cuotasPagadasPorComp = [];
      if (c_c1_p && String(c_c1_p).trim() !== "") cuotasPagadasPorComp.push("mp_cuota_1");
      if (c_c2_p && String(c_c2_p).trim() !== "") cuotasPagadasPorComp.push("mp_cuota_2");
      if (c_c3_p && String(c_c3_p).trim() !== "") cuotasPagadasPorComp.push("mp_cuota_3");

      let cantidadCuotasReg = parseInt(
        filaActualizadaPrincipal[COL_CANTIDAD_CUOTAS - 1] // AI
      );
      if (isNaN(cantidadCuotasReg) || cantidadCuotasReg < 1)
        cantidadCuotasReg = metodoPagoHoja === "Pago en Cuotas" ? 3 : 0;
      // Ajustar según cantidad registrada
      let cuotasPagadasFinalP = cuotasPagadasPorComp.slice();
      if (cantidadCuotasReg === 2)
        cuotasPagadasFinalP = cuotasPagadasFinalP.filter(
          (c) => c !== "mp_cuota_3"
        );
      if (cantidadCuotasReg === 1)
        cuotasPagadasFinalP = cuotasPagadasFinalP.filter(
          (c) => c === "mp_cuota_1"
        );
      const pagadasCountP = cuotasPagadasFinalP.length;
      const pendientesByCompP = Math.max(0, cantidadCuotasReg - pagadasCountP);
      const comprobantesCompletosResp =
        (cantidadCuotasReg > 0 && pagadasCountP >= cantidadCuotasReg) ||
        Boolean(c_total_p);
      
      // (CORRECCIÓN FINAL v2) Calcular el estado de la familia para el refresco de la UI
      const idFamiliarPrincipal = filaActualizadaPrincipal[COL_VINCULO_PRINCIPAL - 1];
      
      // Esta lógica ahora se ejecuta siempre para asegurar que 'familiaPagos' esté definido.
      try {
        const rangoVinculos = hoja.getRange(2, COL_VINCULO_PRINCIPAL, hoja.getLastRow() - 1, 1);
        const rangoCompletoHoja = hoja.getRange(1, 1, hoja.getLastRow(), hoja.getLastColumn()).getValues();

        function obtenerEstadoCuota(fila, numeroCuota) {
          const colIndex = (COL_COMPROBANTE_MANUAL_CUOTA1 - 1) + (numeroCuota - 1);
          const comprobante = fila[colIndex];
          return (comprobante && String(comprobante).trim() !== "") ? "pagada" : "pendiente";
        }

        if (idFamiliarPrincipal) {
            const filasFamiliaEncontradas = rangoVinculos.createTextFinder(idFamiliarPrincipal).matchEntireCell(true).findAll();
            filasFamiliaEncontradas.forEach(celda => {
              const filaIndex = celda.getRow() - 1;
              const filaMiembro = rangoCompletoHoja[filaIndex];
              const dniMiembro = limpiarDNI(filaMiembro[COL_DNI_INSCRIPTO - 1]);
              if (dniMiembro) {
                familiaPagos[dniMiembro] = {
                  c1: obtenerEstadoCuota(filaMiembro, 1),
                  c2: obtenerEstadoCuota(filaMiembro, 2),
                  c3: obtenerEstadoCuota(filaMiembro, 3)
                };
              }
            });
        } else {
            // Si no hay vínculo familiar, al menos devolver el estado del usuario actual
            familiaPagos[dniLimpio] = {
              c1: obtenerEstadoCuota(filaActualizadaPrincipal, 1),
              c2: obtenerEstadoCuota(filaActualizadaPrincipal, 2),
              c3: obtenerEstadoCuota(filaActualizadaPrincipal, 3)
            };
        }
      } catch (e) {
        Logger.log("Error recalculando familiaPagos al final de subirComprobanteManual: " + e.toString());
        // Si todo falla, al menos enviar el estado del principal para evitar el error de 'undefined'
        familiaPagos[dniLimpio] = { c1: "pendiente", c2: "pendiente", c3: "pendiente" };
      }

      return {
        status: "OK",
        message: mensajeExito,
        estadoPago: resultadoPrincipal.nuevoEstado,
        comprobantesCompletos: comprobantesCompletosResp,
        cuotasPagadas: cuotasPagadasFinalP, // (CORRECCIÓN) El nombre de la propiedad debe ser 'cuotasPagadas'
        familiaPagos: familiaPagos, // (CORRECCIÓN) Añadir el estado de la familia para el refresco de la UI
        cuotasPendientes: pendientesByCompP,
      };
    } else {
      Logger.log(
        `No se encontró DNI ${dniLimpio} para subir comprobante manual.`
      );
      return {
        status: "ERROR",
        message: `No se encontró el registro para el DNI ${dniLimpio}. Asegúrese de que el DNI del inscripto sea correcto.`,
      };
    }
  } catch (e) {
    Logger.log(
      "Error en subirComprobanteManual: " + e.toString() + " Stack: " + e.stack
    );
    return { status: "ERROR", message: "Error en el servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}