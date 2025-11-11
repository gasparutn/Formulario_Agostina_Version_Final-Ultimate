// Archivo: Comprobantes.js
// Contiene la lógica de subida y procesamiento de comprobantes
// Migrado desde Código.js para mayor modularidad.

/**
 * (MODIFICADO - v9 - CORRECCIÓN BUGS 1, 2, 0>=0 y Formato AN/AO)
 * - Lógica de estado (AI) y prefijo (AN/AO) REFACTORIZADA.
 * - BUG 1 (Prefijo): Prefijos (C1, C2) ELIMINADOS de AN/AO.
 * - BUG 2 (Hermanos): El estado de pago AHORA se recalcula por fila, en lugar de copiarse desde el principal.
 * - BUG 0>=0: La lógica de 'esTotal' ahora comprueba que 'cantidadCuotas' sea > 0.
 * - FORMATO AN/AO: Columna AN guarda "Nombre Apellido", Columna AO guarda "DNI".
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
          rangoFila[COL_CUOTA_1 - 1],
          rangoFila[COL_CUOTA_2 - 1],
          rangoFila[COL_CUOTA_3 - 1],
        ];
        const estadoAIActual = String(rangoFila[COL_ESTADO_PAGO - 1] || "");

        // (Corrección Bug 0>=0)
        let cantidadCuotasRegistrada = parseInt(
          rangoFila[COL_CANTIDAD_CUOTAS - 1]
        );
        if (
          metodoPago === "Pago en Cuotas" &&
          (isNaN(cantidadCuotasRegistrada) || cantidadCuotasRegistrada < 1)
        ) {
          cantidadCuotasRegistrada = 3;
          hoja.getRange(filaAfectada, COL_CANTIDAD_CUOTAS).setValue(3);
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
        const comp1 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA1 - 1];
        const comp2 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA2 - 1];
        const comp3 = rangoFila[COL_COMPROBANTE_MANUAL_CUOTA3 - 1];
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
          ["mp_total", "externo"].includes(cuotasSeleccionadas[0])
        ) {
          esTotal = true; // Pago total para Transferencia/Efectivo
        }

        // 4. DETERMINAR ESTADO DE PAGO (Columna AI) (Corrección Bug 2)
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

        // 5. ACUMULAR DATOS PAGADOR (Columnas AN/AO)
        // --- (INICIO CORRECCIÓN v9 - Formato "Nombre Apellido" y "DNI") ---
        const datosNuevosNombre = nombrePagador; // Formato: "Nombre Apellido"
        const datosNuevosDNI = dniPagador; // Columna AO solo DNI
        // --- (FIN CORRECCIÓN v9) ---

        const celdaNombre = hoja.getRange(
          filaAfectada,
          COL_PAGADOR_NOMBRE_MANUAL
        ); // Columna AN
        const celdaDNI = hoja.getRange(filaAfectada, COL_PAGADOR_DNI_MANUAL); // Columna AO
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

        // 6. SETEAR ESTADO DE PAGO (Columna AI)
        hoja.getRange(filaAfectada, COL_ESTADO_PAGO).setValue(nuevoEstadoPago);
        Logger.log(
          `aplicarCambios FIN fila:${filaAfectada} nuevoEstado:'${nuevoEstadoPago}' pagadasCount:${pagadasCount} cantidadCuotas:${cantidadCuotasRegistrada}`
        );

        // 7. SETEAR LINK COMPROBANTE (Columnas AQ, AR, AS, AT)
        // Escribir solo si se está pagando (fileUrl no está vacío)
        if (fileUrl) {
          if (
            metodoPago === "Transferencia" ||
            metodoPago === "Pago Efectivo (Adm del Club)"
          ) {
            hoja
              .getRange(filaAfectada, COL_COMPROBANTE_MANUAL_TOTAL_EXT)
              .setValue(fileUrl); // AQ
          } else {
            // 'Pago en Cuotas' -> escribir en todas las cuotas seleccionadas (no usar else-if)
            cuotasPagadasAhora.forEach((cuota) => {
              if (cuota === "mp_cuota_1")
                hoja
                  .getRange(filaAfectada, COL_COMPROBANTE_MANUAL_CUOTA1)
                  .setValue(fileUrl); // AR
              if (cuota === "mp_cuota_2")
                hoja
                  .getRange(filaAfectada, COL_COMPROBANTE_MANUAL_CUOTA2)
                  .setValue(fileUrl); // AS
              if (cuota === "mp_cuota_3")
                hoja
                  .getRange(filaAfectada, COL_COMPROBANTE_MANUAL_CUOTA3)
                  .setValue(fileUrl); // AT
            });
          }
        }

        // 8. SETEAR ESTADO CUOTAS (Columnas AE, AF, AG)
        // (CORRECCIÓN) Solo modificar columnas de cuotas si el método de pago es "Pago en Cuotas"
        if (metodoPago === "Pago en Cuotas") {
          if (esTotal) {
            hoja
              .getRange(filaAfectada, COL_CUOTA_1, 1, 3)
              .setValues([["Pagada", "Pagada", "Pagada"]]);
          } else {
            // Solo marcar las cuotas pagadas AHORA
            cuotasPagadasAhora.forEach((cuota) => {
              if (cuota === "mp_cuota_1")
                hoja
                  .getRange(filaAfectada, COL_CUOTA_1)
                  .setValue("Pagada (En revisión)"); // AE
              if (cuota === "mp_cuota_2")
                hoja
                  .getRange(filaAfectada, COL_CUOTA_2)
                  .setValue("Pagada (En revisión)"); // AF
              if (cuota === "mp_cuota_3")
                hoja
                  .getRange(filaAfectada, COL_CUOTA_3)
                  .setValue("Pagada (En revisión)"); // AG
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
      // (Se usa el estado del principal para el nombre del archivo, ANTES de aplicar cambios)
      const { nuevoEstado: estadoParaNombre } = aplicarCambios(
        fila,
        metodoPagoHoja,
        ""
      ); // Simulación
      hoja
        .getRange(fila, 1, 1, hoja.getLastColumn())
        .setValues([rangoFilaPrincipal]); // Revertir simulación

      let baseNombreArchivo = "";
      const metodoPagoSimple = metodoPagoHoja.replace(/[\s()]/g, "");
      const estadoPagoSimple = estadoParaNombre.replace(/[\s(),]/g, "_");

      if (esPagoFamiliar) {
        baseNombreArchivo = `${dniHoja}_${apellidoHoja}_${metodoPagoSimple}_${estadoPagoSimple}`;
      } else {
        baseNombreArchivo = `${dniHoja}_${apellidoHoja}_${nombreHoja}_${metodoPagoSimple}_${estadoPagoSimple}`;
      }
      if (metodoPagoHoja === "Pago en Cuotas") {
        const prefijoCuotas = cuotasSeleccionadas
          .map((c) => c.replace("mp_", ""))
          .join("-");
        baseNombreArchivo = `${prefijoCuotas}_${baseNombreArchivo}`;
      }

      const nombreArchivoLimpio = baseNombreArchivo.replace(/[^\w.-]/g, "_");
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
        const idFamiliar = rangoFilaPrincipal[COL_VINCULO_PRINCIPAL - 1]; // AV (48)
        if (!idFamiliar) {
          Logger.log(
            `Pago Familiar marcado, pero no se encontró ID Familiar en fila ${fila}. Aplicando solo al DNI ${dniLimpio}.`
          );
          resultadoPrincipal = aplicarCambios(fila, metodoPagoHoja, fileUrl);
        } else {
          const rangoVinculos = hoja.getRange(
            2,
            COL_VINCULO_PRINCIPAL,
            hoja.getLastRow() - 1,
            1
          );
          const todasLasFilas = rangoVinculos
            .createTextFinder(idFamiliar)
            .matchEntireCell(true)
            .findAll();
          let nombresActualizados = [];

          // Helper específico para hermanos: solo marca la(s) cuota(s) seleccionada(s)
          // y recalcula el estado de pago a partir de AE/AF/AG y COL_CANTIDAD_CUOTAS.
          const aplicarCambiosHermano = (filaHermano, fileUrlHermano) => {
            Logger.log(
              `aplicarCambiosHermano INICIO. fila:${filaHermano} fileUrl:${
                fileUrlHermano ? "SI" : "NO"
              } cuotasActuales:${Array.from(cuotasPagadasAhora).join(
                "|"
              )} esPagoFamiliar:${esPagoFamiliar} pagador:${nombrePagador}/${dniPagador}`
            );
            // Leer la fila actual
            let filaDatos = hoja
              .getRange(filaHermano, 1, 1, hoja.getLastColumn())
              .getValues()[0];
            const metodoPagoHermano = filaDatos[COL_METODO_PAGO - 1]; // Leer el método de pago del hermano
            const cantidadCuotasHermano =
              parseInt(filaDatos[COL_CANTIDAD_CUOTAS - 1]) || 0;

            // 1) Append pagador manual (AN/AO) si corresponde
            if (fileUrlHermano) {
              const celdaNombreH = hoja.getRange(
                filaHermano,
                COL_PAGADOR_NOMBRE_MANUAL
              );
              const celdaDNIH = hoja.getRange(
                filaHermano,
                COL_PAGADOR_DNI_MANUAL
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

            // 2) Marcar las cuotas pagadas AHORA (solo modificar AE/AF/AG)
            // (CORRECCIÓN) Solo ejecutar si el método de pago del hermano es "Pago en Cuotas"
            if (metodoPagoHermano === "Pago en Cuotas") {
              if (cuotasPagadasAhora.has("mp_cuota_1"))
                hoja
                  .getRange(filaHermano, COL_CUOTA_1)
                  .setValue("Pagada (En revisión)");
              if (cuotasPagadasAhora.has("mp_cuota_2"))
                hoja
                  .getRange(filaHermano, COL_CUOTA_2)
                  .setValue("Pagada (En revisión)");
              if (cuotasPagadasAhora.has("mp_cuota_3"))
                hoja
                  .getRange(filaHermano, COL_CUOTA_3)
                  .setValue("Pagada (En revisión)");
            }

            // 3) Setear comprobantes (AQ/AR/AS/AT) según la cuota pagada y fileUrl
            if (fileUrlHermano) {
              // Para cada cuota seleccionada intentar escribir en la misma cuota del hermano;
              // si ya existe un comprobante en esa cuota, buscar la primera cuota disponible
              const selected = Array.from(cuotasPagadasAhora)
                .map((s) => parseInt(s.split("_").pop()))
                .filter((n) => !isNaN(n));
              const maxCuotas = cantidadCuotasHermano || 3;
              selected.forEach((idx) => {
                // índice objetivo inicial (1..3)
                let target = idx;
                const getCompAt = (i) => {
                  if (i === 1)
                    return hoja
                      .getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA1)
                      .getValue();
                  if (i === 2)
                    return hoja
                      .getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA2)
                      .getValue();
                  if (i === 3)
                    return hoja
                      .getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA3)
                      .getValue();
                  return null;
                };
                const setCompAt = (i, val) => {
                  if (i === 1)
                    hoja
                      .getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA1)
                      .setValue(val);
                  if (i === 2)
                    hoja
                      .getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA2)
                      .setValue(val);
                  if (i === 3)
                    hoja
                      .getRange(filaHermano, COL_COMPROBANTE_MANUAL_CUOTA3)
                      .setValue(val);
                };

                // If target slot already has a comprobante, find first empty slot within cantidadCuotasHermano
                if (
                  getCompAt(target) &&
                  String(getCompAt(target)).toString().trim() !== ""
                ) {
                  let found = false;
                  for (let j = 1; j <= maxCuotas; j++) {
                    if (
                      !getCompAt(j) ||
                      String(getCompAt(j)).toString().trim() === ""
                    ) {
                      target = j;
                      found = true;
                      break;
                    }
                  }
                  if (!found) {
                    // no empty slot -> skip to avoid overwriting
                    return;
                  }
                }
                // Set comprobante into target slot
                setCompAt(target, fileUrlHermano);
              });
            }

            // 4) Releer la fila AHORA que las escrituras han ocurrido y recalcular el estado
            // Esto evita usar valores antiguos (filaDatos) y prevenir marcar cuotas como pagadas
            // por lecturas previas inconsistentes.
            let filaActualizada = hoja
              .getRange(filaHermano, 1, 1, hoja.getLastColumn())
              .getValues()[0];
            const prevC1 = String(filaActualizada[COL_CUOTA_1 - 1])
              .toString()
              .trim();
            const prevC2 = String(filaActualizada[COL_CUOTA_2 - 1])
              .toString()
              .trim();
            const prevC3 = String(filaActualizada[COL_CUOTA_3 - 1])
              .toString()
              .trim();

            Logger.log(
              `filaHermano=${filaHermano} prevC1:'${prevC1}' prevC2:'${prevC2}' prevC3:'${prevC3}' comp1:'${
                filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA1 - 1]
              }' comp2:'${
                filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA2 - 1]
              }' comp3:'${filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA3 - 1]}'`
            );

            const comp1h = filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA1 - 1];
            const comp2h = filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA2 - 1];
            const comp3h = filaActualizada[COL_COMPROBANTE_MANUAL_CUOTA3 - 1];
            const estadoAIHermano = String(
              filaActualizada[COL_ESTADO_PAGO - 1] || ""
            );
            // Solo considerar una cuota como pagada previamente si existe un comprobante asociado.
            const prevPagada1 =
              comp1h && String(comp1h).toString().trim() !== "";
            const prevPagada2 =
              comp2h && String(comp2h).toString().trim() !== "";
            const prevPagada3 =
              comp3h && String(comp3h).toString().trim() !== "";

            const ahoraPagada1 =
              prevPagada1 || cuotasPagadasAhora.has("mp_cuota_1");
            const ahoraPagada2 =
              prevPagada2 || cuotasPagadasAhora.has("mp_cuota_2");
            const ahoraPagada3 =
              prevPagada3 || cuotasPagadasAhora.has("mp_cuota_3");

            // Construir textos de estado; si es pago familiar, usar la etiqueta 'Familiar Pagada' solo para cuotas
            // que se pagaron ahora como familiar o que ya tenían la marca familiar previamente.
            let estados = [];
            let pagadasCountHermano = 0;

            // Cuota 1
            if (ahoraPagada1) {
              pagadasCountHermano++;
              if (cuotasPagadasAhora.has("mp_cuota_1") && esPagoFamiliar)
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
              if (cuotasPagadasAhora.has("mp_cuota_2") && esPagoFamiliar)
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
              if (cuotasPagadasAhora.has("mp_cuota_3") && esPagoFamiliar)
                estados.push("C3 Familiar Pagada");
              else if (prevPagada3 && estadoAIHermano.includes("C3 Familiar"))
                estados.push("C3 Familiar Pagada");
              else estados.push("C3 Pagada");
            } else {
              estados.push("C3 Pendiente");
            }

            // Asegurar que las cuotas seleccionadas en la operación familiar se marquen como 'Familiar Pagada'
            const selectedIdxHermano = Array.from(cuotasPagadasAhora)
              .map((s) => parseInt(s.split("_").pop()))
              .filter((n) => !isNaN(n));
            selectedIdxHermano.forEach((idx) => {
              if (idx >= 1 && idx <= (cantidadCuotasHermano || 3)) {
                // Ajustar el texto correspondiente (índice - 1)
                const pos = idx - 1;
                // Si la posición existe en el array estados, forzar 'CN Familiar Pagada'
                if (estados[pos] !== undefined) {
                  estados[pos] = `C${idx} Familiar Pagada`;
                }
              }
            });

            if (cantidadCuotasHermano === 2) estados = [estados[0], estados[1]];
            if (cantidadCuotasHermano === 1) estados = [estados[0]];

            // Recalcular pagadasCountHermano tras forzar etiquetas familiares
            pagadasCountHermano = estados.filter(
              (s) => !s.toLowerCase().includes("pendiente")
            ).length;

            // Si el método es Transferencia/Efectivo (cantidadCuotasHermano === 0) y se adjuntó comprobante,
            // tratar como Pago Total Familiar y escribir comprobante total en AQ.
            if (cantidadCuotasHermano === 0 && fileUrlHermano) {
              hoja
                .getRange(filaHermano, COL_COMPROBANTE_MANUAL_TOTAL_EXT)
                .setValue(fileUrlHermano);
              hoja
                .getRange(filaHermano, COL_ESTADO_PAGO)
                .setValue("Pagado Total Familiar");
              return { esTotal: true, nuevoEstado: "Pagado Total Familiar" };
            }

            let nuevoEstadoH = "";
            if (pagadasCountHermano === 0) {
              nuevoEstadoH = `Pendiente (${cantidadCuotasHermano} Cuotas)`;
            } else {
              nuevoEstadoH = estados.join(", ");
            }

            hoja.getRange(filaHermano, COL_ESTADO_PAGO).setValue(nuevoEstadoH);
            Logger.log(
              `aplicarCambiosHermano FIN. fila:${filaHermano} nuevoEstado:${nuevoEstadoH} pagadasCount:${pagadasCountHermano} cantidadCuotas:${cantidadCuotasHermano}`
            );

            return {
              esTotal:
                cantidadCuotasHermano > 0 &&
                pagadasCountHermano >= cantidadCuotasHermano,
              nuevoEstado: nuevoEstadoH,
            };
          };

          // Aplicar para cada miembro: principal con la función completa, hermanos con la función ligera
          todasLasFilas.forEach((celda) => {
            const rowNum = celda.getRow();
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
        filaActualizadaPrincipal[COL_COMPROBANTE_MANUAL_TOTAL_EXT - 1];
      const c_c1_p =
        filaActualizadaPrincipal[COL_COMPROBANTE_MANUAL_CUOTA1 - 1];
      const c_c2_p =
        filaActualizadaPrincipal[COL_COMPROBANTE_MANUAL_CUOTA2 - 1];
      const c_c3_p =
        filaActualizadaPrincipal[COL_COMPROBANTE_MANUAL_CUOTA3 - 1];
      const cuotasPagadasPorComp = [];
      if (c_c1_p && String(c_c1_p).trim() !== "")
        cuotasPagadasPorComp.push("mp_cuota_1");
      if (c_c2_p && String(c_c2_p).trim() !== "")
        cuotasPagadasPorComp.push("mp_cuota_2");
      if (c_c3_p && String(c_c3_p).trim() !== "")
        cuotasPagadasPorComp.push("mp_cuota_3");
      let cantidadCuotasReg = parseInt(
        filaActualizadaPrincipal[COL_CANTIDAD_CUOTAS - 1]
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

      return {
        status: "OK",
        message: mensajeExito,
        estadoPago: resultadoPrincipal.nuevoEstado,
        comprobantesCompletos: comprobantesCompletosResp,
        cuotasPagadas: cuotasPagadasFinalP,
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

/**
 * (MODIFICADO)
 * Sube un archivo a Drive con un nombre de archivo específico.
 * Devuelve un =HYPERLINK() para la hoja de cálculo.
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

// Simulador de pruebas eliminado: si necesita volver a realizar pruebas localmente,
// puede reactivar una copia de este bloque en su entorno local o mantenerlo en una rama separada.
