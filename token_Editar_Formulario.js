
// Archivo: token_Editar_Formulario.gs
// Contiene la lógica para la validación por token de un solo uso para la edición de datos.

const TOKEN_EXPIRATION_MINUTES = 2; // El token expira en 2 minutos

/**
 * Genera un token de 6 dígitos, lo guarda y lo envía por email al responsable.
 * @param {string} dni - El DNI del inscripto.
 * @returns {{status: string, message: string}}
 */
function generarYEnviarToken(dni) {
  try {
    const dniLimpio = limpiarDNI(dni);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);

    if (!hojaRegistro) {
      throw new Error("La hoja de registros no fue encontrada.");
    }

    const rangoDni = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
    const celdaEncontrada = rangoDni.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

    if (!celdaEncontrada) {
      return { status: 'ERROR', message: 'No se encontró el DNI en los registros.' };
    }

    const fila = celdaEncontrada.getRow();
    const emailResponsable = hojaRegistro.getRange(fila, COL_EMAIL).getValue();
    const nombreInscripto = hojaRegistro.getRange(fila, COL_NOMBRE).getValue();

    if (!emailResponsable || !emailResponsable.includes('@')) {
      return { status: 'ERROR', message: 'No se encontró un email válido para este registro. Por favor, contacte a la administración.' };
    }

    // Generar token numérico de 6 dígitos
    const token = Math.floor(100000 + Math.random() * 900000).toString();

    // Guardar el token usando PropertiesService (expira en 10 minutos)
    const scriptProperties = PropertiesService.getScriptProperties();
    const propertyKey = `token_${dniLimpio}`;
    const tokenData = {
      token: token,
      timestamp: new Date().getTime() // Guardar la hora de creación
    };
    scriptProperties.setProperty(propertyKey, JSON.stringify(tokenData));

    // Enviar el email
    const asunto = "Tu Código de Verificación para Editar Datos";
    const cuerpo = `
      <p>Hola,</p>
      <p>Has solicitado editar los datos de <strong>${nombreInscripto}</strong>.</p>
      <p>Usa el siguiente código de verificación para continuar. Este código es válido por ${TOKEN_EXPIRATION_MINUTES} minutos.</p>
      <h2 style="text-align:center; letter-spacing: 5px; font-size: 28px;">${token}</h2>
      <p>Si no solicitaste este cambio, puedes ignorar este correo.</p>
      <p>Saludos,<br>Escuela Hípico Mendoza</p>
    `;

    MailApp.sendEmail({
      to: emailResponsable,
      subject: asunto,
      htmlBody: cuerpo,
      name: 'Escuela Hípico Mendoza'
    });

    Logger.log(`Token ${token} enviado a ${emailResponsable} para DNI ${dniLimpio}.`);
    return { status: 'OK', message: `Se ha enviado un código de 6 dígitos a <strong>${emailResponsable}</strong>. Por favor, ingréselo para continuar.` };

  } catch (e) {
    Logger.log("Error en generarYEnviarToken: " + e.message);
    return { status: 'ERROR', message: 'Error en el servidor al generar el token: ' + e.message };
  }
}

/**
 * Valida el token ingresado por el usuario.
 * @param {string} dni - El DNI del inscripto.
 * @param {string} tokenIngresado - El token de 6 dígitos.
 * @returns {{status: string, message: string}}
 */
function validarToken(dni, tokenIngresado) {
  const dniLimpio = limpiarDNI(dni);
  const scriptProperties = PropertiesService.getScriptProperties();
  const propertyKey = `token_${dniLimpio}`;
  const tokenDataString = scriptProperties.getProperty(propertyKey);

  // Eliminar el token inmediatamente para que sea de un solo uso
  scriptProperties.deleteProperty(propertyKey);

  if (!tokenDataString) {
    Logger.log(`Intento de validación de token fallido para DNI ${dniLimpio}.`);
    return { status: 'ERROR', message: 'El código ingresado es incorrecto o ha expirado. Por favor, intente de nuevo.' };
  }

  const tokenData = JSON.parse(tokenDataString);
  const tokenGuardado = tokenData.token;
  const timestampGuardado = tokenData.timestamp;
  const ahora = new Date().getTime();

  if (tokenGuardado === tokenIngresado && (ahora - timestampGuardado) < (TOKEN_EXPIRATION_MINUTES * 60 * 1000)) {
    Logger.log(`Token validado con éxito para DNI ${dniLimpio}.`);
    return { status: 'OK', message: 'Token correcto.' };
  } else {
    Logger.log(`Intento de validación de token fallido para DNI ${dniLimpio}. Token expirado o incorrecto.`);
    return { status: 'ERROR', message: 'El código ingresado es incorrecto o ha expirado. Por favor, intente de nuevo.' };
  }
}

/**
 * Función de mantenimiento para eliminar tokens antiguos.
 */
function eliminarTokenAntiguo(e) {
  // Esta función es llamada por un trigger, pero la lógica principal de eliminación
  // ya está en validarToken para asegurar que sea de un solo uso.
}
