const SPREADSHEET_ID = '1Ru-XGng2hYJbUvl-H2IA7aYQx7Ju-jk1LT1fkYOnG0w';
/* */
const NOMBRE_HOJA_BUSQUEDA = 'Base de Datos';
const NOMBRE_HOJA_REGISTRO = 'Registros';
const NOMBRE_HOJA_CONFIG = 'Config';

const FOLDER_ID_FOTOS = '1S2SbkuYdvcLFZYoHacfgwEU80kAN094l';
const FOLDER_ID_FICHAS = '1aDsTTDWHiDFUeZ8ByGp8_LY3fdzVQomu';
const FOLDER_ID_COMPROBANTES = '169EISq4RsDetQ0H3B17ViZFfe25xPcMM';

// =========================================================
// CONSTANTES "Base de Datos" (Sin cambios)
// =========================================================
const COL_HABILITADO_BUSQUEDA = 2; // Col B
const COL_NOMBRE_BUSQUEDA = 3; // Col C
const COL_APELLIDO_BUSQUEDA = 4; // Col D
const COL_FECHA_NACIMIENTO_BUSQUEDA = 5; // Col E
const COL_DNI_BUSQUEDA = 7; // Col G
const COL_OBRASOCIAL_BUSQUEDA = 8; // Col H
const COL_COLEGIO_BUSQUEDA = 9; // Col I
const COL_RESPONSABLE_BUSQUEDA = 10; // Col J
const COL_TELEFONO_BUSQUEDA = 11; // Col K
// =========================================================
// (MODIFICADO v15-CORREGIDO) CONSTANTES "Registros" (47 columnas A-AU)
// Basado en la lista FINAL del usuario.
// =========================================================
const COL_NUMERO_TURNO = 1; // A: N° de Turno
const COL_MARCA_TEMPORAL = 2; // B: Marca temporal
const COL_MARCA_N_E_A = 3; // C: N(Normal) E(Extendida)
const COL_ESTADO_NUEVO_ANT = 4; // D: Inscripto
const COL_EMAIL = 5; // E: Email
const COL_NOMBRE = 6; // F: Nombres
const COL_APELLIDO = 7; // G: Apellidos
const COL_FECHA_NACIMIENTO_REGISTRO = 8; // H: Fecha de Nacimiento
const COL_GRUPOS = 9; // I: GRUPOS
const COL_DNI_INSCRIPTO = 10; // J: DNI INSCRIPTO (CORREGIDO)
const COL_OBRA_SOCIAL = 11; // K: Obra Social (CORREGIDO)
const COL_COLEGIO_JARDIN = 12; // L: Colegio/Jardín (CORREGIDO)
const COL_ADULTO_RESPONSABLE_1 = 13; // M: Responsable 1
const COL_DNI_RESPONSABLE_1 = 14; // N: DNI Responsable 1
const COL_TEL_RESPONSABLE_1 = 15; // O: Teléfono Responsable 1
const COL_ADULTO_RESPONSABLE_2 = 16; // P: Responsable 2
const COL_TEL_RESPONSABLE_2 = 17; // Q: Teléfono Responsable 2
const COL_PERSONAS_AUTORIZADAS = 18; // R: Personas Autorizadas
const COL_PRACTICA_DEPORTE = 19; // S: Actividad extracurricular
const COL_ESPECIFIQUE_DEPORTE = 20; // T: Especifique cual actividad
const COL_TIENE_ENFERMEDAD = 21; // U: Enfermedades preexistentes
const COL_ESPECIFIQUE_ENFERMEDAD = 22; // V: Enfermedad medicaciones especiales
const COL_ES_ALERGICO = 23; // W: Alérgica/o
const COL_ESPECIFIQUE_ALERGIA = 24; // X: Especifique la Alergia
const COL_APTITUD_FISICA = 25; // Y: Certificado Aptitud Física
const COL_FOTO_CARNET = 26; // Z: Foto Carnet
const COL_JORNADA = 27; // AA: Jornada
const COL_SOCIO = 28; // AB: SOCIO?
const COL_METODO_PAGO = 29; // AC: Método de Pago
const COL_MODO_PAGO_CUOTA = 30; // AD: Modo Pago Cuotas
const COL_PRECIO = 31; // AE: Precio total
const COL_CUOTA_1 = 32; // AF: Cuota 1 (Valor)
const COL_CUOTA_2 = 33; // AG: Cuota 2 (Valor)
const COL_CUOTA_3 = 34; // AH: Cuota 3 (Valor)
const COL_CANTIDAD_CUOTAS = 35; // AI: Cant. Cuotas
const COL_ESTADO_PAGO = 36; // AJ: Estado de Pago
const COL_MONTO_A_PAGAR = 37; // AK: Se paga (Monto a pagar actual)
const COL_PAGADOR_NOMBRE_MANUAL = 38; // AL: Nombre y Apellido (Pagador)
const COL_PAGADOR_DNI_MANUAL = 39; // AM: Pagador (DNI)
const COL_COMPROBANTE_MANUAL_TOTAL_EXT = 40; // AN: Comprobante único
const COL_COMPROBANTE_MANUAL_CUOTA1 = 41; // AO: Comprobante C1
const COL_COMPROBANTE_MANUAL_CUOTA2 = 42; // AP: Comprobante C2
const COL_COMPROBANTE_MANUAL_CUOTA3 = 43; // AQ: Comprobante C3
const COL_VINCULO_PRINCIPAL = 44; // AR: VÍNCULO
const COL_ENVIAR_EMAIL_MANUAL = 45; // AS: Enviar Email (Checkbox)
const COL_ESTADO_EMAIL_ENVIADO = 46; // AT: False Envía/Email True Enviado
const COL_EDIT_TOKEN = 47; // AU: Token
// --- FIN DE LISTA (47 COLUMNAS) ---


// (Punto 25) CONSTANTES PARA LA NUEVA HOJA "Preventa" (Sin cambios)
const NOMBRE_HOJA_PREVENTA = 'PRE-VENTA';
const COL_PREVENTA_EMAIL = 3;       // Col C
const COL_PREVENTA_NOMBRE = 4;      // Col D
const COL_PREVENTA_APELLIDO = 5;    // Col E
const COL_PREVENTA_DNI = 6;         // Col F
const COL_PREVENTA_FECHA_NAC = 7;   // Col G
const COL_PREVENTA_GUARDA = 9;      // Col I
