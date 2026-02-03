/**
 * MÓDULO: UTILIDADES GENERALES
 *
 * Funciones auxiliares y utilidades comunes usadas en todo el sistema.
 */

/**
 * Formatea una fecha al formato español (dd/mm/yyyy)
 * @param {Date} fecha - Objeto Date
 * @return {string} Fecha formateada
 */
function formatearFecha(fecha) {
  if (!(fecha instanceof Date)) {
    fecha = new Date(fecha);
  }

  const dia = String(fecha.getDate()).padStart(2, '0');
  const mes = String(fecha.getMonth() + 1).padStart(2, '0');
  const año = fecha.getFullYear();

  return `${dia}/${mes}/${año}`;
}

/**
 * Formatea una fecha con hora
 * @param {Date} fecha - Objeto Date
 * @return {string} Fecha y hora formateadas
 */
function formatearFechaHora(fecha) {
  if (!(fecha instanceof Date)) {
    fecha = new Date(fecha);
  }

  const fechaStr = formatearFecha(fecha);
  const hora = String(fecha.getHours()).padStart(2, '0');
  const minutos = String(fecha.getMinutes()).padStart(2, '0');

  return `${fechaStr} ${hora}:${minutos}`;
}

/**
 * Valida un email
 * @param {string} email - Email a validar
 * @return {boolean} True si es válido
 */
function validarEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

/**
 * Valida un número de teléfono (formato flexible)
 * @param {string} telefono - Teléfono a validar
 * @return {boolean} True si es válido
 */
function validarTelefono(telefono) {
  // Eliminar espacios y guiones
  const telLimpio = telefono.replace(/[\s\-]/g, '');

  // Debe tener entre 8 y 15 dígitos
  return /^\+?\d{8,15}$/.test(telLimpio);
}

/**
 * Limpia y normaliza un string
 * @param {string} str - String a limpiar
 * @return {string} String limpio
 */
function limpiarString(str) {
  if (!str) return '';

  return str.toString()
    .trim()
    .replace(/\s+/g, ' '); // Reemplazar múltiples espacios por uno solo
}

/**
 * Convierte un string a Title Case (Primera Letra Mayúscula)
 * @param {string} str - String a convertir
 * @return {string} String en Title Case
 */
function toTitleCase(str) {
  if (!str) return '';

  return str.toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
}

/**
 * Genera un ID único
 * @return {string} ID único
 */
function generarIdUnico() {
  return Utilities.getUuid();
}

/**
 * Envía una notificación por email con formato HTML
 * @param {string} destinatario - Email del destinatario
 * @param {string} asunto - Asunto del email
 * @param {string} mensaje - Mensaje en HTML
 */
function enviarEmail(destinatario, asunto, mensaje) {
  try {
    if (!validarEmail(destinatario)) {
      Logger.log('Email inválido: ' + destinatario);
      return false;
    }

    MailApp.sendEmail({
      to: destinatario,
      subject: asunto,
      htmlBody: mensaje
    });

    Logger.log(`Email enviado a: ${destinatario}`);
    return true;

  } catch (error) {
    Logger.log('Error al enviar email: ' + error);
    return false;
  }
}

/**
 * Envía notificación de nuevos graduados importados
 * @param {number} cantidad - Cantidad de graduados nuevos
 */
function enviarNotificacionNuevosGraduados(cantidad) {
  const config = obtenerConfiguracion();

  if (!config.emailNotificaciones) {
    return;
  }

  const mensaje = `
    <h2>📊 Nuevos Graduados Importados</h2>
    <p>Se han importado <strong>${cantidad}</strong> graduado(s) nuevo(s) desde KoboToolbox.</p>
    <p>Por favor, revisa y clasifica los nuevos graduados en el sistema.</p>
    <p><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}">Ir a la Hoja de Seguimiento</a></p>
  `;

  enviarEmail(
    config.emailNotificaciones,
    `📊 ${cantidad} Nuevo(s) Graduado(s) - Equipito Empleabilidad`,
    mensaje
  );
}

/**
 * Registra un evento en el log del sistema
 * @param {string} tipo - Tipo de evento
 * @param {string} descripcion - Descripción del evento
 * @param {Object} datos - Datos adicionales (opcional)
 */
function registrarEvento(tipo, descripcion, datos = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaLog = ss.getSheetByName('Log del Sistema');

  if (!hojaLog) {
    hojaLog = ss.insertSheet('Log del Sistema');

    // Crear headers
    const headers = ['Fecha/Hora', 'Tipo', 'Descripción', 'Usuario', 'Datos'];
    hojaLog.appendRow(headers);

    const headerRange = hojaLog.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#607d8b');
    headerRange.setFontColor('#ffffff');
  }

  const fila = [
    formatearFechaHora(new Date()),
    tipo,
    descripcion,
    Session.getActiveUser().getEmail(),
    JSON.stringify(datos)
  ];

  hojaLog.appendRow(fila);
}

/**
 * Crea un backup de una hoja específica
 * @param {string} nombreHoja - Nombre de la hoja a respaldar
 * @return {string} Nombre de la hoja de backup creada
 */
function crearBackupHoja(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);

  if (!hoja) {
    throw new Error(`Hoja no encontrada: ${nombreHoja}`);
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const nombreBackup = `${nombreHoja}_backup_${timestamp}`;

  const hojaBackup = hoja.copyTo(ss);
  hojaBackup.setName(nombreBackup);

  Logger.log(`Backup creado: ${nombreBackup}`);

  return nombreBackup;
}

/**
 * Obtiene estadísticas generales del sistema
 * @return {Object} Objeto con estadísticas
 */
function obtenerEstadisticasGenerales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const stats = {
    totalGraduados: 0,
    porClasificacion: {},
    empleados: 0,
    seguimientosPendientes: 0
  };

  // Total de graduados
  const hojaGraduados = ss.getSheetByName('Graduados');
  if (hojaGraduados) {
    stats.totalGraduados = hojaGraduados.getLastRow() - 1; // -1 por header

    // Contar por clasificación
    const datos = hojaGraduados.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const clasificacion = datos[i][10]; // Columna de clasificación

      if (clasificacion) {
        stats.porClasificacion[clasificacion] = (stats.porClasificacion[clasificacion] || 0) + 1;
      }

      if (datos[i][11] === 'Sí') { // Columna empleado
        stats.empleados++;
      }
    }
  }

  // Seguimientos pendientes
  const seguimientos = obtenerSeguimientosPendientes();
  stats.seguimientosPendientes = seguimientos.length;

  return stats;
}

/**
 * Formatea un número con separadores de miles
 * @param {number} numero - Número a formatear
 * @return {string} Número formateado
 */
function formatearNumero(numero) {
  return numero.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

/**
 * Calcula la diferencia en días entre dos fechas
 * @param {Date} fecha1 - Primera fecha
 * @param {Date} fecha2 - Segunda fecha
 * @return {number} Diferencia en días
 */
function calcularDiferenciaDias(fecha1, fecha2) {
  const unDia = 24 * 60 * 60 * 1000; // milisegundos en un día
  return Math.round(Math.abs((fecha1 - fecha2) / unDia));
}

/**
 * Verifica si una cadena está vacía o solo contiene espacios
 * @param {string} str - Cadena a verificar
 * @return {boolean} True si está vacía
 */
function estaVacio(str) {
  return !str || str.trim().length === 0;
}

/**
 * Obtiene el color asociado a un estado
 * @param {string} estado - Estado del graduado
 * @return {string} Código de color hexadecimal
 */
function obtenerColorEstado(estado) {
  const colores = {
    'Nuevo': '#fff3e0',
    'En Proceso': '#e3f2fd',
    'Activo': '#e8f5e9',
    'Empleado': '#c8e6c9',
    'Completado': '#f1f8e9',
    'Inactivo': '#ffebee'
  };

  return colores[estado] || '#ffffff';
}

/**
 * Convierte un array a formato CSV
 * @param {Array} datos - Array bidimensional con los datos
 * @return {string} Contenido CSV
 */
function arrayACSV(datos) {
  return datos.map(fila =>
    fila.map(celda => {
      // Escapar comillas y envolver en comillas si contiene comas
      const valor = String(celda).replace(/"/g, '""');
      return valor.includes(',') ? `"${valor}"` : valor;
    }).join(',')
  ).join('\n');
}

/**
 * Muestra un mensaje de progreso al usuario
 * @param {string} mensaje - Mensaje a mostrar
 */
function mostrarProgreso(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, 'Procesando...', 3);
}

/**
 * Muestra un mensaje de éxito al usuario
 * @param {string} mensaje - Mensaje a mostrar
 */
function mostrarExito(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, '✅ Éxito', 3);
}

/**
 * Muestra un mensaje de error al usuario
 * @param {string} mensaje - Mensaje a mostrar
 */
function mostrarError(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, '❌ Error', 5);
}
