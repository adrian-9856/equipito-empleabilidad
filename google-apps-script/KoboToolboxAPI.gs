/**
 * MÓDULO: INTEGRACIÓN CON KOBOTOOLBOX API
 *
 * Maneja toda la comunicación con la API de KoboToolbox
 * para obtener los datos de los formularios de graduados.
 */

/**
 * Obtiene los datos desde KoboToolbox usando la URL de exportación
 * @return {Array} Array de objetos con los datos de los graduados
 */
function obtenerDatosKoboToolbox() {
  try {
    const config = obtenerConfiguracion();

    if (!config.koboExportUrl) {
      throw new Error('URL de exportación de KoboToolbox no configurada');
    }

    // Preparar opciones para la petición HTTP
    const opciones = {
      method: 'get',
      headers: {},
      muteHttpExceptions: true
    };

    // Si hay token de autenticación, agregarlo
    if (config.koboToken) {
      opciones.headers['Authorization'] = 'Token ' + config.koboToken;
    }

    Logger.log('Obteniendo datos de KoboToolbox...');

    // Hacer la petición a KoboToolbox
    const respuesta = UrlFetchApp.fetch(config.koboExportUrl, opciones);
    const codigo = respuesta.getResponseCode();

    if (codigo !== 200) {
      throw new Error(`Error HTTP ${codigo}: ${respuesta.getContentText()}`);
    }

    // Parsear el CSV
    const csvData = respuesta.getContentText();
    const datosParseados = parsearCSV(csvData);

    Logger.log(`Se obtuvieron ${datosParseados.length} registros de KoboToolbox`);

    return datosParseados;

  } catch (error) {
    Logger.log('Error al obtener datos de KoboToolbox: ' + error);
    throw error;
  }
}

/**
 * Parsea el contenido CSV y lo convierte en array de objetos
 * @param {string} csvContent - Contenido del CSV
 * @return {Array} Array de objetos con los datos
 */
function parsearCSV(csvContent) {
  try {
    const lineas = csvContent.split('\n');

    if (lineas.length < 2) {
      return [];
    }

    // Primera línea son los headers
    const headers = parsearLineaCSV(lineas[0]);
    const datos = [];

    // Procesar cada línea de datos
    for (let i = 1; i < lineas.length; i++) {
      if (lineas[i].trim() === '') continue;

      const valores = parsearLineaCSV(lineas[i]);

      // Crear objeto con los datos
      const objeto = {};
      headers.forEach((header, index) => {
        objeto[header] = valores[index] || '';
      });

      datos.push(objeto);
    }

    return datos;

  } catch (error) {
    Logger.log('Error al parsear CSV: ' + error);
    throw error;
  }
}

/**
 * Parsea una línea de CSV respetando comillas y comas
 * @param {string} linea - Línea del CSV
 * @return {Array} Array con los valores
 */
function parsearLineaCSV(linea) {
  const valores = [];
  let valorActual = '';
  let dentroDeComillas = false;

  for (let i = 0; i < linea.length; i++) {
    const char = linea[i];

    if (char === '"') {
      dentroDeComillas = !dentroDeComillas;
    } else if (char === ',' && !dentroDeComillas) {
      valores.push(valorActual.trim());
      valorActual = '';
    } else {
      valorActual += char;
    }
  }

  // Agregar el último valor
  valores.push(valorActual.trim());

  return valores;
}

/**
 * Valida y limpia los datos obtenidos de KoboToolbox
 * @param {Object} dato - Objeto con datos de un graduado
 * @return {Object} Objeto validado y limpio
 */
function validarDatosGraduado(dato) {
  return {
    // Identificación
    id: dato.id || dato._id || '',
    nombre: dato.nombre || dato.name || '',

    // Datos de contacto
    telefono: dato.telefono || dato.phone || '',
    email: dato.email || dato.correo || '',

    // Datos de entrevista
    fechaEntrevista: dato.fecha_entrevista || dato.interview_date || '',
    formacion: dato.formacion || dato.training || '',

    // Datos adicionales
    habilidades: dato.habilidades || dato.skills || '',
    experiencia: dato.experiencia || dato.experience || '',

    // Metadatos
    fechaRegistro: dato._submission_time || new Date().toISOString(),
    estado: 'Nuevo',

    // Campos del flujo (inicialmente vacíos)
    clasificacion: '',
    empleado: false,
    seguimientoActual: '',
    proximaLlamada: ''
  };
}

/**
 * Verifica si un graduado ya existe en el sistema
 * @param {string} id - ID del graduado
 * @return {boolean} True si existe, false si no
 */
function graduadoExiste(id) {
  const hoja = obtenerHoja('Graduados');
  const datos = hoja.getDataRange().getValues();

  // Buscar en la columna de ID (asumiendo que está en la primera columna)
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === id) {
      return true;
    }
  }

  return false;
}

/**
 * Sincroniza datos periódicamente (puede ser usado con un trigger)
 */
function sincronizacionAutomatica() {
  try {
    Logger.log('Iniciando sincronización automática...');

    const datos = obtenerDatosKoboToolbox();
    const resultado = procesarDatosGraduados(datos);

    Logger.log(`Sincronización completada: ${resultado.nuevos} nuevos, ${resultado.total} total`);

    // Enviar notificación si hay nuevos graduados
    if (resultado.nuevos > 0) {
      enviarNotificacionNuevosGraduados(resultado.nuevos);
    }

  } catch (error) {
    Logger.log('Error en sincronización automática: ' + error);
    // Opcionalmente enviar email de error
  }
}

/**
 * Configura un trigger para sincronización automática cada hora
 */
function configurarSincronizacionAutomatica() {
  // Eliminar triggers existentes
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sincronizacionAutomatica') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Crear nuevo trigger cada hora
  ScriptApp.newTrigger('sincronizacionAutomatica')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Sincronización automática configurada para ejecutarse cada hora');
}
