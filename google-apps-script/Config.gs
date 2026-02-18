/**
 * MÓDULO: CONFIGURACIÓN DEL SISTEMA
 *
 * Gestiona la configuración y credenciales del sistema:
 * - Credenciales de KoboToolbox
 * - Configuración de notificaciones
 * - Preferencias del sistema
 */

/**
 * Verifica si la configuración está completa
 * @return {boolean} True si está configurado correctamente
 */
function verificarConfiguracion() {
  const config = obtenerConfiguracion();

  return !!(config.koboExportUrl && config.koboToken);
}

/**
 * Obtiene la configuración actual del sistema
 * @return {Object} Objeto con la configuración
 */
function obtenerConfiguracion() {
  const props = PropertiesService.getScriptProperties();

  return {
    koboExportUrl: props.getProperty('KOBO_EXPORT_URL') || '',
    koboToken: props.getProperty('KOBO_TOKEN') || '',
    emailNotificaciones: props.getProperty('EMAIL_NOTIFICACIONES') || '',
    sincronizacionAuto: props.getProperty('SINCRONIZACION_AUTO') === 'true'
  };
}

/**
 * Guarda la configuración de KoboToolbox
 * @param {string} exportUrl - URL de exportación CSV de KoboToolbox
 * @param {string} token - Token de autenticación de KoboToolbox
 * @return {boolean} True si se guardó correctamente
 */
function guardarConfiguracionKobo(exportUrl, token) {
  try {
    const props = PropertiesService.getScriptProperties();

    // Validar URL
    if (!exportUrl || !exportUrl.includes('kobotoolbox.org')) {
      throw new Error('URL de exportación inválida');
    }

    // Guardar configuración
    props.setProperty('KOBO_EXPORT_URL', exportUrl);
    props.setProperty('KOBO_TOKEN', token);

    Logger.log('Configuración de KoboToolbox guardada correctamente');

    return true;

  } catch (error) {
    Logger.log('Error al guardar configuración: ' + error);
    return false;
  }
}

/**
 * Guarda la configuración de notificaciones
 * @param {string} email - Email para recibir notificaciones
 * @param {boolean} activar - Activar o desactivar notificaciones automáticas
 */
function guardarConfiguracionNotificaciones(email, activar) {
  const props = PropertiesService.getScriptProperties();

  props.setProperty('EMAIL_NOTIFICACIONES', email);

  if (activar) {
    configurarNotificacionesAutomaticas(email);
  } else {
    desactivarNotificacionesAutomaticas();
  }

  Logger.log(`Notificaciones ${activar ? 'activadas' : 'desactivadas'} para: ${email}`);
}

/**
 * Desactiva las notificaciones automáticas
 */
function desactivarNotificacionesAutomaticas() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'enviarNotificacionesDiarias') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('Notificaciones automáticas desactivadas');
}

/**
 * Activa o desactiva la sincronización automática con KoboToolbox
 * @param {boolean} activar - True para activar, false para desactivar
 */
function configurarSincronizacionAutomatica(activar) {
  const props = PropertiesService.getScriptProperties();

  props.setProperty('SINCRONIZACION_AUTO', activar.toString());

  if (activar) {
    configurarSincronizacionAutomatica();
  } else {
    desactivarSincronizacionAutomatica();
  }

  Logger.log(`Sincronización automática ${activar ? 'activada' : 'desactivada'}`);
}

/**
 * Desactiva la sincronización automática
 */
function desactivarSincronizacionAutomatica() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sincronizacionAutomatica') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('Sincronización automática desactivada');
}

/**
 * Obtiene todas las configuraciones para mostrar en la UI
 * @return {Object} Objeto con todas las configuraciones
 */
function obtenerTodasLasConfiguraciones() {
  const config = obtenerConfiguracion();

  // Ocultar token parcialmente por seguridad
  if (config.koboToken) {
    config.koboTokenOculto = config.koboToken.substring(0, 4) + '***' + config.koboToken.substring(config.koboToken.length - 4);
  }

  return config;
}

/**
 * Reinicia toda la configuración del sistema
 */
function reiniciarConfiguracion() {
  const props = PropertiesService.getScriptProperties();

  // Eliminar todas las propiedades
  props.deleteAllProperties();

  // Eliminar todos los triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  Logger.log('Configuración reiniciada completamente');
}

/**
 * Exporta la configuración actual (sin token por seguridad)
 * @return {string} JSON con la configuración
 */
function exportarConfiguracion() {
  const config = obtenerConfiguracion();

  // Remover token por seguridad
  delete config.koboToken;

  return JSON.stringify(config, null, 2);
}

/**
 * Prueba la conexión con KoboToolbox
 * @return {Object} Resultado de la prueba
 */
function probarConexionKobo() {
  try {
    const config = obtenerConfiguracion();

    if (!config.koboExportUrl || !config.koboToken) {
      return {
        exito: false,
        mensaje: 'Configuración incompleta. Por favor configura la URL y el token.'
      };
    }

    const opciones = {
      method: 'get',
      headers: {
        'Authorization': 'Token ' + config.koboToken
      },
      muteHttpExceptions: true
    };

    const respuesta = UrlFetchApp.fetch(config.koboExportUrl, opciones);
    const codigo = respuesta.getResponseCode();

    if (codigo === 200) {
      return {
        exito: true,
        mensaje: 'Conexión exitosa con KoboToolbox'
      };
    } else {
      return {
        exito: false,
        mensaje: `Error de conexión. Código: ${codigo}`
      };
    }

  } catch (error) {
    return {
      exito: false,
      mensaje: 'Error: ' + error.message
    };
  }
}

/**
 * Guarda configuración desde el formulario HTML
 * Esta función es llamada desde el frontend
 */
function guardarConfiguracionDesdeFormulario(datos) {
  try {
    // Guardar configuración de Kobo
    if (datos.koboExportUrl && datos.koboToken) {
      guardarConfiguracionKobo(datos.koboExportUrl, datos.koboToken);
    }

    // Guardar configuración de notificaciones
    if (datos.emailNotificaciones) {
      guardarConfiguracionNotificaciones(datos.emailNotificaciones, datos.activarNotificaciones);
    }

    // Configurar sincronización automática
    if (datos.sincronizacionAuto !== undefined) {
      configurarSincronizacionAutomatica(datos.sincronizacionAuto);
    }

    return {
      exito: true,
      mensaje: 'Configuración guardada correctamente'
    };

  } catch (error) {
    return {
      exito: false,
      mensaje: 'Error al guardar: ' + error.message
    };
  }
}
