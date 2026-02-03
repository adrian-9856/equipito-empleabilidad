/**
 * MÓDULO: SEGUIMIENTO Y LLAMADAS PROGRAMADAS
 *
 * Gestiona el sistema de seguimientos automáticos para graduados empleados:
 * - Llamada 1: Semanal (primera semana)
 * - Llamada 2: 3 meses después
 * - Llamada 3: 6 meses después
 */

/**
 * Programa los seguimientos automáticos para un graduado empleado
 * @param {string} graduadoId - ID del graduado
 * @param {string} nombreGraduado - Nombre del graduado
 */
function programarSeguimientos(graduadoId, nombreGraduado) {
  const hojaSeguimientos = obtenerHoja('Seguimientos');

  // Crear headers si es necesario
  if (hojaSeguimientos.getLastRow() === 0) {
    const headers = [
      'ID Graduado',
      'Nombre',
      'Tipo Seguimiento',
      'Fecha Programada',
      'Fecha Realizada',
      'Estado',
      'Resultado',
      'Notas',
      'Próximo Paso'
    ];
    hojaSeguimientos.appendRow(headers);

    // Formatear headers
    const headerRange = hojaSeguimientos.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#673ab7');
    headerRange.setFontColor('#ffffff');
  }

  const fechaBase = new Date();

  // Crear los tres tipos de seguimiento
  const seguimientos = [
    {
      tipo: 'Llamada 1 - Semanal',
      diasDespues: 7,
      descripcion: 'Primera semana de trabajo'
    },
    {
      tipo: 'Llamada 2 - 3 Meses',
      diasDespues: 90,
      descripcion: 'Seguimiento a los 3 meses'
    },
    {
      tipo: 'Llamada 3 - 6 Meses',
      diasDespues: 180,
      descripcion: 'Seguimiento a los 6 meses'
    }
  ];

  seguimientos.forEach(seg => {
    const fechaProgramada = new Date(fechaBase);
    fechaProgramada.setDate(fechaProgramada.getDate() + seg.diasDespues);

    const fila = [
      graduadoId,
      nombreGraduado,
      seg.tipo,
      fechaProgramada.toLocaleDateString('es-ES'),
      '', // Fecha realizada (vacío)
      'Pendiente',
      '', // Resultado (vacío)
      seg.descripcion,
      '' // Próximo paso (vacío)
    ];

    hojaSeguimientos.appendRow(fila);
  });

  Logger.log(`Seguimientos programados para: ${nombreGraduado}`);

  // Actualizar en hoja de Graduados
  actualizarProximaLlamada(graduadoId, 'Llamada 1 - Semanal');
}

/**
 * Obtiene los seguimientos pendientes para hoy o fechas pasadas
 * @return {Array} Array de objetos con seguimientos pendientes
 */
function obtenerSeguimientosPendientes() {
  const hojaSeguimientos = obtenerHoja('Seguimientos');
  const datos = hojaSeguimientos.getDataRange().getValues();

  if (datos.length <= 1) {
    return [];
  }

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const pendientes = [];

  for (let i = 1; i < datos.length; i++) {
    const estado = datos[i][5]; // Columna Estado
    const fechaProgramadaStr = datos[i][3];

    if (estado === 'Pendiente' && fechaProgramadaStr) {
      const fechaProgramada = parsearFecha(fechaProgramadaStr);

      if (fechaProgramada && fechaProgramada <= hoy) {
        pendientes.push({
          fila: i + 1,
          id: datos[i][0],
          nombre: datos[i][1],
          tipo: datos[i][2],
          fechaProgramada: fechaProgramadaStr,
          notas: datos[i][7]
        });
      }
    }
  }

  return pendientes;
}

/**
 * Marca un seguimiento como realizado
 * @param {number} fila - Número de fila en la hoja de seguimientos
 * @param {string} resultado - Resultado del seguimiento
 * @param {string} notas - Notas adicionales
 * @param {string} proximoPaso - Próximo paso a seguir
 */
function marcarSeguimientoRealizado(fila, resultado, notas, proximoPaso) {
  const hojaSeguimientos = obtenerHoja('Seguimientos');

  // Actualizar campos
  hojaSeguimientos.getRange(fila, 5).setValue(new Date().toLocaleDateString('es-ES')); // Fecha realizada
  hojaSeguimientos.getRange(fila, 6).setValue('Realizado'); // Estado
  hojaSeguimientos.getRange(fila, 7).setValue(resultado); // Resultado
  hojaSeguimientos.getRange(fila, 8).setValue(notas); // Notas
  hojaSeguimientos.getRange(fila, 9).setValue(proximoPaso); // Próximo paso

  // Colorear la fila en verde
  const rangoFila = hojaSeguimientos.getRange(fila, 1, 1, 9);
  rangoFila.setBackground('#d9ead3');

  Logger.log(`Seguimiento marcado como realizado en fila ${fila}`);
}

/**
 * Actualiza la próxima llamada programada en la hoja de Graduados
 * @param {string} graduadoId - ID del graduado
 * @param {string} tipoLlamada - Tipo de la próxima llamada
 */
function actualizarProximaLlamada(graduadoId, tipoLlamada) {
  const hojaGraduados = obtenerHoja('Graduados');
  const datos = hojaGraduados.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === graduadoId) {
      hojaGraduados.getRange(i + 1, 13).setValue(tipoLlamada); // Columna Próxima Llamada
      break;
    }
  }
}

/**
 * Genera un reporte de seguimientos realizados
 * @param {Date} fechaInicio - Fecha de inicio del reporte
 * @param {Date} fechaFin - Fecha de fin del reporte
 * @return {Object} Objeto con estadísticas del reporte
 */
function generarReporteSeguimientos(fechaInicio, fechaFin) {
  const hojaSeguimientos = obtenerHoja('Seguimientos');
  const datos = hojaSeguimientos.getDataRange().getValues();

  const stats = {
    total: 0,
    realizados: 0,
    pendientes: 0,
    vencidos: 0,
    porTipo: {}
  };

  for (let i = 1; i < datos.length; i++) {
    const estado = datos[i][5];
    const tipo = datos[i][2];
    const fechaProgramada = parsearFecha(datos[i][3]);

    if (!fechaProgramada) continue;

    // Filtrar por rango de fechas
    if (fechaProgramada >= fechaInicio && fechaProgramada <= fechaFin) {
      stats.total++;

      if (estado === 'Realizado') {
        stats.realizados++;
      } else if (estado === 'Pendiente') {
        stats.pendientes++;

        // Verificar si está vencido
        if (fechaProgramada < new Date()) {
          stats.vencidos++;
        }
      }

      // Conteo por tipo
      if (!stats.porTipo[tipo]) {
        stats.porTipo[tipo] = 0;
      }
      stats.porTipo[tipo]++;
    }
  }

  return stats;
}

/**
 * Envía notificaciones por email de seguimientos pendientes
 * @param {string} emailDestinatario - Email del responsable
 */
function enviarNotificacionesSeguimientos(emailDestinatario) {
  const pendientes = obtenerSeguimientosPendientes();

  if (pendientes.length === 0) {
    return;
  }

  let mensaje = '<h2>Seguimientos Pendientes - Equipito Empleabilidad</h2>';
  mensaje += `<p>Tienes <strong>${pendientes.length}</strong> seguimiento(s) pendiente(s):</p>`;
  mensaje += '<ul>';

  pendientes.forEach(seg => {
    mensaje += `<li><strong>${seg.nombre}</strong> - ${seg.tipo} (${seg.fechaProgramada})</li>`;
  });

  mensaje += '</ul>';
  mensaje += '<p>Por favor, realiza estos seguimientos lo antes posible.</p>';

  MailApp.sendEmail({
    to: emailDestinatario,
    subject: `⏰ ${pendientes.length} Seguimiento(s) Pendiente(s) - Equipito Empleabilidad`,
    htmlBody: mensaje
  });

  Logger.log(`Notificación enviada a: ${emailDestinatario}`);
}

/**
 * Configura notificaciones automáticas diarias
 * @param {string} emailDestinatario - Email del responsable
 */
function configurarNotificacionesAutomaticas(emailDestinatario) {
  // Eliminar triggers existentes
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'enviarNotificacionesDiarias') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Crear nuevo trigger diario a las 9 AM
  ScriptApp.newTrigger('enviarNotificacionesDiarias')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  // Guardar email en propiedades
  PropertiesService.getScriptProperties().setProperty('emailNotificaciones', emailDestinatario);

  Logger.log('Notificaciones automáticas configuradas para las 9 AM');
}

/**
 * Función que se ejecuta diariamente para enviar notificaciones
 */
function enviarNotificacionesDiarias() {
  const email = PropertiesService.getScriptProperties().getProperty('emailNotificaciones');

  if (email) {
    enviarNotificacionesSeguimientos(email);
  }
}

/**
 * Parsea una fecha en formato español a objeto Date
 * @param {string} fechaStr - Fecha en formato dd/mm/yyyy
 * @return {Date} Objeto Date o null si no es válida
 */
function parsearFecha(fechaStr) {
  if (!fechaStr) return null;

  try {
    const partes = fechaStr.split('/');
    if (partes.length === 3) {
      return new Date(partes[2], partes[1] - 1, partes[0]);
    }
  } catch (e) {
    Logger.log('Error parseando fecha: ' + fechaStr);
  }

  return null;
}

/**
 * Genera HTML con la lista de seguimientos pendientes
 * @param {Array} seguimientos - Array de seguimientos
 * @return {string} HTML formateado
 */
function generarHTMLSeguimientos(seguimientos) {
  let html = '<style>';
  html += 'body { font-family: Arial, sans-serif; padding: 20px; }';
  html += 'table { width: 100%; border-collapse: collapse; }';
  html += 'th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }';
  html += 'th { background-color: #673ab7; color: white; }';
  html += 'tr:hover { background-color: #f5f5f5; }';
  html += '</style>';

  html += '<h2>📞 Seguimientos Pendientes</h2>';
  html += `<p>Total: <strong>${seguimientos.length}</strong> seguimiento(s)</p>`;

  html += '<table>';
  html += '<tr><th>Nombre</th><th>Tipo</th><th>Fecha Programada</th><th>Notas</th></tr>';

  seguimientos.forEach(seg => {
    html += '<tr>';
    html += `<td><strong>${seg.nombre}</strong></td>`;
    html += `<td>${seg.tipo}</td>`;
    html += `<td>${seg.fechaProgramada}</td>`;
    html += `<td>${seg.notas}</td>`;
    html += '</tr>';
  });

  html += '</table>';

  return html;
}
