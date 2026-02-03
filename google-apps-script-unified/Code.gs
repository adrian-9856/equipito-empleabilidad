/**
 * ═══════════════════════════════════════════════════════════════════════════
 * SISTEMA DE SEGUIMIENTO DE GRADUADOS - EQUIPITO EMPLEABILIDAD
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * Sistema completo de seguimiento de graduados con integración a KoboToolbox
 * Archivo unificado con todas las funcionalidades
 *
 * Autor: Equipito Empleabilidad
 * Fecha: 2026-02-03
 * ═══════════════════════════════════════════════════════════════════════════
 */


// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 1: MENÚ PRINCIPAL Y FUNCIONES DE INTERFAZ
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Función que se ejecuta al abrir la hoja de cálculo
 * Crea el menú personalizado en la interfaz
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Seguimiento Graduados')
    .addItem('🔄 Importar desde KoboToolbox', 'importarDatosKobo')
    .addSeparator()
    .addItem('📝 Clasificar Graduados', 'mostrarFormularioClasificacion')
    .addItem('📞 Ver Seguimientos Pendientes', 'mostrarSeguimientosPendientes')
    .addSeparator()
    .addItem('⚙️ Configurar Credenciales', 'mostrarConfiguracion')
    .addItem('📋 Crear Estructura de Hojas', 'crearEstructuraHojas')
    .addToUi();
}

/**
 * Función principal para importar datos desde KoboToolbox
 */
function importarDatosKobo() {
  try {
    const ui = SpreadsheetApp.getUi();

    // Verificar configuración
    if (!verificarConfiguracion()) {
      ui.alert('⚠️ Configuración Incompleta',
               'Por favor configura las credenciales de KoboToolbox primero.\n' +
               'Ve a: Seguimiento Graduados > Configurar Credenciales',
               ui.ButtonSet.OK);
      return;
    }

    ui.alert('🔄 Importando Datos',
             'Iniciando importación desde KoboToolbox...\n' +
             'Esto puede tardar unos momentos.',
             ui.ButtonSet.OK);

    // Obtener datos de KoboToolbox
    const datos = obtenerDatosKoboToolbox();

    if (!datos || datos.length === 0) {
      ui.alert('⚠️ Sin Datos',
               'No se encontraron datos nuevos para importar.',
               ui.ButtonSet.OK);
      return;
    }

    // Procesar y almacenar datos
    const resultado = procesarDatosGraduados(datos);

    ui.alert('✅ Importación Completada',
             `Se importaron ${resultado.nuevos} graduados nuevos.\n` +
             `Total de graduados: ${resultado.total}`,
             ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error',
                                  'Error al importar datos: ' + error.message,
                                  SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error en importarDatosKobo: ' + error);
  }
}

/**
 * Muestra el formulario para clasificar graduados según el flujo
 */
function mostrarFormularioClasificacion() {
  const html = HtmlService.createHtmlOutputFromFile('FormularioClasificacion')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Clasificar Graduados');

  SpreadsheetApp.getUi().showModalDialog(html, 'Clasificación de Graduados');
}

/**
 * Muestra los seguimientos pendientes (llamadas programadas)
 */
function mostrarSeguimientosPendientes() {
  const seguimientos = obtenerSeguimientosPendientes();

  if (seguimientos.length === 0) {
    SpreadsheetApp.getUi().alert('✅ Todo al Día',
                                  'No hay seguimientos pendientes por el momento.',
                                  SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const html = HtmlService.createHtmlOutput(generarHTMLSeguimientos(seguimientos))
    .setWidth(700)
    .setHeight(500)
    .setTitle('Seguimientos Pendientes');

  SpreadsheetApp.getUi().showModalDialog(html, 'Seguimientos Pendientes');
}

/**
 * Inicializa la configuración del proyecto
 */
function mostrarConfiguracion() {
  const html = HtmlService.createHtmlOutputFromFile('ConfiguracionKobo')
    .setWidth(500)
    .setHeight(400)
    .setTitle('Configuración de KoboToolbox');

  SpreadsheetApp.getUi().showModalDialog(html, 'Configuración');
}

/**
 * Crea la estructura completa de hojas necesarias para el sistema
 */
function crearEstructuraHojas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const respuesta = ui.alert('Crear Estructura de Hojas',
                                '¿Deseas crear todas las hojas necesarias para el sistema?\n\n' +
                                'Se crearán las siguientes hojas:\n' +
                                '• Graduados (datos generales)\n' +
                                '• Entrevista/Seguimiento General\n' +
                                '• Aliados\n' +
                                '• Plataforma\n' +
                                '• Conexiones Laborales\n' +
                                '• Por su Cuenta\n' +
                                '• Busca Trabajo (Fito)\n' +
                                '• Empleados\n' +
                                '• Seguimientos (llamadas programadas)\n' +
                                '• Reportes Mensuales\n' +
                                '• Configuración',
                                ui.ButtonSet.YES_NO);

    if (respuesta !== ui.Button.YES) {
      return;
    }

    // Crear todas las hojas necesarias según el flujo
    const hojasNecesarias = [
      'Graduados',
      'Entrevista/Seguimiento General',
      'Aliados',
      'Plataforma',
      'Conexiones Laborales',
      'Por su Cuenta',
      'Busca Trabajo (Fito)',
      'Empleados',
      'Seguimientos',
      'Reportes Mensuales',
      'Configuración'
    ];

    let hojasCreadas = 0;

    hojasNecesarias.forEach(nombreHoja => {
      try {
        let hoja = ss.getSheetByName(nombreHoja);
        if (!hoja) {
          hoja = ss.insertSheet(nombreHoja);
          configurarHoja(hoja, nombreHoja);
          hojasCreadas++;
        }
      } catch (e) {
        Logger.log(`Error al crear hoja ${nombreHoja}: ${e}`);
      }
    });

    ui.alert('✅ Estructura Creada',
             `Se crearon ${hojasCreadas} hojas nuevas.\n\n` +
             'El sistema está listo para usar.',
             ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error',
                                  'Error al crear estructura: ' + error.message,
                                  SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error en crearEstructuraHojas: ' + error);
  }
}

/**
 * Obtiene la lista de graduados que aún no han sido clasificados
 * @return {Array} Array de objetos con información de graduados sin clasificar
 */
function obtenerGraduadosSinClasificar() {
  try {
    const hoja = obtenerHoja('Graduados');
    const datos = hoja.getDataRange().getValues();

    if (datos.length <= 1) {
      return [];
    }

    const graduadosSinClasificar = [];

    // Columna 10 (índice 10) es la de clasificación
    for (let i = 1; i < datos.length; i++) {
      const clasificacion = datos[i][10];

      if (!clasificacion || clasificacion.trim() === '') {
        graduadosSinClasificar.push({
          id: datos[i][0],
          nombre: datos[i][1],
          email: datos[i][3]
        });
      }
    }

    return graduadosSinClasificar;

  } catch (error) {
    Logger.log('Error al obtener graduados sin clasificar: ' + error);
    return [];
  }
}


// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 2: INTEGRACIÓN CON KOBOTOOLBOX API
// ═══════════════════════════════════════════════════════════════════════════

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


// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 3: GESTIÓN DE HOJAS Y DATOS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Procesa los datos de graduados y los almacena en la hoja correspondiente
 * @param {Array} datos - Array de objetos con datos de graduados
 * @return {Object} Objeto con resultado del procesamiento
 */
function procesarDatosGraduados(datos) {
  const hojaGraduados = obtenerHoja('Graduados');
  let nuevos = 0;
  let actualizados = 0;

  datos.forEach(dato => {
    const graduadoValidado = validarDatosGraduado(dato);

    if (!graduadoExiste(graduadoValidado.id)) {
      agregarGraduado(graduadoValidado);
      nuevos++;
    } else {
      actualizarGraduado(graduadoValidado);
      actualizados++;
    }
  });

  const total = hojaGraduados.getLastRow() - 1; // -1 por el header

  return {
    nuevos: nuevos,
    actualizados: actualizados,
    total: total
  };
}

/**
 * Agrega un nuevo graduado a la hoja de Graduados
 * @param {Object} graduado - Datos del graduado
 */
function agregarGraduado(graduado) {
  const hoja = obtenerHoja('Graduados');

  // Si la hoja está vacía, agregar headers
  if (hoja.getLastRow() === 0) {
    const headers = [
      'ID',
      'Nombre',
      'Teléfono',
      'Email',
      'Fecha Entrevista',
      'Formación',
      'Habilidades',
      'Experiencia',
      'Fecha Registro',
      'Estado',
      'Clasificación',
      'Empleado',
      'Seguimiento Actual',
      'Próxima Llamada',
      'Notas'
    ];
    hoja.appendRow(headers);

    // Formatear headers
    const headerRange = hoja.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
  }

  // Agregar el graduado
  const fila = [
    graduado.id,
    graduado.nombre,
    graduado.telefono,
    graduado.email,
    graduado.fechaEntrevista,
    graduado.formacion,
    graduado.habilidades,
    graduado.experiencia,
    graduado.fechaRegistro,
    graduado.estado,
    graduado.clasificacion,
    graduado.empleado ? 'Sí' : 'No',
    graduado.seguimientoActual,
    graduado.proximaLlamada,
    ''
  ];

  hoja.appendRow(fila);
  Logger.log(`Graduado agregado: ${graduado.nombre} (${graduado.id})`);
}

/**
 * Actualiza los datos de un graduado existente
 * @param {Object} graduado - Datos actualizados del graduado
 */
function actualizarGraduado(graduado) {
  const hoja = obtenerHoja('Graduados');
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === graduado.id) {
      // Actualizar solo campos que pueden cambiar
      hoja.getRange(i + 1, 2).setValue(graduado.nombre);
      hoja.getRange(i + 1, 3).setValue(graduado.telefono);
      hoja.getRange(i + 1, 4).setValue(graduado.email);

      Logger.log(`Graduado actualizado: ${graduado.nombre} (${graduado.id})`);
      break;
    }
  }
}

/**
 * Clasifica un graduado y lo mueve a la hoja correspondiente
 * @param {string} graduadoId - ID del graduado
 * @param {string} clasificacion - Tipo de clasificación del flujo
 * @param {Object} datosAdicionales - Datos adicionales según el tipo
 */
function clasificarGraduado(graduadoId, clasificacion, datosAdicionales = {}) {
  const hojaGraduados = obtenerHoja('Graduados');
  const datos = hojaGraduados.getDataRange().getValues();

  // Buscar el graduado
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === graduadoId) {
      // Actualizar clasificación en hoja principal
      hojaGraduados.getRange(i + 1, 11).setValue(clasificacion); // Columna de clasificación

      // Copiar a la hoja correspondiente según el flujo
      copiarAHojaClasificacion(datos[i], clasificacion, datosAdicionales);

      // Si la clasificación es "Empleado", programar seguimientos
      if (clasificacion === 'Empleado') {
        programarSeguimientos(graduadoId, datos[i][1]); // nombre en columna 1
      }

      Logger.log(`Graduado ${datos[i][1]} clasificado como: ${clasificacion}`);
      break;
    }
  }
}

/**
 * Copia los datos del graduado a la hoja de su clasificación
 * @param {Array} datosGraduado - Fila con datos del graduado
 * @param {string} clasificacion - Tipo de clasificación
 * @param {Object} datosAdicionales - Datos adicionales
 */
function copiarAHojaClasificacion(datosGraduado, clasificacion, datosAdicionales) {
  const nombreHoja = obtenerNombreHojaClasificacion(clasificacion);
  const hoja = obtenerHoja(nombreHoja);

  // Crear headers si la hoja está vacía
  if (hoja.getLastRow() === 0) {
    const headers = crearHeadersParaClasificacion(clasificacion);
    hoja.appendRow(headers);

    // Formatear headers
    const headerRange = hoja.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground(obtenerColorClasificacion(clasificacion));
    headerRange.setFontColor('#ffffff');
  }

  // Preparar fila con datos
  const fila = prepararFilaClasificacion(datosGraduado, clasificacion, datosAdicionales);
  hoja.appendRow(fila);
}

/**
 * Obtiene el nombre de la hoja según la clasificación
 * @param {string} clasificacion - Tipo de clasificación
 * @return {string} Nombre de la hoja
 */
function obtenerNombreHojaClasificacion(clasificacion) {
  const mapeo = {
    'Entrevista/Seguimiento General': 'Entrevista/Seguimiento General',
    'Aliados': 'Aliados',
    'Plataforma': 'Plataforma',
    'Conexiones Laborales': 'Conexiones Laborales',
    'Por su Cuenta': 'Por su Cuenta',
    'Busca Trabajo': 'Busca Trabajo (Fito)',
    'Fito': 'Busca Trabajo (Fito)',
    'Empleado': 'Empleados'
  };

  return mapeo[clasificacion] || 'Graduados';
}

/**
 * Crea los headers apropiados según la clasificación
 * @param {string} clasificacion - Tipo de clasificación
 * @return {Array} Array con los nombres de las columnas
 */
function crearHeadersParaClasificacion(clasificacion) {
  const headersBase = [
    'ID',
    'Nombre',
    'Teléfono',
    'Email',
    'Fecha Clasificación',
    'Estado'
  ];

  const headersEspecificos = {
    'Entrevista/Seguimiento General': ['Tipo Entrevista', 'Resultado', 'Siguiente Paso', 'Notas'],
    'Aliados': ['Aliado Asignado', 'Contacto del Aliado', 'Fecha Derivación', 'Notas'],
    'Plataforma': ['Plataforma Asignada', 'Usuario', 'Fecha Registro', 'Notas'],
    'Conexiones Laborales': ['Contacto', 'Empresa/Área', 'Tipo Conexión', 'Notas'],
    'Por su Cuenta': ['Actividad', 'Progreso', 'Última Actualización', 'Notas'],
    'Busca Trabajo': ['Derivado a Fito', 'Fecha Derivación', 'Motivo', 'Notas'],
    'Fito': ['Derivado a Fito', 'Fecha Derivación', 'Motivo', 'Notas'],
    'Empleado': ['Empresa', 'Puesto', 'Fecha Contratación', 'Salario', 'Notas']
  };

  return headersBase.concat(headersEspecificos[clasificacion] || ['Notas']);
}

/**
 * Prepara la fila de datos para insertar en la hoja de clasificación
 * @param {Array} datosGraduado - Datos del graduado
 * @param {string} clasificacion - Tipo de clasificación
 * @param {Object} datosAdicionales - Datos adicionales específicos
 * @return {Array} Fila preparada
 */
function prepararFilaClasificacion(datosGraduado, clasificacion, datosAdicionales) {
  const filaBase = [
    datosGraduado[0], // ID
    datosGraduado[1], // Nombre
    datosGraduado[2], // Teléfono
    datosGraduado[3], // Email
    new Date().toLocaleDateString('es-ES'),
    'Activo'
  ];

  // Agregar campos específicos según la clasificación
  const camposEspecificos = [
    datosAdicionales.campo1 || '',
    datosAdicionales.campo2 || '',
    datosAdicionales.campo3 || '',
    datosAdicionales.notas || ''
  ];

  return filaBase.concat(camposEspecificos);
}

/**
 * Obtiene el color asociado a cada clasificación
 * @param {string} clasificacion - Tipo de clasificación
 * @return {string} Código de color hexadecimal
 */
function obtenerColorClasificacion(clasificacion) {
  const colores = {
    'Entrevista/Seguimiento General': '#9c27b0',
    'Aliados': '#3f51b5',
    'Plataforma': '#00bcd4',
    'Conexiones Laborales': '#009688',
    'Por su Cuenta': '#4285f4',
    'Busca Trabajo': '#ff9800',
    'Fito': '#ff9800',
    'Empleado': '#0f9d58'
  };

  return colores[clasificacion] || '#4285f4';
}

/**
 * Obtiene o crea una hoja por nombre
 * @param {string} nombreHoja - Nombre de la hoja
 * @return {Sheet} Objeto de la hoja
 */
function obtenerHoja(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(nombreHoja);

  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
    Logger.log(`Hoja creada: ${nombreHoja}`);
  }

  return hoja;
}

/**
 * Configura el formato y estructura de una hoja según su tipo
 * @param {Sheet} hoja - Objeto de la hoja
 * @param {string} nombreHoja - Nombre de la hoja para determinar su tipo
 */
function configurarHoja(hoja, nombreHoja) {
  // Configuraciones generales
  hoja.setFrozenRows(1); // Congelar primera fila (headers)
  hoja.setColumnWidth(1, 120); // ID
  hoja.setColumnWidth(2, 200); // Nombre
  hoja.setColumnWidth(3, 150); // Teléfono
  hoja.setColumnWidth(4, 250); // Email

  // Agregar filtros
  const lastColumn = hoja.getMaxColumns();
  if (lastColumn > 0) {
    const range = hoja.getRange(1, 1, 1, lastColumn);
    range.createFilter();
  }

  Logger.log(`Hoja configurada: ${nombreHoja}`);
}


// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 4: SEGUIMIENTOS Y LLAMADAS PROGRAMADAS
// ═══════════════════════════════════════════════════════════════════════════

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


// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 5: CONFIGURACIÓN Y CREDENCIALES
// ═══════════════════════════════════════════════════════════════════════════

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


// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 6: UTILIDADES Y FUNCIONES AUXILIARES
// ═══════════════════════════════════════════════════════════════════════════

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


// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 7: FUNCIONES DE PRUEBA Y DEPURACIÓN
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Función de prueba para verificar que todo funciona
 */
function testSistema() {
  Logger.log('Iniciando prueba del sistema...');

  try {
    // Verificar hojas
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Hojas actuales: ' + ss.getSheets().map(s => s.getName()).join(', '));

    // Verificar configuración
    const config = verificarConfiguracion();
    Logger.log('Configuración válida: ' + config);

    Logger.log('✅ Sistema funcionando correctamente');
  } catch (error) {
    Logger.log('❌ Error en prueba: ' + error);
  }
}
