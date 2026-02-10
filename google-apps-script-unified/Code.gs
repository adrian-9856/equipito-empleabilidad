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
    .addSeparator()
    .addItem('🚀 Instalar Sistema (primera vez)', 'instalarSistema')
    .addItem('🔁 Reinstalar Sistema (borra todo)', 'reinstalarSistema')
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

// ─────────────────────────────────────────────────────────────────────────────
// ESTRUCTURA EXACTA DE COLUMNAS POR HOJA
// ─────────────────────────────────────────────────────────────────────────────

const ESTRUCTURA_HOJAS = {

  // ── GRADUADOS ──────────────────────────────────────────────────────────────
  // Datos vienen de KoboToolbox + se completan manualmente en el flujo
  'Graduados': {
    color: '#1a73e8',
    columnas: [
      // — Desde KoboToolbox —
      { nombre: 'ID Kobo',           ancho: 120, nota: 'Auto: KoboToolbox' },
      { nombre: 'Nombre Completo',   ancho: 200, nota: 'Auto: KoboToolbox' },
      { nombre: 'Número de Teléfono',ancho: 150, nota: 'Auto: KoboToolbox' },
      { nombre: 'Email',             ancho: 220, nota: 'Auto: KoboToolbox' },
      { nombre: 'Formación',         ancho: 180, nota: 'Auto: KoboToolbox' },
      { nombre: 'Cohorte',           ancho: 100, nota: 'Auto: KoboToolbox' },
      { nombre: 'Fecha Entrevista',  ancho: 140, nota: 'Auto: KoboToolbox' },
      { nombre: 'Entrevistador',     ancho: 160, nota: 'Auto: KoboToolbox' },
      // — Seguimiento General (se llena manualmente) —
      { nombre: 'Resultado Entrevista', ancho: 220, nota: 'Manual' },
      { nombre: 'Siguiente Paso',    ancho: 200, nota: 'Manual' },
      // — Estado en el flujo —
      { nombre: 'Clasificación',     ancho: 180, nota: 'Auto: al clasificar' },
      { nombre: 'Fecha Clasificación',ancho: 150, nota: 'Auto: al clasificar' },
      { nombre: 'Empleado',          ancho: 80,  nota: 'Auto: al clasificar' },
      { nombre: 'Próxima Llamada',   ancho: 180, nota: 'Auto: al emplearse' },
      { nombre: 'Notas',             ancho: 300, nota: 'Manual' }
    ]
  },

  // ── ALIADOS ────────────────────────────────────────────────────────────────
  'Aliados': {
    color: '#3f51b5',
    columnas: [
      { nombre: 'ID Kobo',           ancho: 120 },
      { nombre: 'Nombre Completo',   ancho: 200 },
      { nombre: 'Número de Teléfono',ancho: 150 },
      { nombre: 'Email',             ancho: 220 },
      { nombre: 'Formación',         ancho: 180 },
      { nombre: 'Cohorte',           ancho: 100 },
      { nombre: 'Aliado Asignado',   ancho: 200 },
      { nombre: 'Contacto Aliado',   ancho: 200 },
      { nombre: 'Fecha Derivación',  ancho: 140 },
      { nombre: 'Estado',            ancho: 120 },
      { nombre: 'Notas',             ancho: 300 }
    ]
  },

  // ── PLATAFORMA ─────────────────────────────────────────────────────────────
  'Plataforma': {
    color: '#00bcd4',
    columnas: [
      { nombre: 'ID Kobo',           ancho: 120 },
      { nombre: 'Nombre Completo',   ancho: 200 },
      { nombre: 'Número de Teléfono',ancho: 150 },
      { nombre: 'Email',             ancho: 220 },
      { nombre: 'Formación',         ancho: 180 },
      { nombre: 'Cohorte',           ancho: 100 },
      { nombre: 'Plataforma',        ancho: 180 },
      { nombre: 'Usuario/Perfil',    ancho: 200 },
      { nombre: 'Fecha Registro',    ancho: 140 },
      { nombre: 'Estado',            ancho: 120 },
      { nombre: 'Notas',             ancho: 300 }
    ]
  },

  // ── CONEXIONES LABORALES ───────────────────────────────────────────────────
  'Conexiones Laborales': {
    color: '#009688',
    columnas: [
      { nombre: 'ID Kobo',           ancho: 120 },
      { nombre: 'Nombre Completo',   ancho: 200 },
      { nombre: 'Número de Teléfono',ancho: 150 },
      { nombre: 'Email',             ancho: 220 },
      { nombre: 'Formación',         ancho: 180 },
      { nombre: 'Cohorte',           ancho: 100 },
      { nombre: 'Contacto',          ancho: 200 },
      { nombre: 'Empresa / Área',    ancho: 200 },
      { nombre: 'Tipo Conexión',     ancho: 160 },
      { nombre: 'Fecha',             ancho: 120 },
      { nombre: 'Estado',            ancho: 120 },
      { nombre: 'Notas',             ancho: 300 }
    ]
  },

  // ── POR SU CUENTA ──────────────────────────────────────────────────────────
  'Por su Cuenta': {
    color: '#4285f4',
    columnas: [
      { nombre: 'ID Kobo',              ancho: 120 },
      { nombre: 'Nombre Completo',      ancho: 200 },
      { nombre: 'Número de Teléfono',   ancho: 150 },
      { nombre: 'Email',                ancho: 220 },
      { nombre: 'Formación',            ancho: 180 },
      { nombre: 'Cohorte',              ancho: 100 },
      { nombre: 'Actividad',            ancho: 220 },
      { nombre: 'Progreso',             ancho: 220 },
      { nombre: 'Última Actualización', ancho: 160 },
      { nombre: 'Estado',               ancho: 120 },
      { nombre: 'Notas',                ancho: 300 }
    ]
  },

  // ── BUSCA TRABAJO (FITO) ───────────────────────────────────────────────────
  'Busca Trabajo (Fito)': {
    color: '#ff9800',
    columnas: [
      { nombre: 'ID Kobo',           ancho: 120 },
      { nombre: 'Nombre Completo',   ancho: 200 },
      { nombre: 'Número de Teléfono',ancho: 150 },
      { nombre: 'Email',             ancho: 220 },
      { nombre: 'Formación',         ancho: 180 },
      { nombre: 'Cohorte',           ancho: 100 },
      { nombre: 'Motivo Derivación', ancho: 250 },
      { nombre: 'Fecha Derivación',  ancho: 140 },
      { nombre: 'Estado en Fito',    ancho: 160 },
      { nombre: 'Notas',             ancho: 300 }
    ]
  },

  // ── EMPLEADOS ──────────────────────────────────────────────────────────────
  'Empleados': {
    color: '#0f9d58',
    columnas: [
      { nombre: 'ID Kobo',              ancho: 120 },
      { nombre: 'Nombre Completo',      ancho: 200 },
      { nombre: 'Número de Teléfono',   ancho: 150 },
      { nombre: 'Email',                ancho: 220 },
      { nombre: 'Formación',            ancho: 180 },
      { nombre: 'Cohorte',              ancho: 100 },
      { nombre: 'Empresa',              ancho: 200 },
      { nombre: 'Puesto',               ancho: 180 },
      { nombre: 'Fecha Contratación',   ancho: 150 },
      { nombre: 'Salario',              ancho: 120 },
      { nombre: 'Llamada 1 - Semanal',  ancho: 160, nota: 'Auto: fecha programada' },
      { nombre: 'Estado Llamada 1',     ancho: 130 },
      { nombre: 'Llamada 2 - 3 Meses',  ancho: 160, nota: 'Auto: fecha programada' },
      { nombre: 'Estado Llamada 2',     ancho: 130 },
      { nombre: 'Llamada 3 - 6 Meses',  ancho: 160, nota: 'Auto: fecha programada' },
      { nombre: 'Estado Llamada 3',     ancho: 130 },
      { nombre: 'Notas',                ancho: 300 }
    ]
  },

  // ── SEGUIMIENTOS ───────────────────────────────────────────────────────────
  'Seguimientos': {
    color: '#673ab7',
    columnas: [
      { nombre: 'ID Graduado',       ancho: 120 },
      { nombre: 'Nombre',            ancho: 200 },
      { nombre: 'Tipo Seguimiento',  ancho: 180 },
      { nombre: 'Fecha Programada',  ancho: 150 },
      { nombre: 'Fecha Realizada',   ancho: 150 },
      { nombre: 'Estado',            ancho: 120 },
      { nombre: 'Resultado',         ancho: 220 },
      { nombre: 'Notas',             ancho: 300 },
      { nombre: 'Próximo Paso',      ancho: 220 }
    ]
  },

  // ── REPORTES MENSUALES ─────────────────────────────────────────────────────
  'Reportes Mensuales': {
    color: '#e91e63',
    columnas: [
      { nombre: 'Mes / Período',         ancho: 150 },
      { nombre: 'Total Graduados',        ancho: 140 },
      { nombre: 'Empleados',              ancho: 120 },
      { nombre: 'Tasa Empleabilidad %',   ancho: 160 },
      { nombre: 'En Aliados',             ancho: 120 },
      { nombre: 'En Plataforma',          ancho: 130 },
      { nombre: 'En Conexiones',          ancho: 130 },
      { nombre: 'Por su Cuenta',          ancho: 130 },
      { nombre: 'En Fito',                ancho: 110 },
      { nombre: 'Llamadas Realizadas',    ancho: 160 },
      { nombre: 'Llamadas Pendientes',    ancho: 160 },
      { nombre: 'Retención 6 Meses %',    ancho: 160 },
      { nombre: 'Notas',                  ancho: 300 }
    ]
  }
};


// ─────────────────────────────────────────────────────────────────────────────
// FUNCIONES DE INSTALACIÓN Y REINSTALACIÓN
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Instalación inicial: crea las hojas si no existen (no borra nada)
 */
function instalarSistema() {
  const ui = SpreadsheetApp.getUi();

  const respuesta = ui.alert(
    '🚀 Instalar Sistema',
    'Esto creará todas las hojas del sistema.\n\n' +
    'Las hojas existentes NO serán modificadas.\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) return;

  _ejecutarInstalacion(false);
}

/**
 * Reinstalación: borra TODAS las hojas existentes y las recrea desde cero
 */
function reinstalarSistema() {
  const ui = SpreadsheetApp.getUi();

  // Primera confirmación
  const confirm1 = ui.alert(
    '⚠️ REINSTALAR SISTEMA',
    '¡ATENCIÓN! Esto borrará TODAS las hojas y datos existentes.\n\n' +
    'Esta acción NO se puede deshacer.\n\n' +
    '¿Estás seguro de que quieres continuar?',
    ui.ButtonSet.YES_NO
  );

  if (confirm1 !== ui.Button.YES) return;

  // Segunda confirmación de seguridad
  const confirm2 = ui.alert(
    '⚠️ CONFIRMAR REINSTALACIÓN',
    'Última confirmación: se borrarán TODOS los datos.\n\n' +
    '¿Confirmar reinstalación completa?',
    ui.ButtonSet.YES_NO
  );

  if (confirm2 !== ui.Button.YES) return;

  _ejecutarInstalacion(true);
}

/**
 * Lógica central de instalación
 * @param {boolean} borrarExistentes - Si true, elimina todas las hojas primero
 */
function _ejecutarInstalacion(borrarExistentes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      borrarExistentes ? 'Eliminando hojas existentes...' : 'Creando hojas...',
      '⏳ Procesando', -1
    );

    // ── 1. Eliminar hojas existentes si se solicita ──────────────────────────
    if (borrarExistentes) {
      const hojasDelSistema = [
        'Graduados', 'Aliados', 'Plataforma', 'Conexiones Laborales',
        'Por su Cuenta', 'Busca Trabajo (Fito)', 'Empleados',
        'Seguimientos', 'Reportes Mensuales', 'Configuración'
      ];

      // Crear hoja temporal para no quedar sin hojas
      let hojaTemp = ss.getSheetByName('_temp_');
      if (!hojaTemp) hojaTemp = ss.insertSheet('_temp_');

      hojasDelSistema.forEach(nombre => {
        const h = ss.getSheetByName(nombre);
        if (h) ss.deleteSheet(h);
      });
    }

    // ── 2. Crear cada hoja con su estructura exacta ──────────────────────────
    const ordenHojas = [
      'Graduados',
      'Aliados',
      'Plataforma',
      'Conexiones Laborales',
      'Por su Cuenta',
      'Busca Trabajo (Fito)',
      'Empleados',
      'Seguimientos',
      'Reportes Mensuales'
    ];

    ordenHojas.forEach(nombreHoja => {
      let hoja = ss.getSheetByName(nombreHoja);

      // Si no existe, crearla
      if (!hoja) {
        hoja = ss.insertSheet(nombreHoja);
      } else if (!borrarExistentes) {
        // Ya existe y no se quiere borrar: saltar
        return;
      }

      _construirHoja(hoja, nombreHoja);
    });

    // ── 3. Crear hoja Configuración ──────────────────────────────────────────
    _crearHojaConfiguracion(ss);

    // ── 4. Eliminar hoja temporal si existe ──────────────────────────────────
    const temp = ss.getSheetByName('_temp_');
    if (temp) ss.deleteSheet(temp);

    // ── 5. Ordenar hojas en el orden correcto ────────────────────────────────
    _ordenarHojas(ss, [...ordenHojas, 'Configuración']);

    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);

    ui.alert(
      '✅ Sistema Instalado',
      `El sistema está listo.\n\n` +
      `Se crearon ${ordenHojas.length + 1} hojas con sus columnas.\n\n` +
      'Próximo paso:\n' +
      '1. Configura las credenciales de KoboToolbox\n' +
      '2. Importa los graduados\n' +
      '3. Comienza a clasificar',
      ui.ButtonSet.OK
    );

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    ui.alert('❌ Error', 'Error durante la instalación:\n' + error.message, ui.ButtonSet.OK);
    Logger.log('Error en _ejecutarInstalacion: ' + error);
  }
}

/**
 * Construye una hoja con sus headers, formato y anchos de columna
 * @param {Sheet} hoja - Objeto de la hoja
 * @param {string} nombreHoja - Nombre para buscar en ESTRUCTURA_HOJAS
 */
function _construirHoja(hoja, nombreHoja) {
  const estructura = ESTRUCTURA_HOJAS[nombreHoja];
  if (!estructura) return;

  hoja.clearContents();
  hoja.clearFormats();

  const columnas = estructura.columnas;
  const headers  = columnas.map(c => c.nombre);

  // ── Escribir headers ──────────────────────────────────────────────────────
  const rangoHeader = hoja.getRange(1, 1, 1, headers.length);
  rangoHeader.setValues([headers]);

  // ── Formato de headers ────────────────────────────────────────────────────
  rangoHeader.setBackground(estructura.color)
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setFontSize(11)
             .setVerticalAlignment('middle')
             .setHorizontalAlignment('center')
             .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  hoja.setRowHeight(1, 36);

  // ── Anchos de columna ─────────────────────────────────────────────────────
  columnas.forEach((col, i) => {
    hoja.setColumnWidth(i + 1, col.ancho);
  });

  // ── Congelar primera fila ─────────────────────────────────────────────────
  hoja.setFrozenRows(1);

  // ── Quitar columnas sobrantes ─────────────────────────────────────────────
  const totalCols = hoja.getMaxColumns();
  if (totalCols > headers.length) {
    hoja.deleteColumns(headers.length + 1, totalCols - headers.length);
  }

  // ── Agregar filtros ───────────────────────────────────────────────────────
  rangoHeader.createFilter();

  // ── Color alterno en filas de datos ──────────────────────────────────────
  const regla = SpreadsheetApp.newBandingTheme ? null : null;
  try {
    hoja.getRange(2, 1, 1000, headers.length)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  } catch(e) { /* Las bandas pueden no estar disponibles en todas las versiones */ }

  Logger.log('Hoja construida: ' + nombreHoja);
}

/**
 * Crea la hoja de Configuración con los campos del sistema
 * @param {Spreadsheet} ss - Objeto del Spreadsheet
 */
function _crearHojaConfiguracion(ss) {
  let hoja = ss.getSheetByName('Configuración');
  if (!hoja) hoja = ss.insertSheet('Configuración');

  hoja.clearContents();
  hoja.clearFormats();

  // Header
  const header = hoja.getRange(1, 1, 1, 3);
  header.setValues([['Parámetro', 'Valor', 'Descripción']]);
  header.setBackground('#607d8b')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontSize(11);
  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 36);

  // Filas de configuración
  const filas = [
    ['KOBO_EXPORT_URL', '', 'URL de exportación CSV de KoboToolbox'],
    ['KOBO_TOKEN',      '', 'Token de autenticación de KoboToolbox'],
    ['EMAIL_NOTIF',     '', 'Email para recibir notificaciones'],
    ['SYNC_AUTO',       'false', 'true = sincronizar cada hora automáticamente'],
    ['NOTIF_AUTO',      'false', 'true = enviar email diario de seguimientos pendientes']
  ];

  hoja.getRange(2, 1, filas.length, 3).setValues(filas);

  hoja.setColumnWidth(1, 200);
  hoja.setColumnWidth(2, 350);
  hoja.setColumnWidth(3, 400);
}

/**
 * Ordena las hojas en el orden definido
 * @param {Spreadsheet} ss
 * @param {Array} orden - Array con los nombres en el orden deseado
 */
function _ordenarHojas(ss, orden) {
  orden.forEach((nombre, posicion) => {
    const hoja = ss.getSheetByName(nombre);
    if (hoja) ss.setActiveSheet(hoja).moveActiveSheet(posicion + 1);
  });
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

  // Agregar el graduado con la estructura exacta de ESTRUCTURA_HOJAS['Graduados']
  // Columnas: ID Kobo | Nombre Completo | Número de Teléfono | Email | Formación |
  //           Cohorte | Fecha Entrevista | Entrevistador | Resultado Entrevista |
  //           Siguiente Paso | Clasificación | Fecha Clasificación | Empleado |
  //           Próxima Llamada | Notas
  const fila = [
    graduado.id,                        // col 1 - ID Kobo
    graduado.nombre,                    // col 2 - Nombre Completo
    graduado.telefono,                  // col 3 - Número de Teléfono
    graduado.email,                     // col 4 - Email
    graduado.formacion,                 // col 5 - Formación
    graduado.cohorte || '',             // col 6 - Cohorte
    graduado.fechaEntrevista,           // col 7 - Fecha Entrevista
    graduado.entrevistador || '',       // col 8 - Entrevistador
    '',                                 // col 9 - Resultado Entrevista (manual)
    '',                                 // col 10 - Siguiente Paso (manual)
    '',                                 // col 11 - Clasificación (auto al clasificar)
    '',                                 // col 12 - Fecha Clasificación (auto)
    'No',                               // col 13 - Empleado
    '',                                 // col 14 - Próxima Llamada (auto al emplearse)
    ''                                  // col 15 - Notas (manual)
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
      // Actualizar campos de KoboToolbox (columnas 2-8)
      hoja.getRange(i + 1, 2).setValue(graduado.nombre);
      hoja.getRange(i + 1, 3).setValue(graduado.telefono);
      hoja.getRange(i + 1, 4).setValue(graduado.email);
      hoja.getRange(i + 1, 5).setValue(graduado.formacion || datos[i][4]);
      hoja.getRange(i + 1, 6).setValue(graduado.cohorte || datos[i][5]);
      hoja.getRange(i + 1, 7).setValue(graduado.fechaEntrevista || datos[i][6]);
      hoja.getRange(i + 1, 8).setValue(graduado.entrevistador || datos[i][7]);

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
      const fechaHoy = new Date().toLocaleDateString('es-ES');
      // col 11 - Clasificación
      hojaGraduados.getRange(i + 1, 11).setValue(clasificacion);
      // col 12 - Fecha Clasificación
      hojaGraduados.getRange(i + 1, 12).setValue(fechaHoy);
      // col 13 - Empleado
      hojaGraduados.getRange(i + 1, 13).setValue(clasificacion === 'Empleado' ? 'Sí' : 'No');

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
  // Headers base compartidos entre todas las sub-hojas
  const headersBase = [
    'ID Kobo',
    'Nombre Completo',
    'Número de Teléfono',
    'Email',
    'Formación',
    'Cohorte',
    'Fecha Entrevista',
    'Fecha Clasificación'
  ];

  const headersEspecificos = {
    'Entrevista/Seguimiento General': ['Tipo Entrevista', 'Resultado', 'Siguiente Paso', 'Notas'],
    'Aliados': ['Aliado Asignado', 'Contacto del Aliado', 'Notas'],
    'Plataforma': ['Plataforma Asignada', 'Usuario/Email', 'Notas'],
    'Conexiones Laborales': ['Contacto', 'Empresa/Área', 'Tipo de Conexión', 'Notas'],
    'Por su Cuenta': ['Actividad', 'Notas'],
    'Busca Trabajo': ['Motivo de Derivación', 'Notas'],
    'Fito': ['Motivo de Derivación', 'Notas'],
    'Empleado': ['Empresa', 'Puesto', 'Fecha Contratación', 'Salario',
                 'Llamada 1 - Semanal', 'Estado Llamada 1',
                 'Llamada 2 - 3 Meses', 'Estado Llamada 2',
                 'Llamada 3 - 6 Meses', 'Estado Llamada 3',
                 'Notas']
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
  // Base: campos de KoboToolbox compartidos entre todas las hojas de clasificación
  // Indices según ESTRUCTURA_HOJAS['Graduados']:
  //   0=ID Kobo, 1=Nombre, 2=Teléfono, 3=Email, 4=Formación, 5=Cohorte,
  //   6=Fecha Entrevista, 7=Entrevistador
  const filaBase = [
    datosGraduado[0], // ID Kobo
    datosGraduado[1], // Nombre Completo
    datosGraduado[2], // Número de Teléfono
    datosGraduado[3], // Email
    datosGraduado[4], // Formación
    datosGraduado[5], // Cohorte
    datosGraduado[6], // Fecha Entrevista
    new Date().toLocaleDateString('es-ES'), // Fecha Clasificación
  ];

  // Agregar campos específicos según la clasificación
  switch (clasificacion) {
    case 'Entrevista/Seguimiento General':
      return filaBase.concat([
        datosAdicionales.tipoEntrevista || '',
        datosAdicionales.resultado || '',
        datosAdicionales.siguientePaso || '',
        datosAdicionales.notas || ''
      ]);
    case 'Aliados':
      return filaBase.concat([
        datosAdicionales.aliado || '',
        datosAdicionales.contactoAliado || '',
        datosAdicionales.notas || ''
      ]);
    case 'Plataforma':
      return filaBase.concat([
        datosAdicionales.plataforma || '',
        datosAdicionales.usuario || '',
        datosAdicionales.notas || ''
      ]);
    case 'Conexiones Laborales':
      return filaBase.concat([
        datosAdicionales.contacto || '',
        datosAdicionales.empresa || '',
        datosAdicionales.tipoConexion || '',
        datosAdicionales.notas || ''
      ]);
    case 'Por su Cuenta':
      return filaBase.concat([
        datosAdicionales.actividad || '',
        datosAdicionales.notas || ''
      ]);
    case 'Busca Trabajo':
    case 'Fito':
      return filaBase.concat([
        datosAdicionales.motivoDerivacion || '',
        datosAdicionales.notas || ''
      ]);
    case 'Empleado':
      return filaBase.concat([
        datosAdicionales.empresa || '',
        datosAdicionales.puesto || '',
        datosAdicionales.fechaContratacion || new Date().toLocaleDateString('es-ES'),
        datosAdicionales.salario || '',
        '', // Llamada 1 - Semanal (auto)
        'Pendiente', // Estado Llamada 1
        '', // Llamada 2 - 3 Meses (auto)
        'Pendiente', // Estado Llamada 2
        '', // Llamada 3 - 6 Meses (auto)
        'Pendiente', // Estado Llamada 3
        datosAdicionales.notas || ''
      ]);
    default:
      return filaBase.concat([datosAdicionales.notas || '']);
  }
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
      hojaGraduados.getRange(i + 1, 14).setValue(tipoLlamada); // col 14 - Próxima Llamada
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
