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
    const datos = obtenerDatosKoboToolbox();
    if (!datos || datos.length === 0) {
      ui.alert('⚠️ Sin Datos',
               'No se encontraron datos nuevos para importar.',
               ui.ButtonSet.OK);
      return;
    }
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
 * Muestra la configuración del proyecto
 */
function mostrarConfiguracion() {
  const html = HtmlService.createHtmlOutputFromFile('ConfiguracionKobo')
    .setWidth(500)
    .setHeight(400)
    .setTitle('Configuración de KoboToolbox');
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuración');
}

// ─────────────────────────────────────────────────────────────────────────────
// COLUMNAS COMUNES — van en las 6 hojas de apilabilidad
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Columnas que aparecen en TODAS las hojas de apilabilidad (información principal)
 * tipo: 'texto' | 'fecha' | 'siNo' | 'checkbox' | 'dropdown'
 */
const COLUMNAS_COMUNES = [
  { nombre: 'Fecha de ingreso',  ancho: 140, tipo: 'fecha' },
  { nombre: 'Creamos ID',        ancho: 130, tipo: 'texto' },  // antes de Nombre
  { nombre: 'Nombre completo',   ancho: 200, tipo: 'texto' },
  { nombre: 'Género',            ancho: 100, tipo: 'texto' },
  { nombre: 'Edad',              ancho: 80,  tipo: 'texto' },
  { nombre: 'Nivel educativo',   ancho: 160, tipo: 'texto' }
];

// Etapas del flujo — usadas como opciones de dropdown en varias hojas
const ETAPAS_FLUJO = [
  'Aliados',
  'Plataforma',
  'Derivaciones',
  'Por su cuenta',
  'No busca trabajo - Fito',
  'Activamente busca trabajo'
];

// Áreas de trabajo — opciones del dropdown Área en Aliados
const AREAS_TRABAJO = [
  'Tecnología',
  'Manufactura',
  'Servicios al cliente',
  'Administración',
  'Ventas',
  'Logística',
  'Salud',
  'Construcción',
  'Educación',
  'Finanzas',
  'Otro'
];

// ─────────────────────────────────────────────────────────────────────────────
// ESTRUCTURA EXACTA DE COLUMNAS POR HOJA
// ─────────────────────────────────────────────────────────────────────────────

const ESTRUCTURA_HOJAS = {

  // ── GRADUADOS ──────────────────────────────────────────────────────────────
  // Columnas:
  //  1=No. | 2=Fecha de envío | 3=Creamos ID | 4=Nombre completo |
  //  5=Número de teléfono | 6=Formación | 7=Cohorte | 8=Fecha de entrevista |
  //  9=Entrevistador | 10=Resultado entrevista | 11=Siguiente paso |
  //  12=Clasificación | 13=Empleado | 14=Próxima llamada | 15=Notas | 16=Etapa
  'Graduados': {
    color: '#1a73e8',
    columnas: [
      { nombre: 'No.',                  ancho: 60,  tipo: 'texto'  },  // era ID Kobo
      { nombre: 'Fecha de envío',       ancho: 140, tipo: 'fecha'  },  // auto al importar
      { nombre: 'Creamos ID',           ancho: 130, tipo: 'texto'  },  // antes de Nombre
      { nombre: 'Nombre completo',      ancho: 200, tipo: 'texto'  },
      { nombre: 'Número de teléfono',   ancho: 150, tipo: 'texto'  },
      { nombre: 'Formación',            ancho: 180, tipo: 'texto'  },
      { nombre: 'Cohorte',              ancho: 100, tipo: 'texto'  },
      { nombre: 'Fecha de entrevista',  ancho: 140, tipo: 'fecha'  },
      { nombre: 'Entrevistador',        ancho: 160, tipo: 'texto'  },
      { nombre: 'Resultado entrevista', ancho: 220, tipo: 'texto'  },
      { nombre: 'Siguiente paso',       ancho: 200, tipo: 'texto'  },
      { nombre: 'Clasificación',        ancho: 180, tipo: 'texto'  },
      { nombre: 'Empleado',             ancho: 90,  tipo: 'siNo'   },
      { nombre: 'Próxima llamada',      ancho: 180, tipo: 'texto'  },
      { nombre: 'Notas',                ancho: 300, tipo: 'texto'  },
      { nombre: 'Etapa',                ancho: 200, tipo: 'dropdown', opciones: ETAPAS_FLUJO }
    ]
  },

  // ── 1. ALIADOS ─────────────────────────────────────────────────────────────
  'Aliados': {
    color: '#3f51b5',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'Compartió el CV',          ancho: 140, tipo: 'siNo'     },
      { nombre: 'Área',                      ancho: 160, tipo: 'dropdown', opciones: AREAS_TRABAJO },
      { nombre: 'Entrevista',                ancho: 130, tipo: 'texto'    },
      { nombre: 'Día de prueba',             ancho: 130, tipo: 'texto'    },
      { nombre: 'Confirmación de recepción', ancho: 200, tipo: 'siNo'     },
      { nombre: 'Notas',                     ancho: 300, tipo: 'texto'    },
      { nombre: 'Activo',                    ancho: 90,  tipo: 'siNo'     }
    ]
  },

  // ── 2. PLATAFORMA ──────────────────────────────────────────────────────────
  'Plataforma': {
    color: '#00bcd4',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'Cita',              ancho: 90,  tipo: 'checkbox' },
      { nombre: 'Creación de perfil',ancho: 160, tipo: 'checkbox' },  // cambio: checkbox
      { nombre: 'Contacto',          ancho: 90,  tipo: 'checkbox' },
      { nombre: 'Trámites',          ancho: 200, tipo: 'texto'    },
      { nombre: 'Entrevista',        ancho: 130, tipo: 'texto'    },
      { nombre: 'Confirmación',      ancho: 130, tipo: 'texto'    },
      { nombre: 'Recepción',         ancho: 130, tipo: 'texto'    },
      { nombre: 'Nota',              ancho: 300, tipo: 'texto'    },
      { nombre: 'Activo',            ancho: 90,  tipo: 'siNo'     }
    ]
  },

  // ── 3. DERIVACIONES ───────────────────────────────────────────────────────
  'Derivaciones': {
    color: '#ff9800',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'Envío 1',   ancho: 140, tipo: 'texto' },
      { nombre: 'Llamada 1', ancho: 120, tipo: 'texto' },
      { nombre: 'Envío 2',   ancho: 140, tipo: 'texto' },
      { nombre: 'Llamada 2', ancho: 120, tipo: 'texto' },
      { nombre: 'Envío 3',   ancho: 140, tipo: 'texto' },
      { nombre: 'Llamada 3', ancho: 120, tipo: 'texto' },
      { nombre: 'Notas',     ancho: 300, tipo: 'texto' },
      { nombre: 'Activo',    ancho: 90,  tipo: 'siNo'  }
    ]
  },

  // ── 4. POR SU CUENTA ──────────────────────────────────────────────────────
  'Por su cuenta': {
    color: '#4285f4',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'Mensaje', ancho: 120, tipo: 'texto' },
      { nombre: 'Llamada', ancho: 120, tipo: 'texto' },
      { nombre: 'Nota',    ancho: 300, tipo: 'texto' },
      { nombre: 'Activo',  ancho: 90,  tipo: 'siNo'  }
    ]
  },

  // ── 5. NO BUSCA TRABAJO / FITO ────────────────────────────────────────────
  'No busca trabajo - Fito': {
    color: '#9e9e9e',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'DPI',                ancho: 140, tipo: 'texto' },
      { nombre: 'Número de teléfono', ancho: 160, tipo: 'texto' },
      { nombre: 'Formación',          ancho: 180, tipo: 'texto' },
      { nombre: 'Cohorte',            ancho: 110, tipo: 'texto' },
      { nombre: 'Nota',               ancho: 300, tipo: 'texto' },
      { nombre: 'Activo',             ancho: 90,  tipo: 'siNo'  }
    ]
  },

  // ── 6. ACTIVAMENTE BUSCA TRABAJO ──────────────────────────────────────────
  'Activamente busca trabajo': {
    color: '#0f9d58',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'Entrevista', ancho: 130, tipo: 'texto' },
      { nombre: 'Trámites',   ancho: 200, tipo: 'texto' },
      { nombre: 'Activo',     ancho: 90,  tipo: 'siNo'  }
    ]
  },

  // ── CONEXIONES LABORALES ──────────────────────────────────────────────────
  'Conexiones Laborales': {
    color: '#e65100',
    columnas: [
      { nombre: 'Creamos ID',                    ancho: 130, tipo: 'texto' },
      { nombre: 'Nombres',                       ancho: 160, tipo: 'texto' },
      { nombre: 'Apellidos',                     ancho: 160, tipo: 'texto' },
      { nombre: 'Edad',                          ancho: 80,  tipo: 'texto' },
      { nombre: 'Nivel de estudios',             ancho: 160, tipo: 'texto' },
      { nombre: 'Sexo',                          ancho: 100, tipo: 'texto' },
      { nombre: 'Tipo',                          ancho: 130, tipo: 'texto' },
      { nombre: 'Programa',                      ancho: 160, tipo: 'texto' },
      { nombre: 'Proyecto',                      ancho: 160, tipo: 'texto' },
      { nombre: 'Especialidad',                  ancho: 180, tipo: 'texto' },
      { nombre: 'Empresa',                       ancho: 200, tipo: 'texto' },
      { nombre: 'Cargo que desempeña',           ancho: 200, tipo: 'texto' },
      { nombre: 'Tipo de duración de contrato',  ancho: 220, tipo: 'texto' },
      { nombre: 'Tipo de contrato',              ancho: 180, tipo: 'texto' },
      { nombre: 'Fecha de inicio',               ancho: 140, tipo: 'fecha' },
      { nombre: 'Fecha de final',                ancho: 140, tipo: 'fecha' },
      { nombre: 'Duración (meses)',              ancho: 140, tipo: 'texto' },
      { nombre: 'Salario mensual',               ancho: 150, tipo: 'texto' }
    ]
  },

  // ── SEGUIMIENTOS ───────────────────────────────────────────────────────────
  'Seguimientos': {
    color: '#673ab7',
    columnas: [
      { nombre: 'No.',              ancho: 60,  tipo: 'texto' },  // número de registro
      { nombre: 'Creamos ID',       ancho: 130, tipo: 'texto' },  // antes de Nombre
      { nombre: 'Nombre',           ancho: 200, tipo: 'texto' },
      { nombre: 'Tipo seguimiento', ancho: 180, tipo: 'texto' },
      { nombre: 'Fecha programada', ancho: 150, tipo: 'fecha' },
      { nombre: 'Fecha realizada',  ancho: 150, tipo: 'fecha' },
      { nombre: 'Estado',           ancho: 120, tipo: 'texto' },
      { nombre: 'Resultado',        ancho: 220, tipo: 'texto' },
      { nombre: 'Notas',            ancho: 300, tipo: 'texto' },
      { nombre: 'Próximo paso',     ancho: 220, tipo: 'texto' }
    ]
  },

  // ── ESTADO ACTUAL DEL PARTICIPANTE ────────────────────────────────────────
  // Historial de movimientos entre etapas + resumen de conteos al final
  'Estado actual del participante': {
    color: '#e91e63',
    columnas: [
      { nombre: 'Fecha de registro', ancho: 150, tipo: 'fecha'    },
      { nombre: 'Creamos ID',        ancho: 130, tipo: 'texto'    },  // antes de Nombre
      { nombre: 'Nombre completo',   ancho: 200, tipo: 'texto'    },
      { nombre: 'Etapa',             ancho: 200, tipo: 'dropdown', opciones: ETAPAS_FLUJO },
      { nombre: 'Nota',              ancho: 300, tipo: 'texto'    }
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
  const confirm1 = ui.alert(
    '⚠️ REINSTALAR SISTEMA',
    '¡ATENCIÓN! Esto borrará TODAS las hojas y datos existentes.\n\n' +
    'Esta acción NO se puede deshacer.\n\n' +
    '¿Estás seguro de que quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirm1 !== ui.Button.YES) return;

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
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  try {
    ss.toast(
      borrarExistentes ? 'Eliminando hojas existentes...' : 'Creando hojas...',
      '⏳ Procesando', -1
    );

    // ── 1. Eliminar hojas existentes si se solicita ──────────────────────────
    if (borrarExistentes) {
      const hojasDelSistema = [
        // Hojas actuales
        'Graduados',
        'Aliados',
        'Plataforma',
        'Derivaciones',
        'Por su cuenta',
        'No busca trabajo - Fito',
        'Activamente busca trabajo',
        'Seguimientos',
        'Estado actual del participante',
        'Conexiones Laborales',
        'Configuración',
        // Nombres legacy (por si acaso existen)
        'Reportes mensuales',
        'Reportes Mensuales',
        'Por su Cuenta',
        'Busca Trabajo (Fito)',
        'No Busca Trabajar',
        'Filtro/Papelería',
        'Revisión y Control',
        'Empleados',
        'Conexiones Laborales',
        'Entrevista/Seguimiento General'
      ];

      // Hoja temporal para no quedar sin hojas
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
      'Derivaciones',
      'Por su cuenta',
      'No busca trabajo - Fito',
      'Activamente busca trabajo',
      'Conexiones Laborales',
      'Seguimientos',
      'Estado actual del participante'
    ];

    ordenHojas.forEach(nombreHoja => {
      let hoja = ss.getSheetByName(nombreHoja);
      if (!hoja) {
        hoja = ss.insertSheet(nombreHoja);
      } else if (!borrarExistentes) {
        return; // Ya existe y no se quiere borrar
      }
      _construirHoja(hoja, nombreHoja);
    });

    // ── 3. Crear hoja Configuración ──────────────────────────────────────────
    _crearHojaConfiguracion(ss);

    // ── 4. Eliminar hoja temporal si existe ──────────────────────────────────
    const temp = ss.getSheetByName('_temp_');
    if (temp) ss.deleteSheet(temp);

    // ── 5. Ordenar hojas ─────────────────────────────────────────────────────
    _ordenarHojas(ss, [...ordenHojas, 'Configuración']);

    ss.toast('', '', 1);

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
    ss.toast('', '', 1);
    ui.alert('❌ Error', 'Error durante la instalación:\n' + error.message, ui.ButtonSet.OK);
    Logger.log('Error en _ejecutarInstalacion: ' + error);
  }
}

/**
 * Construye una hoja: headers, formato, anchos, validaciones y filtros
 * @param {Sheet}  hoja       - Objeto de la hoja
 * @param {string} nombreHoja - Clave en ESTRUCTURA_HOJAS
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
  rangoHeader
    .setBackground(estructura.color)
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

  // ── Validaciones por tipo de columna ─────────────────────────────────────
  columnas.forEach((col, i) => {
    const colNum = i + 1;
    const rango  = hoja.getRange(2, colNum, 999);

    if (col.tipo === 'siNo') {
      const regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Sí', 'No'], true)
        .setAllowInvalid(false)
        .build();
      rango.setDataValidation(regla);

    } else if (col.tipo === 'dropdown' && col.opciones && col.opciones.length) {
      // Dropdown con lista de opciones personalizada
      const regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(col.opciones, true)
        .setAllowInvalid(false)
        .build();
      rango.setDataValidation(regla);

    } else if (col.tipo === 'checkbox') {
      const regla = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      rango.setDataValidation(regla);

    } else if (col.tipo === 'fecha') {
      rango.setNumberFormat('dd/mm/yyyy');
    }
  });

  // ── Color alterno en filas de datos ──────────────────────────────────────
  try {
    hoja.getRange(2, 1, 1000, headers.length)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  } catch (e) { /* ignorar si no está disponible */ }

  Logger.log('Hoja construida: ' + nombreHoja);
}

/**
 * Crea la hoja de Configuración con los campos del sistema
 * @param {Spreadsheet} ss
 */
function _crearHojaConfiguracion(ss) {
  let hoja = ss.getSheetByName('Configuración');
  if (!hoja) hoja = ss.insertSheet('Configuración');

  hoja.clearContents();
  hoja.clearFormats();

  const header = hoja.getRange(1, 1, 1, 3);
  header.setValues([['Parámetro', 'Valor', 'Descripción']]);
  header.setBackground('#607d8b')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontSize(11);

  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 36);

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
 * @param {Array}       orden - nombres en orden deseado
 */
function _ordenarHojas(ss, orden) {
  orden.forEach((nombre, posicion) => {
    const hoja = ss.getSheetByName(nombre);
    if (hoja) {
      ss.setActiveSheet(hoja);
      ss.moveActiveSheet(posicion + 1);
    }
  });
}

/**
 * Obtiene graduados que aún no han sido clasificados
 * @return {Array}
 */
function obtenerGraduadosSinClasificar() {
  try {
    const hoja  = obtenerHoja('Graduados');
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const sin = [];
    for (let i = 1; i < datos.length; i++) {
      const clasificacion = datos[i][11]; // col 12 = Clasificación
      if (!clasificacion || clasificacion.trim() === '') {
        sin.push({ id: datos[i][0], nombre: datos[i][3] }); // [3] = Nombre completo
      }
    }
    return sin;
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
 * @return {Array}
 */
function obtenerDatosKoboToolbox() {
  try {
    const config = obtenerConfiguracion();
    if (!config.koboExportUrl) {
      throw new Error('URL de exportación de KoboToolbox no configurada');
    }
    const opciones = {
      method: 'get',
      headers: {},
      muteHttpExceptions: true
    };
    if (config.koboToken) {
      opciones.headers['Authorization'] = 'Token ' + config.koboToken;
    }
    Logger.log('Obteniendo datos de KoboToolbox...');
    const respuesta = UrlFetchApp.fetch(config.koboExportUrl, opciones);
    const codigo    = respuesta.getResponseCode();
    if (codigo !== 200) {
      throw new Error(`Error HTTP ${codigo}: ${respuesta.getContentText()}`);
    }
    const datosParseados = parsearCSV(respuesta.getContentText());
    Logger.log(`Se obtuvieron ${datosParseados.length} registros de KoboToolbox`);
    return datosParseados;
  } catch (error) {
    Logger.log('Error al obtener datos de KoboToolbox: ' + error);
    throw error;
  }
}

/**
 * Parsea el contenido CSV y lo convierte en array de objetos
 * @param {string} csvContent
 * @return {Array}
 */
function parsearCSV(csvContent) {
  try {
    const lineas = csvContent.split('\n');
    if (lineas.length < 2) return [];
    const headers = parsearLineaCSV(lineas[0]);
    const datos   = [];
    for (let i = 1; i < lineas.length; i++) {
      if (lineas[i].trim() === '') continue;
      const valores = parsearLineaCSV(lineas[i]);
      const objeto  = {};
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
 * Parsea una línea de CSV respetando comillas
 * @param {string} linea
 * @return {Array}
 */
function parsearLineaCSV(linea) {
  const valores = [];
  let valorActual      = '';
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
  valores.push(valorActual.trim());
  return valores;
}

/**
 * Valida y limpia los datos obtenidos de KoboToolbox
 * @param {Object} dato
 * @return {Object}
 */
function validarDatosGraduado(dato) {
  return {
    id:              dato.id || dato._id || '',
    nombre:          dato.nombre || dato.name || '',
    telefono:        dato.telefono || dato.phone || '',
    email:           dato.email || dato.correo || '',
    fechaEntrevista: dato.fecha_entrevista || dato.interview_date || '',
    formacion:       dato.formacion || dato.training || '',
    cohorte:         dato.cohorte || '',
    entrevistador:   dato.entrevistador || '',
    fechaRegistro:   dato._submission_time || new Date().toISOString(),
    estado:          'Nuevo',
    clasificacion:   '',
    empleado:        false,
    proximaLlamada:  ''
  };
}

/**
 * Verifica si un graduado ya existe en el sistema
 * @param {string} id
 * @return {boolean}
 */
function graduadoExiste(id) {
  const hoja  = obtenerHoja('Graduados');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === id) return true;
  }
  return false;
}

/**
 * Sincronización automática (puede usarse con trigger)
 */
function sincronizacionAutomatica() {
  try {
    Logger.log('Iniciando sincronización automática...');
    const datos     = obtenerDatosKoboToolbox();
    const resultado = procesarDatosGraduados(datos);
    Logger.log(`Sincronización completada: ${resultado.nuevos} nuevos, ${resultado.total} total`);
    if (resultado.nuevos > 0) enviarNotificacionNuevosGraduados(resultado.nuevos);
  } catch (error) {
    Logger.log('Error en sincronización automática: ' + error);
  }
}

/**
 * Configura un trigger para sincronización automática cada hora
 */
function configurarTriggerSincronizacion() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sincronizacionAutomatica') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizacionAutomatica').timeBased().everyHours(1).create();
  Logger.log('Sincronización automática configurada para ejecutarse cada hora');
}

// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 3: GESTIÓN DE HOJAS Y DATOS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Procesa los datos de graduados y los almacena
 * @param {Array} datos
 * @return {Object}
 */
function procesarDatosGraduados(datos) {
  let nuevos      = 0;
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

  const total = obtenerHoja('Graduados').getLastRow() - 1;
  return { nuevos, actualizados, total };
}

/**
 * Agrega un nuevo graduado a la hoja Graduados
 * Columnas (según ESTRUCTURA_HOJAS['Graduados']):
 *   1=No. | 2=Fecha de envío (auto) | 3=Creamos ID | 4=Nombre completo |
 *   5=Número de teléfono | 6=Formación | 7=Cohorte |
 *   8=Fecha de entrevista | 9=Entrevistador | 10=Resultado entrevista |
 *   11=Siguiente paso | 12=Clasificación | 13=Empleado |
 *   14=Próxima llamada | 15=Notas | 16=Etapa
 * @param {Object} graduado
 */
function agregarGraduado(graduado) {
  const hoja = obtenerHoja('Graduados');
  const fila = [
    graduado.id,                            // No. (referencia KoboToolbox)
    new Date().toLocaleDateString('es-ES'), // Fecha de envío — auto al importar
    '',                                     // Creamos ID (se llena manualmente)
    graduado.nombre,                        // Nombre completo
    graduado.telefono,
    graduado.formacion,
    graduado.cohorte || '',
    graduado.fechaEntrevista,
    graduado.entrevistador || '',
    '',    // Resultado entrevista (manual)
    '',    // Siguiente paso (manual)
    '',    // Clasificación (auto al clasificar)
    'No',  // Empleado
    '',    // Próxima llamada (auto al emplearse)
    '',    // Notas (manual)
    ''     // Etapa (dropdown — se llena al clasificar)
  ];
  hoja.appendRow(fila);
  Logger.log(`Graduado agregado: ${graduado.nombre} (${graduado.id})`);
}

/**
 * Actualiza los datos de KoboToolbox de un graduado existente
 * Nota: la Fecha de envío (col 2) NO se actualiza — conserva la original
 * @param {Object} graduado
 */
function actualizarGraduado(graduado) {
  const hoja  = obtenerHoja('Graduados');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === graduado.id) {
      // col 1=No. | col 2=Fecha de envío | col 3=Creamos ID → no tocar
      hoja.getRange(i + 1, 4).setValue(graduado.nombre);
      hoja.getRange(i + 1, 5).setValue(graduado.telefono);
      hoja.getRange(i + 1, 6).setValue(graduado.formacion       || datos[i][5]);
      hoja.getRange(i + 1, 7).setValue(graduado.cohorte         || datos[i][6]);
      hoja.getRange(i + 1, 8).setValue(graduado.fechaEntrevista || datos[i][7]);
      hoja.getRange(i + 1, 9).setValue(graduado.entrevistador   || datos[i][8]);
      Logger.log(`Graduado actualizado: ${graduado.nombre} (${graduado.id})`);
      break;
    }
  }
}

/**
 * Clasifica un graduado y lo copia a la hoja correspondiente
 * @param {string} graduadoId
 * @param {string} clasificacion
 * @param {Object} datosAdicionales
 */
function clasificarGraduado(graduadoId, clasificacion, datosAdicionales = {}) {
  const hojaGraduados = obtenerHoja('Graduados');
  const datos         = hojaGraduados.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === graduadoId) {
      // Índices con nueva estructura (No. | Fecha envío | Creamos ID | Nombre...):
      //   col 12=Clasificación | col 13=Empleado | col 14=Próxima llamada |
      //   col 15=Notas | col 16=Etapa
      hojaGraduados.getRange(i + 1, 12).setValue(clasificacion);             // Clasificación
      hojaGraduados.getRange(i + 1, 13).setValue(
        clasificacion === 'Activamente busca trabajo' ? 'Sí' : 'No'          // Empleado
      );
      hojaGraduados.getRange(i + 1, 16).setValue(clasificacion);             // Etapa (dropdown)
      copiarAHojaClasificacion(datos[i], clasificacion, datosAdicionales);
      registrarMovimientoEtapa(datos[i][2], datos[i][3], clasificacion, datosAdicionales.nota || '');
      if (clasificacion === 'Activamente busca trabajo') {
        programarSeguimientos(graduadoId, datos[i][3]); // [3] = Nombre completo
      }
      Logger.log(`Graduado ${datos[i][3]} clasificado como: ${clasificacion}`);
      break;
    }
  }
}

/**
 * Copia los datos del graduado a su hoja de clasificación
 * @param {Array}  datosGraduado
 * @param {string} clasificacion
 * @param {Object} datosAdicionales
 */
function copiarAHojaClasificacion(datosGraduado, clasificacion, datosAdicionales) {
  const nombreHoja = obtenerNombreHojaClasificacion(clasificacion);
  const hoja       = obtenerHoja(nombreHoja);
  const fila       = prepararFilaClasificacion(datosGraduado, clasificacion, datosAdicionales);
  hoja.appendRow(fila);
}

/**
 * Mapeo de clasificación → nombre de hoja
 * @param {string} clasificacion
 * @return {string}
 */
function obtenerNombreHojaClasificacion(clasificacion) {
  const mapeo = {
    'Aliados':                    'Aliados',
    'Plataforma':                 'Plataforma',
    'Derivaciones':               'Derivaciones',
    'Por su cuenta':              'Por su cuenta',
    'No busca trabajo':           'No busca trabajo - Fito',
    'Fito':                       'No busca trabajo - Fito',
    'No busca trabajo - Fito':    'No busca trabajo - Fito',
    'Activamente busca trabajo':  'Activamente busca trabajo',
    // legacy
    'Por su Cuenta':              'Por su cuenta',
    'Busca Trabajo':              'Derivaciones',
    'Empleado':                   'Activamente busca trabajo'
  };
  return mapeo[clasificacion] || 'Graduados';
}

/**
 * Prepara la fila de datos para insertar en la hoja de clasificación.
 * Base = columnas comunes (información principal) mapeadas desde Graduados.
 *
 * Índices en datosGraduado (hoja Graduados):
 *   0=No. | 1=Fecha de envío | 2=Creamos ID | 3=Nombre completo |
 *   4=Teléfono | 5=Formación | 6=Cohorte | 7=Fecha entrevista | 8=Entrevistador
 *
 * Columnas comunes que van en las 6 hojas de apilabilidad:
 *   Fecha de ingreso | Creamos ID | Nombre completo | Género | Edad | Nivel educativo
 *
 * @param {Array}  datosGraduado
 * @param {string} clasificacion
 * @param {Object} datosAdicionales
 * @return {Array}
 */
function prepararFilaClasificacion(datosGraduado, clasificacion, datosAdicionales) {
  const filaBase = [
    new Date().toLocaleDateString('es-ES'), // Fecha de ingreso
    datosGraduado[2],                        // Creamos ID (índice 2)
    datosGraduado[3],                        // Nombre completo (índice 3)
    datosAdicionales.genero   || '',         // Género
    datosAdicionales.edad     || '',         // Edad
    datosAdicionales.nivelEdu || ''          // Nivel educativo
  ];

  switch (clasificacion) {

    case 'Aliados':
      return filaBase.concat([
        datosAdicionales.pendiente             || 'No',
        datosAdicionales.tramites              || '',
        datosAdicionales.entrevista            || '',
        datosAdicionales.diaDePrueba           || '',
        datosAdicionales.confirmacionRecepcion || '',
        datosAdicionales.notas                 || '',
        datosAdicionales.activo                || 'Sí'
      ]);

    case 'Plataforma':
      return filaBase.concat([
        false,                                 // Cita (checkbox)
        datosAdicionales.creacionPerfil        || '',
        false,                                 // Contacto (checkbox)
        datosAdicionales.tramites              || '',
        datosAdicionales.entrevista            || '',
        datosAdicionales.confirmacion          || '',
        datosAdicionales.recepcion             || '',
        datosAdicionales.nota                  || '',
        datosAdicionales.activo                || 'Sí'
      ]);

    case 'Derivaciones':
      return filaBase.concat([
        datosAdicionales.envio1   || '',
        datosAdicionales.llamada1 || '',
        datosAdicionales.envio2   || '',
        datosAdicionales.llamada2 || '',
        datosAdicionales.envio3   || '',
        datosAdicionales.llamada3 || '',
        datosAdicionales.notas    || '',
        datosAdicionales.activo   || 'Sí'
      ]);

    case 'Por su cuenta':
    case 'Por su Cuenta':
      return filaBase.concat([
        datosAdicionales.mensaje || '',
        datosAdicionales.llamada || '',
        datosAdicionales.nota    || '',
        datosAdicionales.activo  || 'Sí'
      ]);

    case 'No busca trabajo - Fito':
    case 'No busca trabajo':
    case 'Fito':
      return filaBase.concat([
        datosAdicionales.dpi       || '',
        datosAdicionales.telefono  || datosGraduado[4] || '',  // índice 4 = Teléfono
        datosAdicionales.formacion || datosGraduado[5] || '',  // índice 5 = Formación
        datosAdicionales.cohorte   || datosGraduado[6] || '',  // índice 6 = Cohorte
        datosAdicionales.nota      || '',
        datosAdicionales.activo    || 'No'
      ]);

    case 'Activamente busca trabajo':
    case 'Empleado':
      return filaBase.concat([
        datosAdicionales.entrevista || '',
        datosAdicionales.tramites   || '',
        datosAdicionales.activo     || 'Sí'
      ]);

    default:
      return filaBase.concat([datosAdicionales.notas || '']);
  }
}

/**
 * Obtiene o crea una hoja por nombre
 * @param {string} nombreHoja
 * @return {Sheet}
 */
function obtenerHoja(nombreHoja) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let   hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
    Logger.log(`Hoja creada: ${nombreHoja}`);
  }
  return hoja;
}

// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 4: SEGUIMIENTOS Y LLAMADAS PROGRAMADAS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Programa los seguimientos automáticos para un graduado empleado/activo
 * @param {string} graduadoId
 * @param {string} nombreGraduado
 */
function programarSeguimientos(graduadoId, nombreGraduado) {
  const hojaSeguimientos = obtenerHoja('Seguimientos');

  const fechaBase   = new Date();
  const seguimientos = [
    { tipo: 'Llamada 1 - Semanal',  diasDespues: 7,   descripcion: 'Primera semana de trabajo' },
    { tipo: 'Llamada 2 - 3 Meses',  diasDespues: 90,  descripcion: 'Seguimiento a los 3 meses' },
    { tipo: 'Llamada 3 - 6 Meses',  diasDespues: 180, descripcion: 'Seguimiento a los 6 meses' }
  ];

  seguimientos.forEach(seg => {
    const fechaProgramada = new Date(fechaBase);
    fechaProgramada.setDate(fechaProgramada.getDate() + seg.diasDespues);

    // Columnas Seguimientos: No. | Creamos ID | Nombre | Tipo | Fecha prog. | Fecha real. | Estado | Resultado | Notas | Próximo paso
    hojaSeguimientos.appendRow([
      graduadoId,     // No. (referencia)
      '',             // Creamos ID (se completa manualmente)
      nombreGraduado, // Nombre
      seg.tipo,
      fechaProgramada.toLocaleDateString('es-ES'),
      '',             // Fecha realizada
      'Pendiente',
      '',             // Resultado
      seg.descripcion,
      ''              // Próximo paso
    ]);
  });

  Logger.log(`Seguimientos programados para: ${nombreGraduado}`);
  actualizarProximaLlamada(graduadoId, 'Llamada 1 - Semanal');
}

/**
 * Obtiene los seguimientos pendientes para hoy o fechas pasadas
 * @return {Array}
 */
function obtenerSeguimientosPendientes() {
  const hojaSeguimientos = obtenerHoja('Seguimientos');
  const datos            = hojaSeguimientos.getDataRange().getValues();
  if (datos.length <= 1) return [];

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const pendientes = [];
  // Seguimientos: [0]=No. | [1]=Creamos ID | [2]=Nombre | [3]=Tipo | [4]=Fecha prog. |
  //              [5]=Fecha real. | [6]=Estado | [7]=Resultado | [8]=Notas | [9]=Próximo paso
  for (let i = 1; i < datos.length; i++) {
    const estado             = datos[i][6];
    const fechaProgramadaStr = datos[i][4];
    if (estado === 'Pendiente' && fechaProgramadaStr) {
      const fechaProgramada = parsearFecha(fechaProgramadaStr);
      if (fechaProgramada && fechaProgramada <= hoy) {
        pendientes.push({
          fila:            i + 1,
          id:              datos[i][0],
          nombre:          datos[i][2],  // [2] = Nombre
          tipo:            datos[i][3],  // [3] = Tipo seguimiento
          fechaProgramada: fechaProgramadaStr,
          notas:           datos[i][8]   // [8] = Notas
        });
      }
    }
  }
  return pendientes;
}

/**
 * Marca un seguimiento como realizado
 * @param {number} fila
 * @param {string} resultado
 * @param {string} notas
 * @param {string} proximoPaso
 */
function marcarSeguimientoRealizado(fila, resultado, notas, proximoPaso) {
  // Seguimientos: col 1=No. | 2=Creamos ID | 3=Nombre | 4=Tipo | 5=Fecha prog. |
  //              6=Fecha real. | 7=Estado | 8=Resultado | 9=Notas | 10=Próximo paso
  const hoja = obtenerHoja('Seguimientos');
  hoja.getRange(fila, 6).setValue(new Date().toLocaleDateString('es-ES')); // Fecha realizada
  hoja.getRange(fila, 7).setValue('Realizado');                             // Estado
  hoja.getRange(fila, 8).setValue(resultado);                               // Resultado
  hoja.getRange(fila, 9).setValue(notas);                                   // Notas
  hoja.getRange(fila, 10).setValue(proximoPaso);                            // Próximo paso
  hoja.getRange(fila, 1, 1, 10).setBackground('#d9ead3');
  Logger.log(`Seguimiento marcado como realizado en fila ${fila}`);
}

/**
 * Actualiza la próxima llamada en la hoja Graduados
 * @param {string} graduadoId
 * @param {string} tipoLlamada
 */
function actualizarProximaLlamada(graduadoId, tipoLlamada) {
  const hoja  = obtenerHoja('Graduados');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === graduadoId) {
      hoja.getRange(i + 1, 13).setValue(tipoLlamada); // col 13 = Próxima llamada
      break;
    }
  }
}

/**
 * Genera HTML con la lista de seguimientos pendientes
 * @param {Array} seguimientos
 * @return {string}
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
    html += `<tr><td><strong>${seg.nombre}</strong></td><td>${seg.tipo}</td>` +
            `<td>${seg.fechaProgramada}</td><td>${seg.notas}</td></tr>`;
  });
  html += '</table>';
  return html;
}

/**
 * Envía notificaciones por email de seguimientos pendientes
 * @param {string} emailDestinatario
 */
function enviarNotificacionesSeguimientos(emailDestinatario) {
  const pendientes = obtenerSeguimientosPendientes();
  if (pendientes.length === 0) return;

  let mensaje = '<h2>Seguimientos Pendientes - Equipito Empleabilidad</h2>';
  mensaje += `<p>Tienes <strong>${pendientes.length}</strong> seguimiento(s) pendiente(s):</p><ul>`;
  pendientes.forEach(seg => {
    mensaje += `<li><strong>${seg.nombre}</strong> - ${seg.tipo} (${seg.fechaProgramada})</li>`;
  });
  mensaje += '</ul><p>Por favor, realiza estos seguimientos lo antes posible.</p>';

  MailApp.sendEmail({
    to: emailDestinatario,
    subject: `⏰ ${pendientes.length} Seguimiento(s) Pendiente(s) - Equipito Empleabilidad`,
    htmlBody: mensaje
  });
  Logger.log(`Notificación enviada a: ${emailDestinatario}`);
}

/**
 * Función que se ejecuta diariamente para enviar notificaciones
 */
function enviarNotificacionesDiarias() {
  const email = PropertiesService.getScriptProperties().getProperty('emailNotificaciones');
  if (email) enviarNotificacionesSeguimientos(email);
}

// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 5: CONFIGURACIÓN Y CREDENCIALES
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Verifica si la configuración está completa
 * @return {boolean}
 */
function verificarConfiguracion() {
  const config = obtenerConfiguracion();
  return !!(config.koboExportUrl && config.koboToken);
}

/**
 * Obtiene la configuración actual del sistema
 * @return {Object}
 */
function obtenerConfiguracion() {
  const props = PropertiesService.getScriptProperties();
  return {
    koboExportUrl:        props.getProperty('KOBO_EXPORT_URL') || '',
    koboToken:            props.getProperty('KOBO_TOKEN')      || '',
    emailNotificaciones:  props.getProperty('EMAIL_NOTIFICACIONES') || '',
    sincronizacionAuto:   props.getProperty('SINCRONIZACION_AUTO') === 'true'
  };
}

/**
 * Guarda la configuración de KoboToolbox
 * @param {string} exportUrl
 * @param {string} token
 * @return {boolean}
 */
function guardarConfiguracionKobo(exportUrl, token) {
  try {
    if (!exportUrl || !exportUrl.includes('kobotoolbox.org')) {
      throw new Error('URL de exportación inválida');
    }
    const props = PropertiesService.getScriptProperties();
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
 * @param {string}  email
 * @param {boolean} activar
 */
function guardarConfiguracionNotificaciones(email, activar) {
  PropertiesService.getScriptProperties().setProperty('EMAIL_NOTIFICACIONES', email);
  if (activar) {
    configurarNotificacionesAutomaticas(email);
  } else {
    desactivarNotificacionesAutomaticas();
  }
  Logger.log(`Notificaciones ${activar ? 'activadas' : 'desactivadas'} para: ${email}`);
}

/**
 * Configura notificaciones automáticas diarias a las 9 AM
 * @param {string} emailDestinatario
 */
function configurarNotificacionesAutomaticas(emailDestinatario) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'enviarNotificacionesDiarias') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarNotificacionesDiarias').timeBased().atHour(9).everyDays(1).create();
  PropertiesService.getScriptProperties().setProperty('emailNotificaciones', emailDestinatario);
  Logger.log('Notificaciones automáticas configuradas para las 9 AM');
}

/**
 * Desactiva las notificaciones automáticas
 */
function desactivarNotificacionesAutomaticas() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarNotificacionesDiarias') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Notificaciones automáticas desactivadas');
}

/**
 * Desactiva la sincronización automática
 */
function desactivarSincronizacionAutomatica() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sincronizacionAutomatica') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Sincronización automática desactivada');
}

/**
 * Prueba la conexión con KoboToolbox
 * @return {Object}
 */
function probarConexionKobo() {
  try {
    const config = obtenerConfiguracion();
    if (!config.koboExportUrl || !config.koboToken) {
      return { exito: false, mensaje: 'Configuración incompleta.' };
    }
    const respuesta = UrlFetchApp.fetch(config.koboExportUrl, {
      method: 'get',
      headers: { 'Authorization': 'Token ' + config.koboToken },
      muteHttpExceptions: true
    });
    const codigo = respuesta.getResponseCode();
    return codigo === 200
      ? { exito: true,  mensaje: 'Conexión exitosa con KoboToolbox' }
      : { exito: false, mensaje: `Error de conexión. Código: ${codigo}` };
  } catch (error) {
    return { exito: false, mensaje: 'Error: ' + error.message };
  }
}

/**
 * Guarda configuración desde el formulario HTML
 * @param {Object} datos
 * @return {Object}
 */
function guardarConfiguracionDesdeFormulario(datos) {
  try {
    if (datos.koboExportUrl && datos.koboToken) {
      guardarConfiguracionKobo(datos.koboExportUrl, datos.koboToken);
    }
    if (datos.emailNotificaciones) {
      guardarConfiguracionNotificaciones(datos.emailNotificaciones, datos.activarNotificaciones);
    }
    return { exito: true, mensaje: 'Configuración guardada correctamente' };
  } catch (error) {
    return { exito: false, mensaje: 'Error al guardar: ' + error.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SECCIÓN 6: UTILIDADES Y FUNCIONES AUXILIARES
// ═══════════════════════════════════════════════════════════════════════════

function formatearFecha(fecha) {
  if (!(fecha instanceof Date)) fecha = new Date(fecha);
  const d = String(fecha.getDate()).padStart(2, '0');
  const m = String(fecha.getMonth() + 1).padStart(2, '0');
  return `${d}/${m}/${fecha.getFullYear()}`;
}

function formatearFechaHora(fecha) {
  if (!(fecha instanceof Date)) fecha = new Date(fecha);
  const h = String(fecha.getHours()).padStart(2, '0');
  const min = String(fecha.getMinutes()).padStart(2, '0');
  return `${formatearFecha(fecha)} ${h}:${min}`;
}

function validarEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function limpiarString(str) {
  if (!str) return '';
  return str.toString().trim().replace(/\s+/g, ' ');
}

function generarIdUnico() {
  return Utilities.getUuid();
}

/**
 * Envía un email HTML
 * @param {string} destinatario
 * @param {string} asunto
 * @param {string} mensaje - HTML
 * @return {boolean}
 */
function enviarEmail(destinatario, asunto, mensaje) {
  try {
    if (!validarEmail(destinatario)) { Logger.log('Email inválido: ' + destinatario); return false; }
    MailApp.sendEmail({ to: destinatario, subject: asunto, htmlBody: mensaje });
    Logger.log(`Email enviado a: ${destinatario}`);
    return true;
  } catch (error) {
    Logger.log('Error al enviar email: ' + error);
    return false;
  }
}

/**
 * Envía notificación de nuevos graduados importados
 * @param {number} cantidad
 */
function enviarNotificacionNuevosGraduados(cantidad) {
  const config = obtenerConfiguracion();
  if (!config.emailNotificaciones) return;

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
 * Parsea una fecha en formato dd/mm/yyyy a objeto Date
 * @param {string} fechaStr
 * @return {Date|null}
 */
function parsearFecha(fechaStr) {
  if (!fechaStr) return null;
  try {
    const p = fechaStr.split('/');
    if (p.length === 3) return new Date(p[2], p[1] - 1, p[0]);
  } catch (e) {
    Logger.log('Error parseando fecha: ' + fechaStr);
  }
  return null;
}

/**
 * Registra un movimiento de etapa en la hoja "Estado actual del participante"
 * y actualiza el resumen de conteos al final de esa hoja.
 * @param {string} creamosId     - ID Creamos del participante (índice [2] en Graduados)
 * @param {string} nombreCompleto - Nombre del participante (índice [3] en Graduados)
 * @param {string} etapa
 * @param {string} nota
 */
function registrarMovimientoEtapa(creamosId, nombreCompleto, etapa, nota) {
  const hoja = obtenerHoja('Estado actual del participante');

  // Agregar la fila de historial
  // Orden: Fecha de registro | Creamos ID | Nombre completo | Etapa | Nota
  hoja.appendRow([
    new Date().toLocaleDateString('es-ES'), // Fecha de registro
    creamosId     || '',                    // Creamos ID
    nombreCompleto,                         // Nombre completo
    etapa,
    nota || ''
  ]);

  // Actualizar el resumen de conteos
  actualizarResumenEstado(hoja);
}

/**
 * Escribe/actualiza el bloque de resumen de conteos por etapa
 * al final de la hoja "Estado actual del participante".
 * Siempre se reescribe para mantenerlo actualizado.
 * @param {Sheet} hoja
 */
function actualizarResumenEstado(hoja) {
  const datos       = hoja.getDataRange().getValues();
  const ultimaFila  = hoja.getLastRow();

  // Contar cuántas personas hay actualmente en cada etapa
  // (se toma la etapa más reciente de cada participante por nombre)
  const ultimaEtapaPorNombre = {};
  // Columnas de Estado actual: [0]=Fecha | [1]=Creamos ID | [2]=Nombre | [3]=Etapa | [4]=Nota
  for (let i = 1; i < datos.length; i++) {
    const nombre = datos[i][2];  // Nombre completo
    const etapa  = datos[i][3];  // Etapa
    // Solo filas válidas (no parte del bloque de resumen)
    if (nombre && etapa && ETAPAS_FLUJO.indexOf(etapa) !== -1) {
      ultimaEtapaPorNombre[nombre] = etapa;
    }
  }

  // Calcular conteos
  const conteos = {};
  ETAPAS_FLUJO.forEach(e => { conteos[e] = 0; });
  Object.values(ultimaEtapaPorNombre).forEach(etapa => {
    if (conteos[etapa] !== undefined) conteos[etapa]++;
  });

  // Buscar dónde empieza el bloque de resumen (si ya existe)
  let inicioResumen = -1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '── RESUMEN ──') {
      inicioResumen = i + 1; // fila en Sheet (1-based)
      break;
    }
  }

  // Borrar resumen anterior si existe
  if (inicioResumen > 0) {
    const filasResumen = ETAPAS_FLUJO.length + 2; // separador + filas + total
    hoja.deleteRows(inicioResumen, filasResumen);
  }

  // Escribir nuevo bloque de resumen al final
  const filaInicio = hoja.getLastRow() + 2; // dejar una fila vacía

  // Separador visual
  const rangoSep = hoja.getRange(filaInicio, 1, 1, 5);
  rangoSep.merge();
  rangoSep.setValue('── RESUMEN ──');
  rangoSep.setBackground('#e91e63')
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');

  // Encabezado del resumen
  hoja.getRange(filaInicio + 1, 1, 1, 2)
      .setValues([['Etapa', 'Total personas']])
      .setFontWeight('bold')
      .setBackground('#fce4ec');

  // Filas de conteo
  let totalGeneral = 0;
  ETAPAS_FLUJO.forEach((etapa, idx) => {
    const fila = filaInicio + 2 + idx;
    hoja.getRange(fila, 1).setValue(etapa);
    hoja.getRange(fila, 2).setValue(conteos[etapa]);
    totalGeneral += conteos[etapa];
  });

  // Fila de total
  const filaTotalIdx = filaInicio + 2 + ETAPAS_FLUJO.length;
  hoja.getRange(filaTotalIdx, 1, 1, 2)
      .setValues([['TOTAL', totalGeneral]])
      .setFontWeight('bold')
      .setBackground('#fce4ec');

  Logger.log('Resumen de estado actualizado');
}

/**
 * Obtiene estadísticas generales del sistema
 * @return {Object}
 */
function obtenerEstadisticasGenerales() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stats = { totalGraduados: 0, porClasificacion: {}, seguimientosPendientes: 0 };

  const hojaGraduados = ss.getSheetByName('Graduados');
  if (hojaGraduados) {
    stats.totalGraduados = hojaGraduados.getLastRow() - 1;
    const datos = hojaGraduados.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const c = datos[i][11]; // col 12 = Clasificación
      if (c) stats.porClasificacion[c] = (stats.porClasificacion[c] || 0) + 1;
    }
  }

  stats.seguimientosPendientes = obtenerSeguimientosPendientes().length;
  return stats;
}

function mostrarProgreso(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, 'Procesando...', 3);
}

function mostrarExito(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, '✅ Éxito', 3);
}

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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Hojas actuales: ' + ss.getSheets().map(s => s.getName()).join(', '));
    Logger.log('Configuración válida: ' + verificarConfiguracion());
    Logger.log('✅ Sistema funcionando correctamente');
  } catch (error) {
    Logger.log('❌ Error en prueba: ' + error);
  }
}
