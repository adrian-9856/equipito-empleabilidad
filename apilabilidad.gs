/**
 * ==========================================================================
 * SISTEMA DE SEGUIMIENTO DE GRADUADOS - EQUIPITO EMPLEABILIDAD
 * ==========================================================================
 *
 * Sistema completo de seguimiento de graduados con integracion a KoboToolbox
 * Archivo unificado con todas las funcionalidades
 *
 * Autor: Equipito Empleabilidad
 * Fecha: 2026-02-03
 * ==========================================================================
 */

// ==========================================================================
// SECCION 1: MENU PRINCIPAL Y FUNCIONES DE INTERFAZ
// ==========================================================================

/**
 * Función que se ejecuta al abrir la hoja de cálculo
 * Crea el menú personalizado en la interfaz
 */
function onOpen(e) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📊 Seguimiento Graduados')
      .addItem('🔄 Importar desde KoboToolbox', 'importarDatosKobo')
      .addSeparator()
      .addItem('📝 Clasificar Graduados', 'mostrarFormularioClasificacion')
      .addItem('📊 Generar Reporte', 'generarReporte')
      .addItem('📞 Ver Seguimientos Pendientes', 'mostrarSeguimientosPendientes')
      .addSeparator()
      .addItem('⚙️ Configurar Credenciales', 'mostrarConfiguracion')
      .addSeparator()
      .addItem('🚀 Instalar Sistema (primera vez)', 'instalarSistema')
      .addItem('🔁 Reinstalar Sistema (borra todo)', 'reinstalarSistema')
      .addToUi();
  } catch (error) {
    Logger.log('onOpen: no se pudo crear el menú — ' + error.message);
  }
}

/**
 * Simple trigger: se ejecuta al editar cualquier celda.
 * Si el usuario cambia la columna "Etapa" (col 12) en Graduados,
 * copia automáticamente al graduado a la hoja de clasificación correspondiente.
 */
/**
 * Simple trigger onEdit: maneja TODAS las etapas EXCEPTO Conexiones Laborales.
 * Conexiones Laborales se maneja en onEditInstalable (trigger instalable)
 * porque el trigger simple NO puede abrir formularios/dialogos.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const hoja = e.range.getSheet();
    if (hoja.getName() !== 'Graduados') return;
    if (e.range.getColumn() !== 15) return;
    if (e.range.getRow() <= 1) return;

    const nuevaEtapa = e.value;
    if (!nuevaEtapa || nuevaEtapa === 'Conexiones Laborales') return;

    const fila       = e.range.getRow();
    const datosGrad  = hoja.getRange(fila, 1, 1, 15).getValues()[0];
    const graduadoId = datosGrad[0];
    const nombre     = datosGrad[3];

    copiarAHojaClasificacion(datosGrad, nuevaEtapa, {});
    generarReporte();

    if (nuevaEtapa === 'Activamente busca trabajo') {
      hoja.getRange(fila, 12).setValue('Si');
      programarSeguimientos(graduadoId, nombre);
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      nombre + ' enviado a: ' + nuevaEtapa, '✅ Clasificado', 3
    );
  } catch (error) {
    Logger.log('onEdit: ' + error);
  }
}

/**
 * Trigger INSTALABLE de onEdit: maneja Conexiones Laborales.
 * Al seleccionar "Conexiones Laborales" en el dropdown de Etapa,
 * limpia la selección y abre el formulario directamente.
 *
 * Los triggers instalables SÍ pueden abrir dialogos (a diferencia del simple onEdit).
 * Se instala automáticamente con instalarSistema() o ejecutando configurarEditTrigger().
 */
function onEditInstalable(e) {
  try {
    if (!e || !e.range) return;
    const hoja = e.range.getSheet();
    if (hoja.getName() !== 'Graduados') return;
    if (e.range.getColumn() !== 15) return;
    if (e.range.getRow() <= 1) return;

    const nuevaEtapa = e.value;
    if (nuevaEtapa !== 'Conexiones Laborales') return;

    const fila      = e.range.getRow();
    const datosGrad = hoja.getRange(fila, 1, 1, 15).getValues()[0];
    const creamosId = datosGrad[2] || '';
    const nombre    = datosGrad[3] || '';

    // Limpiar el dropdown
    e.range.setValue('');

    if (!nombre) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'La fila no tiene nombre.', '⚠️ Sin datos', 3
      );
      return;
    }

    // Datos personales del graduado (Graduados: 4=Género, 5=Edad, 6=Nivel edu, 7=Teléfono)
    const datosExtra = {
      genero:         datosGrad[4] || '',
      edad:           datosGrad[5] || '',
      nivelEducativo: datosGrad[6] || '',
      telefono:       datosGrad[7] || ''
    };

    // Abrir formulario de Conexiones Laborales
    const html = HtmlService.createHtmlOutput(
      _generarHTMLFormConexionLaboral(fila, creamosId, nombre, 'Graduados', datosExtra)
    )
      .setWidth(560)
      .setHeight(720)
      .setTitle('Conexión Laboral');
    SpreadsheetApp.getUi().showModalDialog(html, '💼 Conexión Laboral — ' + nombre);
  } catch (error) {
    Logger.log('onEditInstalable: ' + error);
  }
}

/**
 * Instala el trigger de onEdit para Conexiones Laborales.
 * También se puede ejecutar manualmente si el trigger no está configurado.
 */
function configurarEditTrigger() {
  // Eliminar triggers anteriores de onEditInstalable
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditInstalable') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditInstalable')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  Logger.log('Trigger instalable de onEditInstalable configurado');
}

/**
 * Trigger instalable: úsalo si onOpen no crea el menú automáticamente.
 * Para instalarlo ejecuta: configurarMenuTrigger()
 */
function configurarMenuTrigger() {
  // Eliminar triggers anteriores de onOpen
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onOpen') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();
  Logger.log('Trigger instalable de onOpen configurado correctamente');
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

// -----------------------------------------------------------------------------
// CONSTANTES DE OPCIONES DE DROPDOWN
// -----------------------------------------------------------------------------

// Géneros — opciones del dropdown Género en todas las hojas
const GENEROS = [
  'Hombre',
  'Mujer',
  'Trans hombre',
  'No binario',
  'Otro'
];

// -----------------------------------------------------------------------------
// COLUMNAS COMUNES — van en las 6 hojas de apilabilidad
// -----------------------------------------------------------------------------

/**
 * Columnas que aparecen en TODAS las hojas de apilabilidad (información principal)
 * tipo: 'texto' | 'fecha' | 'siNo' | 'checkbox' | 'dropdown'
 */
const COLUMNAS_COMUNES = [
  { nombre: 'Fecha de ingreso',    ancho: 140, tipo: 'fecha' },
  { nombre: 'Creamos ID',          ancho: 130, tipo: 'texto' },
  { nombre: 'Nombre completo',     ancho: 200, tipo: 'texto' },
  { nombre: 'Número de teléfono',  ancho: 150, tipo: 'texto' },
  { nombre: 'Género',              ancho: 120, tipo: 'dropdown', opciones: GENEROS },
  { nombre: 'Edad',                ancho: 80,  tipo: 'texto' },
  { nombre: 'Nivel educativo',     ancho: 160, tipo: 'texto' }
];

// Etapas del flujo — usadas como opciones de dropdown en varias hojas
const ETAPAS_FLUJO = [
  'Aliados',
  'Plataforma',
  'Derivaciones',
  'Por su cuenta',
  'Paso a paso',
  'Activamente busca trabajo',
  'Conexiones Laborales'
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

// -----------------------------------------------------------------------------
// ESTRUCTURA EXACTA DE COLUMNAS POR HOJA
// -----------------------------------------------------------------------------

const ESTRUCTURA_HOJAS = {

  // -- GRADUADOS --------------------------------------------------------------
  // Columnas:
  //  1=No. | 2=Fecha de envío | 3=Creamos ID | 4=Nombre completo |
  //  5=Género | 6=Edad | 7=Nivel educativo | 8=Número de teléfono |
  //  9=Formación | 10=Cohorte | 11=Fecha de entrevista |
  //  12=Empleado | 13=Próxima llamada | 14=Notas | 15=Etapa
  'Graduados': {
    color: '#1a73e8',
    columnas: [
      { nombre: 'No.',                  ancho: 60,  tipo: 'texto'  },
      { nombre: 'Fecha de envío',       ancho: 140, tipo: 'fecha'  },
      { nombre: 'Creamos ID',           ancho: 130, tipo: 'texto'  },
      { nombre: 'Nombre completo',      ancho: 200, tipo: 'texto'  },
      { nombre: 'Género',               ancho: 120, tipo: 'dropdown', opciones: GENEROS },
      { nombre: 'Edad',                 ancho: 80,  tipo: 'texto'  },
      { nombre: 'Nivel educativo',      ancho: 160, tipo: 'texto'  },
      { nombre: 'Número de teléfono',   ancho: 150, tipo: 'texto'  },
      { nombre: 'Formación',            ancho: 180, tipo: 'texto'  },
      { nombre: 'Cohorte',              ancho: 100, tipo: 'texto'  },
      { nombre: 'Fecha de entrevista',  ancho: 140, tipo: 'fecha'  },
      { nombre: 'Empleado',             ancho: 90,  tipo: 'siNo'   },
      { nombre: 'Próxima llamada',      ancho: 180, tipo: 'texto'  },
      { nombre: 'Notas',                ancho: 300, tipo: 'texto'  },
      { nombre: 'Etapa',                ancho: 200, tipo: 'dropdown', opciones: ETAPAS_FLUJO }
    ]
  },

  // -- 1. ALIADOS -------------------------------------------------------------
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

  // -- 2. PLATAFORMA ----------------------------------------------------------
  'Plataforma': {
    color: '#00bcd4',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'Cita',              ancho: 90,  tipo: 'checkbox' },
      { nombre: 'Creación de perfil',ancho: 160, tipo: 'siNo'     },
      { nombre: 'Contacto',          ancho: 90,  tipo: 'checkbox' },
      { nombre: 'Trámites',          ancho: 200, tipo: 'texto'    },
      { nombre: 'Entrevista',        ancho: 130, tipo: 'texto'    },
      { nombre: 'Confirmación',      ancho: 130, tipo: 'texto'    },
      { nombre: 'Recepción',         ancho: 130, tipo: 'texto'    },
      { nombre: 'Nota',              ancho: 300, tipo: 'texto'    },
      { nombre: 'Activo',            ancho: 90,  tipo: 'siNo'     }
    ]
  },

  // -- 3. DERIVACIONES -------------------------------------------------------
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

  // -- 4. POR SU CUENTA ------------------------------------------------------
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

  // -- 5. PASO A PASO --------------------------------------------------------
  'Paso a paso': {
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

  // -- 6. ACTIVAMENTE BUSCA TRABAJO ------------------------------------------
  'Activamente busca trabajo': {
    color: '#0f9d58',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'Entrevista', ancho: 130, tipo: 'texto' },
      { nombre: 'Trámites',   ancho: 200, tipo: 'texto' },
      { nombre: 'Activo',     ancho: 90,  tipo: 'siNo'  }
    ]
  },

  // -- CONEXIONES LABORALES --------------------------------------------------
  'Conexiones Laborales': {
    color: '#e65100',
    columnas: [
      { nombre: 'Creamos ID',                    ancho: 130, tipo: 'texto' },
      { nombre: 'Nombre completo',               ancho: 200, tipo: 'texto' },
      { nombre: 'Número de teléfono',            ancho: 150, tipo: 'texto' },
      { nombre: 'Género',                         ancho: 120, tipo: 'dropdown', opciones: GENEROS },
      { nombre: 'Edad',                          ancho: 80,  tipo: 'texto' },
      { nombre: 'Nivel educativo',               ancho: 160, tipo: 'texto' },
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

  // -- SEGUIMIENTOS -----------------------------------------------------------
  'Seguimientos': {
    color: '#673ab7',
    columnas: [
      { nombre: 'No.',                ancho: 60,  tipo: 'texto' },
      { nombre: 'Creamos ID',         ancho: 130, tipo: 'texto' },
      { nombre: 'Nombre completo',    ancho: 200, tipo: 'texto' },
      { nombre: 'Número de teléfono', ancho: 150, tipo: 'texto' },
      { nombre: 'Género',             ancho: 120, tipo: 'dropdown', opciones: GENEROS },
      { nombre: 'Edad',               ancho: 80,  tipo: 'texto' },
      { nombre: 'Nivel educativo',    ancho: 160, tipo: 'texto' },
      { nombre: 'Tipo seguimiento', ancho: 180, tipo: 'texto' },
      { nombre: 'Fecha programada', ancho: 150, tipo: 'fecha' },
      { nombre: 'Fecha realizada',  ancho: 150, tipo: 'fecha' },
      { nombre: 'Estado',           ancho: 120, tipo: 'texto' },
      { nombre: 'Resultado',        ancho: 220, tipo: 'texto' },
      { nombre: 'Notas',            ancho: 300, tipo: 'texto' },
      { nombre: 'Próximo paso',     ancho: 220, tipo: 'texto' }
    ]
  },

  // -- SEGUIMIENTO BOT (para n8n + WhatsApp) --------------------------------
  'Seguimiento Bot': {
    color: '#00897b',
    columnas: [
      { nombre: 'Creamos ID',         ancho: 130, tipo: 'texto' },
      { nombre: 'Nombre',             ancho: 200, tipo: 'texto' },
      { nombre: 'Telefono',           ancho: 150, tipo: 'texto' },
      { nombre: 'Empresa',            ancho: 200, tipo: 'texto' },
      { nombre: 'Cargo',              ancho: 200, tipo: 'texto' },
      { nombre: 'Fecha Empleo',       ancho: 140, tipo: 'fecha' },
      { nombre: 'Fecha Seg1',         ancho: 140, tipo: 'fecha' },
      { nombre: 'Fecha Seg2',         ancho: 140, tipo: 'fecha' },
      { nombre: 'Fecha Recordatorio', ancho: 160, tipo: 'fecha' },
      { nombre: 'Estado',             ancho: 160, tipo: 'texto' },
      { nombre: 'Etapa Actual',       ancho: 130, tipo: 'texto' },
      { nombre: 'Resp S1 P1',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S1 P2',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S1 P3',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S2 P1',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S2 P2',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S2 P3',         ancho: 300, tipo: 'texto' },
      { nombre: 'Email Enviado',      ancho: 130, tipo: 'siNo'  }
    ]
  },

  // -- REPORTE ---------------------------------------------------------------
  // Hoja de reporte automático — no tiene columnas editables por el usuario
  // Se genera/actualiza automáticamente con generarReporte()
  'Reporte': {
    color: '#e91e63',
    columnas: []
  }
};

// -----------------------------------------------------------------------------
// FUNCIONES DE INSTALACIÓN Y REINSTALACIÓN
// -----------------------------------------------------------------------------

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

    // -- 1. Eliminar hojas existentes si se solicita --------------------------
    if (borrarExistentes) {
      const hojasDelSistema = [
        // Hojas actuales
        'Graduados',
        'Aliados',
        'Plataforma',
        'Derivaciones',
        'Por su cuenta',
        'Paso a paso',
        'Activamente busca trabajo',
        'Seguimientos',
        'Seguimiento Bot',
        'Reporte',
        'Conexiones Laborales',
        'Configuración',
        // Nombres legacy (por si acaso existen)
        'Reportes mensuales',
        'Reportes Mensuales',
        'Por su Cuenta',
        'Busca Trabajo (Fito)',
        'No Busca Trabajar',
        'No busca trabajo - Fito',
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

    // -- 2. Crear cada hoja con su estructura exacta --------------------------
    const ordenHojas = [
      'Graduados',
      'Aliados',
      'Plataforma',
      'Derivaciones',
      'Por su cuenta',
      'Paso a paso',
      'Activamente busca trabajo',
      'Conexiones Laborales',
      'Seguimientos',
      'Seguimiento Bot',
      'Reporte'
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

    // -- 3. Crear hoja Configuración ------------------------------------------
    _crearHojaConfiguracion(ss);

    // -- 4. Eliminar hoja temporal si existe ----------------------------------
    const temp = ss.getSheetByName('_temp_');
    if (temp) ss.deleteSheet(temp);

    // -- 5. Ordenar hojas -----------------------------------------------------
    _ordenarHojas(ss, [...ordenHojas, 'Configuración']);

    // -- 6. Instalar trigger de onEdit para Conexiones Laborales -------------
    configurarEditTrigger();

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
  if (!columnas || columnas.length === 0) return; // Hoja sin columnas (ej. Reporte)

  const headers  = columnas.map(c => c.nombre);

  // -- Escribir headers ------------------------------------------------------
  const rangoHeader = hoja.getRange(1, 1, 1, headers.length);
  rangoHeader.setValues([headers]);

  // -- Formato de headers ----------------------------------------------------
  rangoHeader
    .setBackground(estructura.color)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  hoja.setRowHeight(1, 36);

  // -- Anchos de columna -----------------------------------------------------
  columnas.forEach((col, i) => {
    hoja.setColumnWidth(i + 1, col.ancho);
  });

  // -- Congelar primera fila -------------------------------------------------
  hoja.setFrozenRows(1);

  // -- Quitar columnas sobrantes ---------------------------------------------
  const totalCols = hoja.getMaxColumns();
  if (totalCols > headers.length) {
    hoja.deleteColumns(headers.length + 1, totalCols - headers.length);
  }

  // -- Agregar filtros -------------------------------------------------------
  rangoHeader.createFilter();

  // -- Validaciones por tipo de columna -------------------------------------
  // Solo aplicar validaciones de dropdown/siNo (NO checkboxes, que llenan
  // las celdas con FALSE y causan problemas con appendRow y conteos)
  columnas.forEach((col, i) => {
    const colNum = i + 1;
    const rango  = hoja.getRange(2, colNum, 100);

    if (col.tipo === 'siNo') {
      const regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Si', 'No'], true)
        .setAllowInvalid(false)
        .build();
      rango.setDataValidation(regla);

    } else if (col.tipo === 'dropdown' && col.opciones && col.opciones.length) {
      const regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(col.opciones, true)
        .setAllowInvalid(false)
        .build();
      rango.setDataValidation(regla);

    } else if (col.tipo === 'fecha') {
      rango.setNumberFormat('dd/mm/yyyy');
    }
    // Checkboxes: NO aplicar requireCheckbox() porque llena con FALSE.
    // Se aplicaran manualmente cuando la fila tenga datos.
  });

  // -- Color alterno en filas de datos --------------------------------------
  try {
    hoja.getRange(2, 1, 100, headers.length)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  } catch (e) { /* ignorar si no esta disponible */ }

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
      const etapa = datos[i][14]; // col 15 = Etapa
      if (!etapa || etapa.toString().trim() === '') {
        sin.push({ id: datos[i][0], nombre: datos[i][3] }); // [3] = Nombre completo
      }
    }
    return sin;
  } catch (error) {
    Logger.log('Error al obtener graduados sin clasificar: ' + error);
    return [];
  }
}

// ===========================================================================
// SECCIÓN 2: INTEGRACIÓN CON KOBOTOOLBOX API
// ===========================================================================

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

// ===========================================================================
// SECCIÓN 3: GESTIÓN DE HOJAS Y DATOS
// ===========================================================================

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
 * Columnas (15):
 *   1=No. | 2=Fecha de envío | 3=Creamos ID | 4=Nombre completo |
 *   5=Género | 6=Edad | 7=Nivel educativo | 8=Número de teléfono |
 *   9=Formación | 10=Cohorte | 11=Fecha de entrevista |
 *   12=Empleado | 13=Próxima llamada | 14=Notas | 15=Etapa
 * @param {Object} graduado
 */
function agregarGraduado(graduado) {
  const hoja = obtenerHoja('Graduados');
  const fila = [
    graduado.id,                            // 1  No.
    new Date().toLocaleDateString('es-ES'), // 2  Fecha de envío
    '',                                     // 3  Creamos ID (manual)
    graduado.nombre,                        // 4  Nombre completo
    graduado.genero || '',                  // 5  Género
    graduado.edad || '',                    // 6  Edad
    graduado.nivelEducativo || '',          // 7  Nivel educativo
    graduado.telefono,                      // 8  Número de teléfono
    graduado.formacion,                     // 9  Formación
    graduado.cohorte || '',                 // 10 Cohorte
    graduado.fechaEntrevista,               // 11 Fecha de entrevista
    'No',                                   // 12 Empleado
    '',                                     // 13 Próxima llamada
    '',                                     // 14 Notas
    ''                                      // 15 Etapa (dropdown)
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
      // col 1-3 (No., Fecha envío, Creamos ID) → no tocar
      hoja.getRange(i + 1, 4).setValue(graduado.nombre);
      hoja.getRange(i + 1, 5).setValue(graduado.genero          || datos[i][4]);
      hoja.getRange(i + 1, 6).setValue(graduado.edad            || datos[i][5]);
      hoja.getRange(i + 1, 7).setValue(graduado.nivelEducativo  || datos[i][6]);
      hoja.getRange(i + 1, 8).setValue(graduado.telefono);
      hoja.getRange(i + 1, 9).setValue(graduado.formacion       || datos[i][8]);
      hoja.getRange(i + 1, 10).setValue(graduado.cohorte        || datos[i][9]);
      hoja.getRange(i + 1, 11).setValue(graduado.fechaEntrevista || datos[i][10]);
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
      // Columnas: 12=Empleado | 15=Etapa
      hojaGraduados.getRange(i + 1, 12).setValue(
        clasificacion === 'Activamente busca trabajo' ? 'Sí' : 'No'
      );
      hojaGraduados.getRange(i + 1, 15).setValue(clasificacion);             // Etapa
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
    'Paso a paso':                'Paso a paso',
    // legacy
    'No busca trabajo':           'Paso a paso',
    'Fito':                       'Paso a paso',
    'No busca trabajo - Fito':    'Paso a paso',
    'Activamente busca trabajo':  'Activamente busca trabajo',
    'Conexiones Laborales':       'Conexiones Laborales',
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
 *   4=Género | 5=Edad | 6=Nivel educativo | 7=Número de teléfono |
 *   8=Formación | 9=Cohorte | 10=Fecha entrevista | 11=Empleado |
 *   12=Próxima llamada | 13=Notas | 14=Etapa
 *
 * Columnas comunes que van en las 6 hojas de apilabilidad (7 cols):
 *   Fecha de ingreso | Creamos ID | Nombre completo | Número de teléfono | Género | Edad | Nivel educativo
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
    datosGraduado[7] || datosAdicionales.telefono || '',  // Número de teléfono (índice 7)
    datosGraduado[4] || datosAdicionales.genero   || '',  // Género (índice 4)
    datosGraduado[5] || datosAdicionales.edad     || '',  // Edad (índice 5)
    datosGraduado[6] || datosAdicionales.nivelEdu || ''   // Nivel educativo (índice 6)
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

    case 'Paso a paso':
    case 'No busca trabajo':
    case 'Fito':
      return filaBase.concat([
        datosAdicionales.dpi       || '',
        datosAdicionales.telefono  || datosGraduado[7] || '',  // índice 7 = Número de teléfono
        datosAdicionales.formacion || datosGraduado[8] || '',  // índice 8 = Formación
        datosAdicionales.cohorte   || datosGraduado[9] || '',  // índice 9 = Cohorte
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

    case 'Conexiones Laborales':
      // Conexiones Laborales: Creamos ID | Nombre completo | Teléfono | Género | Edad | Nivel edu |
      //   Tipo | Programa | Proyecto | Especialidad | Empresa | Cargo | ...
      return [
        datosGraduado[2] || '',                       // Creamos ID
        datosGraduado[3] || '',                       // Nombre completo
        datosGraduado[7] || datosAdicionales.telefono || '', // Número de teléfono
        datosGraduado[4] || datosAdicionales.genero   || '', // Género
        datosGraduado[5] || datosAdicionales.edad     || '', // Edad
        datosGraduado[6] || datosAdicionales.nivelEdu || '', // Nivel educativo
        '', '', '', '',                               // Tipo, Programa, Proyecto, Especialidad
        datosAdicionales.empresa  || '',              // Empresa
        datosAdicionales.cargo    || '',              // Cargo
        '', '',                                       // Tipo duracion, Tipo contrato
        '', '', '', ''                                // Fechas, Duracion, Salario
      ];

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

// ===========================================================================
// SECCIÓN 3B: CONEXIONES LABORALES
// ===========================================================================

/**
 * Abre el formulario de Conexiones Laborales para el graduado seleccionado.
 * El usuario debe tener seleccionada una fila en la hoja "Graduados".
 */
function enviarAConexionesLaborales() {
  const ui   = SpreadsheetApp.getUi();
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const nombreHoja = hoja.getName();

  // Hojas válidas para enviar a Conexiones Laborales
  const hojasValidas = [
    'Graduados', 'Aliados', 'Plataforma', 'Derivaciones',
    'Por su cuenta', 'Paso a paso', 'Activamente busca trabajo'
  ];

  if (hojasValidas.indexOf(nombreHoja) === -1) {
    ui.alert('⚠️ Hoja incorrecta',
             'Debes estar en una hoja de participantes y seleccionar una fila.\n\nHojas válidas: ' + hojasValidas.join(', '),
             ui.ButtonSet.OK);
    return;
  }

  const filaActiva = hoja.getActiveRange().getRow();
  if (filaActiva <= 1) {
    ui.alert('⚠️ Selección inválida',
             'Selecciona la fila de un participante (no la fila de encabezados).',
             ui.ButtonSet.OK);
    return;
  }

  // Leer datos según la hoja activa
  // Graduados: 2=Creamos ID, 3=Nombre, 4=Género, 5=Edad, 6=Nivel edu, 7=Teléfono
  // Hojas clasificación (COLUMNAS_COMUNES): 1=Creamos ID, 2=Nombre, 3=Teléfono, 4=Género, 5=Edad, 6=Nivel edu
  var creamosId, nombreCompleto, datosExtra;
  if (nombreHoja === 'Graduados') {
    const datos = hoja.getRange(filaActiva, 1, 1, 15).getValues()[0];
    creamosId      = datos[2] || '';
    nombreCompleto = datos[3] || '';
    datosExtra = {
      genero:         datos[4] || '',
      edad:           datos[5] || '',
      nivelEducativo: datos[6] || '',
      telefono:       datos[7] || ''
    };
  } else {
    const datos = hoja.getRange(filaActiva, 1, 1, 7).getValues()[0];
    creamosId      = datos[1] || '';  // Col 2 = Creamos ID
    nombreCompleto = datos[2] || '';  // Col 3 = Nombre completo
    datosExtra = {
      telefono:       datos[3] || '',
      genero:         datos[4] || '',
      edad:           datos[5] || '',
      nivelEducativo: datos[6] || ''
    };
  }

  if (!nombreCompleto) {
    ui.alert('⚠️ Sin datos', 'La fila seleccionada no tiene nombre.', ui.ButtonSet.OK);
    return;
  }

  // Generar y mostrar el formulario HTML
  const html = HtmlService.createHtmlOutput(
    _generarHTMLFormConexionLaboral(filaActiva, creamosId, nombreCompleto, nombreHoja, datosExtra)
  )
    .setWidth(560)
    .setHeight(720)
    .setTitle('Conexión Laboral');
  ui.showModalDialog(html, '💼 Conexión Laboral — ' + nombreCompleto);
}

/**
 * Genera el HTML del formulario de Conexiones Laborales
 * @param {number} filaGraduado
 * @param {string} creamosId
 * @param {string} nombre
 * @param {string} hojaOrigen - nombre de la hoja desde donde se abrió
 * @param {Object} datosExtra - { telefono, genero, edad, nivelEducativo }
 * @return {string}
 */
function _generarHTMLFormConexionLaboral(filaGraduado, creamosId, nombre, hojaOrigen, datosExtra) {
  datosExtra = datosExtra || {};
  var telefono       = datosExtra.telefono       || '';
  var genero         = datosExtra.genero         || '';
  var edad           = datosExtra.edad           || '';
  var nivelEducativo = datosExtra.nivelEducativo || '';

  // Fecha de hoy en formato yyyy-mm-dd para el input date
  var hoy = new Date();
  var hoyStr = hoy.getFullYear() + '-' +
    String(hoy.getMonth() + 1).padStart(2, '0') + '-' +
    String(hoy.getDate()).padStart(2, '0');

  return `
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; background: #fafafa; overflow-y: auto; }
      h3 { color: #e65100; margin-top: 0; }
      .info { background: #fff3e0; padding: 10px; border-radius: 6px; margin-bottom: 14px; }
      .info strong { color: #e65100; }
      .info-dato { display: inline-block; margin-right: 16px; margin-top: 4px; }
      label { display: block; font-weight: bold; margin: 10px 0 4px; font-size: 13px; }
      input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;
                      box-sizing: border-box; font-size: 13px; }
      input:focus, select:focus { border-color: #e65100; outline: none; }
      input[readonly] { background: #f0f0f0; color: #666; }
      .row2 { display: flex; gap: 12px; }
      .row2 > div { flex: 1; }
      .seccion { background: #f5f5f5; padding: 10px; border-radius: 6px; margin: 14px 0 8px;
                 border-left: 3px solid #e65100; }
      .seccion-titulo { font-weight: bold; color: #e65100; font-size: 13px; margin: 0; }
      .btn { margin-top: 18px; text-align: right; }
      .btn button { padding: 10px 24px; border: none; border-radius: 4px; font-size: 14px;
                    cursor: pointer; }
      .btn-ok  { background: #e65100; color: #fff; }
      .btn-ok:hover { background: #bf360c; }
      .btn-cancel { background: #eee; color: #333; margin-right: 8px; }
      .msg { padding: 10px; margin-top: 10px; border-radius: 4px; display: none; }
      .msg-ok  { background: #c8e6c9; color: #2e7d32; }
      .msg-err { background: #ffcdd2; color: #c62828; }
      .duracion-auto { font-size: 12px; color: #666; margin-top: 2px; }
    </style>

    <h3>Registro de Conexion Laboral</h3>
    <div class="info">
      <strong>Graduado:</strong> ${nombre}<br>
      <strong>Creamos ID:</strong> ${creamosId || '(sin asignar)'}<br>
      <span class="info-dato"><strong>Tel:</strong> ${telefono || '-'}</span>
      <span class="info-dato"><strong>Genero:</strong> ${genero || '-'}</span>
      <span class="info-dato"><strong>Edad:</strong> ${edad || '-'}</span>
      <span class="info-dato"><strong>Nivel edu:</strong> ${nivelEducativo || '-'}</span>
    </div>

    <div class="seccion"><p class="seccion-titulo">Informacion del programa</p></div>

    <div class="row2">
      <div>
        <label>Tipo</label>
        <input id="tipo" placeholder="Ej: Insercion, Emprendimiento">
      </div>
      <div>
        <label>Programa</label>
        <input id="programa" placeholder="Nombre del programa">
      </div>
    </div>

    <div class="row2">
      <div>
        <label>Proyecto</label>
        <input id="proyecto" placeholder="Nombre del proyecto">
      </div>
      <div>
        <label>Especialidad</label>
        <input id="especialidad" placeholder="Especialidad o area">
      </div>
    </div>

    <div class="seccion"><p class="seccion-titulo">Datos del empleo</p></div>

    <label>Empresa *</label>
    <input id="empresa" placeholder="Nombre de la empresa">

    <label>Cargo que desempena *</label>
    <input id="cargo" placeholder="Puesto o cargo">

    <div class="row2">
      <div>
        <label>Tipo de duracion de contrato</label>
        <input id="tipoDuracion" placeholder="Ej: Temporal, Indefinido">
      </div>
      <div>
        <label>Tipo de contrato</label>
        <input id="tipoContrato" placeholder="Ej: Tiempo completo, Medio tiempo">
      </div>
    </div>

    <div class="row2">
      <div>
        <label>Fecha de inicio</label>
        <input id="fechaInicio" type="date" value="${hoyStr}" onchange="calcularDuracion()">
      </div>
      <div>
        <label>Fecha de final (opcional)</label>
        <input id="fechaFinal" type="date" onchange="calcularDuracion()">
      </div>
    </div>

    <div class="row2">
      <div>
        <label>Duracion (meses)</label>
        <input id="duracion" type="text" readonly placeholder="Se calcula automaticamente">
        <div class="duracion-auto" id="duracionInfo">Se calcula con las fechas de inicio y final</div>
      </div>
      <div>
        <label>Salario mensual</label>
        <input id="salario" placeholder="Q 0.00">
      </div>
    </div>

    <div id="mensaje" class="msg"></div>

    <div class="btn">
      <button class="btn-cancel" onclick="google.script.host.close()">Cancelar</button>
      <button class="btn-ok" id="btnGuardar" onclick="guardar()">Guardar</button>
    </div>

    <script>
      function calcularDuracion() {
        var inicio = document.getElementById('fechaInicio').value;
        var final  = document.getElementById('fechaFinal').value;
        var campo  = document.getElementById('duracion');
        var info   = document.getElementById('duracionInfo');

        if (!inicio) {
          campo.value = '';
          info.textContent = 'Ingresa la fecha de inicio';
          return;
        }

        if (!final) {
          // Si no hay fecha final, calcular desde inicio hasta hoy
          var hoy = new Date();
          var fi  = new Date(inicio);
          if (fi > hoy) {
            campo.value = '0';
            info.textContent = 'Aun no ha iniciado';
            return;
          }
          var meses = (hoy.getFullYear() - fi.getFullYear()) * 12 + (hoy.getMonth() - fi.getMonth());
          var dias  = hoy.getDate() - fi.getDate();
          if (dias < 0) meses--;
          if (meses < 1) {
            var diffDias = Math.floor((hoy - fi) / (1000 * 60 * 60 * 24));
            var semanas  = Math.floor(diffDias / 7);
            campo.value = meses <= 0 ? '0' : String(meses);
            info.textContent = 'Sin fecha final: ' + diffDias + ' dias (' + semanas + ' semanas) hasta hoy';
          } else {
            campo.value = String(meses);
            info.textContent = 'Sin fecha final: calculado hasta hoy (' + meses + ' meses)';
          }
          return;
        }

        // Ambas fechas disponibles
        var fi = new Date(inicio);
        var ff = new Date(final);
        if (ff < fi) {
          campo.value = '';
          info.textContent = 'La fecha final no puede ser antes de la de inicio';
          return;
        }
        var meses = (ff.getFullYear() - fi.getFullYear()) * 12 + (ff.getMonth() - fi.getMonth());
        var dias  = ff.getDate() - fi.getDate();
        if (dias < 0) meses--;
        if (meses < 1) {
          var diffDias = Math.floor((ff - fi) / (1000 * 60 * 60 * 24));
          campo.value = '0';
          info.textContent = diffDias + ' dias (menos de 1 mes)';
        } else {
          campo.value = String(meses);
          info.textContent = meses + ' meses calculados automaticamente';
        }
      }

      // Calcular al cargar el formulario
      calcularDuracion();

      function guardar() {
        var empresa = document.getElementById('empresa').value.trim();
        var cargo   = document.getElementById('cargo').value.trim();
        if (!empresa || !cargo) {
          mostrarMsg('Empresa y Cargo son obligatorios.', true);
          return;
        }
        document.getElementById('btnGuardar').disabled = true;
        document.getElementById('btnGuardar').textContent = 'Guardando...';

        var datos = {
          filaGraduado:   ${filaGraduado},
          hojaOrigen:     '${hojaOrigen}',
          creamosId:      '${creamosId}',
          nombreCompleto: '${nombre}',
          telefono:       '${telefono}',
          genero:         '${genero}',
          edad:           '${edad}',
          nivelEducativo: '${nivelEducativo}',
          tipo:           document.getElementById('tipo').value.trim(),
          programa:       document.getElementById('programa').value.trim(),
          proyecto:       document.getElementById('proyecto').value.trim(),
          especialidad:   document.getElementById('especialidad').value.trim(),
          empresa:        empresa,
          cargo:          cargo,
          tipoDuracion:   document.getElementById('tipoDuracion').value.trim(),
          tipoContrato:   document.getElementById('tipoContrato').value.trim(),
          fechaInicio:    document.getElementById('fechaInicio').value,
          fechaFinal:     document.getElementById('fechaFinal').value,
          duracion:       document.getElementById('duracion').value,
          salario:        document.getElementById('salario').value.trim()
        };

        google.script.run
          .withSuccessHandler(function(r) {
            if (r.exito) {
              mostrarMsg(r.mensaje, false);
              setTimeout(function() { google.script.host.close(); }, 1500);
            } else {
              mostrarMsg(r.mensaje, true);
              document.getElementById('btnGuardar').disabled = false;
              document.getElementById('btnGuardar').textContent = 'Guardar';
            }
          })
          .withFailureHandler(function(e) {
            mostrarMsg('Error: ' + e.message, true);
            document.getElementById('btnGuardar').disabled = false;
            document.getElementById('btnGuardar').textContent = 'Guardar';
          })
          .guardarConexionLaboral(datos);
      }

      function mostrarMsg(texto, esError) {
        var el = document.getElementById('mensaje');
        el.textContent = texto;
        el.className = 'msg ' + (esError ? 'msg-err' : 'msg-ok');
        el.style.display = 'block';
      }
    </script>`;
}

/**
 * Guarda los datos del formulario en la hoja "Conexiones Laborales".
 *
 * Columnas Conexiones Laborales (18):
 *  1=Creamos ID | 2=Nombre completo | 3=Teléfono | 4=Género | 5=Edad | 6=Nivel educativo |
 *  7=Tipo | 8=Programa | 9=Proyecto | 10=Especialidad | 11=Empresa |
 *  12=Cargo | 13=Tipo duracion contrato | 14=Tipo contrato |
 *  15=Fecha inicio | 16=Fecha final | 17=Duracion (meses) | 18=Salario mensual
 *
 * @param {Object} datos - datos del formulario
 * @return {Object}
 */
function guardarConexionLaboral(datos) {
  try {
    const creamosId      = datos.creamosId || '';
    const nombreCompleto = datos.nombreCompleto || '';

    // Formatear fechas de yyyy-mm-dd (input date) a dd/mm/yyyy
    var fechaInicio = '';
    if (datos.fechaInicio) {
      const p = datos.fechaInicio.split('-');
      fechaInicio = p[2] + '/' + p[1] + '/' + p[0];
    }
    var fechaFinal = '';
    if (datos.fechaFinal) {
      const p = datos.fechaFinal.split('-');
      fechaFinal = p[2] + '/' + p[1] + '/' + p[0];
    }

    const fila = [
      creamosId,                        // 1  Creamos ID
      nombreCompleto,                   // 2  Nombre completo
      datos.telefono       || '',       // 3  Número de teléfono
      datos.genero         || '',       // 4  Género
      datos.edad           || '',       // 5  Edad
      datos.nivelEducativo || '',       // 6  Nivel educativo
      datos.tipo           || '',       // 7  Tipo
      datos.programa       || '',       // 8  Programa
      datos.proyecto       || '',       // 9  Proyecto
      datos.especialidad   || '',       // 10 Especialidad
      datos.empresa,                    // 11 Empresa
      datos.cargo,                      // 12 Cargo que desempena
      datos.tipoDuracion   || '',       // 13 Tipo de duracion de contrato
      datos.tipoContrato   || '',       // 14 Tipo de contrato
      fechaInicio,                      // 15 Fecha de inicio
      fechaFinal,                       // 16 Fecha de final
      datos.duracion       || '',       // 17 Duracion (meses)
      datos.salario        || ''        // 18 Salario mensual
    ];

    obtenerHoja('Conexiones Laborales').appendRow(fila);
    Logger.log('Conexión laboral registrada: ' + nombreCompleto + ' en ' + datos.empresa);

    // Crear fila en Seguimiento Bot para n8n/WhatsApp
    crearFilaSeguimientoBot(datos, fechaInicio);

    return { exito: true, mensaje: 'Conexión laboral guardada. Seguimiento WhatsApp programado.' };
  } catch (error) {
    Logger.log('Error al guardar conexión laboral: ' + error);
    return { exito: false, mensaje: 'Error al guardar: ' + error.message };
  }
}

/**
 * Crea una fila en "Seguimiento Bot" con fechas calculadas.
 * Esta hoja es leída por n8n para disparar los mensajes de WhatsApp.
 *
 * Día 0   -> WhatsApp Seguimiento 1 (3 preguntas)
 * Día 14  -> WhatsApp Seguimiento 2 (3 preguntas)
 * Día 90  -> Email a adrian@creamosguatemala.org para llamada de cierre
 */
function crearFilaSeguimientoBot(datos, fechaEmpleo) {
  try {
    const hoja = obtenerHoja('Seguimiento Bot');

    const hoy         = new Date();
    const fechaSeg1   = new Date(hoy);                      // Día 0 = hoy
    const fechaSeg2   = new Date(hoy);
    fechaSeg2.setDate(fechaSeg2.getDate() + 14);            // Día 14
    const fechaRecord = new Date(hoy);
    fechaRecord.setDate(fechaRecord.getDate() + 90);        // Día 90

    function fmt(d) {
      return ('0' + d.getDate()).slice(-2) + '/' +
             ('0' + (d.getMonth() + 1)).slice(-2) + '/' +
             d.getFullYear();
    }

    // Asegurar formato +502XXXXXXXX
    var telefono = (datos.telefono || '').toString().trim();
    if (telefono && !telefono.startsWith('+')) {
      telefono = '+' + telefono;
    }

    const fila = [
      datos.creamosId      || '',   // Creamos ID
      datos.nombreCompleto || '',   // Nombre
      telefono,                     // Telefono
      datos.empresa        || '',   // Empresa
      datos.cargo          || '',   // Cargo
      fechaEmpleo,                  // Fecha Empleo
      fmt(fechaSeg1),               // Fecha Seg1 (hoy)
      fmt(fechaSeg2),               // Fecha Seg2 (+14 días)
      fmt(fechaRecord),             // Fecha Recordatorio (+90 días)
      'Nuevo',                      // Estado (n8n detecta esto)
      '',                           // Etapa Actual
      '', '', '',                   // Resp S1 P1, P2, P3
      '', '', '',                   // Resp S2 P1, P2, P3
      'No'                          // Email Enviado
    ];

    hoja.appendRow(fila);
    Logger.log('Seguimiento Bot creado para: ' + datos.nombreCompleto);

  } catch (error) {
    Logger.log('Error al crear fila en Seguimiento Bot: ' + error);
    // No lanzar el error para no interrumpir el guardado de Conexiones Laborales
  }
}

// ===========================================================================
// SECCIÓN 4: SEGUIMIENTOS Y LLAMADAS PROGRAMADAS
// ===========================================================================

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

    // Columnas Seguimientos: No. | Creamos ID | Nombre completo | Teléfono | Género | Edad | Nivel educativo | Tipo | Fecha prog. | Fecha real. | Estado | Resultado | Notas | Próximo paso
    hojaSeguimientos.appendRow([
      graduadoId,     // No.
      '',             // Creamos ID (manual)
      nombreGraduado, // Nombre completo
      '',             // Número de teléfono (manual)
      '',             // Género (manual)
      '',             // Edad (manual)
      '',             // Nivel educativo (manual)
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
  // Seguimientos: [0]=No. | [1]=Creamos ID | [2]=Nombre | [3]=Género | [4]=Edad |
  //   [5]=Nivel edu | [6]=Tipo | [7]=Fecha prog. | [8]=Fecha real. | [9]=Estado |
  //   [10]=Resultado | [11]=Notas | [12]=Próximo paso
  for (let i = 1; i < datos.length; i++) {
    const estado             = datos[i][9];
    const fechaProgramadaStr = datos[i][7];
    if (estado === 'Pendiente' && fechaProgramadaStr) {
      const fechaProgramada = parsearFecha(fechaProgramadaStr);
      if (fechaProgramada && fechaProgramada <= hoy) {
        pendientes.push({
          fila:            i + 1,
          id:              datos[i][0],
          nombre:          datos[i][2],
          tipo:            datos[i][6],
          fechaProgramada: fechaProgramadaStr,
          notas:           datos[i][11]
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
  // Seguimientos: col 1=No. | 2=Creamos ID | 3=Nombre | 4=Género | 5=Edad |
  //   6=Nivel edu | 7=Tipo | 8=Fecha prog. | 9=Fecha real. | 10=Estado |
  //   11=Resultado | 12=Notas | 13=Próximo paso
  const hoja = obtenerHoja('Seguimientos');
  hoja.getRange(fila, 9).setValue(new Date().toLocaleDateString('es-ES'));  // Fecha realizada
  hoja.getRange(fila, 10).setValue('Realizado');                            // Estado
  hoja.getRange(fila, 11).setValue(resultado);                              // Resultado
  hoja.getRange(fila, 12).setValue(notas);                                  // Notas
  hoja.getRange(fila, 13).setValue(proximoPaso);                            // Próximo paso
  hoja.getRange(fila, 1, 1, 13).setBackground('#d9ead3');
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

// ===========================================================================
// SECCIÓN 5: CONFIGURACIÓN Y CREDENCIALES
// ===========================================================================

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

// ===========================================================================
// SECCIÓN 6: UTILIDADES Y FUNCIONES AUXILIARES
// ===========================================================================

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
 * Registra un movimiento de etapa — ahora solo actualiza el reporte.
 * Se mantiene la firma para compatibilidad con clasificarGraduado().
 */
function registrarMovimientoEtapa(creamosId, nombreCompleto, etapa, nota) {
  generarReporte();
}

/**
 * Genera el reporte completo en la hoja "Reporte".
 * Lee datos directamente de las hojas de clasificación y Graduados.
 * Se puede ejecutar desde el menú o se llama automáticamente.
 */
function generarReporte() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('Reporte');
  if (!hoja) {
    hoja = ss.getSheetByName('Estado actual del participante');
    if (!hoja) {
      hoja = ss.insertSheet('Reporte');
    } else {
      hoja.setName('Reporte');
    }
  }

  // Limpiar toda la hoja
  hoja.clear();
  hoja.setTabColor('#e91e63');

  // -- Colores ------------------------------------------------------------
  const C_TITULO     = '#1a237e';
  const C_TITULO_FG  = '#ffffff';
  const C_HEADER     = '#3949ab';
  const C_HEADER_FG  = '#ffffff';
  const C_PAR        = '#e8eaf6';
  const C_IMPAR      = '#ffffff';
  const C_TOTAL      = '#c5cae9';
  const C_SECCION    = '#283593';
  const C_ACCENT     = '#0d47a1';
  const C_CONEXIONES = '#e65100';
  const MESES_NOMBRE = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const BORDER_STYLE = SpreadsheetApp.BorderStyle.SOLID;

  // -- Leer conteos de cada hoja de clasificacion -------------------------
  // Contamos filas que tengan Nombre completo (no filas vacias con checkboxes)
  // Hojas clasificacion: col 3 (indice 2) = Nombre completo
  // Conexiones Laborales: col 2 (indice 1) = Nombre completo
  const conteos = {};
  let totalParticipantes = 0;
  ETAPAS_FLUJO.forEach(nombreHoja => {
    const h = ss.getSheetByName(nombreHoja);
    if (!h || h.getLastRow() <= 1) { conteos[nombreHoja] = 0; return; }
    const datos = h.getDataRange().getValues();
    var cant = 0;
    var colNombre = (nombreHoja === 'Conexiones Laborales') ? 1 : 2; // indice de Nombre completo
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][colNombre] && datos[i][colNombre].toString().trim() !== '') {
        cant++;
      }
    }
    conteos[nombreHoja] = cant;
    totalParticipantes += cant;
  });

  // Conexiones Laborales (ya incluida arriba si esta en ETAPAS_FLUJO)
  var totalConexiones = conteos['Conexiones Laborales'] || 0;

  // Total graduados (col 4, indice 3 = Nombre completo)
  const hojaGrad = ss.getSheetByName('Graduados');
  var totalGraduados = 0;
  if (hojaGrad && hojaGrad.getLastRow() > 1) {
    var datosGr = hojaGrad.getDataRange().getValues();
    for (var i = 1; i < datosGr.length; i++) {
      if (datosGr[i][3] && datosGr[i][3].toString().trim() !== '') totalGraduados++;
    }
  }

  // Seguimientos pendientes
  const totalSeguimientos = obtenerSeguimientosPendientes().length;

  // -- Leer fechas de ingreso por mes desde cada hoja clasificación -------
  const ingresosPorMes = {}; // { 'Ene 2026': { 'Aliados': 2, ... } }
  ETAPAS_FLUJO.forEach(nombreHoja => {
    const h = ss.getSheetByName(nombreHoja);
    if (!h || h.getLastRow() <= 1) return;
    const datos = h.getDataRange().getValues();
    // Col 1 (indice 0) = Fecha de ingreso en todas las hojas de clasificacion
    var colNom = (nombreHoja === 'Conexiones Laborales') ? 1 : 2;
    for (let i = 1; i < datos.length; i++) {
      // Ignorar filas sin nombre (vacias por checkboxes/validaciones)
      if (!datos[i][colNom] || datos[i][colNom].toString().trim() === '') continue;
      const fechaRaw = datos[i][0];
      var fecha;
      if (fechaRaw instanceof Date) {
        fecha = fechaRaw;
      } else if (typeof fechaRaw === 'string' && fechaRaw.indexOf('/') !== -1) {
        const p = fechaRaw.split('/');
        fecha = new Date(p[2], parseInt(p[1]) - 1, p[0]);
      } else {
        continue;
      }
      const claveMes = MESES_NOMBRE[fecha.getMonth()] + ' ' + fecha.getFullYear();
      if (!ingresosPorMes[claveMes]) {
        ingresosPorMes[claveMes] = { _orden: fecha.getTime() };
        ETAPAS_FLUJO.forEach(e => { ingresosPorMes[claveMes][e] = 0; });
      }
      ingresosPorMes[claveMes][nombreHoja]++;
    }
  });
  const mesesOrdenados = Object.keys(ingresosPorMes).sort(
    (a, b) => ingresosPorMes[a]._orden - ingresosPorMes[b]._orden
  );

  // -- Ajustar anchos de columna ------------------------------------------
  hoja.setColumnWidth(1, 200);
  hoja.setColumnWidth(2, 120);
  hoja.setColumnWidth(3, 80);
  for (let c = 4; c <= 10; c++) hoja.setColumnWidth(c, 110);

  let fila = 1;

  // ======================================================================
  // BLOQUE 1: RESUMEN GENERAL
  // ======================================================================
  const anchoBloq1 = 3;
  hoja.getRange(fila, 1, 1, anchoBloq1).merge()
      .setValue('REPORTE DE PARTICIPANTES')
      .setBackground(C_TITULO).setFontColor(C_TITULO_FG)
      .setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center');
  fila++;

  hoja.getRange(fila, 1, 1, anchoBloq1).merge()
      .setValue('Actualizado: ' + new Date().toLocaleDateString('es-ES') + ' ' +
                new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' }))
      .setFontColor('#666666').setFontStyle('italic')
      .setHorizontalAlignment('center').setFontSize(10);
  fila += 2;

  // Indicadores clave
  const indicadores = [
    ['Total graduados registrados', totalGraduados],
    ['Participantes clasificados', totalParticipantes],
    ['Conexiones laborales', totalConexiones],
    ['Seguimientos pendientes', totalSeguimientos]
  ];
  hoja.getRange(fila, 1, 1, 2)
      .setValues([['Indicador', 'Cantidad']])
      .setBackground(C_ACCENT).setFontColor(C_TITULO_FG)
      .setFontWeight('bold').setHorizontalAlignment('center');
  fila++;
  indicadores.forEach((ind, idx) => {
    const bg = idx % 2 === 0 ? C_PAR : C_IMPAR;
    hoja.getRange(fila, 1).setValue(ind[0]).setBackground(bg).setFontSize(11);
    hoja.getRange(fila, 2).setValue(ind[1]).setBackground(bg)
        .setHorizontalAlignment('center').setFontWeight('bold').setFontSize(13);
    fila++;
  });
  hoja.getRange(fila - indicadores.length - 1, 1, indicadores.length + 1, 2)
      .setBorder(true, true, true, true, true, true, '#cccccc', BORDER_STYLE);
  fila += 2;

  // ======================================================================
  // BLOQUE 2: DISTRIBUCIÓN POR ETAPA
  // ======================================================================
  hoja.getRange(fila, 1, 1, anchoBloq1).merge()
      .setValue('DISTRIBUCIÓN POR ETAPA')
      .setBackground(C_SECCION).setFontColor(C_TITULO_FG)
      .setFontWeight('bold').setFontSize(12)
      .setHorizontalAlignment('center');
  fila++;

  hoja.getRange(fila, 1, 1, 3)
      .setValues([['Etapa', 'Participantes', '%']])
      .setBackground(C_HEADER).setFontColor(C_HEADER_FG)
      .setFontWeight('bold').setHorizontalAlignment('center');
  fila++;

  ETAPAS_FLUJO.forEach((etapa, idx) => {
    const cant = conteos[etapa];
    const pct  = totalParticipantes > 0 ? Math.round((cant / totalParticipantes) * 100) : 0;
    const bg   = idx % 2 === 0 ? C_PAR : C_IMPAR;
    hoja.getRange(fila, 1).setValue(etapa).setBackground(bg).setFontSize(11);
    hoja.getRange(fila, 2).setValue(cant).setBackground(bg)
        .setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12);
    hoja.getRange(fila, 3).setValue(pct + '%').setBackground(bg)
        .setHorizontalAlignment('center').setFontSize(11);
    fila++;
  });

  hoja.getRange(fila, 1, 1, 3)
      .setValues([['TOTAL', totalParticipantes, totalParticipantes > 0 ? '100%' : '0%']])
      .setBackground(C_TOTAL).setFontWeight('bold')
      .setHorizontalAlignment('center').setFontSize(12);
  hoja.getRange(fila, 1).setHorizontalAlignment('left');
  fila++;

  hoja.getRange(fila - ETAPAS_FLUJO.length - 2, 1, ETAPAS_FLUJO.length + 2, 3)
      .setBorder(true, true, true, true, true, true, '#cccccc', BORDER_STYLE);
  fila += 2;

  // ======================================================================
  // BLOQUE 3: INGRESOS POR MES
  // ======================================================================
  if (mesesOrdenados.length > 0) {
    const headerMes = ['Mes'];
    ETAPAS_FLUJO.forEach(e => {
      const corto = e === 'Activamente busca trabajo' ? 'Act. busca' :
                    e === 'Por su cuenta' ? 'Por su cta.' :
                    e === 'Derivaciones' ? 'Deriv.' : e;
      headerMes.push(corto);
    });
    headerMes.push('Total');
    const numColsMes = headerMes.length;

    hoja.getRange(fila, 1, 1, numColsMes).merge()
        .setValue('INGRESOS POR MES')
        .setBackground(C_SECCION).setFontColor(C_TITULO_FG)
        .setFontWeight('bold').setFontSize(12)
        .setHorizontalAlignment('center');
    fila++;

    hoja.getRange(fila, 1, 1, numColsMes)
        .setValues([headerMes])
        .setBackground(C_HEADER).setFontColor(C_HEADER_FG)
        .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
    fila++;

    const totalesPorEtapa = {};
    ETAPAS_FLUJO.forEach(e => { totalesPorEtapa[e] = 0; });
    let granTotal = 0;

    mesesOrdenados.forEach((mes, idx) => {
      const bg = idx % 2 === 0 ? C_PAR : C_IMPAR;
      const filaMes = [mes];
      let totalMes = 0;
      ETAPAS_FLUJO.forEach(e => {
        const cant = ingresosPorMes[mes][e];
        filaMes.push(cant);
        totalesPorEtapa[e] += cant;
        totalMes += cant;
      });
      filaMes.push(totalMes);
      granTotal += totalMes;

      hoja.getRange(fila, 1, 1, numColsMes)
          .setValues([filaMes]).setBackground(bg)
          .setHorizontalAlignment('center').setFontSize(10);
      hoja.getRange(fila, 1).setHorizontalAlignment('left');
      fila++;
    });

    const filaTotales = ['TOTAL'];
    ETAPAS_FLUJO.forEach(e => { filaTotales.push(totalesPorEtapa[e]); });
    filaTotales.push(granTotal);
    hoja.getRange(fila, 1, 1, numColsMes)
        .setValues([filaTotales])
        .setBackground(C_TOTAL).setFontWeight('bold')
        .setHorizontalAlignment('center').setFontSize(11);
    hoja.getRange(fila, 1).setHorizontalAlignment('left');

    hoja.getRange(fila - mesesOrdenados.length - 1, 1, mesesOrdenados.length + 2, numColsMes)
        .setBorder(true, true, true, true, true, true, '#cccccc', BORDER_STYLE);
    fila += 2;
  }

  // ======================================================================
  // BLOQUE 4: CONEXIONES LABORALES
  // ======================================================================
  hoja.getRange(fila, 1, 1, 2).merge()
      .setValue('CONEXIONES LABORALES')
      .setBackground(C_CONEXIONES).setFontColor(C_TITULO_FG)
      .setFontWeight('bold').setFontSize(12)
      .setHorizontalAlignment('center');
  fila++;

  hoja.getRange(fila, 1).setValue('Total conexiones registradas').setFontSize(11);
  hoja.getRange(fila, 2).setValue(totalConexiones)
      .setHorizontalAlignment('center').setFontWeight('bold').setFontSize(13);
  hoja.getRange(fila, 1, 1, 2)
      .setBorder(true, true, true, true, true, true, '#cccccc', BORDER_STYLE);

  Logger.log('Reporte generado exitosamente');
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
      const c = datos[i][14]; // col 15 = Etapa
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

// ===========================================================================
// SECCIÓN 7: FUNCIONES DE PRUEBA Y DEPURACIÓN
// ===========================================================================

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
