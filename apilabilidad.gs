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

    // ── Submenú: Importar datos ────────────────────────────────────────────
    const submenuImport = ui.createMenu('📂 Importar datos')
      .addItem('📥 Importar todos (Kobo + externo)',         'importarTodosLosDatos')
      .addItem('📥 Graduados (externo)',                     'importarGraduadosDesdeExterno')
      .addSeparator()
      .addItem('🔄 Solo Graduados (Kobo)',                   'importarDatosKobo')
      .addItem('📋 Clasificación de Perfiles',               'importarClasificacionPerfiles')
      .addItem('😊 Satisfacción Empleo (IL-06)',             'importarSatisfaccionEmpleo')
      .addItem('🤝 Sesiones Acompañamiento (IL-08)',         'importarSesionesAcompanamiento')
      .addSeparator()
      .addItem('♻️ Reimportar Sesiones desde cero',          'reimportarSesionesDesdeCero')
      .addItem('♻️ Reimportar Satisfacción desde cero',      'reimportarSatisfaccionDesdeCero')
      .addItem('🧹 Limpiar duplicados (Clasificación)',      'limpiarDuplicadosClasificacion');

    // ── Submenú: Herramientas avanzadas ───────────────────────────────────
    const submenuAvanzado = ui.createMenu('⚙️ Herramientas avanzadas')
      .addItem('🔍 Diagnosticar Clasificación de Perfiles',  'diagnosticarClasificacionPerfiles')
      .addItem('🔧 Limpiar triggers duplicados',             'limpiarTriggersDuplicados')
      .addItem('⏱ Activar auto-import (cada hora)',          'activarAutoImport')
      .addItem('⏹ Desactivar auto-import',                   'desactivarAutoImport')
      .addItem('⏹ Desactivar sincronización horaria',        'desactivarSincronizacionAutomatica')
      .addSeparator()
      .addItem('⏰ Instalar trigger Graduados',               'instalarTriggerGraduados')
      .addItem('⏰ Activar verificación Creamos ID',          'instalarTriggerVerificacionCreamos')
      .addSeparator()
      .addItem('🚀 Instalar Sistema (primera vez)',           'instalarSistema')
      .addItem('🔁 Reinstalar Sistema (borra todo)',          'reinstalarSistema');

    ui.createMenu('📊 Equipito Empleabilidad')
      // ── 1. Importar ──────────────────────────────────
      .addSubMenu(submenuImport)
      .addSeparator()
      // ── 2. Ver y clasificar ──────────────────────────
      .addItem('📊 Generar Reporte',                         'generarReporte')
      .addItem('📝 Clasificar Graduados',                    'mostrarFormularioClasificacion')
      .addItem('🎨 Aplicar colores y desplegables (Graduados)','aplicarColoresGraduados')
      .addSeparator()
      // ── 3. Creamos ID / Salesforce ───────────────────
      .addItem('🔄 Autocompletar con Creamos ID',            'autocompletarConCreamos')
      .addItem('🔍 Verificar Creamos ID ahora',              'verificarYCompletarCreamos')
      .addItem('🩺 Diagnosticar errores Creamos ID',         'diagnosticarErroresCreamos')
      .addItem('🔒 Proteger base datos Salesforce',          'protegerBaseDatosSalesforce')
      .addSeparator()
      // ── 4. Avanzado ───────────────────────────────────
      .addSubMenu(submenuAvanzado)
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
    const hoja       = e.range.getSheet();
    const nombreHoja = hoja.getName();
    const col        = e.range.getColumn();
    const fila       = e.range.getRow();
    if (fila <= 1) return;

    // ── Hoja Graduados: columna Etapa (col 15) ─────────────────────────────
    if (nombreHoja === 'Graduados' && col === 15) {
      const nuevaEtapa = e.value;
      if (!nuevaEtapa || nuevaEtapa === 'Conexiones Laborales') return;

      const datosGrad = hoja.getRange(fila, 1, 1, 15).getValues()[0];
      const nombre    = datosGrad[3];

      copiarAHojaClasificacion(datosGrad, nuevaEtapa, {});
      generarReporte();

      if (nuevaEtapa === 'Activamente busca trabajo') {
        hoja.getRange(fila, 12).setValue('Si');
        crearSeguimientoBusquedaActiva(
          datosGrad[2],  // Creamos ID
          datosGrad[3],  // Nombre completo
          datosGrad[7]   // Teléfono
        );
      }

      SpreadsheetApp.getActiveSpreadsheet().toast(
        nombre + ' enviado a: ' + nuevaEtapa, '✅ Clasificado', 3
      );
      return;
    }

    // ── Hoja Sesiones Acompañamiento: columna Acción (col 16) ─────────────
    // Solo maneja "→ Seguimiento Bot" porque no abre diálogos.
    // "→ Conexiones Laborales" lo maneja onEditInstalable (requiere diálogo).
    if (nombreHoja === 'Sesiones Acompañamiento' && col === 16) {
      const accion = (e.value || '').toString().trim();
      if (!accion) return;

      // Siempre limpiar la celda de acción
      e.range.setValue('');

      if (accion === '→ Seguimiento Bot') {
        const datos        = hoja.getRange(fila, 1, 1, 15).getValues()[0];
        const creamosId    = (datos[0] || '').toString().trim();
        const nombre       = (datos[4] || '').toString().trim();
        const apellidos    = (datos[5] || '').toString().trim();
        const telefono     = (datos[6] || '').toString().trim();
        const nombreComp   = (nombre + ' ' + apellidos).trim();

        if (!nombreComp) {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            'La fila no tiene nombre.', '⚠️ Sin datos', 3
          );
          return;
        }

        crearSeguimientoBusquedaActiva(creamosId, nombreComp, telefono);
        SpreadsheetApp.getActiveSpreadsheet().toast(
          nombreComp + ' agregado a Seguimiento Bot', '✅ Seguimiento creado', 4
        );
      }
      // "→ Conexiones Laborales" se maneja en onEditInstalable
    }

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
    const hoja       = e.range.getSheet();
    const nombreHoja = hoja.getName();
    const col        = e.range.getColumn();
    const fila       = e.range.getRow();
    if (fila <= 1) return;

    // ── Graduados: Etapa = "Conexiones Laborales" ─────────────────────────
    if (nombreHoja === 'Graduados' && col === 15) {
      if (e.value !== 'Conexiones Laborales') return;

      const datosGrad = hoja.getRange(fila, 1, 1, 15).getValues()[0];
      const creamosId = datosGrad[2] || '';
      const nombre    = datosGrad[3] || '';

      e.range.setValue('');

      if (!nombre) {
        SpreadsheetApp.getActiveSpreadsheet().toast('La fila no tiene nombre.', '⚠️ Sin datos', 3);
        return;
      }

      const datosExtra = {
        genero:         datosGrad[4] || '',
        edad:           datosGrad[5] || '',
        nivelEducativo: datosGrad[6] || '',
        telefono:       datosGrad[7] || ''
      };

      const html = HtmlService.createHtmlOutput(
        _generarHTMLFormConexionLaboral(fila, creamosId, nombre, 'Graduados', datosExtra)
      ).setWidth(560).setHeight(720).setTitle('Conexión Laboral');
      SpreadsheetApp.getUi().showModalDialog(html, '💼 Conexión Laboral — ' + nombre);
      return;
    }

    // ── Sesiones Acompañamiento: columna Acción (col 16) = "→ Conexiones Laborales" ──
    if (nombreHoja === 'Sesiones Acompañamiento' && col === 16) {
      if ((e.value || '').toString().trim() !== '→ Conexiones Laborales') return;

      e.range.setValue('');

      // Columnas Sesiones: 1=Creamos ID, 5=Nombre, 6=Apellidos, 7=Teléfono, 9=Edad, 10=Género
      const datos      = hoja.getRange(fila, 1, 1, 15).getValues()[0];
      const creamosId  = (datos[0] || '').toString().trim();
      const nombre     = (datos[4] || '').toString().trim();
      const apellidos  = (datos[5] || '').toString().trim();
      const nombreComp = (nombre + ' ' + apellidos).trim();

      if (!nombreComp) {
        SpreadsheetApp.getActiveSpreadsheet().toast('La fila no tiene nombre.', '⚠️ Sin datos', 3);
        return;
      }

      const datosExtra = {
        telefono:       (datos[6] || '').toString().trim(),
        edad:           (datos[8] || '').toString().trim(),
        genero:         (datos[9] || '').toString().trim(),
        nivelEducativo: ''
      };

      const html = HtmlService.createHtmlOutput(
        _generarHTMLFormConexionLaboral(fila, creamosId, nombreComp, 'Sesiones Acompañamiento', datosExtra)
      ).setWidth(560).setHeight(720).setTitle('Conexión Laboral');
      SpreadsheetApp.getUi().showModalDialog(html, '💼 Conexión Laboral — ' + nombreComp);
    }

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
    SpreadsheetApp.getActiveSpreadsheet().toast('Descargando datos de Graduados...', '🔄 Importando', -1);
    const datos = obtenerDatosKoboToolbox();
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    if (!datos || datos.length === 0) {
      ui.alert('⚠️ Sin Datos', 'No se encontraron datos para importar.', ui.ButtonSet.OK);
      return;
    }
    const resultado = procesarDatosGraduados(datos);
    ui.alert('✅ Importación Completada',
             'Graduados nuevos: ' + resultado.nuevos + '\n' +
             'Actualizados: ' + resultado.actualizados + '\n' +
             'Total: ' + resultado.total,
             ui.ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    SpreadsheetApp.getUi().alert('❌ Error', 'Error al importar Graduados: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error en importarDatosKobo: ' + error);
  }
}

/**
 * Importa todos los formularios de KoboToolbox en una sola acción.
 * Muestra un resumen al final con los nuevos registros de cada fuente.
 */
function importarTodosLosDatos() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var res = [];

  ss.toast('Importando todos los datos...', '📥 Importando', -1);

  // 1. Graduados
  try {
    var d = obtenerDatosKoboToolbox();
    var g = procesarDatosGraduados(d);
    res.push('✅ Graduados: ' + g.nuevos + ' nuevos (total ' + g.total + ')');
  } catch (e) { res.push('⚠️ Graduados: ' + e.message); }

  // 2. Clasificación de Perfiles
  try {
    var c = _syncClasificacionSoloNuevos();
    res.push('✅ Clasificación de Perfiles: ' + c.nuevos + ' nuevos (total ' + c.total + ')');
  } catch (e) { res.push('⚠️ Clasificación: ' + e.message); }

  // 3. Satisfacción Empleo (IL-06)
  try {
    var s = _syncSatisfaccionSoloNuevos();
    res.push('✅ Satisfacción Empleo: ' + s.nuevos + ' nuevos (total ' + s.total + ')');
  } catch (e) { res.push('⚠️ Satisfacción: ' + e.message); }

  // 4. Sesiones Acompañamiento (IL-08)
  try {
    var se = _syncSesionesSoloNuevos();
    res.push('✅ Sesiones Acompañamiento: ' + se.nuevos + ' nuevos (total ' + se.total + ')');
  } catch (e) { res.push('⚠️ Sesiones: ' + e.message); }

  ss.toast('', '', 1);
  ui.alert('📥 Importación completada', res.join('\n'), ui.ButtonSet.OK);
}

/**
 * Muestra instrucciones para clasificar graduados.
 * La clasificación se hace directamente desde la hoja Graduados:
 * cambia la columna "Etapa" y el sistema copia automáticamente.
 */
function mostrarFormularioClasificacion() {
  SpreadsheetApp.getUi().alert(
    '📝 Cómo Clasificar Graduados',
    'Para clasificar un graduado:\n\n' +
    '1. Abre la hoja "Graduados"\n' +
    '2. Busca al graduado que quieres clasificar\n' +
    '3. En la columna "Etapa de flujo" selecciona la etapa del desplegable\n' +
    '4. El sistema lo copiará automáticamente a la hoja correspondiente\n\n' +
    'Etapas disponibles:\n' +
    '  • Aliados\n' +
    '  • Plataforma\n' +
    '  • Derivaciones\n' +
    '  • Activamente busca trabajo (incluye "Por su cuenta")\n' +
    '  • Paso a paso\n' +
    '  • Conexiones Laborales (abre formulario de conexión)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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
 * Muestra la configuración del proyecto (ya no se necesita, todo está en el código)
 */
function mostrarConfiguracion() {
  SpreadsheetApp.getUi().alert('⚙️ Configuración',
    'Las credenciales ya están configuradas en el código.\n\n' +
    'Token: ****' + KOBO_TOKEN.slice(-6) + '\n' +
    'URL Clasificación: configurada',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// -----------------------------------------------------------------------------
// CONFIGURACIÓN DE KOBOTOOLBOX
// -----------------------------------------------------------------------------

const KOBO_TOKEN = '64cc018b88067397addd36b09288be8b6539cf39';

// URLs de exportación de KoboToolbox
// TODO: Agregar URL de Graduados cuando esté disponible
const URL_GRADUADOS              = ''; // pendiente — pegar URL cuando la tengas
const URL_GRADUADOS_FALLBACK     = '';
const URL_CLASIFICACION_PERFILES  = 'https://kf.kobotoolbox.org/api/v2/assets/aSH2JhYXLqn4o66z8L3RmK/export-settings/esQRaR2qPsjyboQtpNEiFy3/data.csv';
const URL_CLASIFICACION_FALLBACK  = 'https://kf.kobotoolbox.org/api/v2/assets/aSH2JhYXLqn4o66z8L3RmK/data/?format=csv';
// IL_06 — Cuestionario de satisfacción-Empleo
const URL_SATISFACCION_EMPLEO     = 'https://kf.kobotoolbox.org/api/v2/assets/aKPtMoCx6tFzagpLoWysuq/export-settings/esoDGnReKRANZHjqdVVZZm8/data.csv';
// IL_08 — Sesiones Acompañamiento profesional
const URL_SESIONES_ACOMPANAMIENTO = 'https://kf.kobotoolbox.org/api/v2/assets/azWEsFBnHSUvfVugsTcgTD/export-settings/esr7u8sHX95HjFXmUpjunam/data.csv';

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

// Perfiles de clasificación según puntaje
const PERFILES_CLASIFICACION = [
  'Perfil 1: Completamente independiente',
  'Perfil 2: Quiere, puede, necesita apoyo',
  'Perfil 2 Bajo: Necesita acompañamiento intensivo',
  'Perfil 3: Quiere pero no puede (Barreras activas)',
  'Perfil 4: No quiere / Otras prioridades'
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
  'Activamente busca trabajo',
  'Paso a paso',
  'Conexiones Laborales'
];

// Niveles educativos — opciones del dropdown "Nivel educativo" en Graduados
const NIVELES_EDUCATIVOS = [
  'Sin escolaridad',
  'Primero primaria',
  'Segundo primaria',
  'Tercero primaria',
  'Cuarto primaria',
  'Quinto primaria',
  'Sexto primaria',
  'Primero básico',
  'Segundo básico',
  'Tercero básico',
  'Cuarto bachillerato',
  'Quinto bachillerato',
  'Diversificado',
  'Universidad'
];

// Colores de fondo por opción de dropdown (formato condicional).
// clave = texto exacto de la opción, valor = { bg, fg }
const COLORES_DROPDOWN = {
  // Etapa (ETAPAS_FLUJO)
  'Aliados':                     { bg: '#4285f4', fg: '#ffffff' },
  'Plataforma':                  { bg: '#00bcd4', fg: '#003c43' },
  'Derivaciones':                { bg: '#ab47bc', fg: '#ffffff' },
  'Activamente busca trabajo':   { bg: '#ff7043', fg: '#ffffff' },
  'Paso a paso':                 { bg: '#ffca28', fg: '#5f4300' },
  'Conexiones Laborales':        { bg: '#66bb6a', fg: '#ffffff' },
  // Empleado (siNo)
  'Si':                          { bg: '#66bb6a', fg: '#ffffff' },
  'No':                          { bg: '#ef5350', fg: '#ffffff' },
  // Género
  'Hombre':                      { bg: '#bbdefb', fg: '#0d47a1' },
  'Mujer':                       { bg: '#f8bbd0', fg: '#880e4f' },
  'Trans hombre':                { bg: '#c5cae9', fg: '#1a237e' },
  'No binario':                  { bg: '#d1c4e9', fg: '#311b92' },
  'Otro':                        { bg: '#e0e0e0', fg: '#424242' },
  // Nivel educativo (gradiente claro de menor a mayor)
  'Sin escolaridad':             { bg: '#eeeeee', fg: '#424242' },
  'Primero primaria':            { bg: '#ffebee', fg: '#b71c1c' },
  'Segundo primaria':            { bg: '#fce4ec', fg: '#880e4f' },
  'Tercero primaria':            { bg: '#f3e5f5', fg: '#4a148c' },
  'Cuarto primaria':             { bg: '#ede7f6', fg: '#311b92' },
  'Quinto primaria':             { bg: '#e8eaf6', fg: '#1a237e' },
  'Sexto primaria':              { bg: '#e3f2fd', fg: '#0d47a1' },
  'Primero básico':              { bg: '#e1f5fe', fg: '#01579b' },
  'Segundo básico':              { bg: '#e0f7fa', fg: '#006064' },
  'Tercero básico':              { bg: '#e0f2f1', fg: '#004d40' },
  'Cuarto bachillerato':         { bg: '#e8f5e9', fg: '#1b5e20' },
  'Quinto bachillerato':         { bg: '#f1f8e9', fg: '#33691e' },
  'Diversificado':               { bg: '#fff8e1', fg: '#e65100' },
  'Universidad':                 { bg: '#fff3e0', fg: '#bf360c' }
};

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
      { nombre: 'Nivel educativo',      ancho: 160, tipo: 'dropdown', opciones: NIVELES_EDUCATIVOS },
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

  // -- 5. PASO A PASO --------------------------------------------------------
  'Paso a paso': {
    color: '#9e9e9e',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'DPI',                ancho: 140, tipo: 'texto' },
      { nombre: 'Formación',          ancho: 180, tipo: 'texto' },
      { nombre: 'Cohorte',            ancho: 110, tipo: 'texto' },
      { nombre: 'Nota',               ancho: 300, tipo: 'texto' },
      { nombre: 'Activo',             ancho: 90,  tipo: 'siNo'  }
    ]
  },

  // -- 6. ACTIVAMENTE BUSCA TRABAJO (fusionado con "Por su cuenta") ----------
  'Activamente busca trabajo': {
    color: '#0f9d58',
    columnas: [
      ...COLUMNAS_COMUNES,
      { nombre: 'Tipo de búsqueda', ancho: 180, tipo: 'dropdown', opciones: ['Activamente busca trabajo', 'Por su cuenta'] },
      { nombre: 'Mensaje',          ancho: 120, tipo: 'texto' },
      { nombre: 'Llamada',          ancho: 120, tipo: 'texto' },
      { nombre: 'Nota',             ancho: 300, tipo: 'texto' },
      { nombre: 'Entrevista',       ancho: 130, tipo: 'texto' },
      { nombre: 'Trámites',         ancho: 200, tipo: 'texto' },
      { nombre: 'Activo',           ancho: 90,  tipo: 'siNo'  }
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

  // -- IL_06: SATISFACCIÓN EMPLEO --------------------------------------------
  // Nueva estructura del formulario (5 preguntas escala 1-5 + texto):
  //   satisfaccion_empleo, cumple_expectativas, ambiente_laboral,
  //   salario_beneficios, permanencia (1-5)
  //   aspecto_mejorar, otro_aspecto, Nota (texto libre)
  'Satisfacción Empleo': {
    color: '#00897b',
    columnas: [
      { nombre: 'Creamos ID',           ancho: 130, tipo: 'texto' },
      { nombre: 'Fecha envío',          ancho: 160, tipo: 'texto' },
      { nombre: 'Satisfacción empleo',  ancho: 110, tipo: 'texto' },
      { nombre: 'Cumple expectativas',  ancho: 110, tipo: 'texto' },
      { nombre: 'Ambiente laboral',     ancho: 110, tipo: 'texto' },
      { nombre: 'Salario y beneficios', ancho: 130, tipo: 'texto' },
      { nombre: 'Permanencia',          ancho: 100, tipo: 'texto' },
      { nombre: 'Aspecto a mejorar',    ancho: 220, tipo: 'texto' },
      { nombre: 'Otro aspecto',         ancho: 200, tipo: 'texto' },
      { nombre: 'Nota',                 ancho: 280, tipo: 'texto' },
      { nombre: 'Promedio (1-5)',       ancho: 110, tipo: 'texto' }
    ]
  },

  // -- IL_08: SESIONES ACOMPAÑAMIENTO PROFESIONAL ----------------------------
  // Columnas exactas del NOMBRES whitelist (excluye columnas redundantes de KoboToolbox)
  // Col 16 = "Acción": dropdown para enviar directamente desde la hoja
  'Sesiones Acompañamiento': {
    color: '#5e35b1',
    columnas: [
      { nombre: 'Creamos ID',          ancho: 130, tipo: 'texto' },
      { nombre: 'Fecha envío',         ancho: 170, tipo: 'texto' },
      { nombre: 'Inicio sesión',       ancho: 160, tipo: 'texto' },
      { nombre: 'Proyecto',            ancho: 150, tipo: 'texto' },
      { nombre: 'Nombre',              ancho: 130, tipo: 'texto' },
      { nombre: 'Apellidos',           ancho: 130, tipo: 'texto' },
      { nombre: 'Teléfono',            ancho: 130, tipo: 'texto' },
      { nombre: 'Fecha nacimiento',    ancho: 130, tipo: 'texto' },
      { nombre: 'Edad',                ancho: 65,  tipo: 'texto' },
      { nombre: 'Género',              ancho: 100, tipo: 'texto' },
      { nombre: 'Año ingreso Creamos', ancho: 150, tipo: 'texto' },
      { nombre: 'En qué año',          ancho: 100, tipo: 'texto' },
      { nombre: 'Grado académico',     ancho: 150, tipo: 'texto' },
      { nombre: 'Tipo servicio',       ancho: 200, tipo: 'texto' },
      { nombre: 'Comentario',          ancho: 280, tipo: 'texto' },
      { nombre: 'Acción',              ancho: 220, tipo: 'dropdown',
        opciones: ['→ Conexiones Laborales', '→ Seguimiento Bot'] }
    ]
  },

  // Hoja "Seguimientos" eliminada — reemplazada por "Seguimiento Bot"

  // -- SEGUIMIENTO BOT (para n8n + WhatsApp) --------------------------------
  // Hoja unificada para dos flujos:
  //   "Post-empleo"     → viene de Conexiones Laborales
  //   "Búsqueda activa" → viene de Activamente busca trabajo
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
      { nombre: 'Estado',             ancho: 180, tipo: 'dropdown', opciones: ['Pendiente', 'Mensaje Enviado', 'Respuesta Recibida', 'Llamada Programada', 'Completado'] },
      { nombre: 'Etapa Actual',       ancho: 130, tipo: 'texto' },
      { nombre: 'Resp S1 P1',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S1 P2',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S1 P3',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S2 P1',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S2 P2',         ancho: 300, tipo: 'texto' },
      { nombre: 'Resp S2 P3',         ancho: 300, tipo: 'texto' },
      { nombre: 'Email Enviado',      ancho: 130, tipo: 'siNo'  },
      { nombre: 'Tipo Seguimiento',   ancho: 180, tipo: 'dropdown', opciones: ['Post-empleo', 'Búsqueda activa'] },
      { nombre: 'Origen',             ancho: 220, tipo: 'texto' }
    ]
  },

  // -- CLASIFICACIÓN DE PERFILES -----------------------------------------------
  // Importada desde KoboToolbox: IL_09_Módulo de Clasificación de Perfiles
  // Estructura dinámica — se reconstruye en procesarClasificacionPerfiles()
  'Clasificación de Perfiles': {
    color: '#5c6bc0',
    columnas: [
      { nombre: 'Fecha de ingreso',    ancho: 140, tipo: 'texto' },
      { nombre: 'Creamos ID',          ancho: 130, tipo: 'texto' },
      { nombre: 'Nombre completo',     ancho: 200, tipo: 'texto' },
      { nombre: 'D1: Cuidado',         ancho: 280, tipo: 'texto' },
      { nombre: 'D1 Comentario',       ancho: 300, tipo: 'texto' },
      { nombre: 'D2: Violencia',       ancho: 280, tipo: 'texto' },
      { nombre: 'D2 Comentario',       ancho: 300, tipo: 'texto' },
      { nombre: 'D3: Movilidad',       ancho: 280, tipo: 'texto' },
      { nombre: 'D3 Comentario',       ancho: 300, tipo: 'texto' },
      { nombre: 'D4: Legal/Salud',     ancho: 280, tipo: 'texto' },
      { nombre: 'D4 Comentario',       ancho: 300, tipo: 'texto' },
      { nombre: 'D5: Motivación',      ancho: 280, tipo: 'texto' },
      { nombre: 'D5 Comentario',       ancho: 300, tipo: 'texto' },
      { nombre: 'D6: Experiencia',     ancho: 280, tipo: 'texto' },
      { nombre: 'D6 Comentario',       ancho: 300, tipo: 'texto' },
      { nombre: 'D7: Autonomía',       ancho: 280, tipo: 'texto' },
      { nombre: 'D7 Comentario',       ancho: 300, tipo: 'texto' },
      { nombre: 'Puntaje Total',       ancho: 110, tipo: 'texto' },
      { nombre: 'Barreras activas',    ancho: 130, tipo: 'texto' },
      { nombre: 'Desmotivación',       ancho: 130, tipo: 'texto' },
      { nombre: 'Perfil Asignado',     ancho: 300, tipo: 'texto' },
      { nombre: 'Notas de observación', ancho: 380, tipo: 'texto' }
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
        'Seguimientos',        // legacy — se elimina si existe
        'Seguimiento Bot',
        'Clasificación de Perfiles',
        'Satisfacción Empleo',
        'Sesiones Acompañamiento',
        'Reporte',
        'Conexiones Laborales',
        'Configuración', // por si existe de versión anterior
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
      'Paso a paso',
      'Activamente busca trabajo',
      'Conexiones Laborales',
      'Seguimiento Bot',
      'Clasificación de Perfiles',
      'Satisfacción Empleo',
      'Sesiones Acompañamiento',
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

    // -- 3. Eliminar hoja temporal si existe ----------------------------------
    const temp = ss.getSheetByName('_temp_');
    if (temp) ss.deleteSheet(temp);

    // -- 4. Ordenar hojas -----------------------------------------------------
    _ordenarHojas(ss, ordenHojas);

    // -- 6. Instalar trigger de onEdit para Conexiones Laborales -------------
    configurarEditTrigger();

    // -- 7. Instalar sincronización automática cada hora ---------------------
    configurarTriggerSincronizacion();

    ss.toast('', '', 1);

    ui.alert(
      '✅ Sistema Instalado',
      `El sistema está listo.\n\n` +
      `Se crearon ${ordenHojas.length} hojas con sus columnas.\n` +
      'Token de KoboToolbox configurado en el código.\n' +
      'URL de Clasificación de Perfiles configurada.\n' +
      'Sincronización automática activada (cada hora).\n\n' +
      'Próximo paso:\n' +
      '1. Usa "Importar Clasificación de Perfiles" para cargar evaluaciones\n' +
      '2. Usa "Importar Graduados" si tienes URL de exportación de Graduados\n' +
      '3. Comienza a clasificar graduados en la hoja Graduados',
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

  // -- Colores por opción de dropdown (formato condicional) -----------------
  _aplicarColoresDropdowns(hoja, columnas);

  Logger.log('Hoja construida: ' + nombreHoja);
}

/**
 * Aplica formato condicional para que cada opción de un dropdown/siNo se
 * muestre con su color (definido en COLORES_DROPDOWN). No borra datos.
 * Conserva las reglas de formato condicional que NO apuntan a estas columnas.
 * @param {Sheet} hoja
 * @param {Array} columnas - definición de columnas (ESTRUCTURA_HOJAS[x].columnas)
 */
function _aplicarColoresDropdowns(hoja, columnas) {
  const maxFilas = Math.max(hoja.getMaxRows() - 1, 1);

  // Columnas objetivo (1-based) que vamos a (re)colorear
  const colsObjetivo = [];
  columnas.forEach((col, i) => {
    if (col.tipo === 'dropdown' || col.tipo === 'siNo') colsObjetivo.push(i + 1);
  });
  if (colsObjetivo.length === 0) return;

  // Conservar reglas existentes que no toquen nuestras columnas objetivo
  const reglasPrevias = hoja.getConditionalFormatRules();
  const reglasConservadas = reglasPrevias.filter(function (regla) {
    const rangos = regla.getRanges();
    for (let r = 0; r < rangos.length; r++) {
      const ini = rangos[r].getColumn();
      const fin = ini + rangos[r].getNumColumns() - 1;
      for (let c = 0; c < colsObjetivo.length; c++) {
        if (colsObjetivo[c] >= ini && colsObjetivo[c] <= fin) return false;
      }
    }
    return true;
  });

  const nuevasReglas = [];
  columnas.forEach((col, i) => {
    if (col.tipo !== 'dropdown' && col.tipo !== 'siNo') return;
    const colNum   = i + 1;
    const opciones = col.tipo === 'siNo' ? ['Si', 'No'] : (col.opciones || []);
    const rango    = hoja.getRange(2, colNum, maxFilas);

    opciones.forEach(function (opcion) {
      const c = COLORES_DROPDOWN[opcion];
      if (!c) return;
      const regla = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(opcion)
        .setBackground(c.bg)
        .setFontColor(c.fg)
        .setRanges([rango])
        .build();
      nuevasReglas.push(regla);
    });
  });

  hoja.setConditionalFormatRules(reglasConservadas.concat(nuevasReglas));
}

/**
 * Aplica el desplegable de "Nivel educativo" y los colores por opción a la
 * hoja "Graduados" YA EXISTENTE, sin borrar datos. Pensado para ejecutarse
 * desde el menú sobre una hoja que ya tiene registros.
 */
function aplicarColoresGraduados() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Graduados');
  if (!hoja) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja "Graduados".');
    return;
  }

  const columnas = ESTRUCTURA_HOJAS['Graduados'].columnas;
  const maxFilas = Math.max(hoja.getMaxRows() - 1, 1);

  // Reaplicar validaciones de dropdown/siNo (incluye el nuevo Nivel educativo)
  columnas.forEach((col, i) => {
    const rango = hoja.getRange(2, i + 1, maxFilas);
    if (col.tipo === 'siNo') {
      rango.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['Si', 'No'], true).setAllowInvalid(false).build()
      );
    } else if (col.tipo === 'dropdown' && col.opciones && col.opciones.length) {
      rango.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(col.opciones, true).setAllowInvalid(false).build()
      );
    }
  });

  _aplicarColoresDropdowns(hoja, columnas);

  SpreadsheetApp.getUi().alert(
    '✅ Listo',
    'Se aplicó el desplegable de "Nivel educativo" y los colores por opción ' +
    'en la hoja Graduados (sin borrar datos).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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
  if (!URL_GRADUADOS) {
    throw new Error('URL de Graduados aún no configurada. Pedila al administrador del formulario KoboToolbox.');
  }
  var opciones = {
    method: 'get',
    headers: { 'Authorization': 'Token ' + KOBO_TOKEN, 'Accept': 'text/csv' },
    muteHttpExceptions: true
  };
  var urls = [URL_GRADUADOS, URL_GRADUADOS_FALLBACK].filter(function(u) { return !!u; });
  var contenido;
  var respuesta;
  for (var u = 0; u < urls.length; u++) {
    respuesta = UrlFetchApp.fetch(urls[u], opciones);
    var codigo = respuesta.getResponseCode();
    if (codigo !== 200) { continue; }
    contenido = respuesta.getContentText();
    if (contenido.trim().charAt(0) === '{' || contenido.trim().charAt(0) === '[') {
      contenido = null;
      continue;
    }
    break;
  }
  if (!contenido) {
    throw new Error('No se pudo obtener CSV de Graduados. HTTP: ' + (respuesta ? respuesta.getResponseCode() : 0));
  }
  return parsearCSV(contenido);
}

/**
 * Parsea el contenido CSV y lo convierte en array de objetos
 * @param {string} csvContent
 * @return {Array}
 */
/**
 * Parsea CSV completo manejando:
 *  - Campos multilínea entre comillas (los textos de KoboToolbox tienen saltos de línea)
 *  - Auto-detección de separador , o ;
 *  - Comillas dobles escapadas ""
 */
function parsearCSV(csvContent) {
  try {
    // Auto-detectar separador mirando la primera línea NO quoted
    var sep = _detectarSeparadorCSV(csvContent);
    Logger.log('Separador CSV detectado: "' + sep + '"');

    // Parsear carácter a carácter para respetar campos multilínea
    var filas    = [];
    var fila     = [];
    var campo    = '';
    var enComilla = false;
    var n        = csvContent.length;

    for (var i = 0; i < n; i++) {
      var c = csvContent[i];

      if (c === '"') {
        // Comilla doble escapada ""
        if (enComilla && i + 1 < n && csvContent[i + 1] === '"') {
          campo += '"';
          i++;
        } else {
          enComilla = !enComilla;
        }
        continue;
      }

      if (!enComilla && c === sep) {
        fila.push(campo.trim());
        campo = '';
        continue;
      }

      if (!enComilla && (c === '\n' || c === '\r')) {
        // Saltar \r de \r\n
        if (c === '\r' && i + 1 < n && csvContent[i + 1] === '\n') i++;
        fila.push(campo.trim());
        campo = '';
        if (fila.some(function(f) { return f !== ''; })) filas.push(fila);
        fila = [];
        continue;
      }

      campo += c;
    }
    // Última celda / fila
    if (campo || fila.length > 0) {
      fila.push(campo.trim());
      if (fila.some(function(f) { return f !== ''; })) filas.push(fila);
    }

    if (filas.length < 2) return [];

    // Normalizar headers: eliminar saltos de línea dentro del nombre de columna
    var headers = filas[0].map(function(h) {
      return h.replace(/[\r\n]+/g, ' ').replace(/  +/g, ' ').trim();
    });
    var datos   = [];
    for (var r = 1; r < filas.length; r++) {
      var vals = filas[r];
      var obj  = {};
      headers.forEach(function(h, idx) { obj[h] = vals[idx] || ''; });
      datos.push(obj);
    }
    Logger.log('CSV parseado: ' + (filas.length - 1) + ' filas, ' + headers.length + ' columnas');
    return datos;
  } catch (error) {
    Logger.log('Error al parsear CSV: ' + error);
    throw error;
  }
}

function _detectarSeparadorCSV(contenido) {
  // Recorre hasta el primer \n fuera de comillas para contar , vs ;
  var enComilla = false;
  var comas = 0, puntoycoma = 0;
  for (var i = 0; i < contenido.length; i++) {
    var c = contenido[i];
    if (c === '"') { enComilla = !enComilla; continue; }
    if (enComilla) continue;
    if (c === '\n') break;
    if (c === ',') comas++;
    if (c === ';') puntoycoma++;
  }
  return puntoycoma > comas ? ';' : ',';
}

// parsearLineaCSV se conserva para compatibilidad con código existente
function parsearLineaCSV(linea, sep) {
  var separador = sep || ',';
  var valores = [];
  var val = '';
  var enComilla = false;
  for (var i = 0; i < linea.length; i++) {
    var c = linea[i];
    if (c === '"') { enComilla = !enComilla; }
    else if (c === separador && !enComilla) { valores.push(val.trim()); val = ''; }
    else { val += c; }
  }
  valores.push(val.trim());
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
/**
 * Sincronización manual desde el menú: importa Graduados + Clasificación de Perfiles
 */
function sincronizarTodoManual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Sincronizando Graduados y Clasificación de Perfiles...', '🔄 Sincronizando', -1);

  var resumen = [];

  try {
    var datos = obtenerDatosKoboToolbox();
    var res   = procesarDatosGraduados(datos);
    resumen.push('Graduados: ' + res.nuevos + ' nuevos, ' + res.total + ' total');
  } catch (e) {
    resumen.push('Graduados: ' + (e.message || 'sin URL configurada'));
  }

  try {
    var res2 = _syncClasificacionPerfiles();
    resumen.push('Clasificación: ' + res2.nuevos + ' nuevos, ' + res2.actualizados + ' actualizados');
  } catch (e) {
    resumen.push('Clasificación: ' + (e.message || 'error'));
  }

  ss.toast('', '', 1);
  ui.alert('✅ Sincronización Completada', resumen.join('\n'), ui.ButtonSet.OK);
}

function sincronizacionAutomatica() {
  try {
    Logger.log('Iniciando sincronización automática...');

    // 1. Sincronizar Graduados (si tiene URL configurada)
    try {
      const datos     = obtenerDatosKoboToolbox();
      const resultado = procesarDatosGraduados(datos);
      Logger.log('Graduados: ' + resultado.nuevos + ' nuevos, ' + resultado.total + ' total');
      if (resultado.nuevos > 0) enviarNotificacionNuevosGraduados(resultado.nuevos);
    } catch (e) {
      Logger.log('Sync graduados omitido: ' + e.message);
    }

    // 2. Sincronizar Clasificación de Perfiles (solo nuevos, sin borrar hoja)
    try {
      var resultClasif = _syncClasificacionSoloNuevos();
      Logger.log('Clasificación: ' + resultClasif.nuevos + ' nuevos, ' + resultClasif.total + ' total');
    } catch (e) {
      Logger.log('Sync clasificación omitido: ' + e.message);
    }

    // 3. Sincronizar Satisfacción Empleo (IL-06)
    try {
      var resultSatisf = _syncSatisfaccionSoloNuevos();
      Logger.log('Satisfacción: ' + resultSatisf.nuevos + ' nuevos, ' + resultSatisf.total + ' total');
    } catch (e) {
      Logger.log('Sync Satisfacción omitido: ' + e.message);
    }

    // 4. Sincronizar Sesiones Acompañamiento (IL-08)
    try {
      var resultSes = _syncSesionesSoloNuevos();
      Logger.log('Sesiones: ' + resultSes.nuevos + ' nuevos, ' + resultSes.total + ' total');
    } catch (e) {
      Logger.log('Sync Sesiones omitido: ' + e.message);
    }

    // 5. Importar Graduados desde archivo externo de Google Sheets
    try {
      importarGraduadosDesdeExterno();
      Logger.log('Import Graduados externo ejecutado');
    } catch (e) {
      Logger.log('Import Graduados externo omitido: ' + e.message);
    }

  } catch (error) {
    Logger.log('Error en sincronización automática: ' + error);
  }
}

/**
 * Sincroniza solo Clasificación de Perfiles (usado por sync automática y manual)
 */
function _syncClasificacionPerfiles() {
  var opciones = {
    method: 'get',
    headers: { 'Authorization': 'Token ' + KOBO_TOKEN, 'Accept': 'text/csv' },
    muteHttpExceptions: true
  };

  var urls = [URL_CLASIFICACION_PERFILES, URL_CLASIFICACION_FALLBACK];
  var respuesta;
  var contenido;

  for (var u = 0; u < urls.length; u++) {
    Logger.log('Intentando URL ' + (u + 1) + ': ' + urls[u]);
    respuesta = UrlFetchApp.fetch(urls[u], opciones);
    var codigo = respuesta.getResponseCode();
    Logger.log('HTTP ' + codigo);

    if (codigo !== 200) {
      Logger.log('Respuesta: ' + respuesta.getContentText().substring(0, 300));
      continue;
    }

    contenido = respuesta.getContentText();
    Logger.log('Primeros 300 chars: ' + contenido.substring(0, 300));

    // Si la respuesta es JSON (KoboToolbox a veces devuelve JSON en vez de CSV)
    if (contenido.trim().charAt(0) === '{' || contenido.trim().charAt(0) === '[') {
      Logger.log('Respuesta en JSON en URL ' + (u + 1) + ', intentando siguiente...');
      contenido = null;
      continue;
    }
    break; // CSV válido
  }

  if (!contenido) {
    var ultimoCodigo = respuesta ? respuesta.getResponseCode() : 0;
    throw new Error('No se pudo obtener CSV de KoboToolbox. Último HTTP: ' + ultimoCodigo +
      '\nVerifica que el token sea correcto y que la forma tenga respuestas.');
  }

  var datos = parsearCSV(contenido);
  Logger.log('Registros parseados: ' + (datos ? datos.length : 0));
  if (!datos || datos.length === 0) return { nuevos: 0, actualizados: 0, total: 0 };

  return procesarClasificacionPerfiles(datos);
}

/**
 * Configura un trigger para sincronización automática cada hora
 */
function configurarTriggerSincronizacion() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'sincronizacionAutomatica') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizacionAutomatica').timeBased().everyHours(1).create();
  Logger.log('Sincronización automática configurada para ejecutarse cada hora');
}

// ===========================================================================
// SECCIÓN 2C: AUTO-IMPORT — solo registros nuevos, cada N minutos
// ===========================================================================

// Claves PropertiesService para recordar la última fecha importada por hoja
var PROP_LAST_CLASIF  = 'LAST_SYNC_CLASIF';
var PROP_LAST_SATISF  = 'LAST_SYNC_SATISF';
var PROP_LAST_SESIONES = 'LAST_SYNC_SESIONES';

function _leerUltimaFecha(clave) {
  return PropertiesService.getScriptProperties().getProperty(clave) || '';
}
function _guardarUltimaFecha(clave, fechaISO) {
  if (fechaISO) PropertiesService.getScriptProperties().setProperty(clave, fechaISO);
}
function _resetearUltimaFecha(clave) {
  PropertiesService.getScriptProperties().deleteProperty(clave);
}

/**
 * Sincroniza Clasificación de Perfiles agregando SOLO filas nuevas (sin borrar).
 * Usa la última fecha importada (PropertiesService) para ignorar registros viejos
 * sin necesidad de leer toda la hoja en cada ejecución.
 */
function _syncClasificacionSoloNuevos() {
  // Prevenir ejecuciones concurrentes que generan duplicados
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('_syncClasificacionSoloNuevos: otra instancia en ejecución, omitiendo.');
    return { nuevos: 0, actualizados: 0, total: 0 };
  }
  try {

  var opciones = {
    method: 'get',
    headers: { 'Authorization': 'Token ' + KOBO_TOKEN, 'Accept': 'text/csv' },
    muteHttpExceptions: true
  };
  var contenido;
  var urls = [URL_CLASIFICACION_PERFILES, URL_CLASIFICACION_FALLBACK];
  for (var u = 0; u < urls.length; u++) {
    var resp = UrlFetchApp.fetch(urls[u], opciones);
    if (resp.getResponseCode() !== 200) continue;
    var body = resp.getContentText();
    var c0 = body.trim().charAt(0);
    if (c0 === '{' || c0 === '[') continue;
    contenido = body;
    break;
  }
  if (!contenido) throw new Error('No se pudo obtener CSV de KoboToolbox.');

  var datos = parsearCSV(contenido);
  if (!datos || datos.length === 0) return { nuevos: 0, actualizados: 0, total: 0 };

  var hoja = obtenerHoja('Clasificación de Perfiles');

  // Leer la última fecha que ya importamos
  var ultimaFechaImportada = _leerUltimaFecha(PROP_LAST_CLASIF);

  // Sin headers → importación completa (ignorar filtro de fecha, importar todo)
  if (hoja.getLastRow() <= 1) {
    var resultado = procesarClasificacionPerfiles(datos);
    // Guardar la fecha máxima de lo que se importó
    var maxFecha = '';
    datos.forEach(function(d) {
      var f = (d['_submission_time'] || '').toString().trim();
      if (f > maxFecha) maxFecha = f;
    });
    _guardarUltimaFecha(PROP_LAST_CLASIF, maxFecha);
    return resultado;
  }

  // Mapa inverso: nombre en hoja → clave KoboToolbox
  var M = 'MÓDULO DE OBSERVACIÓN - Evaluación de Perfil/';
  var INVERSO = {
    'Creamos ID':           'Creamos ID del participante:',
    'Fecha evaluación':     '_submission_time',
    'D1: Cuidado':          M+'DIMENSIÓN 1: Barreras de cuidado Basado en la conversación sobre responsabilidades en casa:',
    'D1 Comentario':        M+'Agrega comentario sobre DIEMENSIÓN 1',
    'D2: Violencia':        M+'DIMENSIÓN 2: Barreras de violencia/control Basado en las respuestas sobre trabajo en horarios variados / grupos mixtos / apoyo en casa:',
    'D2 Comentario':        M+'Agrega comentario sobre DIMENSIÓN 2',
    'D3: Movilidad':        M+'DIMENSIÓN 3: Barreras de movilidad/seguridad Basado en las respuestas sobre transporte y movilidad en la ciudad:',
    'D3 Comentario':        M+'Agrega comentario sobre DIMENSIÓN 3',
    'D4: Legal/Salud':      M+'DIMENSIÓN 4: Barreras legales/salud Basado en respuestas sobre antecedentes penales, casos legales, salud:',
    'D4 Comentario':        M+'Agrega comentario sobre DIMENSIÓN 4',
    'D5: Motivación':       M+'DIMENSIÓN 5: Motivación real / Prioridades Basado en las respuestas sobre qué quiere hacer en los próximos meses y qué tan importante es conseguir empleo:',
    'D5 Comentario':        M+'Agrega comentario sobre DIMENSIÓN 5',
    'D6: Experiencia':      M+'DIMENSIÓN 6: Experiencia previa en búsqueda de empleo Basado en si ha trabajado antes / buscado empleo / sabe qué hacer:',
    'D6 Comentario':        M+'Agrega comentario sobre DIMENSIÓN 6',
    'D7: Autonomía':        M+'DIMENSIÓN 7: Autonomía / Autoeficacia percibida Basado en el tono general, lenguaje corporal, y respuestas sobre planes y capacidad:',
    'D7 Comentario':        M+'Agrega comentario sobre DIMENSIÓN 7',
    'Puntaje Total':        M+'puntaje_total',
    'Barreras activas':     M+'tiene_barreras_activas',
    'Desmotivación':        M+'tiene_desmotivacion',
    'Perfil Asignado':      M+'perfil_asignado',
    'Notas de observación': M+'Notas de observación (opcional): Frases textuales, lenguaje corporal, o contexto adicional que influyó en tu evaluación:'
  };

  // Traducir headers de la hoja → claves KoboToolbox para leer cada fila
  var numCols  = hoja.getLastColumn();
  var ordenKobo = hoja.getRange(1, 1, 1, numCols).getValues()[0].map(function(h) {
    return INVERSO[h] || h;
  });

  // Registros ya existentes: clave = "Creamos ID || submission_time"
  // Solo se lee la hoja si no tenemos una fecha de referencia guardada
  var existentes = {};
  if (!ultimaFechaImportada) {
    var ultimaFila = hoja.getLastRow();
    if (ultimaFila > 1) {
      hoja.getRange(2, 1, ultimaFila - 1, 2).getValues().forEach(function(r) {
        var id = (r[0] || '').toString().trim();
        var f  = (r[1] || '').toString().trim();
        if (id) existentes[id + '||' + f] = true;
      });
    }
  }

  // Agregar solo registros nuevos (más nuevos que la última fecha importada)
  var nuevos = 0;
  var maxFecha = ultimaFechaImportada;
  datos.forEach(function(dato) {
    var id    = (dato['Creamos ID del participante:'] || '').toString().trim();
    if (!id) return;
    var fecha = (dato['_submission_time'] || '').toString().trim();

    // Saltar registros que ya teníamos (filtro por fecha)
    if (ultimaFechaImportada && fecha <= ultimaFechaImportada) return;

    // Segunda comprobación: clave compuesta (seguridad extra)
    var key = id + '||' + fecha;
    if (existentes[key]) return;

    hoja.appendRow(ordenKobo.map(function(k) { return dato[k] || ''; }));
    existentes[key] = true;
    nuevos++;
    if (fecha > maxFecha) maxFecha = fecha;
  });

  // Guardar la nueva fecha máxima importada
  _guardarUltimaFecha(PROP_LAST_CLASIF, maxFecha);

  // Auto-limpieza silenciosa: elimina cualquier duplicado que haya quedado
  _limpiarDuplicadosSilencioso(hoja);

  return { nuevos: nuevos, actualizados: 0, total: hoja.getLastRow() - 1 };

  } finally {
    lock.releaseLock();
  }
}

/**
 * Elimina filas duplicadas de una hoja sin mostrar alertas al usuario.
 * Clave: columna 1 (Creamos ID) + columna 2 (Fecha envío).
 * Llamada automáticamente al final de cada sync de Clasificación de Perfiles.
 */
function _limpiarDuplicadosSilencioso(hoja) {
  try {
    var uf = hoja.getLastRow();
    if (uf <= 2) return;
    var todos   = hoja.getRange(2, 1, uf - 1, hoja.getLastColumn()).getValues();
    var vistos  = {};
    var limpias = [];
    todos.forEach(function(fila) {
      var id    = (fila[0] || '').toString().trim();
      var fecha = (fila[1] || '').toString().trim();
      var clave = id ? id + '||' + fecha : (fecha ? 'SIN_ID||' + fecha : null);
      if (!clave) return;
      if (!vistos[clave]) { vistos[clave] = true; limpias.push(fila); }
    });
    if (limpias.length === todos.length) return;
    hoja.getRange(2, 1, uf - 1, hoja.getLastColumn()).clearContent();
    if (limpias.length > 0) {
      hoja.getRange(2, 1, limpias.length, hoja.getLastColumn()).setValues(limpias);
    }
    Logger.log('Auto-limpieza: eliminados ' + (todos.length - limpias.length) +
               ' duplicados en "' + hoja.getName() + '"');
  } catch (e) {
    Logger.log('_limpiarDuplicadosSilencioso error: ' + e.message);
  }
}

// ===========================================================================
// SECCIÓN 2D: SYNC IL_06 — SATISFACCIÓN EMPLEO
// ===========================================================================

/**
 * Sincroniza la hoja "Satisfacción Empleo" desde KoboToolbox (IL_06).
 * Solo agrega filas nuevas. Clave: Creamos ID + _submission_time.
 */
function _syncSatisfaccionSoloNuevos() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('_syncSatisfaccionSoloNuevos: otra instancia en ejecución, omitiendo.');
    return { nuevos: 0, actualizados: 0, total: 0 };
  }
  try {
    var opciones = {
      method: 'get',
      headers: { 'Authorization': 'Token ' + KOBO_TOKEN, 'Accept': 'text/csv' },
      muteHttpExceptions: true
    };
    var resp = UrlFetchApp.fetch(URL_SATISFACCION_EMPLEO, opciones);
    if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
    var body = resp.getContentText();
    if (body.trim().charAt(0) === '{' || body.trim().charAt(0) === '[') {
      throw new Error('KoboToolbox devolvió JSON en vez de CSV.');
    }
    var datos = parsearCSV(body);
    if (!datos || datos.length === 0) return { nuevos: 0, actualizados: 0, total: 0 };

    // Mapeo nuevo formulario → columnas de la hoja
    var NOMBRES = {
      'intro/Creamos_ID':     'Creamos ID',
      '_submission_time':     'Fecha envío',
      'satisfaccion_empleo':  'Satisfacción empleo',
      'cumple_expectativas':  'Cumple expectativas',
      'ambiente_laboral':     'Ambiente laboral',
      'salario_beneficios':   'Salario y beneficios',
      'permanencia':          'Permanencia',
      'aspecto_mejorar':      'Aspecto a mejorar',
      'otro_aspecto':         'Otro aspecto',
      'Nota':                 'Nota'
    };
    // Las 5 preguntas con escala 1-5 (para calcular promedio)
    var CAMPOS_NUMERICOS = [
      'satisfaccion_empleo', 'cumple_expectativas', 'ambiente_laboral',
      'salario_beneficios',  'permanencia'
    ];
    var ORDEN       = Object.keys(NOMBRES);
    var CAMPO_ID    = 'intro/Creamos_ID';
    var CAMPO_FECHA = '_submission_time';

    // Helper: calcular promedio de las 5 preguntas (ignorando vacíos)
    function calcularPromedio(d) {
      var suma = 0, n = 0;
      CAMPOS_NUMERICOS.forEach(function(c) {
        var v = parseFloat((d[c] || '').toString().trim());
        if (!isNaN(v)) { suma += v; n++; }
      });
      return n > 0 ? (suma / n).toFixed(2) : '';
    }

    // Helper: arma la fila completa (NOMBRES + Promedio)
    function armarFila(d) {
      var row = ORDEN.map(function(k) { return d[k] || ''; });
      row.push(calcularPromedio(d));
      return row;
    }

    var hoja = obtenerHoja('Satisfacción Empleo');
    var ultimaFechaSatisf = _leerUltimaFecha(PROP_LAST_SATISF);

    // --- Primera importación (hoja vacía) ----------------------------------
    if (hoja.getLastRow() <= 1) {
      var headers = ORDEN.map(function(k) { return NOMBRES[k]; });
      headers.push('Promedio (1-5)');
      hoja.clearContents();
      var hr = hoja.getRange(1, 1, 1, headers.length);
      hr.setValues([headers]);
      hr.setBackground('#00897b').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
      hoja.setFrozenRows(1);
      var nuevos = 0;
      var maxF = '';
      datos.forEach(function(d) {
        // No filtrar por Creamos ID — importar TODOS los registros (histórico)
        hoja.appendRow(armarFila(d));
        nuevos++;
        var f = (d[CAMPO_FECHA] || '').toString().trim();
        if (f > maxF) maxF = f;
      });
      _guardarUltimaFecha(PROP_LAST_SATISF, maxF);
      return { nuevos: nuevos, actualizados: 0, total: hoja.getLastRow() - 1 };
    }

    // --- Importación incremental (solo registros más nuevos que la última fecha) ---
    var existentes = {};
    if (!ultimaFechaSatisf) {
      var uf = hoja.getLastRow();
      if (uf > 1) {
        hoja.getRange(2, 1, uf - 1, 2).getValues().forEach(function(r) {
          var id = (r[0] || '').toString().trim();
          var f  = (r[1] || '').toString().trim();
          var k  = id ? (id + '||' + f) : ('noId||' + f);
          existentes[k] = true;
        });
      }
    }
    var nuevos = 0;
    var maxFecha = ultimaFechaSatisf;
    datos.forEach(function(d) {
      var id    = (d[CAMPO_ID] || '').toString().trim();
      var fecha = (d[CAMPO_FECHA] || '').toString().trim();
      if (ultimaFechaSatisf && fecha <= ultimaFechaSatisf) return;
      var key   = id ? (id + '||' + fecha) : ('noId||' + fecha);
      if (existentes[key]) return;
      hoja.appendRow(armarFila(d));
      existentes[key] = true;
      nuevos++;
      if (fecha > maxFecha) maxFecha = fecha;
    });
    _guardarUltimaFecha(PROP_LAST_SATISF, maxFecha);
    return { nuevos: nuevos, actualizados: 0, total: hoja.getLastRow() - 1 };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Importa manualmente la hoja "Satisfacción Empleo" desde el menú.
 */
function importarSatisfaccionEmpleo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Descargando Satisfacción Empleo...', '🔄 Importando', -1);
  try {
    var res = _syncSatisfaccionSoloNuevos();
    ss.toast('', '', 1);
    ui.alert('✅ Satisfacción Empleo importada',
      'Nuevos: ' + res.nuevos + '\nTotal en hoja: ' + res.total, ui.ButtonSet.OK);
  } catch (e) {
    ss.toast('', '', 1);
    ui.alert('❌ Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Limpia la hoja "Satisfacción Empleo" y la reimporta desde cero
 * (histórico completo) con las columnas del nuevo formulario.
 */
function reimportarSatisfaccionDesdeCero() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var confirmar = ui.alert(
    '🔄 Reimportar Satisfacción Empleo',
    'Esto borrará toda la hoja "Satisfacción Empleo" y la reimportará\n' +
    'desde KoboToolbox con la nueva estructura del formulario.\n\n' +
    'Se traerán TODOS los registros históricos.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;
  var hoja = ss.getSheetByName('Satisfacción Empleo');
  if (hoja) { hoja.clearContents(); hoja.clearFormats(); }
  _resetearUltimaFecha(PROP_LAST_SATISF);
  ss.toast('Reimportando Satisfacción Empleo...', '🔄', -1);
  try {
    var res = _syncSatisfaccionSoloNuevos();
    ss.toast('', '', 1);
    ui.alert('✅ Reimportación completada',
      'Registros importados: ' + res.nuevos + '\nTotal en hoja: ' + res.total,
      ui.ButtonSet.OK);
  } catch (e) {
    ss.toast('', '', 1);
    ui.alert('❌ Error', e.message, ui.ButtonSet.OK);
  }
}

// ===========================================================================
// SECCIÓN 2E: SYNC IL_08 — SESIONES ACOMPAÑAMIENTO PROFESIONAL
// ===========================================================================

/**
 * Sincroniza la hoja "Sesiones Acompañamiento" desde KoboToolbox (IL_08).
 * Solo agrega filas nuevas. Clave: Creamos ID + _submission_time.
 */
function _syncSesionesSoloNuevos() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('_syncSesionesSoloNuevos: otra instancia en ejecución, omitiendo.');
    return { nuevos: 0, actualizados: 0, total: 0 };
  }
  try {
    var opciones = {
      method: 'get',
      headers: { 'Authorization': 'Token ' + KOBO_TOKEN, 'Accept': 'text/csv' },
      muteHttpExceptions: true
    };
    var resp = UrlFetchApp.fetch(URL_SESIONES_ACOMPANAMIENTO, opciones);
    if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
    var body = resp.getContentText();
    if (body.trim().charAt(0) === '{' || body.trim().charAt(0) === '[') {
      throw new Error('KoboToolbox devolvió JSON en vez de CSV.');
    }
    var datos = parsearCSV(body);
    if (!datos || datos.length === 0) return { nuevos: 0, actualizados: 0, total: 0 };

    var NOMBRES = {
      'Creamos ID':                 'Creamos ID',
      '_submission_time':           'Fecha envío',
      'start':                      'Inicio sesión',
      'Proyecto':                   'Proyecto',
      'Nombre':                     'Nombre',
      'Apellidos':                  'Apellidos',
      'Teléfono':                   'Teléfono',
      'Fecha de nacimiento':        'Fecha nacimiento',
      'Edad_001':                   'Edad',
      'Género':                     'Género',
      'Año que ingreso a Creamos':  'Año ingreso Creamos',
      'En que año':                 'En qué año',
      'Grado académico':            'Grado académico',
      'Tipo de servicio':           'Tipo servicio',
      'Comentario':                 'Comentario'
    };
    // Lista blanca: SOLO las columnas de NOMBRES, en ese orden exacto
    var ORDEN  = Object.keys(NOMBRES);
    var CAMPO_ID    = 'Creamos ID';
    var CAMPO_FECHA = '_submission_time';

    var hoja = obtenerHoja('Sesiones Acompañamiento');

    var ultimaFechaSesiones = _leerUltimaFecha(PROP_LAST_SESIONES);

    // --- Primera importación (hoja vacía) ----------------------------------
    if (hoja.getLastRow() <= 1) {
      var headers = ORDEN.map(function(k) { return NOMBRES[k]; });
      hoja.clearContents();
      var hr = hoja.getRange(1, 1, 1, headers.length);
      hr.setValues([headers]);
      hr.setBackground('#5e35b1').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
      hoja.setFrozenRows(1);
      var nuevos = 0;
      var maxF = '';
      datos.forEach(function(d) {
        hoja.appendRow(ORDEN.map(function(k) { return d[k] || ''; }));
        nuevos++;
        var f = (d[CAMPO_FECHA] || '').toString().trim();
        if (f > maxF) maxF = f;
      });
      _guardarUltimaFecha(PROP_LAST_SESIONES, maxF);
      return { nuevos: nuevos, actualizados: 0, total: hoja.getLastRow() - 1 };
    }

    // --- Importación incremental (solo registros más nuevos que la última fecha) ---
    var existentes = {};
    if (!ultimaFechaSesiones) {
      var uf = hoja.getLastRow();
      if (uf > 1) {
        hoja.getRange(2, 1, uf - 1, 2).getValues().forEach(function(r) {
          var id = (r[0] || '').toString().trim();
          var f  = (r[1] || '').toString().trim();
          if (id) existentes[id + '||' + f] = true;
        });
      }
    }
    var nuevos = 0;
    var maxFecha = ultimaFechaSesiones;
    datos.forEach(function(d) {
      var id    = (d[CAMPO_ID] || '').toString().trim();
      var fecha = (d[CAMPO_FECHA] || '').toString().trim();
      if (ultimaFechaSesiones && fecha <= ultimaFechaSesiones) return;
      // Use Creamos ID if available; otherwise fall back to date-only for deduplication
      var key   = (id ? id + '||' : 'noId||') + fecha;
      if (existentes[key]) return;
      hoja.appendRow(ORDEN.map(function(k) { return d[k] || ''; }));
      existentes[key] = true;
      nuevos++;
      if (fecha > maxFecha) maxFecha = fecha;
    });
    _guardarUltimaFecha(PROP_LAST_SESIONES, maxFecha);
    return { nuevos: nuevos, actualizados: 0, total: hoja.getLastRow() - 1 };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Importa manualmente "Sesiones Acompañamiento" desde el menú.
 */
function importarSesionesAcompanamiento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Descargando Sesiones Acompañamiento...', '🔄 Importando', -1);
  try {
    var res = _syncSesionesSoloNuevos();
    ss.toast('', '', 1);
    ui.alert('✅ Sesiones Acompañamiento importadas',
      'Nuevos: ' + res.nuevos + '\nTotal en hoja: ' + res.total, ui.ButtonSet.OK);
  } catch (e) {
    ss.toast('', '', 1);
    ui.alert('❌ Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Limpia la hoja "Sesiones Acompañamiento" y la reimporta desde cero
 * con solo las columnas limpias (sin duplicados ni sub-columnas de Kobo).
 */
function reimportarSesionesDesdeCero() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var confirmar = ui.alert(
    '🔄 Reimportar Sesiones Acompañamiento',
    'Esto borrará toda la hoja "Sesiones Acompañamiento" y la volverá a importar\n' +
    'desde KoboToolbox con las columnas correctas (sin columnas repetidas).\n\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;
  var hoja = ss.getSheetByName('Sesiones Acompañamiento');
  if (hoja) { hoja.clearContents(); hoja.clearFormats(); }
  _resetearUltimaFecha(PROP_LAST_SESIONES); // reiniciar marca de tiempo para reimportar todo
  ss.toast('Reimportando Sesiones Acompañamiento...', '🔄', -1);
  try {
    var res = _syncSesionesSoloNuevos();
    ss.toast('', '', 1);
    ui.alert('✅ Reimportación completada',
      'Registros importados: ' + res.nuevos + '\nTotal en hoja: ' + res.total,
      ui.ButtonSet.OK);
  } catch (e) {
    ss.toast('', '', 1);
    ui.alert('❌ Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Función llamada por el trigger de tiempo. Sin alertas — solo toast y Logger.
 */
function autoImportarNuevos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var msgs = [];
  try {
    ss.toast('Revisando nuevos registros en KoboToolbox...', '🔄 Auto-import', 5);

    try {
      var datos = obtenerDatosKoboToolbox();
      var resG  = procesarDatosGraduados(datos);
      if (resG.nuevos > 0) msgs.push('Graduados: +' + resG.nuevos);
    } catch(e) { Logger.log('Auto-import Graduados omitido: ' + e.message); }

    try {
      var resC = _syncClasificacionSoloNuevos();
      if (resC.nuevos > 0) msgs.push('Clasificación: +' + resC.nuevos);
    } catch(e) { Logger.log('Auto-import Clasificación omitido: ' + e.message); }

    try {
      var resS = _syncSatisfaccionSoloNuevos();
      if (resS.nuevos > 0) msgs.push('Satisfacción: +' + resS.nuevos);
    } catch(e) { Logger.log('Auto-import Satisfacción omitido: ' + e.message); }

    try {
      var resSe = _syncSesionesSoloNuevos();
      if (resSe.nuevos > 0) msgs.push('Sesiones: +' + resSe.nuevos);
    } catch(e) { Logger.log('Auto-import Sesiones omitido: ' + e.message); }

    if (msgs.length > 0) {
      ss.toast(msgs.join(' | '), '✅ Nuevos registros importados', 10);
    }
    Logger.log('Auto-import: ' + (msgs.join(' | ') || 'sin cambios'));
  } catch(e) {
    Logger.log('Error en auto-import: ' + e.message);
  }
}

/**
 * Activa el trigger automático. Le pregunta al usuario el intervalo en minutos.
 */
function activarAutoImport() {
  var ui  = SpreadsheetApp.getUi();
  var res = ui.prompt(
    '⏱ Activar Auto-import',
    'Cada cuántos minutos revisar si hay nuevos registros?\n(mínimo 1 — recomendado: 5)',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var mins = parseInt(res.getResponseText()) || 5;
  if (mins < 1) mins = 1;

  // Eliminar triggers previos del mismo handler
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoImportarNuevos') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('autoImportarNuevos').timeBased().everyMinutes(mins).create();
  PropertiesService.getScriptProperties().setProperty('AUTO_IMPORT_MINS', mins.toString());

  ui.alert(
    '✅ Auto-import activado',
    'Cada ' + mins + ' minutos se revisará si hay nuevos registros en KoboToolbox.\n\n' +
    '• Solo agrega filas nuevas — no borra ni modifica las existentes.\n' +
    '• Para detenerlo: menú → ⏹ Desactivar auto-import.',
    ui.ButtonSet.OK
  );
}

/**
 * Desactiva el trigger automático de auto-import.
 */
function desactivarAutoImport() {
  var ui = SpreadsheetApp.getUi();
  var encontrado = false;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoImportarNuevos') {
      ScriptApp.deleteTrigger(t);
      encontrado = true;
    }
  });
  PropertiesService.getScriptProperties().deleteProperty('AUTO_IMPORT_MINS');
  ui.alert(
    encontrado ? '⏹ Auto-import desactivado' : 'ℹ Auto-import',
    encontrado ? 'Ya no se importarán datos automáticamente.' : 'El auto-import no estaba activado.',
    ui.ButtonSet.OK
  );
}

// ===========================================================================
// SECCIÓN 2B: IMPORTACIÓN DE CLASIFICACIÓN DE PERFILES
// ===========================================================================

/**
 * Elimina filas duplicadas de "Clasificación de Perfiles".
 * Clave de unicidad: Creamos ID (col 1) + Fecha evaluación (col 2).
 * Conserva la primera aparición de cada registro y elimina las repetidas.
 */
function limpiarDuplicadosClasificacion() {
  var ui   = SpreadsheetApp.getUi();
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clasificación de Perfiles');
  if (!hoja) {
    ui.alert('❌ Error', 'No se encontró la hoja "Clasificación de Perfiles".', ui.ButtonSet.OK);
    return;
  }

  var ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) {
    ui.alert('ℹ Sin datos', 'La hoja no tiene registros para limpiar.', ui.ButtonSet.OK);
    return;
  }

  var datos   = hoja.getRange(1, 1, ultimaFila, hoja.getLastColumn()).getValues();
  var headers = datos[0];
  var filas   = datos.slice(1);

  var vistos    = {};
  var limpias   = [];
  var eliminadas = 0;

  filas.forEach(function(fila) {
    var id    = (fila[0] || '').toString().trim();
    var fecha = (fila[1] || '').toString().trim();
    var clave;
    if (id) {
      clave = id + '||' + fecha;
    } else if (fecha) {
      clave = 'SIN_ID||' + fecha; // deduplicar por fecha cuando no hay ID
    } else {
      return; // fila completamente vacía, descartar
    }
    if (vistos[clave]) {
      eliminadas++;
    } else {
      vistos[clave] = true;
      limpias.push(fila);
    }
  });

  if (eliminadas === 0) {
    ui.alert('✅ Sin duplicados', 'No se encontraron registros duplicados.', ui.ButtonSet.OK);
    return;
  }

  var confirmar = ui.alert(
    '🗑 Eliminar duplicados',
    'Se encontraron ' + eliminadas + ' fila(s) duplicada(s).\n' +
    'Quedarán ' + limpias.length + ' registros únicos.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;

  // Reescribir hoja con datos limpios
  hoja.clearContents();
  hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (limpias.length > 0) {
    hoja.getRange(2, 1, limpias.length, headers.length).setValues(limpias);
  }

  ui.alert(
    '✅ Limpieza completada',
    'Se eliminaron ' + eliminadas + ' duplicado(s).\nRegistros únicos: ' + limpias.length,
    ui.ButtonSet.OK
  );
}

/**
 * Importa datos de Clasificación de Perfiles desde KoboToolbox
 */
function importarClasificacionPerfiles() {
  try {
    const ui = SpreadsheetApp.getUi();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Descargando Clasificación de Perfiles...', '🔄 Importando', -1
    );

    const resultado = _syncClasificacionSoloNuevos();
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    ui.alert('✅ Importación Completada',
      'Clasificaciones nuevas: ' + resultado.nuevos + '\n' +
      'Total en hoja: ' + resultado.total,
      ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    SpreadsheetApp.getUi().alert('❌ Error',
      'Error al importar clasificaciones: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error en importarClasificacionPerfiles: ' + error);
  }
}

/**
 * Diagnóstico: muestra qué devuelve KoboToolbox para Clasificación de Perfiles
 */
function diagnosticarClasificacionPerfiles() {
  var ui = SpreadsheetApp.getUi();
  try {
    var opciones = {
      method: 'get',
      headers: { 'Authorization': 'Token ' + KOBO_TOKEN, 'Accept': 'text/csv' },
      muteHttpExceptions: true
    };

    var urls = [URL_CLASIFICACION_PERFILES, URL_CLASIFICACION_FALLBACK];
    var msg = '🔍 Diagnóstico KoboToolbox — Clasificación de Perfiles\n\n';

    for (var u = 0; u < urls.length; u++) {
      var resp = UrlFetchApp.fetch(urls[u], opciones);
      var codigo = resp.getResponseCode();
      var cuerpo = resp.getContentText();
      var esJSON = cuerpo.trim().charAt(0) === '{' || cuerpo.trim().charAt(0) === '[';

      msg += 'URL ' + (u + 1) + ': HTTP ' + codigo + (esJSON ? ' (JSON — no CSV)' : ' (CSV OK)') + '\n';

      if (codigo === 200 && !esJSON) {
        var lineas = cuerpo.split('\n').filter(function(l) { return l.trim(); });
        msg += '  Filas totales: ' + (lineas.length - 1) + '\n';

        if (lineas.length > 0) {
          var headers = lineas[0].split(',').map(function(h) { return h.replace(/"/g, '').trim(); });
          var tieneCreamos = headers.some(function(h) { return h.indexOf('Creamos ID') !== -1; });
          msg += '  Columnas encontradas: ' + headers.length + '\n';
          msg += '  Tiene "Creamos ID": ' + (tieneCreamos ? 'SÍ ✅' : 'NO ❌') + '\n';

          // Mostrar columnas que contienen "Creamos" o "MÓDULO"
          var relevantes = headers.filter(function(h) {
            return h.indexOf('Creamos') !== -1 || h.indexOf('MÓDULO') !== -1 || h.indexOf('_submission') !== -1;
          });
          if (relevantes.length > 0) {
            msg += '  Cols relevantes:\n';
            relevantes.forEach(function(c) { msg += '    • ' + c.substring(0, 60) + '\n'; });
          }
        }
        break;
      } else if (codigo === 200 && esJSON) {
        try {
          var json = JSON.parse(cuerpo);
          msg += '  Respuesta JSON — count: ' + (json.count || '?') + '\n';
        } catch(e) {}
      } else {
        msg += '  Error: ' + cuerpo.substring(0, 100) + '\n';
      }
    }

    ui.alert('🔍 Diagnóstico', msg, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error en diagnóstico', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Procesa los datos de clasificación y los inserta en la hoja
 * @param {Array} datos - Array de objetos del CSV
 * @return {Object}
 */
function procesarClasificacionPerfiles(datos) {
  if (!datos || datos.length === 0) return { nuevos: 0, actualizados: 0, total: 0 };

  var hoja = obtenerHoja('Clasificación de Perfiles');

  // ── Columnas a EXCLUIR (vacías o solo texto del formulario sin datos reales) ──
  var _excluir = function(col) {
    var c = col;
    if (c.indexOf('GUIÓN') !== -1 || c.indexOf('Guión') !== -1) return true;
    if (c.indexOf('INSTRUCCIONES') !== -1) return true;
    if (c.indexOf('RESULTADO DE EVALUACIÓN') !== -1) return true;
    if (c.indexOf('ACEPTA') !== -1) return true;    // cubre ✅ y ⛔
    if (c === '_id')                 return true;
    if (c === '_uuid')               return true;
    if (c === '_validation_status')  return true;
    if (c === '_notes')              return true;
    if (c === '_status')             return true;
    if (c === '_submitted_by')       return true;
    if (c === '__version__')         return true;
    if (c === '_tags')               return true;
    if (c === 'meta/rootUuid')       return true;
    if (c === '_index')              return true;
    if (/^Columna\s+\d+$/i.test(c)) return true;   // Columna 46, etc.
    return false;
  };

  // ── Mapeo: nombre KoboToolbox → nombre corto legible ──
  var M = 'MÓDULO DE OBSERVACIÓN - Evaluación de Perfil/';
  var NOMBRES = {
    'Creamos ID del participante:': 'Creamos ID',
    '_submission_time': 'Fecha evaluación',
    [M+'DIMENSIÓN 1: Barreras de cuidado Basado en la conversación sobre responsabilidades en casa:']: 'D1: Cuidado',
    [M+'Agrega comentario sobre DIEMENSIÓN 1']: 'D1 Comentario',
    [M+'DIMENSIÓN 2: Barreras de violencia/control Basado en las respuestas sobre trabajo en horarios variados / grupos mixtos / apoyo en casa:']: 'D2: Violencia',
    [M+'Agrega comentario sobre DIMENSIÓN 2']: 'D2 Comentario',
    [M+'DIMENSIÓN 3: Barreras de movilidad/seguridad Basado en las respuestas sobre transporte y movilidad en la ciudad:']: 'D3: Movilidad',
    [M+'Agrega comentario sobre DIMENSIÓN 3']: 'D3 Comentario',
    [M+'DIMENSIÓN 4: Barreras legales/salud Basado en respuestas sobre antecedentes penales, casos legales, salud:']: 'D4: Legal/Salud',
    [M+'Agrega comentario sobre DIMENSIÓN 4']: 'D4 Comentario',
    [M+'DIMENSIÓN 5: Motivación real / Prioridades Basado en las respuestas sobre qué quiere hacer en los próximos meses y qué tan importante es conseguir empleo:']: 'D5: Motivación',
    [M+'Agrega comentario sobre DIMENSIÓN 5']: 'D5 Comentario',
    [M+'DIMENSIÓN 6: Experiencia previa en búsqueda de empleo Basado en si ha trabajado antes / buscado empleo / sabe qué hacer:']: 'D6: Experiencia',
    [M+'Agrega comentario sobre DIMENSIÓN 6']: 'D6 Comentario',
    [M+'DIMENSIÓN 7: Autonomía / Autoeficacia percibida Basado en el tono general, lenguaje corporal, y respuestas sobre planes y capacidad:']: 'D7: Autonomía',
    [M+'Agrega comentario sobre DIMENSIÓN 7']: 'D7 Comentario',
    [M+'puntaje_total']: 'Puntaje Total',
    [M+'tiene_barreras_activas']: 'Barreras activas',
    [M+'tiene_desmotivacion']: 'Desmotivación',
    [M+'perfil_asignado']: 'Perfil Asignado',
    [M+'Notas de observación (opcional): Frases textuales, lenguaje corporal, o contexto adicional que influyó en tu evaluación:']: 'Notas de observación'
  };

  // Obtener columnas del CSV y filtrar
  var todasColumnas     = Object.keys(datos[0]);
  var columnasFiltradas = todasColumnas.filter(function(c) { return !_excluir(c); });

  // ── Cargar nombres desde base de datos Salesforce (fuente de verdad) ──
  var mapaGrads = _cargarMapaCreamos_();

  var colId    = 'Creamos ID del participante:';
  var colFecha = '_submission_time';
  // Columnas de evaluación: todo excepto ID y fecha (ya van en las 6 base)
  var evalCols = columnasFiltradas.filter(function(c) { return c !== colId && c !== colFecha; });

  // Headers: 3 columnas base + columnas de evaluación
  var headersLimpios = [
    'Fecha de ingreso',
    'Creamos ID',
    'Nombre completo'
  ].concat(evalCols.map(function(col) { return NOMBRES[col] || col; }));

  // ── Limpiar hoja y escribir headers con el mismo estilo ──
  hoja.clearContents();
  hoja.clearFormats();
  hoja.clearConditionalFormatRules();
  hoja.getDataRange().clearDataValidations();
  var headerRange = hoja.getRange(1, 1, 1, headersLimpios.length);
  headerRange.setValues([headersLimpios]);
  headerRange.setBackground('#5c6bc0')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setFontSize(11)
             .setWrap(false);
  hoja.setRowHeight(1, 36);
  hoja.setFrozenRows(1);

  // ── Escribir filas: 6 columnas base + datos de evaluación ──
  var nuevos = 0;
  datos.forEach(function(dato) {
    var cid = (dato[colId] || '').toString().trim();
    if (!cid) return;
    var matchGrad = _buscarEnMapaFuzzy_(cid, mapaGrads);
    var grad = matchGrad ? matchGrad.rec : {};
    var fila = [
      dato[colFecha] || '',   // Fecha de ingreso (= fecha evaluación KoboToolbox)
      cid,                    // Creamos ID
      grad.nombre   || ''     // Nombre completo (desde Graduados)
    ].concat(evalCols.map(function(col) { return dato[col] || ''; }));
    hoja.appendRow(fila);
    nuevos++;
  });

  return { nuevos: nuevos, actualizados: 0, total: hoja.getLastRow() - 1 };
}

/**
 * Extrae el número de puntaje desde un texto como "(4 pts) No tiene personas..."
 * @return {number}
 */
function _puntajeDesdeTexto(valor) {
  if (!valor) return 0;
  var match = valor.toString().match(/\((\d+)\s*pt/i);
  if (match) return parseInt(match[1]);
  var num = parseInt(valor);
  return isNaN(num) ? 0 : num;
}

/**
 * Mapea el valor perfil_asignado de KoboToolbox a nuestros nombres estándar.
 * Si no coincide, calcula por puntaje.
 */
function _mapearPerfil(perfilKobo, puntaje) {
  var p = (perfilKobo || '').toString().toLowerCase();
  // "Perfil 2 Bajo" antes que "Perfil 2" para evitar falsa coincidencia
  if (p.indexOf('perfil 1') !== -1 || p.indexOf('completamente independiente') !== -1) return PERFILES_CLASIFICACION[0];
  if (p.indexOf('perfil 2 bajo') !== -1 || p.indexOf('acompañamiento intensivo') !== -1) return PERFILES_CLASIFICACION[2];
  if (p.indexOf('perfil 2') !== -1 || p.indexOf('necesita apoyo') !== -1) return PERFILES_CLASIFICACION[1];
  if (p.indexOf('perfil 3') !== -1 || p.indexOf('barreras activas') !== -1 || p.indexOf('no puede') !== -1) return PERFILES_CLASIFICACION[3];
  if (p.indexOf('perfil 4') !== -1 || p.indexOf('no quiere') !== -1 || p.indexOf('otras prioridades') !== -1) return PERFILES_CLASIFICACION[4];
  // Fallback: calcular por puntaje
  if (puntaje >= 22) return PERFILES_CLASIFICACION[0];
  if (puntaje >= 15) return PERFILES_CLASIFICACION[1];
  if (puntaje >= 8)  return PERFILES_CLASIFICACION[2];
  return PERFILES_CLASIFICACION[3];
}

/**
 * Busca datos personales de un graduado por Creamos ID en la hoja Graduados.
 * Graduados: col3=Creamos ID, col4=Nombre, col5=Género, col6=Edad, col7=Nivel edu, col8=Teléfono
 */
function _buscarGraduadoPorCreamosId(creamosId) {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Graduados');
    if (!hoja || hoja.getLastRow() <= 1) return {};
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][2] && datos[i][2].toString().trim() === creamosId.trim()) {
        return {
          nombre:   datos[i][3] || '',
          genero:   datos[i][4] || '',
          edad:     datos[i][5] || '',
          nivelEdu: datos[i][6] || '',
          telefono: datos[i][7] || ''
        };
      }
    }
    return {};
  } catch (e) { return {}; }
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

  // Asegurar que la hoja tiene headers correctos antes de agregar datos.
  // Esto previene el bug donde una hoja sin headers recibe datos en fila 1
  // (haciéndola parecer "vacía" o "corrupta").
  const estructura = ESTRUCTURA_HOJAS[nombreHoja];
  if (estructura && estructura.columnas && estructura.columnas.length > 0) {
    const headersEsperados = estructura.columnas.map(c => c.nombre);
    const lastRow = hoja.getLastRow();
    const lastCol = hoja.getLastColumn();

    let necesitaReconstruir = false;
    if (lastRow < 1 || lastCol < 1) {
      // Hoja completamente vacía
      necesitaReconstruir = true;
    } else {
      // Verificar que el primer header coincide
      const primerHeader = hoja.getRange(1, 1).getValue();
      if (!primerHeader || primerHeader.toString().trim() === '') {
        necesitaReconstruir = true;
      }
    }

    if (necesitaReconstruir) {
      Logger.log('Reconstruyendo hoja sin headers: ' + nombreHoja);
      _construirHoja(hoja, nombreHoja);
    }
  }

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
    'Paso a paso':                'Paso a paso',
    // legacy
    'No busca trabajo':           'Paso a paso',
    'Fito':                       'Paso a paso',
    'No busca trabajo - Fito':    'Paso a paso',
    'Activamente busca trabajo':  'Activamente busca trabajo',
    'Conexiones Laborales':       'Conexiones Laborales',
    // legacy / fusionado
    'Por su cuenta':              'Activamente busca trabajo',
    'Por su Cuenta':              'Activamente busca trabajo',
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
        datosAdicionales.compartioCv           || 'No',
        datosAdicionales.area                  || '',
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
        'Por su cuenta',              // Tipo de búsqueda
        datosAdicionales.mensaje || '',
        datosAdicionales.llamada || '',
        datosAdicionales.nota    || '',
        '',                           // Entrevista
        '',                           // Trámites
        datosAdicionales.activo  || 'Sí'
      ]);

    case 'Paso a paso':
    case 'No busca trabajo':
    case 'Fito':
      return filaBase.concat([
        datosAdicionales.dpi       || '',
        datosAdicionales.formacion || datosGraduado[8] || '',  // índice 8 = Formación
        datosAdicionales.cohorte   || datosGraduado[9] || '',  // índice 9 = Cohorte
        datosAdicionales.nota      || '',
        datosAdicionales.activo    || 'No'
      ]);

    case 'Activamente busca trabajo':
    case 'Empleado':
      return filaBase.concat([
        'Activamente busca trabajo',  // Tipo de búsqueda
        '',                           // Mensaje
        '',                           // Llamada
        '',                           // Nota
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
    'Paso a paso', 'Activamente busca trabajo', 'Sesiones Acompañamiento'
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
  // Sesiones Acompañamiento: 1=Creamos ID, 2=Fecha envío, 3=Inicio sesión, 4=Proyecto, 5=Nombre, 6=Apellidos, 7=Teléfono, 8=Fecha nacimiento, 9=Edad, 10=Género
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
  } else if (nombreHoja === 'Sesiones Acompañamiento') {
    const datos = hoja.getRange(filaActiva, 1, 1, 10).getValues()[0];
    creamosId      = datos[0] || '';  // Col 1 = Creamos ID
    nombreCompleto = ((datos[4] || '') + ' ' + (datos[5] || '')).trim();  // Col 5 = Nombre, Col 6 = Apellidos
    datosExtra = {
      telefono:       datos[6] || '',  // Col 7 = Teléfono
      edad:           datos[8] || '',  // Col 9 = Edad
      genero:         datos[9] || '',  // Col 10 = Género
      nivelEducativo: ''
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
 * Envía a la persona de la sesión seleccionada al bot de seguimiento de búsqueda activa.
 * El usuario debe tener seleccionada una fila en la hoja "Sesiones Acompañamiento".
 */
function enviarSesionASeguimientoBot() {
  const ui   = SpreadsheetApp.getUi();
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (hoja.getName() !== 'Sesiones Acompañamiento') {
    ui.alert('⚠️ Hoja incorrecta',
             'Debes estar en la hoja "Sesiones Acompañamiento" y seleccionar una fila.',
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

  // Sesiones Acompañamiento: Col1=Creamos ID, Col5=Nombre, Col6=Apellidos, Col7=Teléfono
  const datos        = hoja.getRange(filaActiva, 1, 1, 7).getValues()[0];
  const creamosId    = (datos[0] || '').toString().trim();
  const nombre       = (datos[4] || '').toString().trim();
  const apellidos    = (datos[5] || '').toString().trim();
  const telefono     = (datos[6] || '').toString().trim();
  const nombreCompleto = (nombre + ' ' + apellidos).trim();

  if (!nombreCompleto) {
    ui.alert('⚠️ Sin datos', 'La fila seleccionada no tiene nombre.', ui.ButtonSet.OK);
    return;
  }

  crearSeguimientoBusquedaActiva(creamosId, nombreCompleto, telefono);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    nombreCompleto + ' enviado al bot de seguimiento.', '🔔 Seguimiento activado', 4
  );
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

    var telefono = (datos.telefono || '').toString().trim();
    if (telefono && !telefono.startsWith('+')) {
      telefono = '+' + telefono;
    }

    const fila = [
      datos.creamosId      || '',   // 1  Creamos ID
      datos.nombreCompleto || '',   // 2  Nombre
      telefono,                     // 3  Telefono
      datos.empresa        || '',   // 4  Empresa
      datos.cargo          || '',   // 5  Cargo
      fechaEmpleo,                  // 6  Fecha Empleo
      fmt(fechaSeg1),               // 7  Fecha Seg1 (hoy)
      fmt(fechaSeg2),               // 8  Fecha Seg2 (+14 días)
      fmt(fechaRecord),             // 9  Fecha Recordatorio (+90 días)
      'Pendiente',                  // 10 Estado
      '',                           // 11 Etapa Actual
      '', '', '',                   // 12-14 Resp S1 P1, P2, P3
      '', '', '',                   // 15-17 Resp S2 P1, P2, P3
      'No',                         // 18 Email Enviado
      'Post-empleo',                // 19 Tipo Seguimiento
      'Conexiones Laborales'        // 20 Origen
    ];

    hoja.appendRow(fila);
    Logger.log('Seguimiento Bot (post-empleo) creado para: ' + datos.nombreCompleto);

  } catch (error) {
    Logger.log('Error al crear fila en Seguimiento Bot: ' + error);
    // No lanzar el error para no interrumpir el guardado de Conexiones Laborales
  }
}

/**
 * Crea una fila en "Seguimiento Bot" cuando alguien entra a "Activamente busca trabajo".
 * Flujo: onEdit detecta Etapa → llama esta función.
 *
 * Día 7  → WhatsApp Mensaje 1 (bot n8n)
 * Día 14 → WhatsApp Mensaje 2 (bot n8n)
 * Día 21 → Llamada por colaborador (marcado como "Llamada Programada")
 */
function crearSeguimientoBusquedaActiva(creamosId, nombre, telefono) {
  try {
    const hoja = obtenerHoja('Seguimiento Bot');

    const hoy = new Date();
    const seg1 = new Date(hoy); seg1.setDate(hoy.getDate() + 7);   // Día 7
    const seg2 = new Date(hoy); seg2.setDate(hoy.getDate() + 14);  // Día 14
    const rec  = new Date(hoy); rec.setDate(hoy.getDate() + 21);   // Día 21 (llamada)

    function fmt(d) {
      return ('0' + d.getDate()).slice(-2) + '/' +
             ('0' + (d.getMonth() + 1)).slice(-2) + '/' +
             d.getFullYear();
    }

    var tel = (telefono || '').toString().trim();
    if (tel && !tel.startsWith('+')) tel = '+' + tel;

    const fila = [
      creamosId || '',              // 1  Creamos ID
      nombre    || '',              // 2  Nombre
      tel,                          // 3  Telefono
      '',                           // 4  Empresa (vacío: aún no tiene)
      '',                           // 5  Cargo   (vacío)
      fmt(hoy),                     // 6  Fecha inicio búsqueda
      fmt(seg1),                    // 7  Fecha Seg1 (+7 días → WhatsApp 1)
      fmt(seg2),                    // 8  Fecha Seg2 (+14 días → WhatsApp 2)
      fmt(rec),                     // 9  Fecha Recordatorio (+21 días → Llamada)
      'Pendiente',                  // 10 Estado
      'Búsqueda Activa',            // 11 Etapa Actual
      '', '', '',                   // 12-14 Resp S1 P1, P2, P3
      '', '', '',                   // 15-17 Resp S2 P1, P2, P3
      'No',                         // 18 Email Enviado
      'Búsqueda activa',            // 19 Tipo Seguimiento
      'Activamente busca trabajo'   // 20 Origen
    ];

    hoja.appendRow(fila);
    Logger.log('Seguimiento Bot (búsqueda activa) creado para: ' + nombre);

  } catch (error) {
    Logger.log('Error en crearSeguimientoBusquedaActiva: ' + error);
  }
}

// ===========================================================================
// SECCIÓN 4: SEGUIMIENTOS Y LLAMADAS PROGRAMADAS
// ===========================================================================

/**
 * Programa los seguimientos automáticos para un graduado empleado/activo.
 * NOTA: Función legacy. El reemplazo es crearSeguimientoBusquedaActiva()
 * que escribe a la hoja "Seguimiento Bot". Esta función solo escribe si
 * la hoja "Seguimientos" ya existe (no la crea automáticamente).
 * @param {string} graduadoId
 * @param {string} nombreGraduado
 */
function programarSeguimientos(graduadoId, nombreGraduado) {
  // Usar getSheetByName (NO obtenerHoja) para evitar crear la hoja legacy.
  const hojaSeguimientos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Seguimientos');
  if (!hojaSeguimientos) {
    Logger.log('Hoja "Seguimientos" no existe; se omite programarSeguimientos (legacy)');
    return;
  }

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
  // Usar getSheetByName (NO obtenerHoja) para evitar crear la hoja legacy "Seguimientos"
  // si no existe. Si no existe, no hay pendientes que reportar.
  const hojaSeguimientos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Seguimientos');
  if (!hojaSeguimientos) return [];
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
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Seguimientos');
  if (!hoja) {
    Logger.log('Hoja "Seguimientos" no existe; se omite marcarSeguimientoRealizado');
    return;
  }
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
  return !!(config.koboExportUrl && KOBO_TOKEN);
}

/**
 * Obtiene la configuración actual del sistema
 * @return {Object}
 */
function obtenerConfiguracion() {
  const props = PropertiesService.getScriptProperties();
  return {
    koboExportUrl:           props.getProperty('KOBO_EXPORT_URL') || '',
    koboClasificacionUrl:    props.getProperty('KOBO_CLASIFICACION_URL') || '',
    koboToken:               props.getProperty('KOBO_TOKEN')      || '',
    emailNotificaciones:     props.getProperty('EMAIL_NOTIFICACIONES') || '',
    sincronizacionAuto:      props.getProperty('SINCRONIZACION_AUTO') === 'true'
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
 * Elimina triggers duplicados: por cada función instalada más de una vez, conserva solo 1.
 * También ofrece la opción de desactivar el trigger horario (sincronizacionAutomatica).
 */
function limpiarTriggersDuplicados() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var grupos = {};
  triggers.forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (!grupos[fn]) grupos[fn] = [];
    grupos[fn].push(t);
  });

  var eliminados = 0;
  var detalle = [];
  Object.keys(grupos).forEach(function(fn) {
    var lista = grupos[fn];
    if (lista.length > 1) {
      // Conservar el primero, eliminar el resto
      for (var i = 1; i < lista.length; i++) {
        ScriptApp.deleteTrigger(lista[i]);
        eliminados++;
      }
      detalle.push('• ' + fn + ': ' + lista.length + ' → 1 (eliminados ' + (lista.length - 1) + ')');
    }
  });

  // Mostrar resumen
  var msg = eliminados > 0
    ? 'Se eliminaron ' + eliminados + ' trigger(s) duplicado(s):\n\n' + detalle.join('\n')
    : 'No se encontraron triggers duplicados. Todo está limpio.';

  // Listar todos los triggers activos
  var activos = ScriptApp.getProjectTriggers().map(function(t) {
    return '• ' + t.getHandlerFunction() + ' (' + t.getTriggerSource() + ')';
  });
  msg += '\n\nTriggers activos:\n' + (activos.length > 0 ? activos.join('\n') : 'Ninguno');

  ui.alert('🧹 Limpiar triggers duplicados', msg, ui.ButtonSet.OK);
  Logger.log('limpiarTriggersDuplicados: eliminados=' + eliminados);
}

/**
 * Prueba la conexión con KoboToolbox
 * @return {Object}
 */
function probarConexionKobo() {
  try {
    const config = obtenerConfiguracion();
    if (!config.koboExportUrl) {
      return { exito: false, mensaje: 'URL de exportación no configurada.' };
    }
    const respuesta = UrlFetchApp.fetch(config.koboExportUrl, {
      method: 'get',
      headers: { 'Authorization': 'Token ' + KOBO_TOKEN },
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('Reporte') || ss.insertSheet('Reporte');
  hoja.clear();
  hoja.setTabColor('#e91e63');

  // ── Paleta de colores ─────────────────────────────────────────────────────
  const C = {
    bannerBg:   '#1a237e', bannerFg:   '#ffffff',
    secBg:      '#283593', secFg:      '#ffffff',
    headerBg:   '#3949ab', headerFg:   '#ffffff',
    kpiBg:      '#0d47a1', kpiFg:      '#ffffff',
    par:        '#e8eaf6', impar:      '#ffffff',
    total:      '#c5cae9', totalFg:    '#1a237e',
    koboBg:     '#e0f2f1', koboAcc:    '#00695c',
    satisfBg:   '#e8f5e9', satisfAcc:  '#2e7d32',
    sessBg:     '#ede7f6', sessAcc:    '#4527a0',
    conBg:      '#fff3e0', conAcc:     '#e65100',
    border:     '#b0bec5'
  };
  const BS    = SpreadsheetApp.BorderStyle.SOLID;
  const MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const W     = 6; // ancho total en columnas

  // ── Anchos de columna ─────────────────────────────────────────────────────
  hoja.setColumnWidth(1, 220);
  hoja.setColumnWidth(2, 110);
  hoja.setColumnWidth(3, 80);
  hoja.setColumnWidth(4, 100);
  hoja.setColumnWidth(5, 100);
  hoja.setColumnWidth(6, 100);

  // ── Helpers internos ──────────────────────────────────────────────────────
  function contarFilas(nombreHoja, colNombre) {
    var h = ss.getSheetByName(nombreHoja);
    if (!h || h.getLastRow() <= 1) return 0;
    var vals = h.getRange(2, colNombre + 1, h.getLastRow() - 1, 1).getValues();
    return vals.filter(function(r) { return r[0] && r[0].toString().trim() !== ''; }).length;
  }
  function titulo(f, texto, bg, fg, sz) {
    hoja.getRange(f, 1, 1, W).merge()
        .setValue(texto)
        .setBackground(bg).setFontColor(fg)
        .setFontWeight('bold').setFontSize(sz || 12)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    hoja.setRowHeight(f, 34);
    return f + 1;
  }
  function fila2(f, label, valor, bg, valBg) {
    hoja.getRange(f, 1, 1, 4).merge()
        .setValue(label).setBackground(bg || C.par).setFontSize(11)
        .setVerticalAlignment('middle');
    hoja.getRange(f, 5, 1, 2).merge()
        .setValue(valor).setBackground(valBg || C.par)
        .setHorizontalAlignment('center').setFontWeight('bold').setFontSize(13)
        .setVerticalAlignment('middle');
    hoja.setRowHeight(f, 30);
    return f + 1;
  }
  function borde(r1, c1, nr, nc) {
    hoja.getRange(r1, c1, nr, nc)
        .setBorder(true, true, true, true, true, true, C.border, BS);
  }

  // ── DATOS ─────────────────────────────────────────────────────────────────
  // Seguimiento (etapas)
  const conteos = {};
  var totalParticipantes = 0;
  ETAPAS_FLUJO.forEach(function(e) {
    var col = (e === 'Conexiones Laborales') ? 1 : 2;
    conteos[e] = contarFilas(e, col);
    totalParticipantes += conteos[e];
  });

  // Graduados (col 4 → índice 3)
  var totalGraduados = contarFilas('Graduados', 3);

  // Seguimientos pendientes
  var totalPendientes = obtenerSeguimientosPendientes().length;

  // Registros KoboToolbox
  var totalClasif   = (function() {
    var h = ss.getSheetByName('Clasificación de Perfiles');
    return (!h || h.getLastRow() <= 1) ? 0 : h.getLastRow() - 1;
  })();
  var totalSatisf   = (function() {
    var h = ss.getSheetByName('Satisfacción Empleo');
    return (!h || h.getLastRow() <= 1) ? 0 : h.getLastRow() - 1;
  })();
  var totalSesiones = (function() {
    var h = ss.getSheetByName('Sesiones Acompañamiento');
    return (!h || h.getLastRow() <= 1) ? 0 : h.getLastRow() - 1;
  })();
  var totalConexiones = conteos['Conexiones Laborales'] || 0;

  // Ingresos por mes (hojas de etapas)
  const ipm = {};
  ETAPAS_FLUJO.forEach(function(nombreHoja) {
    var h = ss.getSheetByName(nombreHoja);
    if (!h || h.getLastRow() <= 1) return;
    var datos   = h.getDataRange().getValues();
    var colNom  = (nombreHoja === 'Conexiones Laborales') ? 1 : 2;
    for (var i = 1; i < datos.length; i++) {
      if (!datos[i][colNom] || datos[i][colNom].toString().trim() === '') continue;
      var raw = datos[i][0];
      var fch;
      if (raw instanceof Date) { fch = raw; }
      else if (typeof raw === 'string' && raw.indexOf('/') !== -1) {
        var p = raw.split('/'); fch = new Date(p[2], parseInt(p[1]) - 1, p[0]);
      } else { continue; }
      var clave = MESES[fch.getMonth()] + ' ' + fch.getFullYear();
      if (!ipm[clave]) {
        ipm[clave] = { _t: fch.getTime() };
        ETAPAS_FLUJO.forEach(function(e) { ipm[clave][e] = 0; });
        ipm[clave]['Sesiones'] = 0;
      }
      ipm[clave][nombreHoja]++;
    }
  });
  // Add Sesiones Acompañamiento data (col1=Creamos ID, col2=Fecha envío ISO "2024-01-15T10:30:00")
  (function() {
    var hSes = ss.getSheetByName('Sesiones Acompañamiento');
    if (!hSes || hSes.getLastRow() <= 1) return;
    var datosSes = hSes.getDataRange().getValues();
    for (var i = 1; i < datosSes.length; i++) {
      var raw = datosSes[i][1]; // col 2 = Fecha envío
      var fch;
      if (raw instanceof Date) {
        fch = raw;
      } else if (typeof raw === 'string' && raw.length >= 10) {
        var isoDate = raw.substring(0, 10); // "2024-01-15"
        var parts = isoDate.split('-');
        if (parts.length === 3) fch = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      }
      if (!fch || isNaN(fch.getTime())) continue;
      var clave = MESES[fch.getMonth()] + ' ' + fch.getFullYear();
      if (!ipm[clave]) {
        ipm[clave] = { _t: fch.getTime() };
        ETAPAS_FLUJO.forEach(function(e) { ipm[clave][e] = 0; });
        ipm[clave]['Sesiones'] = 0;
      }
      if (!ipm[clave].hasOwnProperty('Sesiones')) ipm[clave]['Sesiones'] = 0;
      ipm[clave]['Sesiones']++;
    }
  })();
  // Ensure all existing ipm entries have the Sesiones key
  Object.keys(ipm).forEach(function(clave) {
    if (!ipm[clave].hasOwnProperty('Sesiones')) ipm[clave]['Sesiones'] = 0;
  });
  var meses = Object.keys(ipm).sort(function(a, b) { return ipm[a]._t - ipm[b]._t; });

  // ── RENDER ────────────────────────────────────────────────────────────────
  var f = 1;
  var ahora = new Date();

  // ╔══════════════════════════════════════════╗
  // ║  BANNER PRINCIPAL                        ║
  // ╚══════════════════════════════════════════╝
  hoja.setRowHeight(f, 46);
  hoja.getRange(f, 1, 1, W).merge()
      .setValue('REPORTE DE SEGUIMIENTO — EMPLEABILIDAD')
      .setBackground(C.bannerBg).setFontColor(C.bannerFg)
      .setFontWeight('bold').setFontSize(16)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  f++;
  hoja.getRange(f, 1, 1, W).merge()
      .setValue('Actualizado: ' + ahora.toLocaleDateString('es-ES') + ' · ' +
                ahora.toLocaleTimeString('es-ES', {hour:'2-digit', minute:'2-digit'}))
      .setFontColor('#9e9e9e').setFontStyle('italic').setFontSize(10)
      .setHorizontalAlignment('center').setBackground('#f5f5f5');
  hoja.setRowHeight(f, 22);
  f += 2;

  // ╔══════════════════════════════════════════╗
  // ║  BLOQUE 1 — KPIs PRINCIPALES            ║
  // ╚══════════════════════════════════════════╝
  f = titulo(f, '📊  RESUMEN GENERAL', C.kpiBg, C.kpiFg, 13);

  // Fila de 3 KPIs lado a lado
  var kpiRows = [
    [['Graduados registrados', totalGraduados, C.headerBg, C.headerFg],
     ['En seguimiento activo', totalParticipantes, '#1565c0', '#ffffff'],
     ['Conexiones laborales',  totalConexiones,  '#e65100', '#ffffff']],
    [['Evaluaciones de perfil', totalClasif,    C.koboAcc, '#ffffff'],
     ['Encuestas satisfacción', totalSatisf,    C.satisfAcc, '#ffffff'],
     ['Sesiones acompañamiento',totalSesiones,  C.sessAcc, '#ffffff']]
  ];
  kpiRows.forEach(function(row) {
    hoja.setRowHeight(f, 26);
    hoja.setRowHeight(f + 1, 42);
    for (var c = 0; c < 3; c++) {
      var col = c * 2 + 1;
      var kpi = row[c];
      hoja.getRange(f, col, 1, 2).merge()
          .setValue(kpi[0]).setBackground(kpi[2]).setFontColor(kpi[3])
          .setFontSize(10).setFontWeight('bold')
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
      hoja.getRange(f + 1, col, 1, 2).merge()
          .setValue(kpi[1]).setBackground(kpi[2]).setFontColor(kpi[3])
          .setFontSize(22).setFontWeight('bold')
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
    borde(f, 1, 2, W);
    f += 2;
  });

  // Seguimientos pendientes (fila completa)
  hoja.setRowHeight(f, 30);
  hoja.getRange(f, 1, 1, 4).merge()
      .setValue('⏰  Seguimientos pendientes')
      .setBackground('#fff8e1').setFontSize(11).setVerticalAlignment('middle');
  hoja.getRange(f, 5, 1, 2).merge()
      .setValue(totalPendientes)
      .setBackground('#fff8e1').setFontColor('#f57f17')
      .setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  borde(f, 1, 1, W);
  f += 2;

  // ╔══════════════════════════════════════════╗
  // ║  BLOQUE 2 — DISTRIBUCIÓN POR ETAPA      ║
  // ╚══════════════════════════════════════════╝
  f = titulo(f, '🗂  DISTRIBUCIÓN POR ETAPA DE SEGUIMIENTO', C.secBg, C.secFg);

  // Header
  hoja.getRange(f, 1, 1, 4).merge()
      .setValue('Etapa').setBackground(C.headerBg).setFontColor(C.headerFg)
      .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(11);
  hoja.getRange(f, 5, 1, 1)
      .setValue('Personas').setBackground(C.headerBg).setFontColor(C.headerFg)
      .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(11);
  hoja.getRange(f, 6, 1, 1)
      .setValue('%').setBackground(C.headerBg).setFontColor(C.headerFg)
      .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(11);
  hoja.setRowHeight(f, 28);
  f++;

  var filaInicioEtapas = f;
  ETAPAS_FLUJO.forEach(function(etapa, idx) {
    var cant = conteos[etapa] || 0;
    var pct  = totalParticipantes > 0 ? (cant / totalParticipantes * 100).toFixed(1) : '0.0';
    var bg   = idx % 2 === 0 ? C.par : C.impar;
    hoja.setRowHeight(f, 26);
    hoja.getRange(f, 1, 1, 4).merge()
        .setValue(etapa).setBackground(bg).setFontSize(11).setVerticalAlignment('middle');
    hoja.getRange(f, 5).setValue(cant).setBackground(bg)
        .setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12)
        .setVerticalAlignment('middle');
    hoja.getRange(f, 6).setValue(pct + '%').setBackground(bg)
        .setHorizontalAlignment('center').setFontSize(11).setVerticalAlignment('middle');
    f++;
  });
  // Fila total
  hoja.setRowHeight(f, 28);
  hoja.getRange(f, 1, 1, 4).merge()
      .setValue('TOTAL').setBackground(C.total).setFontColor(C.totalFg)
      .setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle');
  hoja.getRange(f, 5).setValue(totalParticipantes)
      .setBackground(C.total).setFontColor(C.totalFg)
      .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(13)
      .setVerticalAlignment('middle');
  hoja.getRange(f, 6).setValue('100%')
      .setBackground(C.total).setFontColor(C.totalFg)
      .setHorizontalAlignment('center').setFontWeight('bold').setVerticalAlignment('middle');
  borde(filaInicioEtapas - 1, 1, ETAPAS_FLUJO.length + 2, W);
  f += 2;

  // ╔══════════════════════════════════════════╗
  // ║  BLOQUE 3 — DATOS KOBO (detalle)        ║
  // ╚══════════════════════════════════════════╝
  f = titulo(f, '📋  FORMULARIOS KOBO — DETALLE', C.secBg, C.secFg);

  var koboItems = [
    ['🧩  Clasificación de Perfiles',  totalClasif,    'Evaluaciones realizadas',   C.koboBg,   C.koboAcc],
    ['😊  Satisfacción Empleo (IL-06)', totalSatisf,   'Encuestas completadas',     C.satisfBg, C.satisfAcc],
    ['🤝  Sesiones Acompañamiento (IL-08)', totalSesiones, 'Sesiones registradas',  C.sessBg,   C.sessAcc]
  ];
  var filaInicioKobo = f;
  koboItems.forEach(function(item) {
    hoja.setRowHeight(f, 28);
    hoja.getRange(f, 1, 1, 3).merge()
        .setValue(item[0]).setBackground(item[3]).setFontSize(11)
        .setFontWeight('bold').setVerticalAlignment('middle');
    hoja.getRange(f, 4, 1, 1)
        .setValue(item[2]).setBackground(item[3]).setFontSize(10)
        .setFontColor('#555555').setVerticalAlignment('middle');
    hoja.getRange(f, 5, 1, 2).merge()
        .setValue(item[1]).setBackground(item[3]).setFontColor(item[4])
        .setFontWeight('bold').setFontSize(16)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    f++;
  });
  borde(filaInicioKobo, 1, koboItems.length, W);
  f += 2;

  // ╔══════════════════════════════════════════╗
  // ║  BLOQUE 4 — INGRESOS POR MES            ║
  // ╚══════════════════════════════════════════╝
  if (meses.length > 0) {
    var etapasCortas = ETAPAS_FLUJO.map(function(e) {
      if (e === 'Activamente busca trabajo') return 'Act.';
      if (e === 'Conexiones Laborales')      return 'Conex.';
      if (e === 'Derivaciones')              return 'Deriv.';
      if (e === 'Paso a paso')               return 'P.Paso';
      return e;
    });
    // Mes + etapas + Sesiones + Total
    var numEtapas  = ETAPAS_FLUJO.length;
    var numColsMes = 1 + numEtapas + 1 + 1; // Mes + etapas + Sesiones + Total

    // Si necesitamos más columnas, ampliar
    for (var cx = W + 1; cx <= numColsMes; cx++) hoja.setColumnWidth(cx, 85);

    f = titulo(f, '📅  INGRESOS POR MES (por etapa)', C.secBg, C.secFg);

    // Header mes
    hoja.setRowHeight(f, 28);
    var headerMes = [['Mes']];
    etapasCortas.forEach(function(e) { headerMes[0].push(e); });
    headerMes[0].push('Sesiones');
    headerMes[0].push('Total');
    hoja.getRange(f, 1, 1, numColsMes).setValues(headerMes)
        .setBackground(C.headerBg).setFontColor(C.headerFg)
        .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
    f++;

    var totEtapa = {};
    ETAPAS_FLUJO.forEach(function(e) { totEtapa[e] = 0; });
    var totSesiones = 0;
    var granTotal = 0;
    var filaInicioMeses = f;

    meses.forEach(function(mes, idx) {
      var bg = idx % 2 === 0 ? C.par : C.impar;
      hoja.setRowHeight(f, 24);
      var row = [mes];
      var totalMes = 0;
      ETAPAS_FLUJO.forEach(function(e) {
        var v = ipm[mes][e] || 0;
        row.push(v);
        totEtapa[e] += v;
        totalMes    += v;
      });
      var vSes = ipm[mes]['Sesiones'] || 0;
      row.push(vSes);
      totSesiones += vSes;
      row.push(totalMes);
      granTotal += totalMes;
      hoja.getRange(f, 1, 1, numColsMes).setValues([row])
          .setBackground(bg).setHorizontalAlignment('center').setFontSize(10);
      hoja.getRange(f, 1).setHorizontalAlignment('left');
      f++;
    });

    // Fila totales
    hoja.setRowHeight(f, 26);
    var rowTot = ['TOTAL'];
    ETAPAS_FLUJO.forEach(function(e) { rowTot.push(totEtapa[e]); });
    rowTot.push(totSesiones);
    rowTot.push(granTotal);
    hoja.getRange(f, 1, 1, numColsMes).setValues([rowTot])
        .setBackground(C.total).setFontColor(C.totalFg)
        .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(11);
    hoja.getRange(f, 1).setHorizontalAlignment('left');
    borde(filaInicioMeses - 1, 1, meses.length + 2, numColsMes);
    f += 2;

    // ╔══════════════════════════════════════════╗
    // ║  BLOQUE 5 — RESUMEN ANUAL               ║
    // ╚══════════════════════════════════════════╝
    // Group months by year and show yearly totals
    var porAnio = {};
    meses.forEach(function(mes) {
      // mes is like "Enero 2024"
      var partes = mes.split(' ');
      var anio = partes[partes.length - 1];
      if (!porAnio[anio]) porAnio[anio] = { _orden: parseInt(anio), etapas: {}, sesiones: 0, total: 0 };
      ETAPAS_FLUJO.forEach(function(e) {
        if (!porAnio[anio].etapas[e]) porAnio[anio].etapas[e] = 0;
        porAnio[anio].etapas[e] += ipm[mes][e] || 0;
        porAnio[anio].total     += ipm[mes][e] || 0;
      });
      porAnio[anio].sesiones += ipm[mes]['Sesiones'] || 0;
    });
    var anios = Object.keys(porAnio).sort(function(a, b) { return porAnio[a]._orden - porAnio[b]._orden; });

    if (anios.length > 0) {
      f = titulo(f, '📆  RESUMEN ANUAL', C.secBg, C.secFg);

      // Header anual
      hoja.setRowHeight(f, 28);
      var headerAnio = [['Año']];
      etapasCortas.forEach(function(e) { headerAnio[0].push(e); });
      headerAnio[0].push('Sesiones');
      headerAnio[0].push('Total');
      hoja.getRange(f, 1, 1, numColsMes).setValues(headerAnio)
          .setBackground(C.headerBg).setFontColor(C.headerFg)
          .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
      f++;

      var filaInicioAnios = f;
      var granTotalAnio = 0;
      var totEtapaAnio = {};
      ETAPAS_FLUJO.forEach(function(e) { totEtapaAnio[e] = 0; });
      var totSesionesAnio = 0;

      anios.forEach(function(anio, idx) {
        var bg = idx % 2 === 0 ? C.par : C.impar;
        hoja.setRowHeight(f, 26);
        var row = [anio];
        ETAPAS_FLUJO.forEach(function(e) {
          var v = porAnio[anio].etapas[e] || 0;
          row.push(v);
          totEtapaAnio[e] += v;
        });
        row.push(porAnio[anio].sesiones);
        totSesionesAnio += porAnio[anio].sesiones;
        row.push(porAnio[anio].total);
        granTotalAnio += porAnio[anio].total;
        hoja.getRange(f, 1, 1, numColsMes).setValues([row])
            .setBackground(bg).setHorizontalAlignment('center').setFontSize(11);
        hoja.getRange(f, 1).setHorizontalAlignment('left').setFontWeight('bold');
        f++;
      });

      // Fila total anual
      hoja.setRowHeight(f, 28);
      var rowTotAnio = ['TOTAL'];
      ETAPAS_FLUJO.forEach(function(e) { rowTotAnio.push(totEtapaAnio[e]); });
      rowTotAnio.push(totSesionesAnio);
      rowTotAnio.push(granTotalAnio);
      hoja.getRange(f, 1, 1, numColsMes).setValues([rowTotAnio])
          .setBackground(C.total).setFontColor(C.totalFg)
          .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);
      hoja.getRange(f, 1).setHorizontalAlignment('left');
      borde(filaInicioAnios - 1, 1, anios.length + 2, numColsMes);
      f += 2;
    }
  }

  // ── Footer ────────────────────────────────────────────────────────────────
  hoja.getRange(f, 1, 1, W).merge()
      .setValue('Generado automáticamente · Sistema Empleabilidad Creamos')
      .setFontColor('#bdbdbd').setFontStyle('italic').setFontSize(9)
      .setHorizontalAlignment('center');

  SpreadsheetApp.flush();
  Logger.log('Reporte generado exitosamente.');
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

// ===========================================================================
// SECCIÓN 8: IMPORTACIÓN DESDE HOJA EXTERNA DE GRADUADOS (pull horario)
// ===========================================================================

// ==========================================================================
// SECCIÓN: BASE DE DATOS SALESFORCE (CREAMOS ID)
// Hoja oculta y protegida que sirve como fuente de verdad para autocompletar
// ==========================================================================

const HOJA_BD_CREAMOS = 'Copy of CREAMOS ID nuevo';

/**
 * Carga el mapa Creamos ID → datos personales desde la hoja Salesforce.
 * Columnas de la hoja: A=Nombre | B=Creamos ID | C=Año ingreso | D=Edad | E=DPI
 * @return {Object} mapa { creamosId: { nombre, anio, edad, dpi } }
 */
function _cargarMapaCreamos_() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_BD_CREAMOS);
  if (!hoja || hoja.getLastRow() < 2) {
    return { porId: {}, porNorm: {}, porDpi: {}, todos: [] };
  }
  var datos   = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
  var porId   = {};
  var porNorm = {};
  var porDpi  = {};
  var todos   = [];
  datos.forEach(function(fila) {
    var cid    = (fila[1] || '').toString().trim(); // col B = Creamos ID
    var nombre = (fila[0] || '').toString().trim(); // col A = Nombre completo
    var dpi    = (fila[4] || '').toString().trim(); // col E = DPI
    if (!cid && !nombre) return;
    var rec = {
      cid:        cid,
      nombre:     nombre,
      normCid:    _normalizarTexto_(cid),
      normNombre: _normalizarTexto_(nombre),
      anio:       fila[2] || '',                    // col C = Año que entró
      edad:       fila[3] || '',                    // col D = Age
      dpi:        dpi
    };
    if (cid) { porId[cid] = rec; porNorm[rec.normCid] = rec; }
    if (dpi) porDpi[dpi] = rec;
    todos.push(rec);
  });
  return { porId: porId, porNorm: porNorm, porDpi: porDpi, todos: todos };
}

/** Elimina acentos, pasa a minúsculas y colapsa espacios para comparar textos. */
function _normalizarTexto_(s) {
  if (!s) return '';
  return s.toString().trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ');
}

/** Distancia Levenshtein entre dos strings. */
function _levenshtein_(a, b) {
  if (a === b) return 0;
  if (!a) return b.length;
  if (!b) return a.length;
  var al = a.length, bl = b.length;
  var dp = [];
  for (var i = 0; i <= al; i++) { dp[i] = [i]; }
  for (var j = 1; j <= bl; j++) { dp[0][j] = j; }
  for (var i = 1; i <= al; i++) {
    for (var j = 1; j <= bl; j++) {
      dp[i][j] = a[i-1] === b[j-1]
        ? dp[i-1][j-1]
        : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  return dp[al][bl];
}

/**
 * Busca un Creamos ID (o nombre) en el mapa con fallback fuzzy.
 * @return {null | { rec, metodo: 'exacto'|'normalizado'|'dist1'|'dist2', idSugerido }}
 */
function _buscarEnMapaFuzzy_(valorBruto, mapaExt) {
  var v = (valorBruto || '').toString().trim();
  if (!v) return null;

  // 1. Coincidencia exacta
  if (mapaExt.porId[v])
    return { rec: mapaExt.porId[v], metodo: 'exacto', idSugerido: null };

  // 2. Coincidencia normalizada (maneja acentos, mayúsculas, espacios)
  var vn = _normalizarTexto_(v);
  if (mapaExt.porNorm[vn])
    return { rec: mapaExt.porNorm[vn], metodo: 'normalizado', idSugerido: mapaExt.porNorm[vn].cid };

  // 3. Fuzzy por Levenshtein sobre el Creamos ID normalizado
  var mejorDist = 999, mejorRec = null;
  for (var i = 0; i < mapaExt.todos.length; i++) {
    var t = mapaExt.todos[i];
    if (!t.normCid) continue;
    var d = _levenshtein_(vn, t.normCid);
    if (d < mejorDist) { mejorDist = d; mejorRec = t; }
  }
  if (mejorDist <= 1) return { rec: mejorRec, metodo: 'dist1', idSugerido: mejorRec.cid };
  if (mejorDist <= 2) return { rec: mejorRec, metodo: 'dist2', idSugerido: mejorRec.cid };

  return null;
}

/**
 * Autocompleta campos faltantes en las hojas del sistema
 * usando la base de datos de Salesforce como fuente de verdad.
 * Solo rellena celdas vacías — nunca sobreescribe datos existentes.
 * Usa fuzzy matching para detectar y corregir errores ortográficos.
 */
function autocompletarConCreamos() {
  var ui      = SpreadsheetApp.getUi();
  var mapaExt = _cargarMapaCreamos_();
  if (mapaExt.todos.length === 0) {
    ui.alert('❌ Error', 'No se encontró la hoja "' + HOJA_BD_CREAMOS + '" o está vacía.', ui.ButtonSet.OK);
    return;
  }

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var cambios  = 0;
  var corregidos = 0;

  function _llenar_(hoja, row, idx, val, campo) {
    if (hoja.getRange(row, idx + 1).getValue()) return;
    hoja.getRange(row, idx + 1).setValue(val);
    cambios++;
  }

  // ── Graduados: Creamos ID = col C (idx 2), Nombre = col D (idx 3), Edad = col F (idx 5) ──
  var hojaGrad = ss.getSheetByName('Graduados');
  if (hojaGrad && hojaGrad.getLastRow() > 1) {
    var filas = hojaGrad.getRange(2, 1, hojaGrad.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < filas.length; i++) {
      var cid   = (filas[i][2] || '').toString().trim();
      if (!cid) continue;
      var match = _buscarEnMapaFuzzy_(cid, mapaExt);
      if (!match) continue;
      var rec = match.rec;
      var row = i + 2;
      if (match.metodo === 'dist1' || match.metodo === 'normalizado') {
        // Corregir el Creamos ID en la hoja y marcar amarillo con nota
        hojaGrad.getRange(row, 3)
          .setValue(rec.cid)
          .setBackground('#FFF9C4')
          .setNote('ID corregido: "' + cid + '" → "' + rec.cid + '"');
        corregidos++;
      }
      if (!filas[i][3] && rec.nombre) _llenar_(hojaGrad, row, 3, rec.nombre);
      if (!filas[i][5] && rec.edad)   _llenar_(hojaGrad, row, 5, rec.edad);
    }
  }

  // ── Hojas de clasificación: Creamos ID = col B (idx 1), Nombre = col C (idx 2), Edad = col F (idx 5) ──
  ['Aliados', 'Plataforma', 'Derivaciones', 'Paso a paso', 'Activamente busca trabajo'].forEach(function(nombre) {
    var h = ss.getSheetByName(nombre);
    if (!h || h.getLastRow() < 2) return;
    var filas = h.getRange(2, 1, h.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < filas.length; i++) {
      var cid   = (filas[i][1] || '').toString().trim();
      if (!cid) continue;
      var match = _buscarEnMapaFuzzy_(cid, mapaExt);
      if (!match) continue;
      var rec = match.rec;
      var row = i + 2;
      if (match.metodo === 'dist1' || match.metodo === 'normalizado') {
        h.getRange(row, 2)
          .setValue(rec.cid)
          .setBackground('#FFF9C4')
          .setNote('ID corregido: "' + cid + '" → "' + rec.cid + '"');
        corregidos++;
      }
      if (!filas[i][2] && rec.nombre) _llenar_(h, row, 2, rec.nombre);
      if (!filas[i][5] && rec.edad)   _llenar_(h, row, 5, rec.edad);
    }
  });

  // ── Clasificación de Perfiles: Creamos ID = col B (idx 1), Nombre = col C (idx 2) ──
  var hojaClasif = ss.getSheetByName('Clasificación de Perfiles');
  if (hojaClasif && hojaClasif.getLastRow() > 1) {
    var filas = hojaClasif.getRange(2, 1, hojaClasif.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < filas.length; i++) {
      var cid   = (filas[i][1] || '').toString().trim();
      if (!cid) continue;
      var match = _buscarEnMapaFuzzy_(cid, mapaExt);
      if (!match) continue;
      var rec = match.rec;
      if (!filas[i][2] && rec.nombre) _llenar_(hojaClasif, i + 2, 2, rec.nombre);
    }
  }

  var msg = 'Se completaron ' + cambios + ' campos faltantes.';
  if (corregidos > 0) msg += '\n⚠️ ' + corregidos + ' Creamos ID corregidos (celdas en amarillo).';
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '✅ Autocompletado listo', 7);
}

/**
 * Oculta y protege la hoja base de datos Salesforce.
 * Solo el usuario que ejecuta esta función puede editarla.
 */
function protegerBaseDatosSalesforce() {
  var ui   = SpreadsheetApp.getUi();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_BD_CREAMOS);
  if (!hoja) {
    ui.alert('❌ Error', 'No se encontró la hoja "' + HOJA_BD_CREAMOS + '".\nVerifica que el nombre sea exactamente ese.', ui.ButtonSet.OK);
    return;
  }

  // Ocultar la hoja
  hoja.hideSheet();

  // Proteger: eliminar todos los editores excepto el usuario actual
  var proteccion = hoja.protect().setDescription('Base de datos Salesforce — solo administrador');
  proteccion.setDomainEdit(false);
  var yo = Session.getEffectiveUser();
  var editores = proteccion.getEditors();
  if (editores.length > 0) proteccion.removeEditors(editores);
  proteccion.addEditor(yo);

  ui.alert(
    '🔒 Hoja protegida',
    '"' + HOJA_BD_CREAMOS + '" está ahora:\n\n' +
    '✅ Oculta (no visible en las pestañas)\n' +
    '✅ Protegida (solo tú puedes editarla)\n\n' +
    'Para verla: clic derecho en cualquier pestaña → "Mostrar hojas".',
    ui.ButtonSet.OK
  );
}

/**
 * Recorre todas las hojas buscando filas sin Creamos ID.
 * - Sin Creamos ID → marca celda en naranja + nota para crear perfil en Salesforce
 * - Con Creamos ID → borra el color de alerta y autocompleta datos faltantes
 * Se ejecuta desde el menú o automáticamente por trigger cada 4 horas.
 */
function verificarYCompletarCreamos() {
  var mapaExt = _cargarMapaCreamos_();
  // Colores de estado en la celda del Creamos ID:
  var C_SIN_ID  = '#FFE0B2'; // naranja    = celda vacía, falta Creamos ID
  var C_NO_BD   = '#FFCDD2'; // rojo claro = tiene ID pero no existe en Salesforce
  var C_TYPO    = '#FFF9C4'; // amarillo   = typo detectado, ID corregido
  var C_SUGER   = '#FFF3E0'; // ámbar claro = posible typo (distancia 2), sin corregir
  var C_OK      = null;      // sin color  = todo correcto

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sinId       = 0, noBD = 0, corregidos = 0, sugerencias = 0, completados = 0;

  function _procesarFila_(hoja, fila, row, colCid, colNombre, colEdad) {
    var cid   = (fila[colCid] || '').toString().trim();
    var celId = hoja.getRange(row, colCid + 1);
    if (!cid) {
      celId.setBackground(C_SIN_ID).setNote('⚠️ Pendiente: crear perfil en Salesforce con este registro');
      sinId++;
      return;
    }
    var match = _buscarEnMapaFuzzy_(cid, mapaExt);
    if (!match) {
      celId.setBackground(C_NO_BD).setNote('❌ Creamos ID "' + cid + '" no existe en Salesforce. Verificar.');
      noBD++;
      return;
    }
    var rec = match.rec;
    if (match.metodo === 'dist1' || match.metodo === 'normalizado') {
      celId.setValue(rec.cid).setBackground(C_TYPO)
        .setNote('✏️ ID corregido automáticamente: "' + cid + '" → "' + rec.cid + '"');
      corregidos++;
    } else if (match.metodo === 'dist2') {
      celId.setBackground(C_SUGER)
        .setNote('❓ Posible typo: ¿quisiste decir "' + rec.cid + '"? Revisar manualmente.');
      sugerencias++;
    } else {
      celId.setBackground(C_OK).clearNote();
    }
    if (colNombre !== null && !fila[colNombre] && rec.nombre) {
      hoja.getRange(row, colNombre + 1).setValue(rec.nombre); completados++;
    }
    if (colEdad !== null && !fila[colEdad] && rec.edad) {
      hoja.getRange(row, colEdad + 1).setValue(rec.edad); completados++;
    }
  }

  // ── Graduados: Creamos ID = col C (idx 2), Nombre = col D (idx 3), Edad = col F (idx 5) ──
  var hojaGrad = ss.getSheetByName('Graduados');
  if (hojaGrad && hojaGrad.getLastRow() > 1) {
    var data = hojaGrad.getRange(2, 1, hojaGrad.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < data.length; i++) _procesarFila_(hojaGrad, data[i], i + 2, 2, 3, 5);
  }

  // ── Hojas de clasificación: Creamos ID = col B (idx 1), Nombre = col C (idx 2), Edad = col F (idx 5) ──
  ['Aliados', 'Plataforma', 'Derivaciones', 'Paso a paso', 'Activamente busca trabajo'].forEach(function(nombre) {
    var h = ss.getSheetByName(nombre);
    if (!h || h.getLastRow() < 2) return;
    var data = h.getRange(2, 1, h.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) _procesarFila_(h, data[i], i + 2, 1, 2, 5);
  });

  var partes = [];
  if (completados > 0)  partes.push(completados  + ' campos completados');
  if (corregidos > 0)   partes.push(corregidos   + ' IDs corregidos (amarillo)');
  if (sugerencias > 0)  partes.push(sugerencias  + ' posibles typos (ámbar) — revisar');
  if (sinId > 0)        partes.push(sinId        + ' sin Creamos ID (naranja)');
  if (noBD > 0)         partes.push(noBD         + ' IDs no encontrados en Salesforce (rojo)');
  var msg = partes.length ? partes.join('\n') : 'Todo está correcto.';
  Logger.log('Verificación Creamos: ' + msg);
  ss.toast(msg, '🔍 Verificación completada', 10);
}

/**
 * Genera un reporte de errores ortográficos / IDs no encontrados en todas las hojas.
 * No modifica datos, solo lee y muestra un resumen en una alerta.
 */
function diagnosticarErroresCreamos() {
  var ui      = SpreadsheetApp.getUi();
  var mapaExt = _cargarMapaCreamos_();
  if (mapaExt.todos.length === 0) {
    ui.alert('❌ Error', 'No se encontró "' + HOJA_BD_CREAMOS + '" o está vacía.', ui.ButtonSet.OK);
    return;
  }
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var problemas = [];

  function _diagnosticar_(hoja, colCid) {
    if (!hoja || hoja.getLastRow() < 2) return;
    var nomH  = hoja.getName();
    var data  = hoja.getRange(2, colCid + 1, hoja.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      var cid = (data[i][0] || '').toString().trim();
      if (!cid) continue;
      var match = _buscarEnMapaFuzzy_(cid, mapaExt);
      if (!match) {
        problemas.push(nomH + ' fila ' + (i + 2) + ': "' + cid + '" — NO encontrado en Salesforce');
      } else if (match.metodo === 'dist1' || match.metodo === 'normalizado') {
        problemas.push(nomH + ' fila ' + (i + 2) + ': "' + cid + '" → corregir a "' + match.idSugerido + '"');
      } else if (match.metodo === 'dist2') {
        problemas.push(nomH + ' fila ' + (i + 2) + ': "' + cid + '" — ¿quisiste decir "' + match.idSugerido + '"?');
      }
    }
  }

  _diagnosticar_(ss.getSheetByName('Graduados'), 2);
  ['Aliados', 'Plataforma', 'Derivaciones', 'Paso a paso', 'Activamente busca trabajo'].forEach(function(n) {
    _diagnosticar_(ss.getSheetByName(n), 1);
  });

  if (problemas.length === 0) {
    ui.alert('✅ Sin errores', 'Todos los Creamos ID encontrados existen en Salesforce.', ui.ButtonSet.OK);
  } else {
    var texto = problemas.slice(0, 30).join('\n');
    if (problemas.length > 30) texto += '\n... y ' + (problemas.length - 30) + ' más.';
    ui.alert(
      '⚠️ ' + problemas.length + ' problema(s) detectado(s)',
      texto + '\n\nUsa "Verificar Creamos ID ahora" para corregir automáticamente.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Instala un trigger que ejecuta verificarYCompletarCreamos() cada 4 horas.
 * Elimina triggers anteriores del mismo nombre para evitar duplicados.
 */
function instalarTriggerVerificacionCreamos() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'verificarYCompletarCreamos') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('verificarYCompletarCreamos')
    .timeBased()
    .everyHours(4)
    .create();

  SpreadsheetApp.getUi().alert(
    '⏰ Verificación automática activada',
    'Cada 4 horas el sistema:\n\n' +
    '🟠 Marcará en naranja registros sin Creamos ID\n' +
    '✅ Completará datos faltantes desde Salesforce\n\n' +
    'También puedes ejecutarlo manualmente desde el menú.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

const CONFIG_GRADUADOS_EXTERNO = {
  FILE_ID:    '1_596FX6yr8tX93UyIks4emSeE2_vxLJMDyw9Zncsnzs',
  SHEET_NAME: 'Graduados',
  SHEET_GID:  676353499,
  DESTINO:    'Graduados Importados',
  NUM_COLS:   16
};

/**
 * Jala (pull) datos desde un archivo externo de Google Sheets hacia la
 * hoja local "Graduados Importados". Solo inserta filas nuevas,
 * deduplicando por "Creamos ID" (col C) o, si está vacío, por Nombre (col D).
 *
 * Diseñada para ser segura ante errores: no lanza excepciones al usuario,
 * solo muestra toasts.
 */
function importarGraduadosDesdeExterno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Abrir archivo externo
  let archivoExterno;
  try {
    archivoExterno = SpreadsheetApp.openById(CONFIG_GRADUADOS_EXTERNO.FILE_ID);
  } catch (e) {
    Logger.log('importarGraduadosDesdeExterno: error abriendo archivo externo — ' + e);
    ss.toast('No se pudo abrir el archivo externo: ' + e.message, '❌ Importar Graduados', 8);
    return;
  }

  // 2. Localizar la hoja (por nombre; si no, por GID de respaldo)
  let hojaExterna = archivoExterno.getSheetByName(CONFIG_GRADUADOS_EXTERNO.SHEET_NAME);
  if (!hojaExterna) {
    const hojas = archivoExterno.getSheets();
    for (let i = 0; i < hojas.length; i++) {
      if (hojas[i].getSheetId() === CONFIG_GRADUADOS_EXTERNO.SHEET_GID) {
        hojaExterna = hojas[i];
        break;
      }
    }
  }
  if (!hojaExterna) {
    ss.toast('No se encontró la hoja "' + CONFIG_GRADUADOS_EXTERNO.SHEET_NAME +
             '" ni el GID ' + CONFIG_GRADUADOS_EXTERNO.SHEET_GID,
             '❌ Importar Graduados', 8);
    return;
  }

  try {
    const ultFilaExt = hojaExterna.getLastRow();
    if (ultFilaExt < 2) {
      ss.toast('Sin registros nuevos', 'ℹ️ Importar Graduados', 4);
      return;
    }
    const datosExt = hojaExterna.getRange(2, 1, ultFilaExt - 1, CONFIG_GRADUADOS_EXTERNO.NUM_COLS).getValues();

    // 3. Hoja local destino (se crea con encabezados si no existe)
    const hojaDestino = _obtenerHojaGraduadosImportados();

    // 4. Construir índices de lo que ya existe localmente
    const existentesId = {};
    const existentesNombre = {};
    const ultFilaLocal = hojaDestino.getLastRow();
    if (ultFilaLocal >= 2) {
      const datosLocal = hojaDestino.getRange(2, 1, ultFilaLocal - 1, CONFIG_GRADUADOS_EXTERNO.NUM_COLS).getValues();
      for (let i = 0; i < datosLocal.length; i++) {
        const creamosId = String(datosLocal[i][2] || '').trim();
        const nombre    = String(datosLocal[i][3] || '').trim().toLowerCase();
        if (creamosId) existentesId[creamosId] = true;
        if (nombre)    existentesNombre[nombre] = true;
      }
    }

    // 5. Filtrar filas nuevas (dedup por Creamos ID; fallback por Nombre)
    const nuevas = [];
    for (let i = 0; i < datosExt.length; i++) {
      const fila = datosExt[i];
      const creamosId = String(fila[2] || '').trim();
      const nombre    = String(fila[3] || '').trim().toLowerCase();

      if (!creamosId && !nombre) continue; // fila vacía

      if (creamosId) {
        if (existentesId[creamosId]) continue;
        existentesId[creamosId] = true;
      } else {
        if (existentesNombre[nombre]) continue;
        existentesNombre[nombre] = true;
      }
      nuevas.push(fila);
    }

    if (nuevas.length === 0) {
      ss.toast('Sin registros nuevos', 'ℹ️ Importar Graduados', 4);
      return;
    }

    const filaInicio = Math.max(hojaDestino.getLastRow() + 1, 2);
    hojaDestino.getRange(filaInicio, 1, nuevas.length, CONFIG_GRADUADOS_EXTERNO.NUM_COLS).setValues(nuevas);

    ss.toast('Se importaron ' + nuevas.length + ' registro(s) nuevo(s)',
             '✅ Importar Graduados', 5);
  } catch (error) {
    Logger.log('Error en importarGraduadosDesdeExterno: ' + error);
    ss.toast('Error al importar: ' + error.message, '❌ Importar Graduados', 8);
  }
}

function _obtenerHojaGraduadosImportados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(CONFIG_GRADUADOS_EXTERNO.DESTINO);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG_GRADUADOS_EXTERNO.DESTINO);
    const encabezados = [
      'No.', 'Fecha de envío', 'Creamos ID', 'Nombre completo', 'Teléfono',
      'Formación / Nivel Educativo', 'Cohorte', 'Fecha de entrevista', 'Entrevistador',
      'Resultado entrevista', 'Siguiente paso', 'Clasificación', 'Empleado',
      'Próxima llamada', 'Notas', 'Etapa'
    ];
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

/**
 * Instala un trigger por tiempo que ejecuta importarGraduadosDesdeExterno
 * cada 1 hora. Elimina triggers previos de la misma función para evitar
 * duplicados.
 */
function instalarTriggerGraduados() {
  const ui = SpreadsheetApp.getUi();
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let eliminados = 0;
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'importarGraduadosDesdeExterno') {
        ScriptApp.deleteTrigger(triggers[i]);
        eliminados++;
      }
    }
    ScriptApp.newTrigger('importarGraduadosDesdeExterno')
      .timeBased()
      .everyHours(1)
      .create();

    ui.alert('✅ Trigger instalado',
             'Se instaló el trigger horario para importarGraduadosDesdeExterno.\n\n' +
             'Frecuencia: cada 1 hora\n' +
             'Triggers previos eliminados: ' + eliminados,
             ui.ButtonSet.OK);
  } catch (error) {
    Logger.log('Error en instalarTriggerGraduados: ' + error);
    ui.alert('❌ Error',
             'No se pudo instalar el trigger: ' + error.message,
             ui.ButtonSet.OK);
  }
}
