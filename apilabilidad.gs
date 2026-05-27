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

    // ── Submenú: Importar datos ───────────────────────────────────────────
    const submenuImport = ui.createMenu('📂 Importar datos')
      .addItem('📥 Importar todos (Kobo + externo)',         'importarTodosLosDatos')
      .addItem('📥 Graduados (externo)',                     'importarGraduadosDesdeExterno')
      .addItem('🔄 Solo Graduados (Kobo)',                   'importarDatosKobo')
      .addSeparator()
      .addItem('📋 Clasificación de Perfiles',               'importarClasificacionPerfiles')
      .addItem('😊 Satisfacción Empleo (IL-06)',             'importarSatisfaccionEmpleo')
      .addItem('🤝 Sesiones Acompañamiento (IL-08)',         'importarSesionesAcompanamiento')
      .addSeparator()
      .addItem('⏱ Activar Auto-import',                     'activarAutoImport')
      .addItem('⏹ Desactivar Auto-import',                  'desactivarAutoImport')
      .addSeparator()
      .addItem('🔍 Diagnosticar import Graduados',           'diagnosticarImportGraduados')
      .addItem('🗑️ Eliminar hoja "Graduados Importados"',   'eliminarHojaGraduadosImportados');

    // ── Submenú: Reportes ─────────────────────────────────────────────────
    const submenuReportes = ui.createMenu('📊 Reportes')
      .addItem('📅 Generar Reporte 2025',                    'generarReporte2025')
      .addItem('📊 Generar Reporte (todo)',                  'generarReporte')
      .addSeparator()
      .addItem('📝 Clasificar Graduados',                    'mostrarFormularioClasificacion');

    // ── Submenú: Power BI ─────────────────────────────────────────────────
    const submenuPowerBI = ui.createMenu('📡 Power BI')
      .addItem('▶ Actualizar ahora',                         'generarExportPowerBI')
      .addItem('⏱ Activar auto-actualización (6 h)',         'configurarTriggerPowerBI');

    // ── Submenú: Configuración ────────────────────────────────────────────
    const submenuConfig = ui.createMenu('⚙️ Configuración')
      .addItem('🔄 Actualizar dropdown Etapa',               'actualizarDropdownEtapa')
      .addItem('🎨 Aplicar colores de fila (Graduados)',     'aplicarColoresFilasGraduados')
      .addItem('🔄 Autocompletar con Creamos ID',            'autocompletarConCreamos')
      .addItem('🔍 Verificar Creamos ID ahora',              'verificarYCompletarCreamos')
      .addSeparator()
      .addItem('⚙️ Instalar triggers',                       'configurarEditTrigger')
      .addItem('💼 Enviar a Conexiones Laborales',           'enviarAConexionesLaborales')
      .addSeparator()
      .addItem('📧 Configurar correos Conexiones Laborales', 'configurarEmailsConexionesLaborales');

    const submenuVerificacion = ui.createMenu('🔍 Verificación de IDs')
      .addItem('Auditar IDs',              'auditarIDs')
      .addItem('Limpiar IDs incorrectos',  'limpiarIDsIncorrectos')
      .addItem('Reporte sin ID',           'reporteSinID');

    ui.createMenu('📊 Equipito Empleabilidad')
      .addItem('🚀 IMPLEMENTAR CAMBIOS NUEVOS',  'implementarCambiosNuevos')
      .addSeparator()
      .addSubMenu(submenuImport)
      .addSubMenu(submenuReportes)
      .addSubMenu(submenuPowerBI)
      .addSeparator()
      .addSubMenu(submenuVerificacion)
      .addItem('⚡ Prueba de Rendimiento',  'pruebaRendimiento')
      .addSeparator()
      .addSubMenu(submenuConfig)
      .addToUi();

  } catch (error) {
    Logger.log('onOpen: no se pudo crear el menú — ' + error.message);
  }
}

/**
 * Aplica todos los cambios nuevos de una sola vez:
 *  1. Colores de fila por Etapa en Graduados
 *  2. Nombre completo en Satisfacción Empleo (desde Graduados)
 *  3. Regenera el Reporte con leyenda de colores
 * Muestra un diálogo de progreso paso a paso.
 */
function implementarCambiosNuevos() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pasos = [];

  // ── Paso 1: Colores de fila en Graduados ──────────────────────────────────
  try {
    ss.toast('Paso 1/3 — Aplicando colores de fila en Graduados...', '🚀 Implementando', -1);
    var hGrad = ss.getSheetByName('Graduados');
    if (hGrad) {
      var numCols = ESTRUCTURA_HOJAS['Graduados'].columnas.length;
      _colorearFilasPorEtapa(hGrad, numCols, 15);
      pasos.push('✅ Colores de fila aplicados en Graduados.');
    } else {
      pasos.push('⚠️ Hoja "Graduados" no encontrada — colores omitidos.');
    }
  } catch (e) {
    pasos.push('❌ Error en colores de fila: ' + e.message);
  }

  // ── Paso 2: Nombre completo en Satisfacción Empleo ────────────────────────
  try {
    ss.toast('Paso 2/3 — Completando nombres en Satisfacción Empleo...', '🚀 Implementando', -1);
    var hSat = ss.getSheetByName('Satisfacción Empleo');
    if (hSat && hSat.getLastRow() > 1) {
      // Verificar si ya tiene la columna "Nombre completo" (col 2)
      var encabezado = hSat.getRange(1, 2).getValue().toString().trim();
      if (encabezado !== 'Nombre completo') {
        // Insertar columna en posición 2
        hSat.insertColumnAfter(1);
        hSat.getRange(1, 2).setValue('Nombre completo')
            .setBackground('#00897b').setFontColor('#ffffff')
            .setFontWeight('bold').setFontSize(11);
        hSat.setColumnWidth(2, 200);
      }
      // Llenar nombres vacíos desde Graduados (búsqueda exacta + normalizada)
      var mapaNombres = {}, mapaNombresNorm = {};
      if (hGrad && hGrad.getLastRow() > 1) {
        hGrad.getRange(2, 3, hGrad.getLastRow() - 1, 2).getValues().forEach(function(r) {
          var cid = (r[0] || '').toString().trim();
          var nom = (r[1] || '').toString().trim();
          if (!cid || !nom) return;
          mapaNombres[cid] = nom;
          mapaNombresNorm[_normalizarTexto_(cid)] = nom;
        });
      }
      var llenados = 0;
      var datSat = hSat.getRange(2, 1, hSat.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < datSat.length; i++) {
        var cid = (datSat[i][0] || '').toString().trim();
        var nom = (datSat[i][1] || '').toString().trim();
        if (cid && !nom) {
          var encontrado = mapaNombres[cid] || mapaNombresNorm[_normalizarTexto_(cid)] || '';
          if (encontrado) {
            hSat.getRange(i + 2, 2).setValue(encontrado);
            llenados++;
          }
        }
      }
      pasos.push('✅ Nombres completados en Satisfacción Empleo: ' + llenados + ' filas actualizadas.');
    } else {
      pasos.push('⚠️ Hoja "Satisfacción Empleo" vacía o no encontrada.');
    }
  } catch (e) {
    pasos.push('❌ Error en nombres Satisfacción Empleo: ' + e.message);
  }

  // ── Paso 3: Regenerar Reporte con leyenda ────────────────────────────────
  try {
    ss.toast('Paso 3/3 — Regenerando Reporte con leyenda de colores...', '🚀 Implementando', -1);
    generarReporte();
    pasos.push('✅ Reporte regenerado con leyenda de colores.');
  } catch (e) {
    pasos.push('❌ Error al generar reporte: ' + e.message);
  }

  ss.toast('', '', 1);

  // ── Resultado final ───────────────────────────────────────────────────────
  ui.alert(
    '🚀 Implementación completada',
    pasos.join('\n') +
    '\n\n📧 Recuerda configurar correos de Conexiones Laborales:\n' +
    '   ⚙️ Configuración → 📧 Configurar correos Conexiones Laborales',
    ui.ButtonSet.OK
  );
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
      if (!nuevaEtapa || nuevaEtapa === 'Conexiones Laborales' || nuevaEtapa === 'Paso a paso - Cierre') return;

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

    // ── Graduados: Etapa = "Paso a paso - Cierre" ────────────────────────
    if (nombreHoja === 'Graduados' && col === 15 && e.value === 'Paso a paso - Cierre') {
      const datosGrad = hoja.getRange(fila, 1, 1, 15).getValues()[0];
      const nombre    = datosGrad[3] || '';

      if (!nombre) {
        SpreadsheetApp.getActiveSpreadsheet().toast('La fila no tiene nombre.', '⚠️ Sin datos', 3);
        return;
      }

      const ui = SpreadsheetApp.getUi();
      const respuesta = ui.prompt(
        '📋 Cierre formal — ' + nombre,
        'Escribe el comentario de cierre formal (motivo, observaciones, etc.):',
        ui.ButtonSet.OK_CANCEL
      );

      if (respuesta.getSelectedButton() !== ui.Button.OK) {
        e.range.setValue('');
        return;
      }

      const comentarioCierre = respuesta.getResponseText().trim() || 'Cierre formal';
      copiarAHojaClasificacion(datosGrad, 'Paso a paso - Cierre', { nota: comentarioCierre });
      generarReporte();
      SpreadsheetApp.getActiveSpreadsheet().toast(
        nombre + ' derivado a Paso a paso (cierre formal)', '✅ Clasificado', 4
      );
      return;
    }

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
 * Actualiza el dropdown de la columna Etapa en la hoja Graduados
 * con las opciones más recientes de ETAPAS_FLUJO.
 * NO borra ni modifica ningún dato existente.
 * Ejecútala una sola vez después de pegar el código actualizado.
 */
function actualizarDropdownEtapa() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Graduados');

  if (!hoja) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja "Graduados".');
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    ss.toast('La hoja Graduados no tiene filas de datos.', '⚠️ Sin datos', 3);
    return;
  }

  const rango = hoja.getRange(2, 15, ultimaFila - 1, 1);
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(ETAPAS_FLUJO, true)
    .setAllowInvalid(false)
    .build();
  rango.setDataValidation(regla);

  ss.toast(
    'Dropdown de Etapa actualizado con ' + ETAPAS_FLUJO.length + ' opciones. Ningún dato fue modificado.',
    '✅ Listo',
    5
  );
  Logger.log('actualizarDropdownEtapa completado. Opciones: ' + ETAPAS_FLUJO.join(', '));
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

  // 1. Graduados externos (AYB + TECH)
  try {
    importarGraduadosDesdeExterno();
    res.push('✅ Graduados externos (AYB + TECH): importados');
  } catch (e) { res.push('⚠️ Graduados externos: ' + e.message); }

  // 2. Graduados Kobo
  try {
    var d = obtenerDatosKoboToolbox();
    var g = procesarDatosGraduados(d);
    res.push('✅ Graduados Kobo: ' + g.nuevos + ' nuevos (total ' + g.total + ')');
  } catch (e) { res.push('⚠️ Graduados Kobo: ' + e.message); }

  // 3. Clasificación de Perfiles
  try {
    var c = _syncClasificacionSoloNuevos();
    res.push('✅ Clasificación de Perfiles: ' + c.nuevos + ' nuevos (total ' + c.total + ')');
  } catch (e) { res.push('⚠️ Clasificación: ' + e.message); }

  // 4. Satisfacción Empleo (IL-06)
  try {
    var s = _syncSatisfaccionSoloNuevos();
    res.push('✅ Satisfacción Empleo: ' + s.nuevos + ' nuevos (total ' + s.total + ')');
  } catch (e) { res.push('⚠️ Satisfacción: ' + e.message); }

  // 5. Sesiones Acompañamiento (IL-08)
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
  'Paso a paso - Cierre',
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
  'Aliados':                     { bg: '#1e88e5', fg: '#ffffff' },
  'Plataforma':                  { bg: '#9e9e9e', fg: '#ffffff' },
  'Derivaciones':                { bg: '#f48fb1', fg: '#880e4f' },
  'Activamente busca trabajo':   { bg: '#ce93d8', fg: '#4a148c' },
  'Paso a paso':                 { bg: '#fb8c00', fg: '#ffffff' },
  'Paso a paso - Cierre':        { bg: '#e53935', fg: '#ffffff' },
  'Conexiones Laborales':        { bg: '#43a047', fg: '#ffffff' },
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
      { nombre: 'Nombre completo',      ancho: 200, tipo: 'texto' },
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

  // -- Coloreado de fila completa por Etapa (solo en Graduados) --------------
  if (nombreHoja === 'Graduados') {
    _colorearFilasPorEtapa(hoja, columnas.length, 15);
  }

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
 * Colorea FILAS COMPLETAS en la hoja indicada según el valor de la columna
 * "Etapa" (columna etapaCol, 1-based). Usa formato condicional basado en
 * fórmula para que toda la fila refleje el color del estado.
 * Las reglas se insertan con PRIORIDAD MÁS ALTA (al inicio del array) para
 * que dominen sobre las reglas de dropdown de columnas individuales.
 *
 * @param {Sheet}  hoja
 * @param {number} numCols   - número de columnas a colorear (ancho de la fila)
 * @param {number} etapaCol  - columna (1-based) que contiene el valor de Etapa
 */
function _colorearFilasPorEtapa(hoja, numCols, etapaCol) {
  var maxFilas = Math.max(hoja.getMaxRows() - 1, 1);

  // Convertir número de columna a letra(s) A1 notation
  function colLetra(n) {
    var s = '';
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  var letra = colLetra(etapaCol);

  // Eliminar reglas anteriores de coloreado de fila (las nuestras cubren col 1 con numCols de ancho)
  var reglasPrevias = hoja.getConditionalFormatRules();
  var reglasConservadas = reglasPrevias.filter(function(regla) {
    var rangos = regla.getRanges();
    for (var r = 0; r < rangos.length; r++) {
      if (rangos[r].getColumn() === 1 && rangos[r].getNumColumns() >= numCols) return false;
    }
    return true;
  });

  // Crear nuevas reglas de fila para cada etapa
  var nuevasReglas = [];
  ETAPAS_FLUJO.forEach(function(etapa) {
    var c = COLORES_DROPDOWN[etapa];
    if (!c) return;
    var rango = hoja.getRange(2, 1, maxFilas, numCols);
    var regla = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + letra + '2="' + etapa + '"')
      .setBackground(c.bg)
      .setFontColor(c.fg)
      .setRanges([rango])
      .build();
    nuevasReglas.push(regla);
  });

  // Insertar con mayor prioridad (al inicio)
  hoja.setConditionalFormatRules(nuevasReglas.concat(reglasConservadas));
}

/**
 * Aplica colores de fila por Etapa en la hoja Graduados (menú / llamada manual).
 */
function aplicarColoresFilasGraduados() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('Graduados');
  if (!hoja) { SpreadsheetApp.getUi().alert('No se encontró la hoja "Graduados".'); return; }
  var numCols = ESTRUCTURA_HOJAS['Graduados'].columnas.length;
  _colorearFilasPorEtapa(hoja, numCols, 15); // col 15 = Etapa
  SpreadsheetApp.getUi().alert('✅ Colores de fila aplicados en Graduados.');
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
 * Verifica si un graduado ya existe en el sistema.
 * Busca por nombre completo (col D, índice 3) ya que col A ahora es número secuencial.
 * @param {string} nombre
 * @return {boolean}
 */
function graduadoExiste(nombre) {
  if (!nombre) return false;
  const normNombre = _normalizarTexto_(nombre);
  const hoja  = obtenerHoja('Graduados');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (_normalizarTexto_(datos[i][3]) === normNombre) return true;
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

    // Mapa Creamos ID → Nombre completo desde la hoja Graduados
    // Se construye con clave exacta Y clave normalizada para tolerar acentos/mayúsculas
    var mapaNombres = {};
    var mapaNombresNorm = {};
    (function() {
      var hGrad = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Graduados');
      if (!hGrad || hGrad.getLastRow() < 2) return;
      hGrad.getRange(2, 3, hGrad.getLastRow() - 1, 2).getValues().forEach(function(r) {
        var cid    = (r[0] || '').toString().trim();
        var nombre = (r[1] || '').toString().trim();
        if (!cid || !nombre) return;
        mapaNombres[cid] = nombre;
        mapaNombresNorm[_normalizarTexto_(cid)] = nombre;
      });
    })();

    // Busca nombre por ID exacto, luego normalizado
    function _buscarNombre_(cid) {
      if (!cid) return '';
      return mapaNombres[cid] || mapaNombresNorm[_normalizarTexto_(cid)] || '';
    }

    // Helper: calcular promedio de las 5 preguntas (ignorando vacíos)
    function calcularPromedio(d) {
      var suma = 0, n = 0;
      CAMPOS_NUMERICOS.forEach(function(c) {
        var v = parseFloat((d[c] || '').toString().trim());
        if (!isNaN(v)) { suma += v; n++; }
      });
      return n > 0 ? (suma / n).toFixed(2) : '';
    }

    // Helper: arma la fila completa (Creamos ID | Nombre completo | resto de NOMBRES | Promedio)
    function armarFila(d) {
      var cid    = (d[CAMPO_ID] || '').toString().trim();
      var nombre = _buscarNombre_(cid);
      var row    = [cid, nombre];
      ORDEN.slice(1).forEach(function(k) { row.push(d[k] || ''); }); // desde Fecha envío en adelante
      row.push(calcularPromedio(d));
      return row;
    }

    var hoja = obtenerHoja('Satisfacción Empleo');
    var ultimaFechaSatisf = _leerUltimaFecha(PROP_LAST_SATISF);

    // --- Primera importación (hoja vacía) ----------------------------------
    if (hoja.getLastRow() <= 1) {
      var headers = ['Creamos ID', 'Nombre completo'];
      ORDEN.slice(1).forEach(function(k) { headers.push(NOMBRES[k]); });
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
    // Col 1 = Creamos ID, Col 3 = Fecha envío (con nueva col Nombre completo en col 2)
    var existentes = {};
    if (!ultimaFechaSatisf) {
      var uf = hoja.getLastRow();
      if (uf > 1) {
        hoja.getRange(2, 1, uf - 1, 3).getValues().forEach(function(r) {
          var id = (r[0] || '').toString().trim();
          var f  = (r[2] || '').toString().trim(); // col 3 = Fecha envío
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
    ss.toast('Revisando nuevos registros...', '🔄 Auto-import', 5);

    // Sheets externos (AYB + TECH)
    try {
      importarGraduadosDesdeExterno();
    } catch(e) { Logger.log('Auto-import externo omitido: ' + e.message); }

    // KoboToolbox
    try {
      var datos = obtenerDatosKoboToolbox();
      var resG  = procesarDatosGraduados(datos);
      if (resG.nuevos > 0) msgs.push('Graduados Kobo: +' + resG.nuevos);
    } catch(e) { Logger.log('Auto-import Graduados Kobo omitido: ' + e.message); }

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
    if (!graduadoExiste(graduadoValidado.nombre)) {
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
    hoja.getLastRow(),                      // 1  No. (número secuencial automático)
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
  const hoja      = obtenerHoja('Graduados');
  const datos     = hoja.getDataRange().getValues();
  const normNombre = _normalizarTexto_(graduado.nombre);
  for (let i = 1; i < datos.length; i++) {
    if (_normalizarTexto_(datos[i][3]) === normNombre) { // col D (índice 3) = Nombre completo
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
    'Paso a paso - Cierre':       'Paso a paso',
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
  // Usar la fecha real del graduado (col B = Fecha de envío), NO la fecha de hoy.
  var _fechaOrigen = datosGraduado[1];
  var _fechaIngreso = _fechaOrigen instanceof Date
    ? _fechaOrigen.toLocaleDateString('es-ES')
    : (_fechaOrigen ? _fechaOrigen.toString() : new Date().toLocaleDateString('es-ES'));

  const filaBase = [
    _fechaIngreso, // Fecha de ingreso (ahora viene de la fila de Graduados, no de hoy)
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

    case 'Paso a paso - Cierre':
      return filaBase.concat([
        datosAdicionales.dpi       || '',
        datosAdicionales.formacion || datosGraduado[8] || '',
        datosAdicionales.cohorte   || datosGraduado[9] || '',
        '[CIERRE FORMAL] ' + (datosAdicionales.nota || ''),
        'No'
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

    // Enviar notificación por correo a los destinatarios configurados
    _enviarNotificacionConexionLaboral(datos, fechaInicio);

    return { exito: true, mensaje: 'Conexión laboral guardada. Seguimiento WhatsApp programado.' };
  } catch (error) {
    Logger.log('Error al guardar conexión laboral: ' + error);
    return { exito: false, mensaje: 'Error al guardar: ' + error.message };
  }
}

// ---------------------------------------------------------------------------
// EMAIL PARA CONEXIONES LABORALES
// ---------------------------------------------------------------------------

const PROP_EMAILS_CONEXIONES = 'EMAILS_CONEXIONES_LABORALES';

/**
 * Envía un correo de notificación a todos los destinatarios configurados
 * cuando se registra una nueva conexión laboral.
 */
function _enviarNotificacionConexionLaboral(datos, fechaInicio) {
  try {
    var emailsRaw = PropertiesService.getScriptProperties().getProperty(PROP_EMAILS_CONEXIONES) || '';
    var emails = emailsRaw.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e.length > 0; });
    if (emails.length === 0) return;

    var nombre   = datos.nombreCompleto || datos.creamosId || '';
    var empresa  = datos.empresa || '(sin empresa)';
    var cargo    = datos.cargo   || '(sin cargo)';
    var telefono = datos.telefono || '';
    var salario  = datos.salario ? 'Q' + datos.salario : '(no indicado)';

    var asunto = '🟢 Nueva Conexión Laboral: ' + nombre + ' en ' + empresa;
    var cuerpo =
      '<h2 style="color:#2e7d32;">✅ Nueva Conexión Laboral Registrada</h2>' +
      '<table style="border-collapse:collapse;font-size:14px;">' +
      '<tr><td style="padding:6px 12px;font-weight:bold;">Nombre:</td><td style="padding:6px 12px;">' + nombre + '</td></tr>' +
      '<tr style="background:#f1f8e9;"><td style="padding:6px 12px;font-weight:bold;">Creamos ID:</td><td style="padding:6px 12px;">' + (datos.creamosId || '') + '</td></tr>' +
      '<tr><td style="padding:6px 12px;font-weight:bold;">Empresa:</td><td style="padding:6px 12px;">' + empresa + '</td></tr>' +
      '<tr style="background:#f1f8e9;"><td style="padding:6px 12px;font-weight:bold;">Cargo:</td><td style="padding:6px 12px;">' + cargo + '</td></tr>' +
      '<tr><td style="padding:6px 12px;font-weight:bold;">Teléfono:</td><td style="padding:6px 12px;">' + telefono + '</td></tr>' +
      '<tr style="background:#f1f8e9;"><td style="padding:6px 12px;font-weight:bold;">Fecha de inicio:</td><td style="padding:6px 12px;">' + (fechaInicio || '(no indicada)') + '</td></tr>' +
      '<tr><td style="padding:6px 12px;font-weight:bold;">Salario mensual:</td><td style="padding:6px 12px;">' + salario + '</td></tr>' +
      '</table>' +
      '<p style="color:#555;font-size:12px;margin-top:16px;">— Sistema Empleabilidad Creamos</p>';

    emails.forEach(function(email) {
      MailApp.sendEmail({ to: email, subject: asunto, htmlBody: cuerpo });
      Logger.log('Correo Conexión Laboral enviado a: ' + email);
    });
  } catch (err) {
    Logger.log('Error al enviar correo de Conexión Laboral: ' + err);
  }
}

/**
 * Abre un diálogo para gestionar los correos fijos que reciben
 * notificaciones de nuevas Conexiones Laborales.
 */
function configurarEmailsConexionesLaborales() {
  var ui      = SpreadsheetApp.getUi();
  var props   = PropertiesService.getScriptProperties();
  var actual  = props.getProperty(PROP_EMAILS_CONEXIONES) || '';

  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial,sans-serif;font-size:13px;padding:16px;}' +
    'label{font-weight:bold;}textarea{width:100%;height:120px;margin:8px 0;padding:6px;font-size:13px;}' +
    'button{padding:8px 16px;margin:4px;font-size:13px;cursor:pointer;}' +
    '.info{color:#555;font-size:11px;margin-bottom:8px;}' +
    '.ok{color:green;} .err{color:red;}</style>' +
    '<h3 style="color:#1565c0;">📧 Correos para Conexiones Laborales</h3>' +
    '<p class="info">Ingresa los correos que recibirán una notificación cada vez que se ' +
    'registre una nueva Conexión Laboral. Separa múltiples correos con coma (,).</p>' +
    '<label>Destinatarios:</label>' +
    '<textarea id="emails">' + actual + '</textarea>' +
    '<div id="msg"></div>' +
    '<button onclick="guardar()">💾 Guardar</button>' +
    '<button onclick="google.script.host.close()">Cerrar</button>' +
    '<script>' +
    'function guardar(){' +
    '  var v=document.getElementById("emails").value.trim();' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){document.getElementById("msg").innerHTML=' +
    '      "<span class=ok>✅ Guardado correctamente.</span>";})' +
    '    .withFailureHandler(function(e){document.getElementById("msg").innerHTML=' +
    '      "<span class=err>❌ "+e.message+"</span>";})' +
    '    ._guardarEmailsConexiones(v);' +
    '}' +
    '</script>'
  ).setWidth(480).setHeight(300);

  ui.showModalDialog(html, '📧 Configurar correos — Conexiones Laborales');
}

/**
 * Guarda la lista de correos en las propiedades del script.
 * Llamada desde el diálogo HTML.
 */
function _guardarEmailsConexiones(emailsTexto) {
  PropertiesService.getScriptProperties().setProperty(PROP_EMAILS_CONEXIONES, emailsTexto);
  Logger.log('Correos Conexiones Laborales guardados: ' + emailsTexto);
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
 * Genera el reporte completo en la hoja "Reporte" (o "Reporte YYYY" si se
 * pasa un año). Lee datos directamente de las hojas de clasificación y Graduados.
 * @param {number} [año] - Si se pasa, filtra solo ese año (ej: 2025)
 */
function generarReporte2025() { generarReporte(2025); }

/**
 * Numera todas las filas de Graduados en la columna A (1, 2, 3...).
 * Útil para corregir filas existentes que no tienen número o tienen el ID de Kobo.
 */
function numerarGraduados() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('Graduados');
  if (!hoja || hoja.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('La hoja Graduados está vacía.'); return;
  }
  var total = hoja.getLastRow() - 1;
  var nums  = [];
  for (var i = 1; i <= total; i++) nums.push([i]);
  hoja.getRange(2, 1, total, 1).setValues(nums);
  SpreadsheetApp.getActiveSpreadsheet()
    .toast(total + ' filas numeradas (1 → ' + total + ')', '✅ Numeración lista', 4);
}

/**
 * Corrige la columna "Fecha de ingreso" en todas las hojas de clasificación.
 * Para cada fila busca la "Fecha de envío" real en Graduados usando el Creamos ID,
 * y reemplaza la fecha incorrecta (ej: 18/5/2026 = hoy) con la fecha correcta de Graduados.
 * Solo modifica filas donde la fecha parece incorrecta (año actual o vacía).
 */
function corregirFechasIngreso() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var hoy = new Date();
  var anioActual = hoy.getFullYear();

  // ── Construir mapa Creamos ID → Fecha de envío desde Graduados ──
  var hojaGrad = ss.getSheetByName('Graduados');
  if (!hojaGrad || hojaGrad.getLastRow() < 2) {
    ui.alert('❌ Error', 'No se encontró la hoja "Graduados" o está vacía.', ui.ButtonSet.OK);
    return;
  }
  var datosGrad = hojaGrad.getRange(2, 1, hojaGrad.getLastRow() - 1, 4).getValues();
  var mapaFechas = {}; // { creamosId: fechaDeEnvio }
  datosGrad.forEach(function(fila) {
    var cid   = (fila[2] || '').toString().trim(); // col C = Creamos ID (índice 2)
    var fecha = fila[1];                            // col B = Fecha de envío (índice 1)
    if (cid && fecha) mapaFechas[cid] = fecha;
  });

  // También mapa normalizado para manejar acentos (ej: GLPÉ vs GLPE)
  var mapaNorm = {};
  Object.keys(mapaFechas).forEach(function(cid) {
    mapaNorm[_normalizarTexto_(cid)] = mapaFechas[cid];
  });

  var totalCorregidos = 0;
  var detalle = [];

  // ── Recorrer hojas de clasificación ──
  // colFechaIngreso = índice 0-based de "Fecha de ingreso"
  // colCreamosId    = índice 0-based de "Creamos ID"
  var hojas = [
    { nombre: 'Aliados',                   colFecha: 0, colCid: 1 },
    { nombre: 'Plataforma',                colFecha: 0, colCid: 1 },
    { nombre: 'Derivaciones',              colFecha: 0, colCid: 1 },
    { nombre: 'Paso a paso',               colFecha: 0, colCid: 1 },
    { nombre: 'Activamente busca trabajo', colFecha: 0, colCid: 1 },
    { nombre: 'Conexiones Laborales',      colFecha: 0, colCid: 1 }
  ];

  hojas.forEach(function(def) {
    var h = ss.getSheetByName(def.nombre);
    if (!h || h.getLastRow() < 2) return;

    var nCols = Math.max(def.colFecha, def.colCid) + 1;
    var filas = h.getRange(2, 1, h.getLastRow() - 1, nCols).getValues();

    for (var i = 0; i < filas.length; i++) {
      var fechaActual = filas[i][def.colFecha];
      var cid         = (filas[i][def.colCid] || '').toString().trim();
      if (!cid) continue;

      // Detectar si la fecha es "incorrecta": año actual (2026) o vacía
      var esFechaIncorrecta = false;
      if (!fechaActual || fechaActual.toString().trim() === '') {
        esFechaIncorrecta = true;
      } else {
        var anioFecha = fechaActual instanceof Date
          ? fechaActual.getFullYear()
          : parseInt((fechaActual.toString().split('/')[2] || '').substring(0, 4), 10);
        if (anioFecha === anioActual) esFechaIncorrecta = true;
      }

      if (!esFechaIncorrecta) continue; // fecha parece correcta, no tocar

      // Buscar la fecha real en el mapa (exacto → normalizado)
      var fechaCorrecta = mapaFechas[cid] || mapaNorm[_normalizarTexto_(cid)];
      if (!fechaCorrecta) continue;

      // Formatear y escribir
      var fechaFormateada = fechaCorrecta instanceof Date
        ? fechaCorrecta.toLocaleDateString('es-ES')
        : fechaCorrecta.toString();

      h.getRange(i + 2, def.colFecha + 1).setValue(fechaCorrecta);
      totalCorregidos++;
      detalle.push(def.nombre + ' fila ' + (i + 2) + ': ' + cid +
                   ' → ' + fechaFormateada);
    }
  });

  if (totalCorregidos === 0) {
    ui.alert('✅ Sin cambios',
      'No se encontraron fechas incorrectas (año ' + anioActual + ') en las hojas de clasificación.\n\n' +
      'Si tienes filas con fechas incorrectas de otro año, usa "Verificar Creamos ID ahora" para diagnosticar.',
      ui.ButtonSet.OK);
  } else {
    var resumen = '✅ Se corrigieron ' + totalCorregidos + ' fecha(s):\n\n' +
      detalle.slice(0, 20).join('\n') +
      (detalle.length > 20 ? '\n... y ' + (detalle.length - 20) + ' más.' : '');
    ui.alert('📅 Fechas corregidas', resumen, ui.ButtonSet.OK);
  }
}

function generarReporte(año) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const nomHoja   = año ? 'Reporte ' + año : 'Reporte';
  var hoja = ss.getSheetByName(nomHoja) || ss.insertSheet(nomHoja);
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

  // Extrae el año de un valor de celda (Date, string dd/mm/yyyy o ISO yyyy-mm-dd)
  function _anioFecha_(raw) {
    if (!raw) return null;
    if (raw instanceof Date) return raw.getFullYear();
    var s = raw.toString().trim();
    if (s.length >= 10 && s.charAt(4) === '-') return parseInt(s.substring(0, 4), 10); // ISO
    var p = s.split('/');
    if (p.length === 3) return parseInt(p[2], 10); // dd/mm/yyyy
    return null;
  }

  // colNombre: índice 0-based de la columna que determina si la fila tiene datos.
  // colFecha: índice 0-based de la columna de fecha (para filtrar por año). Si es null, no filtra.
  function contarFilas(nombreHoja, colNombre, colFecha) {
    var h = ss.getSheetByName(nombreHoja);
    if (!h || h.getLastRow() <= 1) return 0;
    var nCols = Math.max(colNombre, colFecha !== undefined && colFecha !== null ? colFecha : 0) + 1;
    var vals  = h.getRange(2, 1, h.getLastRow() - 1, nCols).getValues();
    return vals.filter(function(r) {
      if (!r[colNombre] || r[colNombre].toString().trim() === '') return false;
      if (año && colFecha !== null && colFecha !== undefined) {
        return _anioFecha_(r[colFecha]) === año;
      }
      return true;
    }).length;
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
  // colFecha=0 → "Fecha de ingreso" en las hojas de clasificación
  // colFecha=1 → "Fecha de envío" en Graduados
  const conteos = {};
  var totalParticipantes = 0;
  ETAPAS_FLUJO.forEach(function(e) {
    var col = (e === 'Conexiones Laborales') ? 1 : 2;
    conteos[e] = contarFilas(e, col, 0);
    totalParticipantes += conteos[e];
  });

  // Graduados (col nombre = índice 3, col fecha = índice 1)
  var totalGraduados = contarFilas('Graduados', 3, 1);

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
      if (año && fch.getFullYear() !== año) continue; // filtrar por año si se pidió
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
      if (año && fch.getFullYear() !== año) continue; // filtrar por año
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
      .setValue(año ? 'REPORTE ' + año + ' — SEGUIMIENTO EMPLEABILIDAD' : 'REPORTE DE SEGUIMIENTO — EMPLEABILIDAD')
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

  // ╔══════════════════════════════════════════╗
  // ║  LEYENDA DE COLORES POR ETAPA           ║
  // ╚══════════════════════════════════════════╝
  f = titulo(f, '🎨  LEYENDA — SIGNIFICADO DE COLORES POR ETAPA', C.secBg, C.secFg);

  var leyendaItems = [
    { etapa: 'Paso a paso - Cierre',      accion: 'Proceso de cierre formal del acompañamiento.' },
    { etapa: 'Paso a paso',               accion: 'Persona en acompañamiento activo paso a paso.' },
    { etapa: 'Conexiones Laborales',      accion: 'Persona con conexión laboral activa (empleada).' },
    { etapa: 'Aliados',                   accion: 'Persona postulando con empresa aliada.' },
    { etapa: 'Plataforma',                accion: 'Persona registrada en plataforma de empleo.' },
    { etapa: 'Derivaciones',              accion: 'Persona derivada a empresa para proceso.' },
    { etapa: 'Activamente busca trabajo', accion: 'Persona buscando empleo activamente / formándose.' }
  ];

  var filaInicioLeyenda = f;
  leyendaItems.forEach(function(item) {
    var c = COLORES_DROPDOWN[item.etapa] || { bg: '#eeeeee', fg: '#000000' };
    hoja.setRowHeight(f, 28);
    // Celda de color (muestra el color de la etapa)
    hoja.getRange(f, 1, 1, 1)
        .setValue('  ' + item.etapa)
        .setBackground(c.bg).setFontColor(c.fg)
        .setFontWeight('bold').setFontSize(10)
        .setVerticalAlignment('middle');
    // Descripción de la acción
    hoja.getRange(f, 2, 1, W - 1).merge()
        .setValue(item.accion)
        .setBackground('#fafafa').setFontSize(10)
        .setVerticalAlignment('middle');
    f++;
  });
  borde(filaInicioLeyenda, 1, leyendaItems.length, W);
  f += 2;

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
 * Busca en el mapa Creamos por nombre completo.
 * Primero intenta coincidencia normalizada exacta (resuelve casos de acentos:
 * "Mendez" == "Méndez" tras normalizar). Luego fuzzy con Levenshtein ≤ 2.
 * Devuelve el mismo formato que _buscarEnMapaFuzzy_.
 */
function _buscarPorNombre_(nombre, mapaExt) {
  var normNombre = _normalizarTexto_(nombre);
  if (!normNombre) return null;

  // 1. Coincidencia normalizada exacta — captura diferencias solo de acentos/mayúsculas
  for (var i = 0; i < mapaExt.todos.length; i++) {
    var t = mapaExt.todos[i];
    if (t.normNombre === normNombre)
      return { rec: t, metodo: 'nombre-exacto', idSugerido: t.cid };
  }

  // 2. Fuzzy sobre nombre normalizado (ej: apellido con 1 letra diferente)
  var mejorDist = 999, mejorRec = null;
  for (var i = 0; i < mapaExt.todos.length; i++) {
    var t = mapaExt.todos[i];
    if (!t.normNombre) continue;
    var d = _levenshtein_(normNombre, t.normNombre);
    if (d < mejorDist) { mejorDist = d; mejorRec = t; }
  }
  if (mejorDist <= 2) return { rec: mejorRec, metodo: 'nombre-dist' + mejorDist, idSugerido: mejorRec.cid };

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

  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var cambios    = 0;
  var corregidos = 0;
  var porNombre  = 0;

  function _llenar_(hoja, row, colIdx1based, val) {
    if (hoja.getRange(row, colIdx1based).getValue()) return;
    hoja.getRange(row, colIdx1based).setValue(val);
    cambios++;
  }

  // Resuelve match por ID (con fuzzy) o, si no hay ID, por nombre normalizado.
  function _resolver_(cid, nombreEnHoja) {
    if (cid) return _buscarEnMapaFuzzy_(cid, mapaExt);
    return _buscarPorNombre_(nombreEnHoja, mapaExt);
  }

  // ── Graduados: Creamos ID = col C (idx 3), Nombre = col D (idx 4), Edad = col F (idx 6) ──
  var hojaGrad = ss.getSheetByName('Graduados');
  if (hojaGrad && hojaGrad.getLastRow() > 1) {
    var filas = hojaGrad.getRange(2, 1, hojaGrad.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < filas.length; i++) {
      var cid    = (filas[i][2] || '').toString().trim();
      var nombre = (filas[i][3] || '').toString().trim();
      var match  = _resolver_(cid, nombre);
      if (!match) continue;
      var rec = match.rec;
      var row = i + 2;
      if (!cid && match.metodo.indexOf('nombre') === 0) {
        // Rellenar Creamos ID encontrado por nombre (acento u ortografía)
        hojaGrad.getRange(row, 3).setValue(rec.cid).setBackground('#FFF9C4')
          .setNote('ID encontrado por nombre: "' + nombre + '" → ' + rec.cid);
        porNombre++;
      } else if (match.metodo === 'dist1' || match.metodo === 'normalizado') {
        hojaGrad.getRange(row, 3).setValue(rec.cid).setBackground('#FFF9C4')
          .setNote('ID corregido: "' + cid + '" → "' + rec.cid + '"');
        corregidos++;
      }
      if (!filas[i][3] && rec.nombre) _llenar_(hojaGrad, row, 4, rec.nombre);
      if (!filas[i][5] && rec.edad)   _llenar_(hojaGrad, row, 6, rec.edad);
    }
  }

  // ── Hojas de clasificación: Creamos ID = col B (idx 2), Nombre = col C (idx 3), Edad = col F (idx 6) ──
  ['Aliados', 'Plataforma', 'Derivaciones', 'Paso a paso', 'Activamente busca trabajo'].forEach(function(nomHoja) {
    var h = ss.getSheetByName(nomHoja);
    if (!h || h.getLastRow() < 2) return;
    var filas = h.getRange(2, 1, h.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < filas.length; i++) {
      var cid    = (filas[i][1] || '').toString().trim();
      var nombre = (filas[i][2] || '').toString().trim();
      var match  = _resolver_(cid, nombre);
      if (!match) continue;
      var rec = match.rec;
      var row = i + 2;
      if (!cid && match.metodo.indexOf('nombre') === 0) {
        h.getRange(row, 2).setValue(rec.cid).setBackground('#FFF9C4')
          .setNote('ID encontrado por nombre: "' + nombre + '" → ' + rec.cid);
        porNombre++;
      } else if (match.metodo === 'dist1' || match.metodo === 'normalizado') {
        h.getRange(row, 2).setValue(rec.cid).setBackground('#FFF9C4')
          .setNote('ID corregido: "' + cid + '" → "' + rec.cid + '"');
        corregidos++;
      }
      if (!filas[i][2] && rec.nombre) _llenar_(h, row, 3, rec.nombre);
      if (!filas[i][5] && rec.edad)   _llenar_(h, row, 6, rec.edad);
    }
  });

  // ── Clasificación de Perfiles: Creamos ID = col B (idx 2), Nombre = col C (idx 3) ──
  var hojaClasif = ss.getSheetByName('Clasificación de Perfiles');
  if (hojaClasif && hojaClasif.getLastRow() > 1) {
    var filas = hojaClasif.getRange(2, 1, hojaClasif.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < filas.length; i++) {
      var cid    = (filas[i][1] || '').toString().trim();
      var nombre = (filas[i][2] || '').toString().trim();
      var match  = _resolver_(cid, nombre);
      if (!match) continue;
      var rec = match.rec;
      if (!filas[i][2] && rec.nombre) _llenar_(hojaClasif, i + 2, 3, rec.nombre);
    }
  }

  // ── Satisfacción Empleo: Creamos ID = col A (idx 0), Nombre = col B (idx 1) ──
  // Completa "Nombre completo" usando la hoja Graduados como fuente.
  // Búsqueda: exacta primero, luego normalizada (sin acentos/mayúsculas) y luego
  // fuzzy para tolerar typos de 1 carácter.
  (function() {
    var hSatEmp = ss.getSheetByName('Satisfacción Empleo');
    if (!hSatEmp || hSatEmp.getLastRow() < 2) return;
    // Construir mapa exacto y normalizado desde Graduados
    var mapaNombres = {}, mapaNorm = {}, listaCids = [];
    if (hojaGrad && hojaGrad.getLastRow() > 1) {
      hojaGrad.getRange(2, 3, hojaGrad.getLastRow() - 1, 2).getValues().forEach(function(r) {
        var cid = (r[0] || '').toString().trim();
        var nom = (r[1] || '').toString().trim();
        if (!cid || !nom) return;
        mapaNombres[cid] = nom;
        mapaNorm[_normalizarTexto_(cid)] = nom;
        listaCids.push({ cid: cid, norm: _normalizarTexto_(cid), nom: nom });
      });
    }
    var filSat = hSatEmp.getRange(2, 1, hSatEmp.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < filSat.length; i++) {
      var cid = (filSat[i][0] || '').toString().trim();
      var nom = (filSat[i][1] || '').toString().trim();
      if (!cid || nom) continue;
      // 1. Exacto
      var encontrado = mapaNombres[cid];
      // 2. Normalizado (sin acentos / mayúsculas)
      if (!encontrado) encontrado = mapaNorm[_normalizarTexto_(cid)];
      // 3. Fuzzy levenshtein ≤ 1 sobre el ID normalizado
      if (!encontrado) {
        var normCid = _normalizarTexto_(cid);
        for (var j = 0; j < listaCids.length; j++) {
          if (_levenshtein_(normCid, listaCids[j].norm) <= 1) {
            encontrado = listaCids[j].nom;
            break;
          }
        }
      }
      if (encontrado) _llenar_(hSatEmp, i + 2, 2, encontrado);
    }
  })();

  var partes = [];
  if (cambios > 0)     partes.push(cambios     + ' campos completados');
  if (corregidos > 0)  partes.push(corregidos  + ' IDs corregidos (amarillo)');
  if (porNombre > 0)   partes.push(porNombre   + ' IDs encontrados por nombre (amarillo)');
  var msg = partes.length ? partes.join('\n') : 'No había campos pendientes.';
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

  var porNombreFnd = 0;

  function _procesarFila_(hoja, fila, row, colCid, colNombre, colEdad) {
    var cid    = (fila[colCid]    || '').toString().trim();
    var nombre = colNombre !== null ? (fila[colNombre] || '').toString().trim() : '';
    var celId  = hoja.getRange(row, colCid + 1);

    if (!cid) {
      // Intentar encontrar el Creamos ID por nombre (maneja acentos y typos)
      var matchNom = _buscarPorNombre_(nombre, mapaExt);
      if (matchNom) {
        var rec = matchNom.rec;
        celId.setValue(rec.cid).setBackground(C_TYPO)
          .setNote('🔎 ID encontrado por nombre: "' + nombre + '" → ' + rec.cid +
                   (matchNom.metodo !== 'nombre-exacto' ? ' (similitud)' : ''));
        porNombreFnd++;
        if (colNombre !== null && !fila[colNombre] && rec.nombre) {
          hoja.getRange(row, colNombre + 1).setValue(rec.nombre); completados++;
        }
        if (colEdad !== null && !fila[colEdad] && rec.edad) {
          hoja.getRange(row, colEdad + 1).setValue(rec.edad); completados++;
        }
      } else {
        celId.setBackground(C_SIN_ID).setNote('⚠️ Pendiente: crear perfil en Salesforce con este registro');
        sinId++;
      }
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
        .setNote('✏️ ID corregido: "' + cid + '" → "' + rec.cid + '"');
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
  if (completados > 0)   partes.push(completados   + ' campos completados');
  if (corregidos > 0)    partes.push(corregidos    + ' IDs corregidos por typo (amarillo)');
  if (porNombreFnd > 0)  partes.push(porNombreFnd  + ' IDs encontrados por nombre (amarillo)');
  if (sugerencias > 0)   partes.push(sugerencias   + ' posibles typos (ámbar) — revisar');
  if (sinId > 0)         partes.push(sinId         + ' sin Creamos ID (naranja)');
  if (noBD > 0)          partes.push(noBD          + ' IDs no encontrados en Salesforce (rojo)');
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
/**
 * Instala un trigger onOpen (instalable, con autorización) que ejecuta
 * autocompletarConCreamos() cada vez que alguien abre el documento.
 * Elimina instalaciones previas para no duplicar.
 */
function instalarTriggerAutocompletarAlAbrir() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autocompletarConCreamos' &&
        t.getEventType() === ScriptApp.EventType.ON_OPEN) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('autocompletarConCreamos')
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  SpreadsheetApp.getUi().alert(
    '✅ Autocomplete activado al abrir',
    'Cada vez que abras este documento, el sistema completará\n' +
    'automáticamente los campos faltantes usando "Copy of CREAMOS ID nuevo".\n\n' +
    'Puedes desactivarlo desde el menú si lo necesitas.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Desinstala el trigger onOpen de autocompletar.
 */
function desactivarTriggerAutocompletarAlAbrir() {
  var eliminados = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autocompletarConCreamos' &&
        t.getEventType() === ScriptApp.EventType.ON_OPEN) {
      ScriptApp.deleteTrigger(t);
      eliminados++;
    }
  });
  SpreadsheetApp.getUi().alert(
    eliminados > 0 ? '✅ Desactivado' : 'ℹ️ No estaba activo',
    'El autocomplete al abrir ha sido ' + (eliminados > 0 ? 'desactivado.' : 'no estaba instalado.'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

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

// Fuentes externas: hoja "Graduadx" de AYB y TECH → llegan a "Graduados"
const CONFIGS_GRADUADOS_EXTERNOS = [
  {
    FUENTE:     'AYB',
    FILE_ID:    '1Ay1z3HdFHTzSjq7891sQVuEpIXBA8g9XGibjI-wFklc',
    SHEET_NAME: 'Graduadx',
    SHEET_GID:  2143517721,
    FORMACION:  'Alimentos y Bebidas',  // se escribe en col "Formación" de Graduados
    FILTRO:     null                    // sin filtro: jala todos
  },
  {
    FUENTE:     'TECH',
    FILE_ID:    '1En60zjrwPTrSMFrLUWr2KgXmH7y3vcopfpQgy6lY3HU',
    SHEET_NAME: 'Graduadx',
    SHEET_GID:  1808109947,
    FORMACION:  'Tech',
    FILTRO:     { COL: 'Cohorte', VALOR: 'SAC' }  // solo personas cuya cohorte contenga "SAC"
  }
];

// Columnas de la hoja "Graduados" (índice 0-based)
// 0=No. | 1=Fecha de envío | 2=Creamos ID | 3=Nombre completo | 4=Género
// 5=Edad | 6=Nivel educativo | 7=Número de teléfono | 8=Formación
// 9=Cohorte | 10=Fecha de entrevista | 11=Empleado | 12=Próxima llamada
// 13=Notas | 14=Etapa
const _MAPA_GRADUADOS_ = {
  'creamos id':          2,
  'nombre completo':     3,
  'nombre':              3,
  'genero':              4,
  'sexo':                4,   // alias por si el externo usa "Sexo" en vez de "Género"
  'edad':                5,
  'nivel educativo':     6,
  'numero de telefono':  7,
  'telefono':            7,
  'formacion':           8,
  'cohorte':             9,
  'fecha de entrevista': 10,
  'fecha graduacion':    1,   // "Fecha Graduación" → Fecha de envío
  'fecha de envio':      1,
  'empleado':            11,
  'proxima llamada':     12,
  'notas':               13,
  'etapa':               14
};

// Normaliza valores de Género del sheet externo a los valores del dropdown local
function _normalizarGenero_(valor) {
  const v = String(valor || '').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  if (v === 'hombre' || v === 'masculino' || v === 'm' || v === 'male')   return 'Hombre';
  if (v === 'mujer'  || v === 'femenino'  || v === 'f' || v === 'female') return 'Mujer';
  if (v.indexOf('trans') !== -1)                                           return 'Trans hombre';
  if (v.indexOf('binario') !== -1 || v === 'nb' || v === 'no binario')    return 'No binario';
  if (v === 'otro' || v === 'other' || v === 'prefiero no decir')          return 'Otro';
  // Intentar por inicial si el valor tiene más de 1 carácter
  if (v.charAt(0) === 'h') return 'Hombre';
  if (v.charAt(0) === 'm' && v.length > 1) return 'Mujer';
  return ''; // desconocido → celda vacía (evita error de validación)
}

// Normaliza valores de Nivel educativo a los valores del dropdown local
function _normalizarNivelEducativo_(valor) {
  const v = String(valor || '').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  if (!v) return '';

  // Sin escolaridad
  if (v === 'sin escolaridad' || v === 'ninguno' || v === 'ninguna' ||
      v === 'analfabeto' || v === 'analfabeta' || v === 'sin educacion') return 'Sin escolaridad';

  // Primaria (grados individuales)
  if (v.indexOf('prim') !== -1) {
    if (v.indexOf('sex') !== -1 || v.indexOf('6') !== -1) return 'Sexto primaria';
    if (v.indexOf('quin') !== -1 || v.indexOf('5') !== -1) return 'Quinto primaria';
    if (v.indexOf('cuar') !== -1 || v.indexOf('4') !== -1) return 'Cuarto primaria';
    if (v.indexOf('ter') !== -1  || v.indexOf('3') !== -1) return 'Tercero primaria';
    if (v.indexOf('seg') !== -1  || v.indexOf('2') !== -1) return 'Segundo primaria';
    if (v.indexOf('pri') !== -1  || v.indexOf('1') !== -1) return 'Primero primaria';
    return 'Sexto primaria'; // "Primaria" genérico → asume completo
  }

  // Básico
  if (v.indexOf('basic') !== -1 || v.indexOf('basico') !== -1) {
    if (v.indexOf('ter') !== -1 || v.indexOf('3') !== -1) return 'Tercero básico';
    if (v.indexOf('seg') !== -1 || v.indexOf('2') !== -1) return 'Segundo básico';
    if (v.indexOf('pri') !== -1 || v.indexOf('1') !== -1) return 'Primero básico';
    return 'Tercero básico'; // "Básico" genérico → asume completo
  }

  // Bachillerato / Diversificado
  if (v.indexOf('bach') !== -1) {
    if (v.indexOf('cuar') !== -1 || v.indexOf('4') !== -1) return 'Cuarto bachillerato';
    return 'Quinto bachillerato'; // genérico → asume 5to
  }
  if (v.indexOf('divers') !== -1 || v.indexOf('carrera') !== -1 ||
      v.indexOf('perito') !== -1  || v.indexOf('secretar') !== -1 ||
      v.indexOf('magist') !== -1  || v.indexOf('maestro') !== -1) return 'Diversificado';

  // Universidad
  if (v.indexOf('univer') !== -1 || v.indexOf('licenc') !== -1 ||
      v.indexOf('ingeni') !== -1  || v.indexOf('tecnico uni') !== -1 ||
      v.indexOf('postgrado') !== -1 || v.indexOf('maestria') !== -1) return 'Universidad';

  return ''; // desconocido → celda vacía
}

/**
 * Convierte una fila del sheet externo a las 15 columnas de "Graduados".
 * formacionDefecto: valor fijo para col Formación (ej. "Tech", "Alimentos y Bebidas")
 */
function _mapearFilaAGraduados_(headers, fila, formacionDefecto) {
  const norm = function(s) {
    return String(s || '').toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  };
  const fila15 = new Array(15).fill('');
  for (let i = 0; i < headers.length; i++) {
    const key = norm(headers[i]);
    if (!(key in _MAPA_GRADUADOS_)) continue;
    const destIdx = _MAPA_GRADUADOS_[key];
    var val = fila[i];
    if (destIdx === 4)  val = _normalizarGenero_(val);         // col E: Género
    if (destIdx === 6)  val = _normalizarNivelEducativo_(val); // col G: Nivel educativo
    fila15[destIdx] = val;
  }
  // Formación siempre viene del config (AYB = "Alimentos y Bebidas", TECH = "Tech")
  if (formacionDefecto) fila15[8] = formacionDefecto;
  return fila15;
}

/**
 * Jala la hoja "Graduadx" de AYB y TECH hacia "Graduados".
 * Solo inserta filas nuevas (dedup por Creamos ID o Nombre normalizado).
 * Seguro de ejecutar varias veces: nunca duplica registros existentes.
 */
function importarGraduadosDesdeExterno() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ss.getSheetByName('Graduados');
  if (!hojaDestino) {
    ss.toast('No se encontró la hoja "Graduados"', '❌ Importar Graduados', 8);
    return;
  }

  // Cargar índices de lo que ya existe en Graduados (cols C y D: Creamos ID y Nombre)
  const existentesId     = {};
  const existentesNombre = {};
  const ultFilaLocal = hojaDestino.getLastRow();
  if (ultFilaLocal >= 2) {
    const datosLocal = hojaDestino.getRange(2, 1, ultFilaLocal - 1, 15).getValues();
    for (let i = 0; i < datosLocal.length; i++) {
      const creamosId = String(datosLocal[i][2] || '').trim();
      const nombre    = _normalizarTexto_(datosLocal[i][3]);
      if (creamosId) existentesId[creamosId] = true;
      if (nombre)    existentesNombre[nombre] = true;
    }
  }

  let totalNuevas = 0;
  const errores   = [];

  for (let c = 0; c < CONFIGS_GRADUADOS_EXTERNOS.length; c++) {
    const cfg = CONFIGS_GRADUADOS_EXTERNOS[c];

    // 1. Abrir archivo externo
    let archivo;
    try {
      archivo = SpreadsheetApp.openById(cfg.FILE_ID);
    } catch (e) {
      errores.push(cfg.FUENTE + ': no se pudo abrir (' + e.message + ')');
      continue;
    }

    // 2. Localizar hoja por nombre; fallback por GID
    let hojaExt = archivo.getSheetByName(cfg.SHEET_NAME);
    if (!hojaExt) {
      const hojas = archivo.getSheets();
      for (let i = 0; i < hojas.length; i++) {
        if (hojas[i].getSheetId() === cfg.SHEET_GID) { hojaExt = hojas[i]; break; }
      }
    }
    if (!hojaExt) {
      errores.push(cfg.FUENTE + ': no se encontró "' + cfg.SHEET_NAME + '"');
      continue;
    }

    try {
      const ultFila = hojaExt.getLastRow();
      if (ultFila < 2) continue;

      const numCols  = hojaExt.getLastColumn();
      // Fila 1 = encabezados, fila 2+ = datos
      const headers  = hojaExt.getRange(1, 1, 1, numCols).getValues()[0];
      const datosExt = hojaExt.getRange(2, 1, ultFila - 1, numCols).getValues();

      // Encontrar índices clave en el sheet externo (por nombre de encabezado)
      const norm = function(s) {
        return String(s || '').toLowerCase()
          .normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
      };
      let idxId     = -1;
      let idxNombre = -1;
      let idxFiltro = -1;
      for (let h = 0; h < headers.length; h++) {
        const n = norm(headers[h]);
        if (n === 'creamos id')                        idxId     = h;
        if (n === 'nombre completo' || n === 'nombre') idxNombre = h;
        if (cfg.FILTRO && n === norm(cfg.FILTRO.COL))  idxFiltro = h;
      }

      const nuevas = [];
      for (let i = 0; i < datosExt.length; i++) {
        const fila      = datosExt[i];

        // Aplicar filtro por columna si está configurado (ej. Cohorte = SAC en TECH)
        if (cfg.FILTRO && idxFiltro >= 0) {
          const valFiltro = String(fila[idxFiltro] || '').toUpperCase();
          if (valFiltro.indexOf(cfg.FILTRO.VALOR.toUpperCase()) === -1) continue;
        }

        const creamosId = idxId     >= 0 ? String(fila[idxId]     || '').trim() : '';
        const nombre    = idxNombre >= 0 ? _normalizarTexto_(fila[idxNombre])    : '';

        if (!creamosId && !nombre) continue; // fila vacía

        if (creamosId) {
          if (existentesId[creamosId]) continue;
          existentesId[creamosId] = true;
        } else {
          if (existentesNombre[nombre]) continue;
          existentesNombre[nombre] = true;
        }
        if (nombre) existentesNombre[nombre] = true;

        nuevas.push(_mapearFilaAGraduados_(headers, fila, cfg.FORMACION));
      }

      if (nuevas.length === 0) continue;

      // Asignar número secuencial en col A
      const filaInicio = Math.max(hojaDestino.getLastRow() + 1, 2);
      for (let j = 0; j < nuevas.length; j++) {
        nuevas[j][0] = (filaInicio - 1) + j;
      }
      hojaDestino.getRange(filaInicio, 1, nuevas.length, 15).setValues(nuevas);
      totalNuevas += nuevas.length;
      Logger.log('importarGraduadosDesdeExterno [' + cfg.FUENTE + ']: ' + nuevas.length + ' nuevas');

    } catch (error) {
      Logger.log('Error [' + cfg.FUENTE + ']: ' + error);
      errores.push(cfg.FUENTE + ': ' + error.message);
    }
  }

  if (errores.length > 0) {
    ss.toast('Errores: ' + errores.join(' | '), '⚠️ Importar Graduados', 10);
  } else if (totalNuevas === 0) {
    ss.toast('Sin registros nuevos en AYB ni TECH', 'ℹ️ Importar Graduados', 4);
  } else {
    ss.toast('✅ ' + totalNuevas + ' registro(s) nuevo(s) agregado(s) a Graduados (AYB + TECH)',
             '✅ Importar Graduados', 6);
  }
}

function _obtenerHojaGraduadosImportados() {
  // Mantenida por compatibilidad con código existente
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Graduados');
}

/** Elimina la hoja auxiliar "Graduados Importados" si existe. */
function eliminarHojaGraduadosImportados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const hoja = ss.getSheetByName('Graduados Importados');
  if (!hoja) {
    ui.alert('ℹ️ Info', 'La hoja "Graduados Importados" no existe.', ui.ButtonSet.OK);
    return;
  }
  const resp = ui.alert(
    '¿Eliminar hoja?',
    'Se eliminará la hoja "Graduados Importados" (' + (hoja.getLastRow() - 1) + ' registros).\n\nEsta acción no se puede deshacer.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  ss.deleteSheet(hoja);
  ss.toast('Hoja "Graduados Importados" eliminada', '✅ Listo', 4);
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

/**
 * Diagnóstico: muestra un reporte de lo que ve el script en cada fuente externa.
 * Ejecutar desde el menú para saber por qué no importa.
 */
function diagnosticarImportGraduados() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  let informe = '🔍 DIAGNÓSTICO IMPORT GRADUADOS\n';
  informe += '================================\n\n';

  for (let c = 0; c < CONFIGS_GRADUADOS_EXTERNOS.length; c++) {
    const cfg = CONFIGS_GRADUADOS_EXTERNOS[c];
    informe += '📂 FUENTE: ' + cfg.FUENTE + '\n';
    informe += '   FILE_ID: ' + cfg.FILE_ID + '\n';

    // 1. ¿Se puede abrir?
    let archivo;
    try {
      archivo = SpreadsheetApp.openById(cfg.FILE_ID);
      informe += '   ✅ Archivo accesible: "' + archivo.getName() + '"\n';
    } catch (e) {
      informe += '   ❌ NO se pudo abrir: ' + e.message + '\n\n';
      continue;
    }

    // 2. ¿Qué hojas tiene?
    const hojas = archivo.getSheets();
    informe += '   Hojas encontradas (' + hojas.length + '):\n';
    let hojaObj = null;
    for (let i = 0; i < hojas.length; i++) {
      const marcador = (hojas[i].getName() === cfg.SHEET_NAME || hojas[i].getSheetId() === cfg.SHEET_GID) ? ' ← ESTA' : '';
      informe += '      · "' + hojas[i].getName() + '" (GID: ' + hojas[i].getSheetId() + ')' + marcador + '\n';
      if (marcador) hojaObj = hojas[i];
    }

    if (!hojaObj) {
      informe += '   ❌ No se encontró la hoja "' + cfg.SHEET_NAME + '" (GID ' + cfg.SHEET_GID + ')\n\n';
      continue;
    }

    // 3. ¿Cuántas filas tiene?
    const ultFila = hojaObj.getLastRow();
    informe += '   ✅ Hoja "' + hojaObj.getName() + '" encontrada\n';
    informe += '   Filas totales: ' + ultFila + ' (incluyendo encabezado)\n';

    if (ultFila < 2) {
      informe += '   ⚠️ La hoja está VACÍA (sin datos)\n\n';
      continue;
    }

    // 4. Encabezados (fila 1)
    const numCols = Math.min(hojaObj.getLastColumn(), 16);
    const headers = hojaObj.getRange(1, 1, 1, numCols).getValues()[0];
    informe += '   Columnas (' + numCols + '): ' + headers.slice(0, 6).join(' | ') + '\n';

    // 5. Primera fila de datos
    const primera = hojaObj.getRange(2, 1, 1, numCols).getValues()[0];
    informe += '   Primera fila: ' + primera.slice(0, 4).join(' | ') + '\n';

    // 6. ¿Cuántas pasarían deduplicación?
    const hojaLocal = ss.getSheetByName('Graduados');
    if (!hojaLocal || hojaLocal.getLastRow() < 2) {
      informe += '   "Graduados" vacía → todas las filas serían nuevas\n\n';
    } else {
      informe += '   "Graduados" tiene ' + (hojaLocal.getLastRow() - 1) + ' registro(s)\n\n';
    }
  }

  Logger.log(informe);
  ui.alert('Diagnóstico Import Graduados', informe, ui.ButtonSet.OK);
}

// ===========================================================================
// SECCIÓN 9: EXPORTACIÓN POWER BI
// ===========================================================================
// Genera y mantiene la hoja "Power BI Export" con datos consolidados de
// todas las hojas operativas. Se actualiza automáticamente cada 6 horas
// mediante un trigger de tiempo instalado con configurarTriggerPowerBI().
//
// Columnas de la hoja de exportación:
//   Última actualización | Fuente | Creamos ID | Nombre completo | Género |
//   Edad | Nivel educativo | Número de teléfono | Formación | Cohorte |
//   Fecha ingreso | Empleado | Etapa / Clasificación | Activo | Nota |
//   Empresa | Cargo | Tipo contrato | Fecha inicio empleo | Salario mensual |
//   Tipo seguimiento | Estado seguimiento | Tipo servicio sesión |
//   Satisfacción empleo | Promedio satisfacción
// ===========================================================================

const NOMBRE_HOJA_POWERBI = 'Power BI Export';

const HEADERS_POWERBI = [
  'Última actualización',
  'Fuente',
  'Creamos ID',
  'Nombre completo',
  'Género',
  'Edad',
  'Nivel educativo',
  'Número de teléfono',
  'Formación',
  'Cohorte',
  'Fecha ingreso',
  'Empleado',
  'Etapa / Clasificación',
  'Activo',
  'Nota',
  'Empresa',
  'Cargo',
  'Tipo contrato',
  'Fecha inicio empleo',
  'Salario mensual',
  'Tipo seguimiento',
  'Estado seguimiento',
  'Tipo servicio sesión',
  'Satisfacción empleo (1-5)',
  'Promedio satisfacción (1-5)'
];

/**
 * Genera (o regenera) la hoja "Power BI Export" con datos de todas las hojas.
 * Se puede ejecutar manualmente o es llamada por el trigger cada 6 horas.
 */
function generarExportPowerBI() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const ahora     = new Date().toLocaleString('es-GT', { timeZone: 'America/Guatemala' });

  // Obtener o crear la hoja de exportación
  var hojaExport = ss.getSheetByName(NOMBRE_HOJA_POWERBI);
  if (!hojaExport) {
    hojaExport = ss.insertSheet(NOMBRE_HOJA_POWERBI);
  }

  // Limpiar contenido anterior (solo datos, no la hoja entera)
  hojaExport.clearContents();
  hojaExport.clearFormats();

  // Escribir encabezados
  var headerRange = hojaExport.getRange(1, 1, 1, HEADERS_POWERBI.length);
  headerRange.setValues([HEADERS_POWERBI]);
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  hojaExport.setFrozenRows(1);

  var filas = [];

  // ── 1. GRADUADOS ─────────────────────────────────────────────────────────
  // Cols: 0=No | 1=Fecha envío | 2=Creamos ID | 3=Nombre | 4=Género |
  //       5=Edad | 6=Nivel edu | 7=Teléfono | 8=Formación | 9=Cohorte |
  //       10=Fecha entrevista | 11=Empleado | 12=Próxima llamada |
  //       13=Notas | 14=Etapa
  var hGrad = ss.getSheetByName('Graduados');
  if (hGrad && hGrad.getLastRow() > 1) {
    var datosGrad = hGrad.getRange(2, 1, hGrad.getLastRow() - 1, 15).getValues();
    datosGrad.forEach(function(r) {
      if (!r[2] && !r[3]) return; // fila vacía
      filas.push(_filaPowerBI(ahora, 'Graduados', {
        creamosId:  r[2],
        nombre:     r[3],
        genero:     r[4],
        edad:       r[5],
        nivelEdu:   r[6],
        telefono:   r[7],
        formacion:  r[8],
        cohorte:    r[9],
        fechaIngreso: r[1],
        empleado:   r[11],
        etapa:      r[14],
        nota:       r[13]
      }));
    });
  }

  // ── 2. ALIADOS ────────────────────────────────────────────────────────────
  // Cols: 0-6=COMUNES | 7=CompartióCV | 8=Área | 9=Entrevista |
  //       10=DíaPrueba | 11=Confirmación | 12=Notas | 13=Activo
  var hAliados = ss.getSheetByName('Aliados');
  if (hAliados && hAliados.getLastRow() > 1) {
    var datosAliados = hAliados.getRange(2, 1, hAliados.getLastRow() - 1, 14).getValues();
    datosAliados.forEach(function(r) {
      if (!r[1] && !r[2]) return;
      filas.push(_filaPowerBI(ahora, 'Aliados', {
        fechaIngreso: r[0], creamosId: r[1], nombre: r[2],
        telefono: r[3], genero: r[4], edad: r[5], nivelEdu: r[6],
        etapa: 'Aliados', activo: r[13], nota: r[12]
      }));
    });
  }

  // ── 3. PLATAFORMA ─────────────────────────────────────────────────────────
  // Cols: 0-6=COMUNES | 7=Cita | 8=CreacionPerfil | 9=Contacto |
  //       10=Tramites | 11=Entrevista | 12=Confirmacion | 13=Recepcion |
  //       14=Nota | 15=Activo
  var hPlat = ss.getSheetByName('Plataforma');
  if (hPlat && hPlat.getLastRow() > 1) {
    var datosPlat = hPlat.getRange(2, 1, hPlat.getLastRow() - 1, 16).getValues();
    datosPlat.forEach(function(r) {
      if (!r[1] && !r[2]) return;
      filas.push(_filaPowerBI(ahora, 'Plataforma', {
        fechaIngreso: r[0], creamosId: r[1], nombre: r[2],
        telefono: r[3], genero: r[4], edad: r[5], nivelEdu: r[6],
        etapa: 'Plataforma', activo: r[15], nota: r[14]
      }));
    });
  }

  // ── 4. DERIVACIONES ───────────────────────────────────────────────────────
  // Cols: 0-6=COMUNES | 7-12=Envíos/Llamadas | 13=Notas | 14=Activo
  var hDeriv = ss.getSheetByName('Derivaciones');
  if (hDeriv && hDeriv.getLastRow() > 1) {
    var datosDeriv = hDeriv.getRange(2, 1, hDeriv.getLastRow() - 1, 15).getValues();
    datosDeriv.forEach(function(r) {
      if (!r[1] && !r[2]) return;
      filas.push(_filaPowerBI(ahora, 'Derivaciones', {
        fechaIngreso: r[0], creamosId: r[1], nombre: r[2],
        telefono: r[3], genero: r[4], edad: r[5], nivelEdu: r[6],
        etapa: 'Derivaciones', activo: r[14], nota: r[13]
      }));
    });
  }

  // ── 5. PASO A PASO ────────────────────────────────────────────────────────
  // Cols: 0-6=COMUNES | 7=DPI | 8=Formación | 9=Cohorte | 10=Nota | 11=Activo
  var hPap = ss.getSheetByName('Paso a paso');
  if (hPap && hPap.getLastRow() > 1) {
    var datosPap = hPap.getRange(2, 1, hPap.getLastRow() - 1, 12).getValues();
    datosPap.forEach(function(r) {
      if (!r[1] && !r[2]) return;
      var esCierre = (r[10] || '').toString().indexOf('[CIERRE FORMAL]') !== -1;
      filas.push(_filaPowerBI(ahora, 'Paso a paso', {
        fechaIngreso: r[0], creamosId: r[1], nombre: r[2],
        telefono: r[3], genero: r[4], edad: r[5], nivelEdu: r[6],
        formacion: r[8], cohorte: r[9],
        etapa: esCierre ? 'Paso a paso - Cierre' : 'Paso a paso',
        activo: r[11], nota: r[10]
      }));
    });
  }

  // ── 6. ACTIVAMENTE BUSCA TRABAJO ──────────────────────────────────────────
  // Cols: 0-6=COMUNES | 7=TipoBúsqueda | 8=Mensaje | 9=Llamada |
  //       10=Nota | 11=Entrevista | 12=Trámites | 13=Activo
  var hAbt = ss.getSheetByName('Activamente busca trabajo');
  if (hAbt && hAbt.getLastRow() > 1) {
    var datosAbt = hAbt.getRange(2, 1, hAbt.getLastRow() - 1, 14).getValues();
    datosAbt.forEach(function(r) {
      if (!r[1] && !r[2]) return;
      filas.push(_filaPowerBI(ahora, 'Activamente busca trabajo', {
        fechaIngreso: r[0], creamosId: r[1], nombre: r[2],
        telefono: r[3], genero: r[4], edad: r[5], nivelEdu: r[6],
        etapa: r[7] || 'Activamente busca trabajo',
        activo: r[13], nota: r[10]
      }));
    });
  }

  // ── 7. CONEXIONES LABORALES ───────────────────────────────────────────────
  // Cols: 0=CreamosID | 1=Nombre | 2=Teléfono | 3=Género | 4=Edad |
  //       5=NivelEdu | 6=Tipo | 7=Programa | 8=Proyecto | 9=Especialidad |
  //       10=Empresa | 11=Cargo | 12=TipoDuración | 13=TipoContrato |
  //       14=FechaInicio | 15=FechaFinal | 16=Duración | 17=Salario
  var hCl = ss.getSheetByName('Conexiones Laborales');
  if (hCl && hCl.getLastRow() > 1) {
    var datosCl = hCl.getRange(2, 1, hCl.getLastRow() - 1, 18).getValues();
    datosCl.forEach(function(r) {
      if (!r[0] && !r[1]) return;
      filas.push(_filaPowerBI(ahora, 'Conexiones Laborales', {
        creamosId: r[0], nombre: r[1], telefono: r[2],
        genero: r[3], edad: r[4], nivelEdu: r[5],
        etapa: 'Conexiones Laborales', empleado: 'Si',
        empresa: r[10], cargo: r[11],
        tipoContrato: r[13], fechaInicioEmpleo: r[14],
        salario: r[17]
      }));
    });
  }

  // ── 8. SATISFACCIÓN EMPLEO ────────────────────────────────────────────────
  // Cols: 0=CreamosID | 1=FechaEnvío | 2=SatEmpleo | 3=CumpleExp |
  //       0=CreamosID | 1=NombreCompleto | 2=FechaEnvío | 3=SatisfEmpleo |
  //       4=CumpleExp | 5=AmbLaboral | 6=SalBeneficios | 7=Permanencia |
  //       8=AspMejorar | 9=OtroAspecto | 10=Nota | 11=Promedio
  var hSat = ss.getSheetByName('Satisfacción Empleo');
  if (hSat && hSat.getLastRow() > 1) {
    var datosSat = hSat.getRange(2, 1, hSat.getLastRow() - 1, 12).getValues();
    datosSat.forEach(function(r) {
      if (!r[0]) return;
      filas.push(_filaPowerBI(ahora, 'Satisfacción Empleo', {
        creamosId: r[0], nombre: r[1], fechaIngreso: r[2],
        etapa: 'Satisfacción Empleo',
        nota: r[10],
        satisfaccionEmpleo: r[3],
        promedioSat: r[11]
      }));
    });
  }

  // ── 9. SESIONES ACOMPAÑAMIENTO ────────────────────────────────────────────
  // Cols: 0=CreamosID | 1=FechaEnvío | 2=InicioSesión | 3=Proyecto |
  //       4=Nombre | 5=Apellidos | 6=Teléfono | 7=FechaNac | 8=Edad |
  //       9=Género | 10=AñoIngreso | 11=EnQueAño | 12=GradoAcad |
  //       13=TipoServicio | 14=Comentario | 15=Acción
  var hSes = ss.getSheetByName('Sesiones Acompañamiento');
  if (hSes && hSes.getLastRow() > 1) {
    var datosSes = hSes.getRange(2, 1, hSes.getLastRow() - 1, 15).getValues();
    datosSes.forEach(function(r) {
      if (!r[0] && !r[4]) return;
      filas.push(_filaPowerBI(ahora, 'Sesiones Acompañamiento', {
        creamosId: r[0], fechaIngreso: r[1],
        nombre: ((r[4] || '') + ' ' + (r[5] || '')).trim(),
        telefono: r[6], edad: r[8], genero: r[9],
        etapa: 'Sesiones Acompañamiento',
        nota: r[14], tipoServicioSesion: r[13]
      }));
    });
  }

  // ── 10. SEGUIMIENTO BOT ───────────────────────────────────────────────────
  // Cols: 0=CreamosID | 1=Nombre | 2=Telefono | 3=Empresa | 4=Cargo |
  //       5=FechaEmpleo | ... | 10=EstadoActual | 18=TipoSeguimiento
  var hBot = ss.getSheetByName('Seguimiento Bot');
  if (hBot && hBot.getLastRow() > 1) {
    var lastColBot = Math.min(hBot.getLastColumn(), 20);
    var datosBot = hBot.getRange(2, 1, hBot.getLastRow() - 1, lastColBot).getValues();
    datosBot.forEach(function(r) {
      if (!r[0] && !r[1]) return;
      filas.push(_filaPowerBI(ahora, 'Seguimiento Bot', {
        creamosId: r[0], nombre: r[1], telefono: r[2],
        etapa: 'Seguimiento Bot',
        empresa: r[3], cargo: r[4],
        tipoSeguimiento: r[18] !== undefined ? r[18] : '',
        estadoSeguimiento: r[9] !== undefined ? r[9] : ''
      }));
    });
  }

  // Escribir todas las filas de una vez (más eficiente que appendRow)
  if (filas.length > 0) {
    hojaExport.getRange(2, 1, filas.length, HEADERS_POWERBI.length).setValues(filas);
  }

  // Formato final
  hojaExport.autoResizeColumns(1, HEADERS_POWERBI.length);
  hojaExport.setFrozenRows(1);

  // Celda de estado en la fila 1, última columna + 2
  var colEstado = HEADERS_POWERBI.length + 2;
  hojaExport.getRange(1, colEstado).setValue('Actualizado: ' + ahora);
  hojaExport.getRange(1, colEstado).setFontStyle('italic').setFontColor('#666666');

  Logger.log('generarExportPowerBI: ' + filas.length + ' filas exportadas. ' + ahora);
  return filas.length;
}

/**
 * Construye una fila de la hoja Power BI Export con los campos en orden.
 * Cualquier campo no provisto queda como cadena vacía.
 */
function _filaPowerBI(ahora, fuente, d) {
  return [
    ahora,
    fuente,
    d.creamosId            || '',
    d.nombre               || '',
    d.genero               || '',
    d.edad                 || '',
    d.nivelEdu             || '',
    d.telefono             || '',
    d.formacion            || '',
    d.cohorte              || '',
    d.fechaIngreso         || '',
    d.empleado             || '',
    d.etapa                || '',
    d.activo               || '',
    d.nota                 || '',
    d.empresa              || '',
    d.cargo                || '',
    d.tipoContrato         || '',
    d.fechaInicioEmpleo    || '',
    d.salario              || '',
    d.tipoSeguimiento      || '',
    d.estadoSeguimiento    || '',
    d.tipoServicioSesion   || '',
    d.satisfaccionEmpleo   || '',
    d.promedioSat          || ''
  ];
}

/**
 * Instala (o reinstala) el trigger de tiempo para actualizar Power BI cada 6 horas.
 * Ejecutar una sola vez desde el editor de Apps Script.
 */
function configurarTriggerPowerBI() {
  // Eliminar triggers previos de este tipo
  var eliminados = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'generarExportPowerBI') {
      ScriptApp.deleteTrigger(t);
      eliminados++;
    }
  });

  // Crear trigger cada 6 horas
  ScriptApp.newTrigger('generarExportPowerBI')
    .timeBased()
    .everyHours(6)
    .create();

  // Ejecutar una vez inmediatamente para poblar la hoja ahora mismo
  var filas = generarExportPowerBI();

  SpreadsheetApp.getUi().alert(
    '✅ Power BI Export configurado',
    'Hoja "Power BI Export" creada con ' + filas + ' filas.\n\n' +
    'Triggers previos eliminados: ' + eliminados + '\n' +
    'Nuevo trigger instalado: cada 6 horas\n\n' +
    'Para conectar a Power BI usa:\n' +
    'Obtener datos → Web → URL pública de la hoja\n' +
    'o usa el conector de Google Sheets.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ===========================================================================
// SECCIÓN 10: VERIFICACIÓN DE IDs Y AUDITORÍA
// ===========================================================================
// Audita la integridad de los Creamos ID en todas las hojas operativas,
// detecta IDs que no existen en Graduados o cuyo nombre no coincide, y
// genera reportes de registros sin ID asignado.
//
// CONSTANTES del sistema — ajustar si cambia la estructura:
const NOMBRE_DIRECTORIO_IDS   = 'Graduados';
const NOMBRE_COLUMNA_ID_IDS   = 'Creamos ID';
// Hojas del sistema que NO deben auditarse
const HOJAS_EXCLUIR_AUDITORIA = [
  'Graduados',
  'Reporte',
  'Reporte 2025',
  'Reportes mensuales',
  'Reportes Mensuales',
  'Power BI Export',
  'Clasificación de Perfiles',
  'Satisfacción Empleo',
  'Sesiones Acompañamiento',
  'Seguimiento Bot',
  'Graduados Importados',
  '🔍 Auditoría IDs',
  '📋 Sin ID',
  '⚡ Rendimiento'
];

// ---------------------------------------------------------------------------
// 1. DETECCIÓN DE HOJAS CON ID (con caché de 1 hora)
// ---------------------------------------------------------------------------

/**
 * Detecta qué hojas del spreadsheet tienen la columna 'Creamos ID'.
 * Usa CacheService para evitar releer encabezados en cada llamada.
 *
 * @param {Spreadsheet} ss
 * @param {string}      nombreDirectorio  Hoja raíz a excluir
 * @param {boolean}     forzarRefresh     Si true, ignora la caché
 * @return {string[]}   Nombres de hojas que tienen la columna ID
 */
function _detectarHojasConID(ss, nombreDirectorio, forzarRefresh) {
  var cacheKey = 'hojas_con_id_' + ss.getId();
  var cache    = CacheService.getScriptCache();

  if (!forzarRefresh) {
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) {}
    }
  }

  var excluir = HOJAS_EXCLUIR_AUDITORIA.slice();
  if (excluir.indexOf(nombreDirectorio) === -1) excluir.push(nombreDirectorio);

  var resultado = [];
  var hojas = ss.getSheets();

  for (var i = 0; i < hojas.length; i++) {
    var hoja = hojas[i];
    var nombre = hoja.getName();

    // Excluir hojas del sistema
    var excluida = false;
    for (var e = 0; e < excluir.length; e++) {
      if (excluir[e] === nombre) { excluida = true; break; }
    }
    if (excluida) continue;

    // Leer solo el encabezado, máximo 60 columnas
    var maxCol = Math.min(hoja.getMaxColumns(), 60);
    if (maxCol === 0) continue;
    var encabezado = hoja.getRange(1, 1, 1, maxCol).getValues()[0];

    for (var c = 0; c < encabezado.length; c++) {
      if ((encabezado[c] || '').toString().trim() === NOMBRE_COLUMNA_ID_IDS) {
        resultado.push(nombre);
        break;
      }
    }
  }

  // Guardar en caché por 1 hora (3600 segundos)
  try { cache.put(cacheKey, JSON.stringify(resultado), 3600); } catch (e) {}
  return resultado;
}

/**
 * Dado el encabezado de una hoja, devuelve el índice (0-based) de una columna
 * o -1 si no existe.
 */
function _indiceColumna(encabezado, nombreColumna) {
  for (var i = 0; i < encabezado.length; i++) {
    if ((encabezado[i] || '').toString().trim() === nombreColumna) return i;
  }
  return -1;
}

// ---------------------------------------------------------------------------
// 6. SIMILITUD DE NOMBRES (declarada antes para uso interno)
// ---------------------------------------------------------------------------

/**
 * Normaliza un nombre: minúsculas, sin tildes, sin caracteres especiales.
 */
function _normalizarNombre(nombre) {
  var s = (nombre || '').toString().toLowerCase();
  var tildes = { 'á':'a','é':'e','í':'i','ó':'o','ú':'u','ü':'u','ñ':'n',
                 'à':'a','è':'e','ì':'i','ò':'o','ù':'u' };
  var resultado = '';
  for (var i = 0; i < s.length; i++) {
    var c = s[i];
    resultado += tildes[c] !== undefined ? tildes[c] : c;
  }
  // Solo letras, números y espacios
  resultado = resultado.replace(/[^a-z0-9 ]/g, ' ');
  // Colapsar espacios múltiples
  resultado = resultado.replace(/\s+/g, ' ').trim();
  return resultado;
}

/**
 * Verifica si dos palabras cortas coinciden con tolerancia de hasta 1 carácter
 * de diferencia (solo aplica a palabras > 4 letras).
 */
function _palabrasCoinciden(p1, p2) {
  if (p1 === p2) return true;
  if (p1.length <= 4 || p2.length <= 4) return false;
  // Diferencia de longitud > 1 → no coinciden con tolerancia 1
  if (Math.abs(p1.length - p2.length) > 1) return false;
  // Contar caracteres diferentes
  var diferencias = 0;
  var largo = Math.max(p1.length, p2.length);
  var i1 = 0; var i2 = 0;
  while (i1 < p1.length || i2 < p2.length) {
    var c1 = i1 < p1.length ? p1[i1] : '';
    var c2 = i2 < p2.length ? p2[i2] : '';
    if (c1 !== c2) {
      diferencias++;
      if (diferencias > 1) return false;
      // Intentar saltar en el más largo para alinear
      if (p1.length > p2.length) { i1++; continue; }
      if (p2.length > p1.length) { i2++; continue; }
    }
    i1++; i2++;
  }
  return diferencias <= 1;
}

/**
 * Calcula similitud entre dos nombres (0–100).
 * Basada en palabras: precision*0.65 + recall*0.35
 *
 * @param {string} nombre1
 * @param {string} nombre2
 * @return {number} 0-100
 */
function similitudNombre(nombre1, nombre2) {
  if (!nombre1 || !nombre2) return 0;
  var n1 = _normalizarNombre(nombre1);
  var n2 = _normalizarNombre(nombre2);
  if (n1 === n2) return 100;

  var palabras1 = n1.split(' ').filter(function(p) { return p.length >= 2; });
  var palabras2 = n2.split(' ').filter(function(p) { return p.length >= 2; });
  if (palabras1.length === 0 || palabras2.length === 0) return 0;

  var matches = 0;
  var usadas2 = [];
  for (var i = 0; i < palabras1.length; i++) {
    for (var j = 0; j < palabras2.length; j++) {
      if (usadas2.indexOf(j) !== -1) continue;
      if (_palabrasCoinciden(palabras1[i], palabras2[j])) {
        matches++;
        usadas2.push(j);
        break;
      }
    }
  }

  var precision = matches / palabras1.length;
  var recall    = matches / palabras2.length;
  return Math.round((precision * 0.65 + recall * 0.35) * 100);
}

// ---------------------------------------------------------------------------
// 2. AUDITORÍA DE IDs
// ---------------------------------------------------------------------------

/**
 * Construye el mapa ID → nombre completo leyendo la hoja Graduados.
 * Retorna { mapaID: {}, totalGrads: N }
 */
function _cargarDirectorio(ss) {
  var hoja = ss.getSheetByName(NOMBRE_DIRECTORIO_IDS);
  if (!hoja || hoja.getLastRow() < 2) return { mapaID: {}, totalGrads: 0 };

  var maxCol = Math.min(hoja.getLastColumn(), 20);
  var datos  = hoja.getRange(2, 1, hoja.getLastRow() - 1, maxCol).getValues();
  var header = hoja.getRange(1, 1, 1, maxCol).getValues()[0];

  var colID     = _indiceColumna(header, NOMBRE_COLUMNA_ID_IDS);
  var colNombre = _indiceColumna(header, 'Nombre completo');
  if (colID === -1) return { mapaID: {}, totalGrads: 0 };

  var mapaID = {};
  for (var i = 0; i < datos.length; i++) {
    var id = (datos[i][colID] || '').toString().trim();
    if (!id) continue;
    mapaID[id] = colNombre !== -1 ? (datos[i][colNombre] || '').toString().trim() : '';
  }
  return { mapaID: mapaID, totalGrads: datos.length };
}

/**
 * Recorre todas las hojas detectadas y devuelve array de problemas.
 * Cada problema: { hoja, fila, nombreHoja, id, nombreDirectorio, problema }
 */
function _detectarProblemas(ss, mapaID) {
  var nombresHojas = _detectarHojasConID(ss, NOMBRE_DIRECTORIO_IDS, false);
  var problemas = [];
  var totalRevisados = 0;

  for (var h = 0; h < nombresHojas.length; h++) {
    var hoja   = ss.getSheetByName(nombresHojas[h]);
    if (!hoja || hoja.getLastRow() < 2) continue;

    var maxCol   = Math.min(hoja.getLastColumn(), 60);
    var header   = hoja.getRange(1, 1, 1, maxCol).getValues()[0];
    var colID     = _indiceColumna(header, NOMBRE_COLUMNA_ID_IDS);
    var colNombre = _indiceColumna(header, 'Nombre completo');
    if (colID === -1) continue;

    var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, maxCol).getValues();

    for (var i = 0; i < datos.length; i++) {
      var id = (datos[i][colID] || '').toString().trim();
      if (!id || id.indexOf('⚠️') === 0) continue;

      totalRevisados++;
      var nombreEnHoja = colNombre !== -1 ? (datos[i][colNombre] || '').toString().trim() : '';

      if (mapaID[id] === undefined) {
        problemas.push({
          hoja:             nombresHojas[h],
          fila:             i + 2,
          colID:            colID,
          nombreHoja:       nombreEnHoja,
          id:               id,
          nombreDirectorio: '',
          problema:         '⚠️ ID no encontrado'
        });
      } else if (nombreEnHoja && mapaID[id]) {
        var sim = similitudNombre(nombreEnHoja, mapaID[id]);
        if (sim < 60) {
          problemas.push({
            hoja:             nombresHojas[h],
            fila:             i + 2,
            colID:            colID,
            nombreHoja:       nombreEnHoja,
            id:               id,
            nombreDirectorio: mapaID[id],
            problema:         '⚠️ Nombre no coincide (' + sim + '%)'
          });
        }
      }
    }
  }
  return { problemas: problemas, totalRevisados: totalRevisados };
}

/**
 * Genera la hoja de resultados de auditoría y la llena con los problemas.
 */
function auditarIDs() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();

  ss.toast('Cargando directorio...', '🔍 Auditoría', -1);

  var dir    = _cargarDirectorio(ss);
  var result = _detectarProblemas(ss, dir.mapaID);

  // Crear / limpiar hoja de auditoría
  var nombreHojaAudit = '🔍 Auditoría IDs';
  var hojaAudit = ss.getSheetByName(nombreHojaAudit);
  if (!hojaAudit) hojaAudit = ss.insertSheet(nombreHojaAudit);
  hojaAudit.clearContents();
  hojaAudit.clearFormats();

  var headers = ['Hoja', 'Fila', 'Nombre en hoja', 'ID', 'Nombre en directorio', 'Problema'];
  var hrng = hojaAudit.getRange(1, 1, 1, headers.length);
  hrng.setValues([headers]);
  hrng.setBackground('#b71c1c').setFontColor('#ffffff').setFontWeight('bold');
  hojaAudit.setFrozenRows(1);

  if (result.problemas.length > 0) {
    var filas = result.problemas.map(function(p) {
      return [p.hoja, p.fila, p.nombreHoja, p.id, p.nombreDirectorio, p.problema];
    });
    hojaAudit.getRange(2, 1, filas.length, headers.length).setValues(filas);
    // Color naranja para filas con problema
    hojaAudit.getRange(2, 1, filas.length, headers.length).setBackground('#fff3e0');
    hojaAudit.autoResizeColumns(1, headers.length);
  } else {
    hojaAudit.getRange(2, 1).setValue('✅ No se encontraron problemas');
  }

  ss.toast('Listo', '🔍 Auditoría', 3);

  var totalOk = result.totalRevisados - result.problemas.length;
  ui.alert(
    '🔍 Resultado de Auditoría',
    'Registros revisados: ' + result.totalRevisados + '\n' +
    'Con problemas:       ' + result.problemas.length + '\n' +
    'Sin problemas:       ' + totalOk + '\n\n' +
    'Ver hoja "' + nombreHojaAudit + '" para el detalle.',
    ui.ButtonSet.OK
  );
}

// ---------------------------------------------------------------------------
// 3. LIMPIEZA DE IDs INCORRECTOS
// ---------------------------------------------------------------------------

/**
 * Muestra los IDs con problemas, pide confirmación y borra SOLO el valor
 * de la celda ID en cada caso, marcándola en naranja.
 */
function limpiarIDsIncorrectos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  ss.toast('Analizando IDs...', '🧹 Limpieza', -1);

  var dir    = _cargarDirectorio(ss);
  var result = _detectarProblemas(ss, dir.mapaID);

  ss.toast('', '', 1);

  if (result.problemas.length === 0) {
    ui.alert('✅ Sin problemas', 'No se encontraron IDs incorrectos.', ui.ButtonSet.OK);
    return;
  }

  // Construir lista para mostrar al usuario (máx 20 líneas)
  var lista = '';
  var mostrar = Math.min(result.problemas.length, 20);
  for (var i = 0; i < mostrar; i++) {
    var p = result.problemas[i];
    lista += '• ' + p.hoja + ' fila ' + p.fila + ' — ' + p.id + ': ' + p.problema + '\n';
  }
  if (result.problemas.length > 20) {
    lista += '... y ' + (result.problemas.length - 20) + ' más.\n';
  }

  var confirm = ui.alert(
    '🧹 Confirmar limpieza',
    'Se borrarán ' + result.problemas.length + ' celdas de ID (solo el valor, no el nombre):\n\n' +
    lista + '\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Agrupar por hoja para minimizar llamadas al servidor
  var porHoja = {};
  for (var j = 0; j < result.problemas.length; j++) {
    var prob = result.problemas[j];
    if (!porHoja[prob.hoja]) porHoja[prob.hoja] = [];
    porHoja[prob.hoja].push({ fila: prob.fila, colID: prob.colID + 1 }); // +1 para GAS (1-based)
  }

  var limpiados = 0;
  var hojasConProblema = Object.keys(porHoja);
  for (var k = 0; k < hojasConProblema.length; k++) {
    var nombreHoja  = hojasConProblema[k];
    var hoja        = ss.getSheetByName(nombreHoja);
    if (!hoja) continue;
    var celdas = porHoja[nombreHoja];
    for (var m = 0; m < celdas.length; m++) {
      var celda = hoja.getRange(celdas[m].fila, celdas[m].colID);
      celda.clearContent();
      celda.setBackground('#ff9800'); // naranja = necesita revisarse
      limpiados++;
    }
  }

  ui.alert(
    '✅ Limpieza completada',
    limpiados + ' celdas de ID borradas y marcadas en naranja.\n' +
    'Busca las celdas naranjas para asignar el ID correcto.',
    ui.ButtonSet.OK
  );
}

// ---------------------------------------------------------------------------
// 4. REPORTE DE REGISTROS SIN ID
// ---------------------------------------------------------------------------

/**
 * Genera la hoja "📋 Sin ID" con todas las filas que tienen nombre pero
 * no tienen ID asignado (celda vacía o que empieza con ⚠️).
 */
function reporteSinID() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  ss.toast('Buscando registros sin ID...', '📋 Sin ID', -1);

  var nombresHojas = _detectarHojasConID(ss, NOMBRE_DIRECTORIO_IDS, false);
  var filasSinID   = [];
  var conteosPorHoja = {};

  for (var h = 0; h < nombresHojas.length; h++) {
    var nombreHoja = nombresHojas[h];
    var hoja       = ss.getSheetByName(nombreHoja);
    if (!hoja || hoja.getLastRow() < 2) continue;

    var maxCol  = Math.min(hoja.getLastColumn(), 60);
    var header  = hoja.getRange(1, 1, 1, maxCol).getValues()[0];
    var colID     = _indiceColumna(header, NOMBRE_COLUMNA_ID_IDS);
    var colNombre = _indiceColumna(header, 'Nombre completo');
    var colDpi    = _indiceColumna(header, 'DPI');
    if (colID === -1 || colNombre === -1) continue;

    var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, maxCol).getValues();
    var countHoja = 0;

    for (var i = 0; i < datos.length; i++) {
      var nombre = (datos[i][colNombre] || '').toString().trim();
      var id     = (datos[i][colID]     || '').toString().trim();

      if (!nombre) continue; // fila vacía
      if (id && id.indexOf('⚠️') !== 0) continue; // tiene ID válido

      var dpi = colDpi !== -1 ? (datos[i][colDpi] || '').toString().trim() : '';
      filasSinID.push([nombreHoja, i + 2, nombre, dpi]);
      countHoja++;
    }

    if (countHoja > 0) conteosPorHoja[nombreHoja] = countHoja;
  }

  // Crear / limpiar hoja de reporte
  var nombreHojaReporte = '📋 Sin ID';
  var hojaReporte = ss.getSheetByName(nombreHojaReporte);
  if (!hojaReporte) hojaReporte = ss.insertSheet(nombreHojaReporte);
  hojaReporte.clearContents();
  hojaReporte.clearFormats();

  var headers = ['Hoja', 'Fila', 'Nombre', 'DPI'];
  var hrng = hojaReporte.getRange(1, 1, 1, headers.length);
  hrng.setValues([headers]);
  hrng.setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
  hojaReporte.setFrozenRows(1);

  if (filasSinID.length > 0) {
    hojaReporte.getRange(2, 1, filasSinID.length, headers.length).setValues(filasSinID);
    // Colorear filas alternadas por hoja para facilitar lectura
    hojaReporte.getRange(2, 1, filasSinID.length, headers.length).setBackground('#fff8e1');
    hojaReporte.autoResizeColumns(1, headers.length);
  } else {
    hojaReporte.getRange(2, 1).setValue('✅ Todos los registros con nombre tienen ID asignado');
  }

  ss.toast('Listo', '📋 Sin ID', 3);

  // Construir resumen por hoja
  var resumen = '';
  var hojasConSinID = Object.keys(conteosPorHoja);
  for (var r = 0; r < hojasConSinID.length; r++) {
    resumen += '  • ' + hojasConSinID[r] + ': ' + conteosPorHoja[hojasConSinID[r]] + '\n';
  }

  ui.alert(
    '📋 Registros sin ID',
    'Total sin ID: ' + filasSinID.length + '\n\n' +
    (resumen ? 'Por hoja:\n' + resumen : '✅ Ninguno') + '\n' +
    'Ver hoja "' + nombreHojaReporte + '" para el detalle.',
    ui.ButtonSet.OK
  );
}

// ---------------------------------------------------------------------------
// 5. PRUEBA DE RENDIMIENTO
// ---------------------------------------------------------------------------

/**
 * Mide el tiempo de las operaciones principales del sistema y genera
 * un reporte visual en la hoja "⚡ Rendimiento".
 */
function pruebaRendimiento() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  ss.toast('Ejecutando prueba de rendimiento...', '⚡', -1);

  var resultados = [];
  var inicio; var duracion;

  // ── 1. Carga del directorio ───────────────────────────────────────────────
  inicio = new Date().getTime();
  var dir = _cargarDirectorio(ss);
  duracion = new Date().getTime() - inicio;
  resultados.push({
    paso: '1. Carga del directorio',
    detalle: dir.totalGrads + ' filas en ' + NOMBRE_DIRECTORIO_IDS,
    ms: duracion
  });

  // ── 2. Detección de hojas con ID (forzar refresh) ─────────────────────────
  inicio = new Date().getTime();
  var hojas = _detectarHojasConID(ss, NOMBRE_DIRECTORIO_IDS, true);
  duracion = new Date().getTime() - inicio;
  resultados.push({
    paso: '2. Detección de hojas con ID',
    detalle: hojas.length + ' hojas detectadas: ' + hojas.join(', '),
    ms: duracion
  });

  // ── 3. Lectura de todas las hojas con getDataRange ────────────────────────
  inicio = new Date().getTime();
  var totalFilas = 0;
  for (var h = 0; h < hojas.length; h++) {
    var hoja = ss.getSheetByName(hojas[h]);
    if (hoja && hoja.getLastRow() > 1) {
      var datos = hoja.getDataRange().getValues();
      totalFilas += datos.length;
    }
  }
  duracion = new Date().getTime() - inicio;
  resultados.push({
    paso: '3. Lectura de todas las hojas',
    detalle: totalFilas + ' filas totales en ' + hojas.length + ' hojas',
    ms: duracion
  });

  // ── 4. Construcción del mapa de búsqueda ──────────────────────────────────
  inicio = new Date().getTime();
  var ids = Object.keys(dir.mapaID);
  var mapaInvertido = {};
  for (var i = 0; i < ids.length; i++) mapaInvertido[ids[i].toLowerCase()] = ids[i];
  duracion = new Date().getTime() - inicio;
  resultados.push({
    paso: '4. Construcción del mapa de búsqueda',
    detalle: ids.length + ' IDs indexados',
    ms: duracion
  });

  // ── 5. 100 comparaciones de similitud ─────────────────────────────────────
  inicio = new Date().getTime();
  var nombres = Object.values ? Object.values(dir.mapaID) : ids.map(function(k) { return dir.mapaID[k]; });
  var len = nombres.length;
  for (var c = 0; c < 100; c++) {
    var a = nombres[c % len] || 'Ana García López';
    var b = nombres[(c + 1) % len] || 'Ana Garcia Lopez';
    similitudNombre(a, b);
  }
  duracion = new Date().getTime() - inicio;
  resultados.push({
    paso: '5. 100 comparaciones de similitud',
    detalle: '~' + Math.round(duracion / 100) + ' ms por comparación',
    ms: duracion
  });

  // ── 6. Conteo de triggers instalados ──────────────────────────────────────
  inicio = new Date().getTime();
  var triggers       = ScriptApp.getProjectTriggers();
  var triggersOnEdit = triggers.filter(function(t) {
    return t.getEventType() === ScriptApp.EventType.ON_EDIT;
  });
  duracion = new Date().getTime() - inicio;
  resultados.push({
    paso: '6. Conteo de triggers instalados',
    detalle: triggers.length + ' triggers totales, ' + triggersOnEdit.length + ' onEdit',
    ms: duracion
  });

  // ── Generar hoja de resultados ────────────────────────────────────────────
  var nombreHoja = '⚡ Rendimiento';
  var hojaR = ss.getSheetByName(nombreHoja);
  if (!hojaR) hojaR = ss.insertSheet(nombreHoja);
  hojaR.clearContents();
  hojaR.clearFormats();

  var headers = ['Paso', 'Detalle', 'Tiempo (ms)', 'Estado'];
  var hrng = hojaR.getRange(1, 1, 1, headers.length);
  hrng.setValues([headers]);
  hrng.setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  hojaR.setFrozenRows(1);

  var tiempoTotal = 0;
  for (var r = 0; r < resultados.length; r++) {
    var res  = resultados[r];
    var fila = r + 2;
    tiempoTotal += res.ms;

    var estado; var colorFondo;
    if (res.ms < 2000)      { estado = '🟢 Rápido';    colorFondo = '#e8f5e9'; }
    else if (res.ms < 5000) { estado = '🟡 Aceptable'; colorFondo = '#fff9c4'; }
    else                    { estado = '🔴 Lento';     colorFondo = '#ffebee'; }

    hojaR.getRange(fila, 1).setValue(res.paso);
    hojaR.getRange(fila, 2).setValue(res.detalle);
    hojaR.getRange(fila, 3).setValue(res.ms);
    hojaR.getRange(fila, 4).setValue(estado);
    hojaR.getRange(fila, 1, 1, 4).setBackground(colorFondo);
  }

  // Fila de total
  var filaTotal = resultados.length + 2;
  hojaR.getRange(filaTotal, 1).setValue('TOTAL');
  hojaR.getRange(filaTotal, 3).setValue(tiempoTotal);
  var veredicto;
  var colorVeredicto;
  if (tiempoTotal < 10000)      { veredicto = '🟢 Rápido (<10s)';    colorVeredicto = '#1b5e20'; }
  else if (tiempoTotal < 20000) { veredicto = '🟡 Aceptable (<20s)'; colorVeredicto = '#f57f17'; }
  else                          { veredicto = '🔴 Lento (>20s)';     colorVeredicto = '#b71c1c'; }
  hojaR.getRange(filaTotal, 4).setValue(veredicto);
  hojaR.getRange(filaTotal, 1, 1, 4)
    .setBackground(colorVeredicto).setFontColor('#ffffff').setFontWeight('bold');

  hojaR.autoResizeColumns(1, headers.length);
  ss.toast('Listo', '⚡ Rendimiento', 3);

  // Alerta si hay más de 1 trigger onEdit (causa diálogos dobles)
  var alertaTriggers = triggersOnEdit.length > 1
    ? '\n\n⚠️ ATENCIÓN: Tienes ' + triggersOnEdit.length + ' triggers onEdit instalados.\n' +
      'Esto puede causar que los diálogos se abran dos veces.\n' +
      'Ejecuta configurarEditTrigger() para corregirlo.'
    : '';

  ui.alert(
    '⚡ Resultado de Rendimiento',
    'Tiempo total: ' + tiempoTotal + ' ms\n' +
    'Veredicto: ' + veredicto + '\n\n' +
    'Triggers onEdit: ' + triggersOnEdit.length +
    alertaTriggers + '\n\n' +
    'Ver hoja "' + nombreHoja + '" para el detalle.',
    ui.ButtonSet.OK
  );
}
