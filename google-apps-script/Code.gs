/**
 * SISTEMA DE SEGUIMIENTO DE GRADUADOS - EQUIPITO EMPLEABILIDAD
 *
 * Este script automatiza el seguimiento de graduados desde KoboToolbox
 * hacia Google Sheets con clasificación automática según el flujo de trabajo.
 *
 * Autor: Sistema Automatizado
 * Fecha: 2026-02-03
 */

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
                                '• Habilidades\n' +
                                '• GlobalGorros\n' +
                                '• Conexión/Relación\n' +
                                '• Por su Cuenta\n' +
                                '• No Hace Falta\n' +
                                '• Empleados\n' +
                                '• Seguimientos (llamadas programadas)\n' +
                                '• Reportes\n' +
                                '• Configuración',
                                ui.ButtonSet.YES_NO);

    if (respuesta !== ui.Button.YES) {
      return;
    }

    // Crear todas las hojas necesarias
    const hojasNecesarias = [
      'Graduados',
      'Habilidades',
      'GlobalGorros',
      'Conexión/Relación',
      'Por su Cuenta',
      'No Hace Falta',
      'Empleados',
      'Seguimientos',
      'Reportes',
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
