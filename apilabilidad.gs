/**
 * HOJAS DE APILABILIDAD - Equipito Empleabilidad
 *
 * Instrucciones:
 * 1. Abre tu Google Sheet
 * 2. Ve a Extensiones > Apps Script
 * 3. Pega este código y ejecuta: configurarHojasApilabilidad()
 *
 * Columnas comunes en TODAS las hojas (Información principal):
 *   Fecha de ingreso | Nombre completo | ID | Género | Edad | Nivel educativo
 *
 * Hojas que se crean:
 *   1. Aliados
 *   2. Plataforma
 *   3. Derivaciones
 *   4. Por su cuenta
 *   5. No busca trabajo - Fito
 *   6. Activamente busca trabajo
 */

// ─────────────────────────────────────────────
// CONFIGURACIÓN DE COLUMNAS POR HOJA
// ─────────────────────────────────────────────

var COLUMNAS_COMUNES = [
  'Fecha de ingreso',
  'Nombre completo',
  'ID',
  'Género',
  'Edad',
  'Nivel educativo'
];

var CONFIGURACION_HOJAS = [
  {
    nombre: 'Aliados',
    columnas: [
      ...COLUMNAS_COMUNES,
      'Pendiente (nombre por definir)',  // Desplegable sí/no — nombre pendiente
      'Trámites',
      'Entrevista',
      'Día de prueba',
      'Confirmación de recepción',
      'Notas',
      'Activo'
    ],
    desplegables: {
      'Pendiente (nombre por definir)': ['Sí', 'No'],
      'Activo': ['Sí', 'No']
    },
    checkboxes: []
  },
  {
    nombre: 'Plataforma',
    columnas: [
      ...COLUMNAS_COMUNES,
      'Cita',
      'Creación de perfil',
      'Contacto',
      'Trámites',
      'Entrevista',
      'Confirmación',
      'Recepción',
      'Nota',
      'Activo'
    ],
    desplegables: {
      'Activo': ['Sí', 'No']
    },
    checkboxes: ['Cita', 'Contacto']
  },
  {
    nombre: 'Derivaciones',
    columnas: [
      ...COLUMNAS_COMUNES,
      'Envío 1',
      'Llamada 1',
      'Envío 2',
      'Llamada 2',
      'Envío 3',
      'Llamada 3',
      'Notas',
      'Activo'
    ],
    desplegables: {
      'Activo': ['Sí', 'No']
    },
    checkboxes: []
  },
  {
    nombre: 'Por su cuenta',
    columnas: [
      ...COLUMNAS_COMUNES,
      'Mensaje',
      'Llamada',
      'Nota',
      'Activo'
    ],
    desplegables: {
      'Activo': ['Sí', 'No']
    },
    checkboxes: []
  },
  {
    nombre: 'No busca trabajo - Fito',
    columnas: [
      ...COLUMNAS_COMUNES,
      'DPI',
      'Número de teléfono',
      'Formación',
      'Cohorte',
      'Nota',
      'Activo'
    ],
    desplegables: {
      'Activo': ['Sí', 'No']
    },
    checkboxes: []
  },
  {
    nombre: 'Activamente busca trabajo',
    columnas: [
      ...COLUMNAS_COMUNES,
      'Entrevista',
      'Trámites',
      'Activo'
    ],
    desplegables: {
      'Activo': ['Sí', 'No']
    },
    checkboxes: []
  }
];

// ─────────────────────────────────────────────
// FUNCIÓN PRINCIPAL — EJECUTAR ESTA
// ─────────────────────────────────────────────

function configurarHojasApilabilidad() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Eliminar todas las hojas existentes
  eliminarTodasLasHojas(ss);

  // 2. Crear las hojas nuevas con sus columnas
  CONFIGURACION_HOJAS.forEach(function(config, index) {
    var hoja = ss.insertSheet(config.nombre, index);
    configurarHoja(hoja, config);
  });

  SpreadsheetApp.getUi().alert('✅ Hojas de apilabilidad creadas correctamente.');
}

// ─────────────────────────────────────────────
// ELIMINAR TODAS LAS HOJAS EXISTENTES
// ─────────────────────────────────────────────

function eliminarTodasLasHojas(ss) {
  var hojas = ss.getSheets();

  // Primero insertar una hoja temporal para no quedarnos sin hojas
  var temporal = ss.insertSheet('__temporal__');

  // Eliminar todas las hojas originales
  hojas.forEach(function(hoja) {
    ss.deleteSheet(hoja);
  });

  // La temporal se elimina al final cuando ya existan las nuevas hojas.
  // Se llama eliminarHojaTemporal() al terminar configurarHoja() en el último ciclo.
}

// ─────────────────────────────────────────────
// CONFIGURAR UNA HOJA INDIVIDUAL
// ─────────────────────────────────────────────

function configurarHoja(hoja, config) {
  var columnas = config.columnas;
  var numColumnas = columnas.length;

  // ── Encabezados ──
  var rangoEncabezados = hoja.getRange(1, 1, 1, numColumnas);
  rangoEncabezados.setValues([columnas]);

  // Estilo de encabezados
  rangoEncabezados
    .setBackground('#1a237e')       // Azul oscuro
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  hoja.setFrozenRows(1);  // Fijar fila de encabezados

  // ── Validaciones: desplegables sí/no ──
  Object.keys(config.desplegables).forEach(function(nombreColumna) {
    var opciones = config.desplegables[nombreColumna];
    var colIndex = columnas.indexOf(nombreColumna) + 1;
    if (colIndex === 0) return;

    var rango = hoja.getRange(2, colIndex, 999);
    var regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(opciones, true)
      .setAllowInvalid(false)
      .build();
    rango.setDataValidation(regla);
  });

  // ── Validaciones: checkboxes ──
  config.checkboxes.forEach(function(nombreColumna) {
    var colIndex = columnas.indexOf(nombreColumna) + 1;
    if (colIndex === 0) return;

    var rango = hoja.getRange(2, colIndex, 999);
    var regla = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    rango.setDataValidation(regla);
  });

  // ── Columna ID: fórmula automática ──
  var idIndex = columnas.indexOf('ID') + 1;
  if (idIndex > 0) {
    // Fórmula que genera ID automático si hay nombre completo
    var nombreIndex = columnas.indexOf('Nombre completo') + 1;
    if (nombreIndex > 0) {
      var nombreLetra = columnIndexToLetter(nombreIndex);
      for (var fila = 2; fila <= 100; fila++) {
        hoja.getRange(fila, idIndex).setFormula(
          '=IF(' + nombreLetra + fila + '<>"","' +
          config.nombre.substring(0, 3).toUpperCase() + '-"&TEXT(ROW()-1,"000"),"")'
        );
      }
    }
  }

  // ── Columna Fecha de ingreso: formato fecha ──
  var fechaIndex = columnas.indexOf('Fecha de ingreso') + 1;
  if (fechaIndex > 0) {
    hoja.getRange(2, fechaIndex, 999)
      .setNumberFormat('dd/mm/yyyy');
  }

  // ── Ancho de columnas ──
  hoja.setColumnWidth(columnas.indexOf('Nombre completo') + 1, 200);
  hoja.setColumnWidth(columnas.indexOf('Notas') + 1 || columnas.indexOf('Nota') + 1, 250);

  // ── Colores alternos en filas de datos ──
  var bandingRange = hoja.getRange(2, 1, 999, numColumnas);
  try {
    bandingRange.applyRowBanding(
      SpreadsheetApp.BandingTheme.BLUE,
      false,
      false
    );
  } catch (e) {
    // Si ya tiene banding, ignorar
  }

  // ── Ajustar el ancho general de columnas ──
  for (var col = 1; col <= numColumnas; col++) {
    if (hoja.getColumnWidth(col) < 120) {
      hoja.setColumnWidth(col, 120);
    }
  }

  // ── Eliminar hoja temporal si es la última hoja creada ──
  var ss = hoja.getParent();
  var temporal = ss.getSheetByName('__temporal__');
  if (temporal) {
    ss.deleteSheet(temporal);
  }
}

// ─────────────────────────────────────────────
// UTILIDAD: número de columna → letra (A, B, ... Z, AA...)
// ─────────────────────────────────────────────

function columnIndexToLetter(index) {
  var letra = '';
  while (index > 0) {
    var rem = (index - 1) % 26;
    letra = String.fromCharCode(65 + rem) + letra;
    index = Math.floor((index - 1) / 26);
  }
  return letra;
}
