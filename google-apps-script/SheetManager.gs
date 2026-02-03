/**
 * MÓDULO: GESTOR DE HOJAS DE CÁLCULO
 *
 * Maneja todas las operaciones relacionadas con las hojas de Google Sheets:
 * - Creación y configuración de hojas
 * - Distribución de datos según el flujo
 * - Actualización de registros
 */

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
    'Habilidades': 'Habilidades',
    'GlobalGorros': 'GlobalGorros',
    'Conexión': 'Conexión/Relación',
    'Por su Cuenta': 'Por su Cuenta',
    'No Hace Falta': 'No Hace Falta',
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
    'Habilidades': ['Habilidad Principal', 'Nivel', 'Certificación', 'Notas'],
    'GlobalGorros': ['Proyecto Asignado', 'Fecha Inicio', 'Responsable', 'Notas'],
    'Conexión': ['Líder Asignado', 'Tipo Relación', 'Fecha Conexión', 'Notas'],
    'Por su Cuenta': ['Actividad', 'Progreso', 'Última Actualización', 'Notas'],
    'No Hace Falta': ['Razón', 'Fecha', 'Observaciones', 'Notas'],
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
    'Habilidades': '#ea4335',
    'GlobalGorros': '#fbbc04',
    'Conexión': '#34a853',
    'Por su Cuenta': '#4285f4',
    'No Hace Falta': '#9e9e9e',
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

/**
 * Exporta los datos de una hoja a CSV
 * @param {string} nombreHoja - Nombre de la hoja a exportar
 * @return {string} Contenido en formato CSV
 */
function exportarHojaACSV(nombreHoja) {
  const hoja = obtenerHoja(nombreHoja);
  const datos = hoja.getDataRange().getValues();

  return datos.map(fila =>
    fila.map(celda => `"${celda}"`).join(',')
  ).join('\n');
}
