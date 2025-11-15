//Funcion para armado de HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

  // Función para guardar el log del mail que ejecuta la app - En desarrollo
function userLog(user){
  Logger.log("El usuario: " + user + " Ha inciado sesión");
}

// Obtener datos de hoja de Sheet y pasar a JSON
function obtenerDatos(sheetId, sheetName) {
  if (!sheetId || !sheetName) {
    throw new Error('Configuración no inicializada');
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('No se encontró la hoja: ' + sheetName);
    }

    const data = sheet.getDataRange().getValues();

    if (data.length === 0) {
      Logger.log('La hoja está vacía o solo tiene encabezados');
      return [];
    }

    const headers = data[0];
    const jsonData = [];

    // VALIDAR QUE EXISTAN LOS ENCABEZADOS REQUERIDOS
    const encabezadosRequeridos = ['Nombre', 'Apellido', 'Sección', 'DNI', "Merged Doc URL"];
    const encabezadosFaltantes = [];
    
    encabezadosRequeridos.forEach(function(encabezado) {
      if (headers.indexOf(encabezado) === -1) {
        encabezadosFaltantes.push(encabezado);
      }
    });
    
    if (encabezadosFaltantes.length > 0) {
      throw new Error('Faltan los siguientes encabezados en la hoja: ' + encabezadosFaltantes.join(', '));
    }

    // Mapear los índices de las columnas una sola vez
    const colIndices = {
      nombre: headers.indexOf('Nombre'),
      apellido: headers.indexOf('Apellido'),
      seccion: headers.indexOf('Sección'),
      dni: headers.indexOf('DNI'),
      docUrl: headers.findIndex(h => 
        typeof h === 'string' && h.toLowerCase().includes('merged doc url')
      ),
      linkToMergedDoc: headers.findIndex(h => 
        typeof h === 'string' && h.toLowerCase().includes('link to merged doc')
      ),
      enviado: headers.indexOf('Enviado'),
      fechaEnvio: headers.indexOf('Fecha Envío'),
      postId: headers.indexOf('Post ID')
    };

    // Procesar las filas (saltar encabezado)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Saltar filas vacías
      if (!row[colIndices.nombre] && !row[colIndices.apellido]) {
        continue;
      }

      const alumno = {
        nombre: row[colIndices.nombre] || '',
        apellido: row[colIndices.apellido] || '',
        seccion: row[colIndices.seccion] || '',
        dni: extraerNumeros(colIndices.dni !== -1 && row[colIndices.dni] ? String(row[colIndices.dni]) : ''),
        docUrl: colIndices.docUrl !== -1 ? (row[colIndices.docUrl] || '') : '',
        linkToMergedDoc: colIndices.linkToMergedDoc !== -1 ? (row[colIndices.linkToMergedDoc] || '') : '',
        mergedDocUrl: colIndices.docUrl !== -1 ? (row[colIndices.docUrl] || '') : (colIndices.linkToMergedDoc !== -1 ? (row[colIndices.linkToMergedDoc] || '') : ''),
        enviado: colIndices.enviado !== -1 ? (row[colIndices.enviado] || false) : false,
        fechaEnvio: colIndices.fechaEnvio !== -1 ? (row[colIndices.fechaEnvio] || '') : '',
        postId: colIndices.postId !== -1 ? (row[colIndices.postId] || '') : ''
      };

      jsonData.push(alumno);
    }

    Logger.log('Datos obtenidos: ' + jsonData.length + ' alumnos');
    return jsonData;

  } catch (error) {
    Logger.log('Error en obtenerDatos: ' + error.toString());
    throw error;
  }
}


// Obtener datos de hoja de Sheet y pasar a JSON
function obtenerHistorial(sheetId, sheetName) {
  if (!sheetId || !sheetName) {
    throw new Error('Configuración no inicializada');
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('No se encontró la hoja: ' + sheetName);
    }

    const data = sheet.getDataRange().getValues();

    if (data.length === 0) {
      Logger.log('La hoja está vacía o solo tiene encabezados');
      return [];
    }

    const headers = data[0];
    const jsonData = [];

    // Mapear los índices de las columnas una sola vez
    const colIndices = {
      nombre: headers.indexOf('Nombre'),
      apellido: headers.indexOf('Apellido'),
      seccion: headers.indexOf('Sección'),
      dni: headers.indexOf('DNI'),
      docUrl: headers.indexOf('PDF URL'),
      enviado: headers.indexOf('Fecha Envío'),
      postId: headers.indexOf('Post ID'),
      postURL: headers.indexOf("Post URL"),
      asunto: headers.indexOf('Asunto'),
    };

    // Procesar las filas (saltar encabezado)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Saltar filas vacías
      if (!row[colIndices.nombre] && !row[colIndices.apellido]) {
        continue;
      }

      const alumno = {
        nombre: row[colIndices.nombre] || '',
        apellido: row[colIndices.apellido] || '',
        seccion: row[colIndices.seccion] || '',
        dni: extraerNumeros(colIndices.dni !== -1 && row[colIndices.dni] ? String(row[colIndices.dni]) : ''),
        docUrl: colIndices.docUrl !== -1 ? (row[colIndices.docUrl] || '') : '',
        enviado: colIndices.enviado !== -1 ? (row[colIndices.enviado] || '') : '',
        postId: colIndices.postId !== -1 ? (row[colIndices.postId] || '') : '',
        postURL: colIndices.postURL !== -1 ? (row[colIndices.postURL] || '') : '',
        asunto: colIndices.asunto !== -1 ? (row[colIndices.asunto] || '') : '',
      };

      jsonData.push(alumno);
    }

    Logger.log('Datos obtenidos: ' + jsonData.length + ' alumnos');
    return jsonData;

  } catch (error) {
    Logger.log('Error en obtenerDatos: ' + error.toString());
    throw error;
  }
}
// Copia los datos a la hoja en donde se ejecuta
function copyData(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data') || ss.getActiveSheet();

  try {
    if (datos && datos.length > 0) {
      // Convertir objetos JSON a array bidimensional
      const headers = ['Nombre', 'Apellido', 'Sección', 'DNI', 'Merged Doc URL', 'Link to Merged Doc', 'Enviado', 'Fecha Envío', 'Post ID'];
      
      const rows = datos.map(alumno => [
        alumno.nombre || '',
        alumno.apellido || '',
        alumno.seccion || '',
        alumno.dni || '',
        alumno.mergedDocUrl || '',
        alumno.linkToMergedDoc || '',
        alumno.enviado || false,
        alumno.fechaEnvio || '',
        alumno.postId || ''
      ]);
      
      // Combinar headers y datos
      const dataParaCopiar = [headers, ...rows];
      
      // Limpiar la hoja antes de copiar
      sheet.clear();
      
      // Escribir los datos en la hoja actual
      sheet.getRange(1, 1, dataParaCopiar.length, dataParaCopiar[0].length).setValues(dataParaCopiar);
      
      // Formatear headers
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
      
      SpreadsheetApp.getUi().alert('Datos copiados exitosamente: ' + datos.length + ' alumnos');
    } else {
      throw new Error('No hay datos para copiar');
    }
    
    return sheet;
    
  } catch(error) {
    Logger.log("Error en copiarData: " + error.toString());
    SpreadsheetApp.getUi().alert('Error al copiar los datos: ' + error.toString());
    throw new Error('Error al copiar los datos: ' + error.toString());
  }
}

// Genera la vista "tabla" con los datos
function generarVista(datos, historial, vista) {
  try {
    const html = HtmlService.createTemplateFromFile('view/' + vista);
    html.datosJSON = JSON.stringify(datos);
    html.historyDataJSON = JSON.stringify(historial);
    
    const output = html.evaluate()
      .setTitle("ECBox apps")
      .setWidth(1800)
      .setHeight(1000);

    SpreadsheetApp.getUi().showModalDialog(output, "ECBox apps");
    
  } catch (error) {
    Logger.log('Error en generarVista: ' + error.toString());
  }
}

/**
 * Extrae solo los números de un string y los devuelve formateados
 */
function extraerNumeros(data) {
  let dni = data.replace(/\D/g, '');
  let numero = parseInt(dni, 10);

  if (isNaN(numero)) return '';

  return numero.toLocaleString('es-AR');
}

/**
 * Registra las respuestas de las APIs en la hoja de Log del cliente
 */
function clientAPILog(response, apiEndpoint, requestData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('API_Log'); 
  
  if (!sheet) {
    sheet = ss.insertSheet('API_Log');
    sheet.appendRow([
      'Fecha/Hora',
      'Endpoint',
      'Status Code',
      'Status',
      'Post ID',
      'Post URL',
      'Student ID',
      'Subject',
      'Link/File',
      'Error Message'
    ]);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#ea4335').setFontColor('#ffffff');
  }
  
  const timestamp = new Date();
  const row = [];
  
  try {
    if (!response) {
      row.push(timestamp, apiEndpoint, 'N/A', 'Error', '', '', 
               requestData ? requestData.uid || '' : '', 
               requestData ? requestData.subject || '' : '', '',
               'No se recibió respuesta de la API');
    } else if (response.status === 'success' && response.post_id) {
      row.push(timestamp, apiEndpoint, 201, 'Success', 
               response.post_id, response.post_url || '',
               requestData ? requestData.uid || '' : '',
               requestData ? requestData.subject || '' : '',
               requestData ? (requestData.report_pdf || '') : '', '');
    } else if (response.status === 'error') {
      row.push(timestamp, apiEndpoint, response.error_code || 400, 'Error', '', '',
               requestData ? requestData.uid || '' : '',
               requestData ? requestData.subject || '' : '', '',
               response.error_message || 'Error desconocido');
    }
    
    if (row.length > 0) {
      sheet.appendRow(row);
      
      // Colorear según status
      const lastRow = sheet.getLastRow();
      const statusCell = sheet.getRange(lastRow, 4);
      if (row[3] === 'Success') {
        statusCell.setBackground('#d9ead3');
      } else if (row[3] === 'Error') {
        statusCell.setBackground('#f4cccc');
      }
    }
    
  } catch (error) {
    Logger.log('Error en clientAPILog: ' + error.toString());
  }
}

/**
 * Extrae el ID de un archivo de Drive desde una URL o ID directo
 */
function extraerDriveFileId(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') return null;
  
  urlOrId = urlOrId.trim();
  
  if (/^[a-zA-Z0-9_-]{25,50}$/.test(urlOrId)) {
    return urlOrId;
  }
  
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/open\?id=([a-zA-Z0-9_-]+)/,
    /\/document\/d\/([a-zA-Z0-9_-]+)/,
    /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/,
    /\/presentation\/d\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/
  ];
  
  for (let pattern of patterns) {
    const match = urlOrId.match(pattern);
    if (match && match[1]) return match[1];
  }
  
  return null;
}

/**
 * Extrae el ID de un spreadsheet desde una URL o ID directo
 */
function extraerSpreadsheetId(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') return null;
  
  urlOrId = urlOrId.trim();
  
  if (/^[a-zA-Z0-9_-]{40,50}$/.test(urlOrId)) {
    return urlOrId;
  }
  
  const pattern = /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/;
  const match = urlOrId.match(pattern);
  
  if (match && match[1]) return match[1];
  
  return null;
}

/**
 * Valida que un ID de Drive sea accesible
 */
function validarDriveFileId(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    
    return {
      valid: true,
      fileName: file.getName(),
      mimeType: file.getMimeType(),
      size: file.getSize(),
      url: file.getUrl(),
      id: fileId
    };
  } catch (error) {
    return {
      valid: false,
      error: 'Archivo no encontrado o sin permisos'
    };
  }
}

/**
 * Actualiza el estado de envío en la hoja Data
 */
function actualizarEstadoEnvio(dni, datosEnvio) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Data') || ss.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const colIndices = {
      dni: headers.indexOf('DNI'),
      enviado: headers.indexOf('Enviado'),
      fechaEnvio: headers.indexOf('Fecha Envío'),
      postId: headers.indexOf('Post ID')
    };
    
    if (colIndices.dni === -1) return false;
    
    for (let i = 1; i < data.length; i++) {
      const rowDni = extraerNumeros(String(data[i][colIndices.dni] || ''));
      
      if (rowDni === dni) {
        if (colIndices.enviado !== -1) {
          sheet.getRange(i + 1, colIndices.enviado + 1).setValue(true);
        }
        if (colIndices.fechaEnvio !== -1 && datosEnvio.fechaEnvio) {
          sheet.getRange(i + 1, colIndices.fechaEnvio + 1).setValue(datosEnvio.fechaEnvio);
        }
        if (colIndices.postId !== -1 && datosEnvio.postId) {
          sheet.getRange(i + 1, colIndices.postId + 1).setValue(datosEnvio.postId);
        }
        
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    Logger.log('Error en actualizarEstadoEnvio: ' + error.toString());
    return false;
  }
}

/**
 * Mueve registros enviados exitosamente a la hoja de historial
 * @param {Array} alumnosEnviados - Array de alumnos que fueron enviados exitosamente
 * @return {Object} Resultado de la operación
 */
function moverAHistorial(alumnosEnviados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historialSheet = ss.getSheetByName('Historial');
    
    // Crear hoja de historial si no existe
    if (!historialSheet) {
      historialSheet = ss.insertSheet('Historial');
      const headers = ['Nombre', 'Apellido', 'Sección', 'DNI', 'PDF URL', 'Fecha Envío', 'Post ID', 'Post URL', 'Asunto'];
      historialSheet.appendRow(headers);
      historialSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    }
    
    // Preparar datos para insertar
    const rows = alumnosEnviados.map(function(alumno) {
      return [
        alumno.nombre || '',
        alumno.apellido || '',
        alumno.seccion || '',
        alumno.dni || '',
        alumno.linkPDF || '',
        alumno.fechaEnvio || new Date().toLocaleString('es-AR'),
        alumno.postId || '',
        alumno.postUrl || '',
        alumno.subject || ''
      ];
    });
    
    // Agregar todos los registros de una vez
    if (rows.length > 0) {
      historialSheet.getRange(historialSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      Logger.log('Movidos ' + rows.length + ' registros al historial');
    }
    
    return {
      registros: rows.length
    };
    
  } catch (error) {
    Logger.log('Error en moverAHistorial: ' + error.toString());
    return {
      error: error.toString()
    };
  }
}

/**
 * Limpia los estados de envío DESPUÉS de moverlos al historial
 */
function cleanRegister() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    
    if (!sheet) {
      return { success: false, error: 'Hoja "Data" no encontrada' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, registros: 0, historial: 0 };
    }

    const headers = data[0];
    const colIndices = {
      nombre: headers.indexOf('Nombre'),
      apellido: headers.indexOf('Apellido'),
      seccion: headers.indexOf('Sección'),
      dni: headers.indexOf('DNI'),
      linkPDF: Math.max(headers.indexOf('Merged Doc URL'), headers.indexOf('Link to Merged Doc')),
      enviado: headers.indexOf('Enviado'),
      fechaEnvio: headers.indexOf('Fecha Envío'),
      postId: headers.indexOf('Post ID')
    };

    // Recopilar enviados para mover al historial
    const alumnosEnviados = [];
    const filasAEliminar = []; // Guardar índices de filas a eliminar
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (colIndices.enviado !== -1 && row[colIndices.enviado] === true) {
        alumnosEnviados.push({
          nombre: row[colIndices.nombre] || '',
          apellido: row[colIndices.apellido] || '',
          seccion: row[colIndices.seccion] || '',
          dni: extraerNumeros(String(row[colIndices.dni] || '')),
          linkPDF: colIndices.linkPDF !== -1 ? (row[colIndices.linkPDF] || '') : '',
          fechaEnvio: colIndices.fechaEnvio !== -1 ? (row[colIndices.fechaEnvio] || '') : '',
          postId: colIndices.postId !== -1 ? (row[colIndices.postId] || '') : '',
          postUrl: '',
          subject: ''
        });
        
        // Guardar el índice de la fila (ajustado para 1-based index de Sheets)
        filasAEliminar.push(i + 1);
      }
    }

    // Mover al historial
    if (alumnosEnviados.length > 0) {
      moverAHistorial(alumnosEnviados);
      
      // Ordenar de mayor a menor para eliminar desde el final
      filasAEliminar.sort(function(a, b) { return b - a; });
      
      for (let i = 0; i < filasAEliminar.length; i++) {
        sheet.deleteRow(filasAEliminar[i]);
      }
      
      Logger.log('Eliminadas ' + filasAEliminar.length + ' filas de Data');
    }

    return {
      success: true,
      registros: filasAEliminar.length,
      historial: alumnosEnviados.length
    };

  } catch (error) {
    Logger.log('Error en cleanRegister: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
