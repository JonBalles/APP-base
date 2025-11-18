function onOpen() {
  // Creamos menú base
  const ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("☁️ Menú");
  menu.addItem('📄 Acerca','about');
  menu.addItem('▶️ Iniciar app', 'showApp');
  menu.addToUi();
}

function about(){
  generarVista(null, null, "about");
}

//Mostrar vista con los datos de la hoja actual
function showApp(userMail) {
  // Función para guardar el log del mail que ejecuta la app - En desarrollo
  userLog(userMail);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  try {
    const datos = obtenerDatos(spreadsheetId, sheetName);
    const historial = obtenerHistorial(spreadsheetId, "Historial");

    generarVista(datos, historial, "main");
  } catch (error) {
    Logger.log("Error en showTable: " + error.toString());
    SpreadsheetApp.getUi().alert('Error al abrir la tabla: ' + error.toString());
  }
}

// Toma los datos de un sheet y las devuelve en la hoja actual
function getData(spreadsheetId, sheetName) {
  if (!spreadsheetId || !sheetName) {
    throw new Error('Debes proporcionar el ID del spreadsheet y el nombre de la hoja');
  }

  try{
    const datos = obtenerDatos(spreadsheetId, sheetName);
    const historial = obtenerHistorial(spreadsheetId, "Historial");
    copyData(datos);
    generarVista(datos, historial, "main");
  } catch(error){
    Logger.log("Error en generarVista: " + error.toString());
    throw new Error('Error al generar la vista: ' + error.toString());
  }
}

// Toma los datos de un sheet y las devuelve en la hoja actual
function modalData(linkOrId, sheetName) {
  try {
    let sheetId = extraerSpreadsheetId(linkOrId);
    
    if (!sheetId) {
      throw new Error('No se pudo extraer el ID del spreadsheet. Verifica el link o ID proporcionado.');
    }
    
    const datos = obtenerDatos(sheetId, sheetName);
    
    if (!datos || datos.length === 0) {
      throw new Error('No se encontraron datos en la hoja especificada');
    }
    
    copyData(datos);
    
    return {
      success: true,
      message: 'Datos cargados exitosamente: ' + datos.length + ' registros',
      datos: datos,
      cantidad: datos.length
    };
    
  } catch (error) {
    return {
      success: false,
      message: error.message || 'Error al cargar los datos',
      error: error.toString()
    };
  }
}

// Toma los datos del sheet, hoja "Data" y actualiza la tabla del front
function reloadData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName('Data');
    
    if (!dataSheet) {
      return {
        success: false,
        message: 'No se encontró la hoja "Data"'
      };
    }
    
    const sheetId = ss.getId();
    const datos = obtenerDatos(sheetId, 'Data');
    
    return datos;

  } catch (error) {
    return {
      success: false,
      message: error.message || 'Error al cargar los datos'
    };
  }
}

/**
 * Envía boletines de forma masiva usando PDFs con historial
 * @param {Array} alumnos - Array de objetos con datos de alumnos
 * @param {string} audienceType - Tipo de audiencia (1-4)
 * @param {string} subject - Asunto del boletín
 * @param {string} contactDescripcion - Descripción de contacto (opcional)
 * @return {Array} Resultado de cada envío
 */
function run(alumnos, audienceType, subject, contactDescripcion) {
  const resultados = [];
  const alumnosExitosos = [];
  
  // Obtener token de autorización
  const token = getAuthorizationToken();
  
  if (!token) {
    throw new Error('No se pudo obtener el token de autorización');
  }

  // Procesar cada alumno
  alumnos.forEach(function(alumno) {
    
    try {
      // Validar que tenga PDF
      if (!alumno.linkPDF) {
        resultados.push({
          success: false,
          dni: alumno.dni,
          error: 'No tiene PDF asignado'
        });
        return;
      }
      
      // Extraer ID del archivo de Drive
      const fileId = extraerDriveFileId(alumno.linkPDF);

      if (!fileId) {
        resultados.push({
          success: false,
          dni: alumno.dni,
          error: 'URL/ID de Drive inválido'
        });
        return;
      }
      
      // Validar que el archivo exista y sea accesible
      const validacion = validarDriveFileId(fileId);
      if (!validacion.valid) {
        resultados.push({
          success: false,
          dni: alumno.dni,
          error: validacion.error
        });
        return;
      }
      
      // Enviar boletín con PDF
      const response = enviarBoletinConPDF(
        token,
        alumno.dni,
        fileId,
        subject,
        parseInt(audienceType),
        contactDescripcion || ''
      );
      
      if (response && response.post_id) {
        const fechaEnvio = new Date().toLocaleString('es-AR', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit'
        });
        
        resultados.push({
          success: true,
          dni: alumno.dni,
          postId: response.post_id,
          postUrl: response.post_url || '',
          fileName: validacion.fileName
        });
        
        // Preparar datos para historial
        alumnosExitosos.push({
          nombre: alumno.nombre,
          apellido: alumno.apellido,
          seccion: alumno.seccion,
          dni: alumno.dni,
          linkPDF: alumno.linkPDF,
          fechaEnvio: fechaEnvio,
          postId: response.post_id,
          postUrl: response.post_url || '',
          subject: subject
        });
        
        // Actualizar estado en la hoja actual
        actualizarEstadoEnvio(alumno.dni, {
          fechaEnvio: fechaEnvio,
          postId: response.post_id
        });
          Logger.log("Creando LOG en API_Log")
        clientAPILog(response, 'EnvioBoletinesPDF', {
          uid: alumno.dni,
          subject: subject,
          report_pdf: validacion.fileName
        });
        
      } else {
        const errMsg = (response && response.error_message) ? response.error_message : 'No se recibió respuesta válida de la API';

        resultados.push({
          success: false,
          dni: alumno.dni,
          error: errMsg
        });
              clientAPILog(null, 'EnvioBoletinesPDF', {
        uid: alumno.dni,
        subject: subject,
        error: errMsg
      });
      }
      
    } catch (error) {
      resultados.push({
        success: false,
        dni: alumno.dni,
        error: error.message || 'Error desconocido'
      });
      
      clientAPILog(null, 'EnvioBoletinesPDF', {
        uid: alumno.dni,
        subject: subject,
        error: error.message
      });
    }
    
    // Pausa entre envíos
    Utilities.sleep(200);
  });
  
  // Mover registros exitosos al historial
  if (alumnosExitosos.length > 0) {
    moverAHistorial(alumnosExitosos);
  }
  
  return resultados;
}

