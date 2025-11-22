/**
 * Datos de Laboratorio
 */
 const URL = "https://integraciones.handing.ar";
 const API_USER_NAME = 'informes-alumnos-beth'
 const APIKEY = 'eabcf8af2d7af48226dcd429a0d24d27ea0e4f743bcad9acd93b9717f76886ef8b0d995d3a9cc9ef4430c537512fe0c65e8b'

/**
 * Datos de Producción
 */
/*
const URL = "https://integraciones.handing.co";
const API_USER_NAME = 'gestion-alumnos-beth'
const APIKEY = 'c75bb813d6bb857315500c225d5397b3c07593f49f07879b8daf5d69b7cfe8f2bf674fd9892396bb27c66334b9d0249eb64e'
*/

// ============= FUNCIONES DE AUTENTICACIÓN =============

function getAuthorizationToken() {

  const url = URL + "/api/v2/auth/application";
  const data = {
    'api_username': API_USER_NAME,
    'api_key': APIKEY, 
  };

  const options = {
    'method': 'POST',
    'contentType': 'application/json',
    'payload': JSON.stringify(data),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    // Logger.log('Handing - authorization token response code ' + responseCode);
    
    if (responseCode === 201) {
      const responseBody = JSON.parse(response.getContentText());
      return responseBody.token;
    } else {
      const responseBody = response.getContentText();
      Logger.log(Utilities.formatString("Request failed. Expected 201, got %d: %s", responseCode, responseBody));
      return null;
    }
  } catch (error) {
    Logger.log('Error en getAuthorizationToken: ' + error.toString());
    return null;
  }
}

// ============= FUNCIONES DE ENVÍO DE BOLETINES =============

function enviarBoletinConPDF(authToken, studentId, reportFileId, subject, audienceType = 2, contactDescription = '') {
  const url = URL + "/api/v2/new-report-card-notification";
  const boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW';

  // Validar parámetros
  if (!authToken || !studentId || !reportFileId || !subject) {
    Logger.log('Error: Faltan parámetros requeridos');
    return null;
  }

  try {
    const file = DriveApp.getFileById(reportFileId);
    
    // Datos del formulario
    const metadata = {
      'uid': studentId,
      'subject': subject,
      'audience_type': audienceType.toString()
    };

    if (contactDescription) {
      metadata['contact_description'] = contactDescription;
    }

    // Construir el contenido multipart
    let data = "";
    for (let key in metadata) {
      data += "--" + boundary + "\r\n";
      data += 'Content-Disposition: form-data; name="' + key + '"\r\n\r\n';
      data += metadata[key] + "\r\n";
    }

    // Agregar el archivo
    data += "--" + boundary + "\r\n";
    data += 'Content-Disposition: form-data; name="report_pdf"; filename="' + file.getName() + '"\r\n';
    data += 'Content-Type: ' + file.getMimeType() + '\r\n\r\n';

    // Generar el payload
    const payload = Utilities.newBlob(data).getBytes()
      .concat(file.getBlob().getBytes())
      .concat(Utilities.newBlob("\r\n--" + boundary + "--").getBytes());

    const options = {
      'method': 'POST',
      'contentType': 'multipart/form-data; boundary=' + boundary,
      'headers': {
        'Authorization': 'Bearer ' + authToken
      },
      'payload': payload,
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 201) {
      const responseBody = JSON.parse(response.getContentText());
      Logger.log('Boletín enviado exitosamente. Post ID: ' + responseBody.post_id);
      return responseBody;
    } else {
      const responseBody = response.getContentText();
      Logger.log(Utilities.formatString("Error al enviar. Expected 201, got %d: %s", responseCode, responseBody));
      return null;
    }
  } catch (error) {
    Logger.log('Error en enviarBoletinConPDF: ' + error.toString());
    return null;
  }
}
