//Datos necesarios para el envio con PDF
const ejemploReal = {
    uid: "44.173.981",
    fileID: "19-4ca9klorEI94Egh-oWx-0hOd_PQdGE",
    subject: "Boletín 1er bimestre",
    audienceType: 1,
    contactDescription: "Ante cualquier duda, podes comunicarte con <strong>Secretaría</strong>."
}

//Datos de ejemplo para pruebas
const ejemploLab = {
  studentId: "123.456.789", 
  fileID: "19-4ca9klorEI94Egh-oWx-0hOd_PQdGE",
  subject: "Boletín 1er bimestre",
  audienceType: 1,
  contactDescription: "Ante cualquier duda, podes comunicarte con <strong>Secretaría</strong>."
}

// ============= FUNCIONES DE PRUEBA =============

/**
 * Función de prueba para enviar boletín con link
 */
function testEnviarBoletinConLink() {
  const token = getAuthorizationToken();
  
  if (!token) {
    Logger.log('Failed to login to Handing.');
    return;
  }

  const studentId = ejemploLink.uid;
  const reportLink = ejemploLink.linkUrl;
  const linkTitle = ejemploLink.linkTitle;
  const subject = ejemploLink.subject;
  const audienceType = 1;
  const resultado = enviarBoletinConLink(token, studentId, reportLink, linkTitle, subject, audienceType);
  
  if (resultado) {
    Logger.log('Boletín enviado con éxito. Post ID: ' + resultado.post_id);
  } else {
    Logger.log('Error al enviar el boletín');
  }
}

/*  ------------------------------------- */

function testEnviarBoletinPDF() {
  const token = getAuthorizationToken();
  
  if (!token) {
    Logger.log('Failed to login to Handing.');
    return;
  }

  const studentId = ejemploPDF.uid;
  const reportFileId = ejemploPDF.fileID;
  const subject = ejemploPDF.subject;
  const audienceType = ejemploPDF.audienceType;
  const contactDescription = ejemploPDF.contactDescription;
  const resultado =  enviarBoletinConPDF(token, studentId, reportFileId, subject, audienceType, contactDescription)
  
  if (resultado) {
    Logger.log('Boletín enviado con éxito. Post ID: ' + resultado.post_id);
  } else {
    Logger.log('Error al enviar el boletín');
  }
}

/**
 * return:
 * Registro de ejecución
    7:33:40 p.m.	Aviso	Se inició la ejecución
    7:33:43 p.m.	Información	Handing - authorization token response code 201
    7:33:43 p.m.	Información	Handing - enviando boletín a 123.456.789
    7:33:43 p.m.	Información	Handing - código de respuesta: 201
    7:33:43 p.m.	Información	Boletín enviado exitosamente. Post ID: 14675
    7:33:44 p.m.	Información	1 registro(s) agregado(s) exitosamente al log
    7:33:44 p.m.	Información	Boletín enviado con éxito. Post ID: 14675
    7:33:44 p.m.	Aviso	Se completó la ejecución
 */