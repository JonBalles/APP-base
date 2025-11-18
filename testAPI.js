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

  const studentId = ejemploLab.uid;
  const reportLink = ejemploLab.linkUrl;
  const linkTitle = ejemploLab.linkTitle;
  const subject = ejemploLab.subject;
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

  const studentId = ejemploLab.uid;
  const reportFileId = ejemploLab.fileID;
  const subject = ejemploLab.subject;
  const audienceType = ejemploLab.audienceType;
  const contactDescription = ejemploLab.contactDescription;
  const resultado =  enviarBoletinConPDF(token, studentId, reportFileId, subject, audienceType, contactDescription)
  
  if (resultado) {
    Logger.log('Boletín enviado con éxito. Post ID: ' + resultado.post_id);
  } else {
    Logger.log('Error al enviar el boletín');
  }
}
