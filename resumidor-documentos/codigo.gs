/**
 * RESUMIDOR AUTOMÁTICO DE DOCUMENTOS DE GOOGLE DRIVE CON GEMINI
 *
 * Qué hace:
 *  1. Revisa una carpeta de Drive ("entrada").
 *  2. Por cada archivo nuevo (Google Docs, PDF o TXT) pide un resumen a Gemini.
 *  3. Guarda el resumen como un Google Doc en otra carpeta ("salida").
 *  4. Recuerda qué archivos ya resumió para no repetirlos.
 *
 * Antes de ejecutar:
 *  - Guarda tu API key en: Configuración del proyecto > Propiedades del script
 *    con el nombre GEMINI_API_KEY (no la pegues en el código).
 *  - Rellena los dos IDs de carpeta en CONFIG.
 */
 
// ============ CONFIGURACIÓN ============
const CONFIG = {
  CARPETA_ENTRADA_ID: 'PEGA_AQUI_EL_ID_DE_LA_CARPETA_DE_ENTRADA',
  CARPETA_SALIDA_ID: 'PEGA_AQUI_EL_ID_DE_LA_CARPETA_DE_SALIDA',
 
  // Modelo de Gemini. Revisa https://ai.google.dev/gemini-api/docs/models
  // por si hay uno más reciente; los modelos antiguos se van retirando.
  MODELO: 'gemini-flash-latest',
 
  IDIOMA: 'español',
  MAX_CARACTERES: 400000,        // recorte de seguridad para textos enormes
  MAX_BYTES_PDF: 15 * 1024 * 1024, // 15 MB
  TIEMPO_MAX_MS: 5 * 60 * 1000,  // Apps Script corta a los 6 min; paramos a los 5
  FRECUENCIA_MINUTOS: 15         // valores válidos: 1, 5, 10, 15, 30
};
 
// ============ FUNCIÓN PRINCIPAL ============
function resumirDocumentosNuevos() {
  const inicio = Date.now();
  const props = PropertiesService.getScriptProperties();
  const entrada = DriveApp.getFolderById(CONFIG.CARPETA_ENTRADA_ID);
  const salida = DriveApp.getFolderById(CONFIG.CARPETA_SALIDA_ID);
  const archivos = entrada.getFiles();
 
  let resumidos = 0;
 
  while (archivos.hasNext()) {
    if (Date.now() - inicio > CONFIG.TIEMPO_MAX_MS) {
      Logger.log('Tiempo límite alcanzado. Continuará en la próxima ejecución.');
      break;
    }
 
    const archivo = archivos.next();
    const clave = 'hecho_' + archivo.getId();
 
    if (props.getProperty(clave)) continue; // ya resumido
 
    try {
      const contenido = obtenerContenido_(archivo);
      if (!contenido) {
        Logger.log('Tipo no soportado, se omite: ' + archivo.getName() + ' (' + archivo.getMimeType() + ')');
        continue;
      }
 
      Logger.log('Resumiendo: ' + archivo.getName());
      const resumen = resumirConGemini_(contenido, archivo.getName());
      guardarResumen_(salida, archivo, resumen);
 
      props.setProperty(clave, new Date().toISOString());
      resumidos++;
    } catch (e) {
      Logger.log('ERROR con "' + archivo.getName() + '": ' + e.message);
    }
  }
 
  Logger.log('Listo. Documentos resumidos en esta ejecución: ' + resumidos);
}
 
// ============ LECTURA DEL CONTENIDO ============
function obtenerContenido_(archivo) {
  const tipo = archivo.getMimeType();
 
  if (tipo === MimeType.GOOGLE_DOCS) {
    const texto = DocumentApp.openById(archivo.getId()).getBody().getText();
    if (!texto.trim()) return null;
    return { texto: texto };
  }
 
  if (tipo === MimeType.PLAIN_TEXT) {
    const texto = archivo.getBlob().getDataAsString('UTF-8');
    if (!texto.trim()) return null;
    return { texto: texto };
  }
 
  if (tipo === MimeType.PDF) {
    if (archivo.getSize() > CONFIG.MAX_BYTES_PDF) {
      throw new Error('El PDF supera el tamaño máximo permitido.');
    }
    return { pdfBase64: Utilities.base64Encode(archivo.getBlob().getBytes()) };
  }
 
  return null; // otros tipos no soportados
}
 
// ============ LLAMADA A GEMINI ============
function resumirConGemini_(contenido, nombreArchivo) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Falta la propiedad GEMINI_API_KEY en las Propiedades del script.');
 
  const instrucciones =
    'Eres un asistente experto en resumir documentos. Resume el documento "' + nombreArchivo + '" en ' + CONFIG.IDIOMA + '.\n' +
    'Estructura la respuesta así:\n' +
    'RESUMEN EJECUTIVO: 3 a 5 frases.\n' +
    'PUNTOS CLAVE: lista con viñetas usando "•".\n' +
    'CONCLUSIONES O ACCIONES: si las hay.\n' +
    'Responde en texto plano, sin asteriscos ni formato Markdown. No inventes datos que no estén en el documento.';
 
  const partes = [{ text: instrucciones }];
 
  if (contenido.texto) {
    partes.push({ text: 'DOCUMENTO:\n' + contenido.texto.substring(0, CONFIG.MAX_CARACTERES) });
  }
  if (contenido.pdfBase64) {
    partes.push({ inlineData: { mimeType: 'application/pdf', data: contenido.pdfBase64 } });
  }
 
  return llamarGemini_({ contents: [{ parts: partes }] }, apiKey);
}
 
function llamarGemini_(payload, apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.MODELO + ':generateContent';
  const opciones = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
 
  const MAX_INTENTOS = 3;
  for (let intento = 1; intento <= MAX_INTENTOS; intento++) {
    const respuesta = UrlFetchApp.fetch(url, opciones);
    const codigo = respuesta.getResponseCode();
    const cuerpo = respuesta.getContentText();
 
    if (codigo === 200) {
      const json = JSON.parse(cuerpo);
      const candidato = json.candidates && json.candidates[0];
      if (!candidato || !candidato.content || !candidato.content.parts) {
        const motivo = json.promptFeedback && json.promptFeedback.blockReason
          ? 'Bloqueado: ' + json.promptFeedback.blockReason
          : 'Respuesta vacía de Gemini';
        throw new Error(motivo);
      }
      return candidato.content.parts.map(function (p) { return p.text || ''; }).join('').trim();
    }
 
    // Errores temporales: reintentar con espera creciente
    if ((codigo === 429 || codigo >= 500) && intento < MAX_INTENTOS) {
      Utilities.sleep(intento * 3000);
      continue;
    }
 
    throw new Error('Gemini respondió ' + codigo + ': ' + cuerpo.substring(0, 300));
  }
}
 
// ============ GUARDAR EL RESUMEN ============
function guardarResumen_(carpetaSalida, archivoOriginal, resumen) {
  const doc = DocumentApp.create('Resumen - ' + archivoOriginal.getName());
  const body = doc.getBody();
 
  body.appendParagraph('Resumen: ' + archivoOriginal.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Documento original: ' + archivoOriginal.getUrl());
  body.appendParagraph('Generado el ' + new Date().toLocaleString('es-ES') + ' con ' + CONFIG.MODELO)
      .setItalic(true);
  body.appendHorizontalRule();
  body.appendParagraph(resumen);
 
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(carpetaSalida);
}
 
// ============ ACTIVADOR AUTOMÁTICO ============
function crearActivador() {
  eliminarActivadores(); // evita duplicados
  ScriptApp.newTrigger('resumirDocumentosNuevos')
    .timeBased()
    .everyMinutes(CONFIG.FRECUENCIA_MINUTOS)
    .create();
  Logger.log('Activador creado: cada ' + CONFIG.FRECUENCIA_MINUTOS + ' minutos.');
}
 
function eliminarActivadores() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'resumirDocumentosNuevos') ScriptApp.deleteTrigger(t);
  });
}
 
// ============ UTILIDADES ============
/** Ejecútala primero para comprobar que la API key y el modelo funcionan. */
function probarConexion() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Falta GEMINI_API_KEY en las Propiedades del script.');
  const respuesta = llamarGemini_({ contents: [{ parts: [{ text: 'Responde solo con la palabra: OK' }] }] }, apiKey);
  Logger.log('Respuesta de Gemini: ' + respuesta);
}
 
/** Borra la memoria de archivos ya resumidos (volverá a resumir todo). */
function reiniciarHistorial() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).forEach(function (k) {
    if (k.indexOf('hecho_') === 0) props.deleteProperty(k);
  });
  Logger.log('Historial reiniciado.');
}
