// ===== CONFIGURACIÓN =====
const CONFIG = {
  GEMINI_MODEL: 'gemini-flash-latest', // rápido y gratuito en su nivel free
  LABEL_ENTRADA: 'Bot-Pendiente',   // etiqueta que marcas en Gmail para que el bot procese
  LABEL_PROCESADO: 'Bot-Respondido',
  MODO_BORRADOR: true,              // true = crea borrador, false = envía directo
  HOJA_LOG: 'Log_Correos',
  MAX_CORREOS_POR_EJECUCION: 10
};

// ===== FUNCIÓN PRINCIPAL (la que se ejecuta con el trigger) =====
function procesarCorreos() {
  const label = GmailApp.getUserLabelByName(CONFIG.LABEL_ENTRADA);
  if (!label) {
    Logger.log('No existe la etiqueta ' + CONFIG.LABEL_ENTRADA);
    return;
  }

  const labelProcesado = getOrCreateLabel(CONFIG.LABEL_PROCESADO);
  const hilos = label.getThreads(0, CONFIG.MAX_CORREOS_POR_EJECUCION);

  hilos.forEach(hilo => {
    try {
      const mensajes = hilo.getMessages();
      const ultimoMensaje = mensajes[mensajes.length - 1];

      const remitente = ultimoMensaje.getFrom();
      const asunto = ultimoMensaje.getSubject();
      const cuerpo = ultimoMensaje.getPlainBody();

      // 1. Generar respuesta con Gemini
      const respuestaIA = generarRespuestaIA(asunto, cuerpo);

      // 2. Registrar en el Sheet
      registrarEnSheet(remitente, asunto, cuerpo, respuestaIA);

      // 3. Responder (borrador o envío directo)
      if (CONFIG.MODO_BORRADOR) {
        ultimoMensaje.createDraftReply(respuestaIA);
      } else {
        ultimoMensaje.reply(respuestaIA);
      }

      // 4. Mover etiquetas
      hilo.removeLabel(label);
      hilo.addLabel(labelProcesado);

    } catch (error) {
      Logger.log('Error procesando hilo: ' + error.message);
    }
  });
}

// ===== LLAMADA A LA API DE GEMINI =====
function generarRespuestaIA(asunto, cuerpoCorreo) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('No hay API Key configurada. Ejecuta guardarApiKey() primero.');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const prompt = `Eres un agente de atención al cliente profesional y amable.
Responde el siguiente correo de un cliente de forma clara, breve y cordial.
No inventes datos que no tengas (precios, plazos, políticas específicas) — si falta información, pide amablemente al cliente que aclare o indica que un agente humano lo contactará.

Asunto: ${asunto}

Correo del cliente:
"""
${cuerpoCorreo}
"""

Escribe solo el cuerpo de la respuesta, sin encabezados tipo "Estimado cliente" repetidos innecesariamente, en español, tono profesional y cercano.`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.4, maxOutputTokens: 500 }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const MAX_REINTENTOS = 5;
  let espera = 2000; // empieza en 2 segundos

  for (let intento = 1; intento <= MAX_REINTENTOS; intento++) {
    const response = UrlFetchApp.fetch(url, options);
    const codigo = response.getResponseCode();

    if (codigo === 200) {
      const data = JSON.parse(response.getContentText());
      return data.candidates[0].content.parts[0].text.trim();
    }

    // Errores temporales que vale la pena reintentar: 429 (rate limit), 503 (sobrecarga), 500
    const esTemporal = [429, 500, 503].includes(codigo);

    Logger.log(`Intento ${intento} falló (código ${codigo}): ${response.getContentText()}`);

    if (!esTemporal || intento === MAX_REINTENTOS) {
      throw new Error(`Fallo la llamada a Gemini tras ${intento} intentos: ${codigo} - ${response.getContentText()}`);
    }

    Utilities.sleep(espera);
    espera *= 2; // backoff exponencial: 2s, 4s, 8s, 16s...
  }
}

// ===== REGISTRO EN SHEET =====
function registrarEnSheet(remitente, asunto, cuerpo, respuesta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(CONFIG.HOJA_LOG);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_LOG);
    hoja.appendRow(['Fecha', 'Remitente', 'Asunto', 'Correo original', 'Respuesta IA', 'Modo']);
  }
  hoja.appendRow([
    new Date(),
    remitente,
    asunto,
    cuerpo.substring(0, 500), // recorta para no saturar la celda
    respuesta,
    CONFIG.MODO_BORRADOR ? 'Borrador' : 'Enviado'
  ]);
}

// ===== UTILIDAD =====
function getOrCreateLabel(nombre) {
  let label = GmailApp.getUserLabelByName(nombre);
  if (!label) label = GmailApp.createLabel(nombre);
  return label;
}

// ===== CREAR EL TRIGGER AUTOMÁTICO (ejecuta una vez) =====
function crearTrigger() {
  // Elimina triggers previos de esta función para no duplicar
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'procesarCorreos') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('procesarCorreos')
    .timeBased()
    .everyMinutes(15) // cada 15 minutos, ajustable
    .create();
}
