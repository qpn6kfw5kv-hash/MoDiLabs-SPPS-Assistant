// Local AI proxy for SPPS Lab Assistant.
// Run with:
//   $env:OPENAI_API_KEY="sk-..."
//   node ai-chat-proxy.js
//
// Optional:
//   $env:OPENAI_MODEL="gpt-4.1-mini"
//   $env:AI_CHAT_PORT="8787"

const http = require('http');

const PORT = Number(process.env.AI_CHAT_PORT || 8787);
const MODEL = process.env.OPENAI_MODEL || 'gpt-4.1-mini';
const API_KEY = process.env.OPENAI_API_KEY;

function sendJson(res, status, payload) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  res.end(JSON.stringify(payload));
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => {
      body += chunk;
      if (body.length > 1_500_000) {
        reject(new Error('Payload too large'));
        req.destroy();
      }
    });
    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (err) {
        reject(err);
      }
    });
    req.on('error', reject);
  });
}

function extractOutputText(response) {
  if (response.output_text) return response.output_text;
  const chunks = [];
  for (const item of response.output || []) {
    for (const content of item.content || []) {
      if (content.type === 'output_text' && content.text) chunks.push(content.text);
    }
  }
  return chunks.join('\n').trim();
}

function buildPrompt(question, context, history) {
  return [
    'Domanda utente:',
    question || '',
    '',
    'Storico recente:',
    JSON.stringify(history || [], null, 2),
    '',
    'Contesto strutturato della sintesi corrente:',
    JSON.stringify(context || {}, null, 2)
  ].join('\n');
}

const server = http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') return sendJson(res, 200, { ok: true });
  if (req.method !== 'POST' || req.url !== '/chat') return sendJson(res, 404, { error: 'Not found' });
  if (!API_KEY) return sendJson(res, 500, { error: 'OPENAI_API_KEY non configurata sul proxy.' });

  try {
    const body = await readJson(req);
    const question = String(body.question || '').trim();
    if (!question) return sendJson(res, 400, { error: 'Domanda mancante.' });

    const apiResponse = await fetch('https://api.openai.com/v1/responses', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${API_KEY}`
      },
      body: JSON.stringify({
        model: MODEL,
        instructions: [
          'Sei un assistente di laboratorio per sintesi peptidica in fase solida (SPPS).',
          'Rispondi in italiano naturale, preciso e non meccanico.',
          'Usa il contesto strutturato fornito come fonte primaria: non inventare locazioni, quantita, CAS, codici, masse o risultati HPLC/MS non presenti.',
          'Se l’utente chiede dove si trova un composto e il contesto contiene inventory.directLookup.exact, rispondi con locazione, quantita, marca, codice, CAS e residuo in modo discorsivo.',
          'Se non hai abbastanza dati, dillo chiaramente e suggerisci quale dato serve.',
          'Quando parli di procedure chimiche, mantieni tono operativo ma prudente: segnala che modifiche sperimentali vanno validate in laboratorio.',
          'Evita risposte a elenco se una risposta breve e ben spiegata e piu utile; usa elenchi solo per dati tecnici.'
        ].join('\n'),
        input: buildPrompt(question, body.context, body.history),
        max_output_tokens: 900
      })
    });

    const data = await apiResponse.json();
    if (!apiResponse.ok) {
      return sendJson(res, apiResponse.status, { error: data.error?.message || 'Errore OpenAI API.' });
    }

    sendJson(res, 200, { answer: extractOutputText(data) || 'Non sono riuscito a generare una risposta utile.' });
  } catch (err) {
    sendJson(res, 500, { error: err.message || 'Errore proxy AI.' });
  }
});

server.listen(PORT, () => {
  console.log(`SPPS AI chat proxy listening on http://localhost:${PORT}/chat`);
});
