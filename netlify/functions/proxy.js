// Proxy serverless: reenvía pedidos del frontend al Apps Script (evita problemas de CORS).
// Configurar la variable de entorno APPS_SCRIPT_URL en Netlify con la URL /exec de tu Apps Script.

exports.handler = async (event) => {
  const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;

  if (!APPS_SCRIPT_URL) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: 'Falta configurar APPS_SCRIPT_URL en Netlify (Site settings → Environment variables)' })
    };
  }

  try {
    let response;
    if (event.httpMethod === 'GET') {
      const qs = event.rawQuery ? '?' + event.rawQuery : '';
      response = await fetch(APPS_SCRIPT_URL + qs);
    } else if (event.httpMethod === 'POST') {
      response = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: event.body
      });
    } else {
      return { statusCode: 405, body: JSON.stringify({ error: 'Método no permitido' }) };
    }

    const text = await response.text();
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      body: text
    };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: String(err) }) };
  }
};
