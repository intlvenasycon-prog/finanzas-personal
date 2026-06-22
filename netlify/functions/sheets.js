export const handler = async (event) => {
  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type, X-Sheets-Url',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  const sheetsUrl = event.headers['x-sheets-url'];
  if (!sheetsUrl) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Falta X-Sheets-Url header' }) };
  }

  try {
    if (event.httpMethod === 'GET') {
      const res = await fetch(sheetsUrl);
      if (!res.ok) throw new Error(`Sheets GET error: ${res.status}`);
      const data = await res.json();
      return { statusCode: 200, headers, body: JSON.stringify(data) };
    }

    if (event.httpMethod === 'POST') {
      const res = await fetch(sheetsUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: event.body,
      });
      if (!res.ok) throw new Error(`Sheets POST error: ${res.status}`);
      const data = await res.json();
      return { statusCode: 200, headers, body: JSON.stringify(data) };
    }

    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Método no permitido' }) };
  } catch (err) {
    return {
      statusCode: 502,
      headers,
      body: JSON.stringify({ error: 'Error con Google Sheets', detail: err.message }),
    };
  }
};
