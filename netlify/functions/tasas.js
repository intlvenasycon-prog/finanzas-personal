export const handler = async (event) => {
  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Cache-Control': 'public, max-age=300',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  const { type = 'oficial' } = event.queryStringParameters || {};
  const valid = ['oficial', 'euro', 'paralelo'];

  if (!valid.includes(type)) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Tipo inválido' }) };
  }

  // EUR lives under /v1/euros/oficial, USD variants under /v1/dolares/<type>
  const url = type === 'euro'
    ? 'https://ve.dolarapi.com/v1/euros/oficial'
    : `https://ve.dolarapi.com/v1/dolares/${type}`;

  try {
    const res = await fetch(url, {
      headers: { 'User-Agent': 'finanzas-personal/1.0' },
    });

    if (!res.ok) throw new Error(`API error: ${res.status}`);

    const data = await res.json();
    return { statusCode: 200, headers, body: JSON.stringify(data) };
  } catch (err) {
    return {
      statusCode: 502,
      headers,
      body: JSON.stringify({ error: 'No se pudo obtener la tasa', detail: err.message }),
    };
  }
};
