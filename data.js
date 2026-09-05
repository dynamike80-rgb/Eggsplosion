// Server-side proxy naar Supabase. De browser praat alleen met dit endpoint
// (zelfde domein als de app zelf), dit stukje praat op de achtergrond met Supabase.
// Dat voorkomt cross-origin fetch-problemen in de browser van de gebruiker.

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;

exports.handler = async (event) => {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Content-Type": "application/json",
  };

  try {
    if (event.httpMethod === "GET") {
      const key = event.queryStringParameters?.key;
      if (!key) return { statusCode: 400, headers, body: JSON.stringify({ error: "key ontbreekt" }) };

      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/eggsplosion_data?key=eq.${encodeURIComponent(key)}&select=value`,
        { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } }
      );
      const text = await res.text();
      if (!res.ok) return { statusCode: res.status, headers, body: text };
      const rows = JSON.parse(text);
      return { statusCode: 200, headers, body: JSON.stringify({ value: rows[0]?.value ?? null }) };
    }

    if (event.httpMethod === "POST") {
      const { key, value } = JSON.parse(event.body || "{}");
      if (!key) return { statusCode: 400, headers, body: JSON.stringify({ error: "key ontbreekt" }) };

      const res = await fetch(`${SUPABASE_URL}/rest/v1/eggsplosion_data`, {
        method: "POST",
        headers: {
          apikey: SUPABASE_KEY,
          Authorization: `Bearer ${SUPABASE_KEY}`,
          "Content-Type": "application/json",
          Prefer: "resolution=merge-duplicates,return=minimal",
        },
        body: JSON.stringify({ key, value, updated_at: new Date().toISOString() }),
      });
      if (!res.ok) {
        const text = await res.text();
        return { statusCode: res.status, headers, body: text };
      }
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) };
    }

    if (event.httpMethod === "OPTIONS") {
      return { statusCode: 204, headers };
    }

    return { statusCode: 405, headers, body: JSON.stringify({ error: "methode niet toegestaan" }) };
  } catch (err) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: String(err) }) };
  }
};
