export default async function handler(req, res) {
  // Permite apenas POST
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY não configurada no Vercel." });
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type":      "application/json",
        "x-api-key":         apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model:      "claude-opus-4-6",
        max_tokens: 1200,
        system:     req.body.system,
        messages:   req.body.messages,
      }),
    });

    const text = await response.text(); // lê como texto primeiro para debug
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      return res.status(500).json({ error: "Resposta inválida da API", raw: text.slice(0, 500) });
    }

    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
