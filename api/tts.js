module.exports = async function handler(request, response) {
  if (request.method !== "POST") {
    response.setHeader("Allow", "POST");
    response.status(405).json({ error: "Method not allowed" });
    return;
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    response.status(500).json({ error: "OPENAI_API_KEY is not configured." });
    return;
  }

  const { text } = request.body || {};
  const input = String(text || "").replace(/\s+/g, " ").trim();

  if (!input) {
    response.status(400).json({ error: "Text is required." });
    return;
  }

  if (input.length > 2500) {
    response.status(400).json({ error: "Text is too long for one speech request." });
    return;
  }

  const speechResponse = await fetch("https://api.openai.com/v1/audio/speech", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: process.env.OPENAI_TTS_MODEL || "gpt-4o-mini-tts",
      voice: process.env.OPENAI_TTS_VOICE || "alloy",
      input,
      response_format: "mp3",
      instructions: "Read this Georgian text clearly and naturally, like a calm lecturer presenting slides.",
    }),
  });

  if (!speechResponse.ok) {
    const message = await speechResponse.text();
    response.status(speechResponse.status).json({ error: message });
    return;
  }

  const audio = Buffer.from(await speechResponse.arrayBuffer());
  response.setHeader("Content-Type", "audio/mpeg");
  response.setHeader("Cache-Control", "no-store");
  response.status(200).send(audio);
};
