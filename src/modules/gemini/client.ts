import { GoogleGenAI } from '@google/genai';

const key = process.env.GEMINI_API_KEY;
if (!key) throw new Error('GEMINI_API_KEY not set');

const ai = new GoogleGenAI({ apiKey: key });

export async function callGemini(systemInstruction: string, userMessage: string): Promise<string> {
  const response = await ai.models.generateContent({
    model: 'gemini-3.5-flash',
    config: { systemInstruction },
    contents: userMessage,
  });
  return response.text ?? '';
}

export function extractJSON(raw: string): string {
  const block = raw.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
  if (block) return block[1];
  const start = raw.indexOf('{');
  const end = raw.lastIndexOf('}');
  if (start !== -1 && end !== -1) return raw.slice(start, end + 1);
  return raw.trim();
}
