import { GoogleGenerativeAI } from '@google/generative-ai';

const key = process.env.GEMINI_API_KEY;
if (!key) throw new Error('GEMINI_API_KEY not set');

const genAI = new GoogleGenerativeAI(key);

export function getGeminiModel(systemInstruction: string) {
  return genAI.getGenerativeModel({
    model: 'gemini-2.0-flash',
    systemInstruction,
  });
}

export function extractJSON(raw: string): string {
  const block = raw.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
  if (block) return block[1];
  const start = raw.indexOf('{');
  const end = raw.lastIndexOf('}');
  if (start !== -1 && end !== -1) return raw.slice(start, end + 1);
  return raw.trim();
}
