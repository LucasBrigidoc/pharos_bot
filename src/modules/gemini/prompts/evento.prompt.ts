export function buildEventoSystem(dataHoje: string): string {
  return `\
Você recebe um texto livre escrito por um consultor da Pharos descrevendo
um evento/reunião que ele quer criar no calendário do Outlook. Extraia as
informações em JSON estruturado. Data de hoje: ${dataHoje} — use isso pra
resolver referências relativas ("amanhã", "sexta que vem", "daqui a duas
semanas", etc.).

═══ REGRA GERAL — NÃO INVENTAR ═══
Nunca crie título, data, horário, local ou participante que não esteja no
texto. Se um campo obrigatório (título, data, horário de início ou de fim)
não estiver claro, marque como "(a confirmar)" — nunca adivinhe, nunca
presuma um horário "padrão" nem uma duração padrão.

═══ HORÁRIOS ═══
Sempre no formato 24h "HH:mm" (ex.: "9h" → "09:00", "14h30" → "14:30",
"2 da tarde" → "14:00"). Se a pessoa disse só o horário de início e não
deu pista nenhuma de duração ou horário de término, marque "hora_fim"
como "(a confirmar)" — não invente 1 hora de duração por padrão.

═══ DATA ═══
Sempre no formato "DD/MM" ou "DD/MM/AAAA". Resolva dias da semana e
expressões relativas a partir da data de hoje. Se não for possível
determinar uma data, marque "(a confirmar)".

═══ LOCAL E DESCRIÇÃO ═══
"local": onde foi dito que vai ser (endereço, "online", nome do escritório
etc.) — "-" se não mencionado. "descricao": qualquer contexto adicional
relevante dito sobre o evento (pauta, motivo) — "-" se não houver.

═══ PARTICIPANTES ═══
Liste os nomes citados como participantes, exatamente como foram ditos
(não invente sobrenome, cargo ou e-mail). Lista vazia se ninguém foi
citado.

Schema de saída (responda APENAS com o JSON, sem texto fora dele):
{
  "titulo": "string",
  "data": "string",
  "hora_inicio": "string",
  "hora_fim": "string",
  "local": "string",
  "participantes": ["string"],
  "descricao": "string"
}`;
}

export function buildEventoUserMessage(textoLivre: string): string {
  return textoLivre;
}

export function buildEventoCorrectionMessage(params: { jsonAtual: string; correcao: string }): string {
  return `O JSON atual do evento é:
${params.jsonAtual}

O usuário pediu a seguinte correção (texto livre):
"${params.correcao}"

Aplique a correção e responda APENAS com o JSON atualizado completo, no mesmo schema. Não altere campos que não têm relação com a correção pedida.`;
}
