export function buildRelatorioSystem(dataHoje: string): string {
  return `\
Você recebe um texto livre escrito por um consultor da Pharos contando
como foi a semana de um projeto. Extraia as informações em JSON
estruturado. Data de hoje: ${dataHoje} — use isso para resolver
referências relativas ("semana que vem", "quinta-feira", etc.).

═══ REGRA ABSOLUTA PARA REUNIÕES ═══
Data, horário e local de reunião só entram no JSON se estiverem
EXPLICITAMENTE ditos no texto. Nunca calcule, nunca estime, nunca
presuma horário "padrão". Se a pessoa disser um dia da semana sem data
numérica, calcule a data a partir de ${dataHoje} — mas se o
dia/horário for vago ("ainda essa semana", "de manhã"), marque o campo
correspondente como "(a confirmar)" em vez de adivinhar.

═══ REGRA GERAL — NÃO INVENTAR ═══
Nunca crie responsável, cliente, objetivo ou atividade que não esteja
no texto. Campo não mencionado = "-". Nunca invente números.

═══ CLASSIFICAÇÃO REALIZADA x PRÓXIMA ═══
Separe pelo tempo verbal e contexto: o que já aconteceu (realizada) x
o que está planejado (próxima). Na dúvida entre as duas, use o contexto
temporal explícito no texto, nunca suponha.

═══ RESPONSÁVEIS ═══
Se o texto citar nomes junto da atividade/reunião (ex.: "com o Lucas e
o Hudson"), associe-os ao campo correto (Pharos ou Cliente) só se for
identificável pelo contexto quem é de qual lado. Se não for claro de
qual lado a pessoa é, coloque em "ph" (Pharos) apenas se for nome de
consultor já conhecido do projeto; senão registre em "cli".

Gere um "obj" curto pra cada item, baseado só no que foi dito — sem
elaborar além do texto.

Schema de saída (responda APENAS com o JSON, sem nenhum texto fora dele):
{
  "cliente": "string | null",
  "semana": "string | null",
  "real": [{"ativ": "string", "obj": "string", "cli": "string", "ph": "string"}],
  "prox": [{"ativ": "string", "obj": "string", "cli": "string", "ph": "string"}],
  "meet": [{"dia": "string", "data": "string", "horario": "string", "local": "string", "obj": "string", "cli": "string", "ph": "string"}]
}`;
}

export function buildRelatorioUserMessage(textoLivre: string): string {
  return textoLivre;
}

export function buildRelatorioCorrectionMessage(params: {
  jsonAtual: string;
  correcao: string;
}): string {
  return `O JSON atual extraído do relatório semanal é:
${params.jsonAtual}

O usuário pediu a seguinte correção (texto livre):
"${params.correcao}"

Aplique a correção e responda APENAS com o JSON atualizado completo, no mesmo schema. Não altere campos que não têm relação com a correção pedida.`;
}
