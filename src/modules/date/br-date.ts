// Helpers de data em português (fuso America/Fortaleza), compartilhados
// entre os fluxos que lidam com datas em texto livre ou determinísticas
// (relatório semanal, turno, evento).

export function dataHojeExtenso(): string {
  return new Intl.DateTimeFormat('pt-BR', {
    timeZone: 'America/Fortaleza',
    weekday: 'long',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(new Date());
}

export function dataHojeCurta(): string {
  return new Intl.DateTimeFormat('pt-BR', {
    timeZone: 'America/Fortaleza',
    day: '2-digit',
    month: '2-digit',
  }).format(new Date());
}

export function capitalize(s: string): string {
  return s.length ? s[0].toUpperCase() + s.slice(1) : s;
}

const WEEKDAYS_PT = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];

export function weekdayPt(date: Date): string {
  return WEEKDAYS_PT[date.getUTCDay()];
}

// Calcula um Date a partir de "DD/MM" ou "DD/MM/AAAA" — nunca confia em
// dia-da-semana dito por texto livre, só na data numérica. Sem ano
// explícito, assume o ano de referência, ou o seguinte se a data já
// ficou mais de 30 dias no passado (ex.: "10/01" dito em dezembro).
export function parseDataBR(dataStr: string, refDate: Date): Date | null {
  const m = dataStr.trim().match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (!m) return null;

  const day = Number(m[1]);
  const month = Number(m[2]) - 1;
  let year = m[3] ? Number(m[3]) : refDate.getFullYear();
  if (m[3] && m[3].length === 2) year += 2000;

  let d = new Date(Date.UTC(year, month, day, 12));
  if (!m[3]) {
    const diffDays = (refDate.getTime() - d.getTime()) / 86_400_000;
    if (diffDays > 30) d = new Date(Date.UTC(year + 1, month, day, 12));
  }
  return Number.isNaN(d.getTime()) ? null : d;
}
