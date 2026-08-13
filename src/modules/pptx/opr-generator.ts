// Portado de opr_pharos_gera_ppt_local.html (função gerarPPT()).
// Diferenças em relação ao original: pres.write() retorna Buffer em vez de
// pres.writeFile() (que baixava no navegador), os cards de cada coluna têm
// altura/espaçamento dinâmicos para nunca ultrapassar o slide, e o cabeçalho
// usa a identidade visual real da Pharos (logo do farol + Playfair
// Display/Helvetica Now, mesmas fontes do módulo de ata).

import fs from 'fs';
import path from 'path';
import pptxgen from 'pptxgenjs';

// Fontes da marca Pharos — mesma dupla usada no template de ata (docx).
const FONT_TITLE = 'Playfair Display';
const FONT_BODY  = 'Helvetica Now';

const LOGO_PATH = path.join(process.cwd(), 'assets', 'pharos-logo.png');
const LOGO_RATIO = 400 / 347; // largura/altura do assets/pharos-logo.png
let logoDataUri: string | undefined;
function getLogoDataUri(): string {
  if (!logoDataUri) {
    logoDataUri = `data:image/png;base64,${fs.readFileSync(LOGO_PATH).toString('base64')}`;
  }
  return logoDataUri;
}

export interface OprItem {
  ativ: string;
  obj:  string;
  cli:  string;
  ph:   string;
}

export interface OprMeeting {
  dia:     string;
  data:    string;
  horario: string;
  local:   string;
  obj:     string;
  cli:     string;
  ph:      string;
}

export interface OprData {
  cliente: string;
  semana:  string;
  real:    OprItem[];
  prox:    OprItem[];
  meet:    OprMeeting[];
}

// ─── Layout responsivo dos cards ─────────────────────────────────────────────
// Dado N itens, uma altura/gap "ideais" e o espaço vertical disponível na
// coluna, calcula a altura/gap reais a usar: primeiro tenta o ideal, depois
// aperta o espaçamento entre cards, e só reduz a altura do card como último
// recurso — sempre garantindo que a pilha de cards caiba no espaço disponível.
function fitCards(n: number, idealH: number, idealGap: number, availH: number): { h: number; gap: number } {
  if (n <= 0) return { h: idealH, gap: idealGap };
  const minGap = 0.03;

  const idealTotal = n * idealH + (n - 1) * idealGap;
  if (idealTotal <= availH) return { h: idealH, gap: idealGap };

  const compactTotal = n * idealH + (n - 1) * minGap;
  if (compactTotal <= availH) {
    const gap = n > 1 ? (availH - n * idealH) / (n - 1) : idealGap;
    return { h: idealH, gap };
  }

  const h = Math.max((availH - (n - 1) * minGap) / n, 0.3);
  return { h, gap: minGap };
}

export async function generateOprPptx(data: OprData): Promise<Buffer> {
  const C = { dg:'072E22', mg:'14663D', gold:'AF7932', off:'F2F5F2', white:'FFFFFF', border:'C2D4C5', muted:'5A7A60', light:'E8F0EA' };
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';

  const s = pres.addSlide();
  s.background = { color: C.off };
  const mkSh = () => ({ type: 'outer' as const, blur: 5, offset: 2, angle: 135, color: C.dg, opacity: 0.08 });

  // ─── Barra superior ──────────────────────────────────────────────────────
  s.addShape(pres.ShapeType.rect, { x:0, y:0, w:10, h:0.55, fill:{color:C.dg}, line:{color:C.dg} });
  s.addShape(pres.ShapeType.rect, { x:0, y:0.55, w:10, h:0.04, fill:{color:C.gold}, line:{color:C.gold} });

  // Logo Pharos (farol) + wordmark
  const logoH = 0.40, logoW = logoH * LOGO_RATIO, logoX = 0.24, logoY = (0.55 - logoH) / 2;
  s.addImage({ data: getLogoDataUri(), x: logoX, y: logoY, w: logoW, h: logoH });

  const wmX = logoX + logoW + 0.10;
  s.addText('PHAROS', { x:wmX, y:0.07, w:1.6, h:0.40, fontSize:14, bold:true, color:C.white, fontFace:FONT_TITLE, charSpacing:4, valign:'middle', margin:0 });
  s.addShape(pres.ShapeType.ellipse, { x:wmX+1.15, y:0.22, w:0.07, h:0.07, fill:{color:C.gold}, line:{color:C.gold} });
  s.addText('Consultoria', { x:wmX+1.26, y:0.21, w:1.0, h:0.16, fontSize:6.5, color:'8FBF9A', fontFace:FONT_BODY, margin:0 });

  // Pill OPR
  s.addShape(pres.ShapeType.rect, { x:4.1, y:0.09, w:1.8, h:0.35, fill:{color:C.gold}, line:{color:C.gold} });
  s.addText('OPR  –  One Page Report', { x:4.1, y:0.09, w:1.8, h:0.35, fontSize:7.5, bold:true, color:C.dg, fontFace:FONT_BODY, align:'center', valign:'middle', margin:0 });

  // Semana + CONFIDENCIAL
  s.addText(`Semana: ${data.semana}  |  Cliente: ${data.cliente}`, { x:6.2, y:0.07, w:3.55, h:0.20, fontSize:7, color:'A8D0B4', fontFace:FONT_BODY, align:'right', margin:0 });
  s.addText('CONFIDENCIAL', { x:6.2, y:0.28, w:3.55, h:0.18, fontSize:6.5, bold:true, color:C.gold, fontFace:FONT_BODY, align:'right', charSpacing:2, margin:0 });

  // ─── 3 colunas ───────────────────────────────────────────────────────────
  const colY=0.66, colH=4.62, colW=3.10, gap=0.10;
  const cols = [0.18, 0.18+colW+gap, 0.18+2*(colW+gap)];
  const headers: [string, string][] = [
    ['✓  ATIVIDADES REALIZADAS', C.dg],
    ['→  PRÓXIMAS ENTREGAS',     C.mg],
    ['◷  REUNIÕES AGENDADAS',    C.dg],
  ];
  cols.forEach((cx, i) => {
    s.addShape(pres.ShapeType.rect, { x:cx, y:colY, w:colW, h:colH, fill:{color:C.white}, line:{color:C.border, pt:0.5}, shadow:mkSh() });
    s.addShape(pres.ShapeType.rect, { x:cx, y:colY, w:colW, h:0.30, fill:{color:headers[i][1]}, line:{color:headers[i][1]} });
    s.addText(headers[i][0], { x:cx, y:colY, w:colW, h:0.30, fontSize:7, bold:true, color:C.white, fontFace:FONT_BODY, align:'center', valign:'middle', margin:0, charSpacing:.5 });
  });

  const startY = colY+0.36;
  const availH = colY + colH - startY;

  // ─── Cards de atividades (Realizadas / Próximas) ────────────────────────────
  const IDEAL_CARD_H = 1.22, IDEAL_CARD_GAP = 0.06;

  function addCard(cx: number, cardY: number, item: OprItem, accentColor: string, num: number, cardH: number) {
    const cW=colW-0.22, cX=cx+0.11, cH=cardH;
    const scale = Math.min(cardH / IDEAL_CARD_H, 1);
    const fz = (v: number) => v * Math.max(scale, 0.6);
    const oy = (v: number) => cardY + v * scale;

    s.addShape(pres.ShapeType.rect,    { x:cX, y:cardY, w:cW, h:cH, fill:{color:C.off}, line:{color:C.border, pt:.5} });
    s.addShape(pres.ShapeType.rect,    { x:cX, y:cardY, w:0.05, h:cH, fill:{color:accentColor}, line:{color:accentColor} });
    s.addShape(pres.ShapeType.ellipse, { x:cX+0.10, y:oy(0.09), w:0.22*scale, h:0.22*scale, fill:{color:accentColor}, line:{color:accentColor} });
    s.addText(String(num),  { x:cX+0.10, y:oy(0.09), w:0.22*scale, h:0.22*scale, fontSize:fz(7), bold:true, color:C.white, fontFace:FONT_BODY, align:'center', valign:'middle', margin:0 });
    s.addText(item.ativ,    { x:cX+0.36, y:oy(0.08), w:cW-0.44, h:0.26*scale, fontSize:fz(7.5), bold:true, color:C.dg, fontFace:FONT_TITLE, valign:'middle', margin:0 });
    s.addText(item.obj,     { x:cX+0.10, y:oy(0.34), w:cW-0.18, h:0.26*scale, fontSize:fz(7), color:C.muted, fontFace:FONT_BODY, valign:'top', margin:0 });
    s.addShape(pres.ShapeType.line, { x:cX+0.10, y:oy(0.64), w:cW-0.20, h:0, line:{color:C.border, pt:.5} });
    s.addText('Cliente:', { x:cX+0.10, y:oy(0.70), w:0.46, h:0.17*scale, fontSize:fz(6.5), bold:true, color:C.gold, fontFace:FONT_BODY, margin:0 });
    s.addText(item.cli,   { x:cX+0.56, y:oy(0.70), w:cW-0.64, h:0.17*scale, fontSize:fz(6.5), color:C.muted, fontFace:FONT_BODY, margin:0 });
    s.addText('Pharos:', { x:cX+0.10, y:oy(0.89), w:0.46, h:0.17*scale, fontSize:fz(6.5), bold:true, color:C.mg, fontFace:FONT_BODY, margin:0 });
    s.addText(item.ph,   { x:cX+0.56, y:oy(0.89), w:cW-0.64, h:0.17*scale, fontSize:fz(6.5), color:C.muted, fontFace:FONT_BODY, margin:0 });
  }

  const realLayout = fitCards(data.real.length, IDEAL_CARD_H, IDEAL_CARD_GAP, availH);
  const proxLayout = fitCards(data.prox.length, IDEAL_CARD_H, IDEAL_CARD_GAP, availH);
  data.real.forEach((it, i) => addCard(cols[0], startY + i*(realLayout.h+realLayout.gap), it, C.gold, i+1, realLayout.h));
  data.prox.forEach((it, i) => addCard(cols[1], startY + i*(proxLayout.h+proxLayout.gap), it, C.mg,   i+1, proxLayout.h));

  // ─── Cards de reuniões ───────────────────────────────────────────────────
  const IDEAL_MEET_H = 1.72, IDEAL_MEET_GAP = 0.14;

  function addMeetCard(cx: number, rY: number, r: OprMeeting, cardH: number) {
    const cW=colW-0.22, cX=cx+0.11, cH=cardH;
    const scale = Math.min(cardH / IDEAL_MEET_H, 1);
    const fz = (v: number) => v * Math.max(scale, 0.6);
    const oy = (v: number) => rY + v * scale;

    s.addShape(pres.ShapeType.rect, { x:cX, y:rY, w:cW, h:cH, fill:{color:C.light}, line:{color:C.border, pt:.5} });
    s.addShape(pres.ShapeType.rect, { x:cX, y:rY, w:0.05, h:cH, fill:{color:C.mg}, line:{color:C.mg} });
    s.addShape(pres.ShapeType.rect, { x:cX+0.10, y:oy(0.09), w:cW-0.20, h:0.30*scale, fill:{color:C.dg}, line:{color:C.dg} });
    s.addText(`${r.dia}  ·  ${r.data}`, { x:cX+0.10, y:oy(0.09), w:cW-0.20, h:0.30*scale, fontSize:fz(7.5), bold:true, color:C.white, fontFace:FONT_BODY, align:'center', valign:'middle', margin:0 });
    s.addText(r.obj,  { x:cX+0.10, y:oy(0.46), w:cW-0.20, h:0.24*scale, fontSize:fz(8),   bold:true, color:C.dg,   fontFace:FONT_TITLE, valign:'middle', margin:0 });
    s.addText(`${r.horario}  |  ${r.local}`, { x:cX+0.10, y:oy(0.72), w:cW-0.20, h:0.18*scale, fontSize:fz(6.5), color:C.muted, fontFace:FONT_BODY, margin:0 });
    s.addShape(pres.ShapeType.line, { x:cX+0.10, y:oy(0.96), w:cW-0.20, h:0, line:{color:C.border, pt:.5} });
    s.addText('Cliente:', { x:cX+0.10, y:oy(1.02), w:0.46, h:0.17*scale, fontSize:fz(6.5), bold:true, color:C.gold, fontFace:FONT_BODY, margin:0 });
    s.addText(r.cli,      { x:cX+0.56, y:oy(1.02), w:cW-0.64, h:0.17*scale, fontSize:fz(6.5), color:C.muted, fontFace:FONT_BODY, margin:0 });
    s.addText('Pharos:', { x:cX+0.10, y:oy(1.22), w:0.46, h:0.17*scale, fontSize:fz(6.5), bold:true, color:C.mg, fontFace:FONT_BODY, margin:0 });
    s.addText(r.ph,       { x:cX+0.56, y:oy(1.22), w:cW-0.64, h:0.17*scale, fontSize:fz(6.5), color:C.muted, fontFace:FONT_BODY, margin:0 });
  }

  const meetLayout = fitCards(data.meet.length, IDEAL_MEET_H, IDEAL_MEET_GAP, availH);
  data.meet.forEach((r, i) => addMeetCard(cols[2], startY + i*(meetLayout.h+meetLayout.gap), r, meetLayout.h));

  // ─── Rodapé ──────────────────────────────────────────────────────────────
  s.addShape(pres.ShapeType.rect, { x:0, y:5.37, w:10, h:0.255, fill:{color:C.dg}, line:{color:C.dg} });
  s.addText('Pharos Consultoria  |  Uso Interno  |  Confidencial', { x:0.28, y:5.37, w:5.5, h:0.255, fontSize:7, color:'6FA882', fontFace:FONT_BODY, valign:'middle', margin:0 });
  s.addText(`Cliente: ${data.cliente}  |  Semana: ${data.semana}`, { x:5.6, y:5.37, w:4.1, h:0.255, fontSize:7, color:C.gold, fontFace:FONT_BODY, align:'right', valign:'middle', margin:0 });

  // ─── Retorna como Buffer (diferença do browser: write() em vez de writeFile()) ─
  return (await pres.write({ outputType: 'nodebuffer' })) as Buffer;
}

// Segue a convenção observada nos arquivos reais da Pharos:
// espaços no cliente viram "_", e "/" ou espaços na semana viram "-"
// (ex.: cliente "Tuckão Autopeças" + semana "03/08 a 07/08" →
//  "OPR_Tuckão_Autopeças_03-08-a-07-08.pptx").
export function buildOprFileName(cliente: string, semana: string): string {
  const stripIllegal = (s: string) => s.replace(/[<>:"|?*\\/]/g, '').trim();
  const cleanCliente = (s: string) => stripIllegal(s.replace(/\s+/g, '_'));
  const cleanSemana  = (s: string) => stripIllegal(s.replace(/[/\s]+/g, '-'));
  return `OPR_${cleanCliente(cliente)}_${cleanSemana(semana)}.pptx`;
}
