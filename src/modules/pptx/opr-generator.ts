// Portado de opr_pharos_gera_ppt_local.html (função gerarPPT)
// Única mudança em relação ao original: pres.write() retorna Buffer
// em vez de pres.writeFile() que baixava no browser.

import pptxgen from 'pptxgenjs';

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

  // Logo PHAROS
  s.addText('PHAROS', { x:0.28, y:0.07, w:1.8, h:0.40, fontSize:14, bold:true, color:C.white, fontFace:'Georgia', charSpacing:4, valign:'middle', margin:0 });
  s.addShape(pres.ShapeType.ellipse, { x:1.43, y:0.22, w:0.07, h:0.07, fill:{color:C.gold}, line:{color:C.gold} });
  s.addText('Consultoria', { x:1.54, y:0.21, w:1.0, h:0.16, fontSize:6.5, color:'8FBF9A', fontFace:'Trebuchet MS', margin:0 });

  // Pill OPR
  s.addShape(pres.ShapeType.rect, { x:4.1, y:0.09, w:1.8, h:0.35, fill:{color:C.gold}, line:{color:C.gold} });
  s.addText('OPR  –  One Page Report', { x:4.1, y:0.09, w:1.8, h:0.35, fontSize:7.5, bold:true, color:C.dg, fontFace:'Trebuchet MS', align:'center', valign:'middle', margin:0 });

  // Semana + CONFIDENCIAL
  s.addText(`Semana: ${data.semana}  |  Cliente: ${data.cliente}`, { x:6.2, y:0.07, w:3.55, h:0.20, fontSize:7, color:'A8D0B4', fontFace:'Trebuchet MS', align:'right', margin:0 });
  s.addText('CONFIDENCIAL', { x:6.2, y:0.28, w:3.55, h:0.18, fontSize:6.5, bold:true, color:C.gold, fontFace:'Trebuchet MS', align:'right', charSpacing:2, margin:0 });

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
    s.addText(headers[i][0], { x:cx, y:colY, w:colW, h:0.30, fontSize:7, bold:true, color:C.white, fontFace:'Trebuchet MS', align:'center', valign:'middle', margin:0, charSpacing:.5 });
  });

  // ─── Cards de atividades ─────────────────────────────────────────────────
  function addCard(cx: number, cardY: number, item: OprItem, accentColor: string, num: number) {
    const cW=colW-0.22, cX=cx+0.11, cH=1.22;
    s.addShape(pres.ShapeType.rect,    { x:cX, y:cardY, w:cW, h:cH, fill:{color:C.off}, line:{color:C.border, pt:.5} });
    s.addShape(pres.ShapeType.rect,    { x:cX, y:cardY, w:0.05, h:cH, fill:{color:accentColor}, line:{color:accentColor} });
    s.addShape(pres.ShapeType.ellipse, { x:cX+0.10, y:cardY+0.09, w:0.22, h:0.22, fill:{color:accentColor}, line:{color:accentColor} });
    s.addText(String(num),  { x:cX+0.10, y:cardY+0.09, w:0.22, h:0.22, fontSize:7, bold:true, color:C.white, fontFace:'Trebuchet MS', align:'center', valign:'middle', margin:0 });
    s.addText(item.ativ,    { x:cX+0.36, y:cardY+0.08, w:cW-0.44, h:0.26, fontSize:7.5, bold:true, color:C.dg, fontFace:'Georgia', valign:'middle', margin:0 });
    s.addText(item.obj,     { x:cX+0.10, y:cardY+0.34, w:cW-0.18, h:0.26, fontSize:7, color:C.muted, fontFace:'Trebuchet MS', valign:'top', margin:0 });
    s.addShape(pres.ShapeType.line, { x:cX+0.10, y:cardY+0.64, w:cW-0.20, h:0, line:{color:C.border, pt:.5} });
    s.addText('Cliente:', { x:cX+0.10, y:cardY+0.70, w:0.46, h:0.17, fontSize:6.5, bold:true, color:C.gold, fontFace:'Trebuchet MS', margin:0 });
    s.addText(item.cli,   { x:cX+0.56, y:cardY+0.70, w:cW-0.64, h:0.17, fontSize:6.5, color:C.muted, fontFace:'Trebuchet MS', margin:0 });
    s.addText('Pharos:', { x:cX+0.10, y:cardY+0.89, w:0.46, h:0.17, fontSize:6.5, bold:true, color:C.mg, fontFace:'Trebuchet MS', margin:0 });
    s.addText(item.ph,   { x:cX+0.56, y:cardY+0.89, w:cW-0.64, h:0.17, fontSize:6.5, color:C.muted, fontFace:'Trebuchet MS', margin:0 });
  }

  const startY = colY+0.36, cardGap = 0.06;
  data.real.forEach((it, i) => addCard(cols[0], startY + i*(1.22+cardGap), it, C.gold, i+1));
  data.prox.forEach((it, i) => addCard(cols[1], startY + i*(1.22+cardGap), it, C.mg,   i+1));

  // ─── Cards de reuniões ───────────────────────────────────────────────────
  data.meet.forEach((r, i) => {
    const cx=cols[2], rY=startY+i*1.86, cW=colW-0.22, cX=cx+0.11, cH=1.72;
    s.addShape(pres.ShapeType.rect, { x:cX, y:rY, w:cW, h:cH, fill:{color:C.light}, line:{color:C.border, pt:.5} });
    s.addShape(pres.ShapeType.rect, { x:cX, y:rY, w:0.05, h:cH, fill:{color:C.mg}, line:{color:C.mg} });
    s.addShape(pres.ShapeType.rect, { x:cX+0.10, y:rY+0.09, w:cW-0.20, h:0.30, fill:{color:C.dg}, line:{color:C.dg} });
    s.addText(`${r.dia}  ·  ${r.data}`, { x:cX+0.10, y:rY+0.09, w:cW-0.20, h:0.30, fontSize:7.5, bold:true, color:C.white, fontFace:'Trebuchet MS', align:'center', valign:'middle', margin:0 });
    s.addText(r.obj,  { x:cX+0.10, y:rY+0.46, w:cW-0.20, h:0.24, fontSize:8,   bold:true, color:C.dg,   fontFace:'Georgia', valign:'middle', margin:0 });
    s.addText(`${r.horario}  |  ${r.local}`, { x:cX+0.10, y:rY+0.72, w:cW-0.20, h:0.18, fontSize:6.5, color:C.muted, fontFace:'Trebuchet MS', margin:0 });
    s.addShape(pres.ShapeType.line, { x:cX+0.10, y:rY+0.96, w:cW-0.20, h:0, line:{color:C.border, pt:.5} });
    s.addText('Cliente:', { x:cX+0.10, y:rY+1.02, w:0.46, h:0.17, fontSize:6.5, bold:true, color:C.gold, fontFace:'Trebuchet MS', margin:0 });
    s.addText(r.cli,      { x:cX+0.56, y:rY+1.02, w:cW-0.64, h:0.17, fontSize:6.5, color:C.muted, fontFace:'Trebuchet MS', margin:0 });
    s.addText('Pharos:', { x:cX+0.10, y:rY+1.22, w:0.46, h:0.17, fontSize:6.5, bold:true, color:C.mg, fontFace:'Trebuchet MS', margin:0 });
    s.addText(r.ph,       { x:cX+0.56, y:rY+1.22, w:cW-0.64, h:0.17, fontSize:6.5, color:C.muted, fontFace:'Trebuchet MS', margin:0 });
  });

  // ─── Rodapé ──────────────────────────────────────────────────────────────
  s.addShape(pres.ShapeType.rect, { x:0, y:5.37, w:10, h:0.255, fill:{color:C.dg}, line:{color:C.dg} });
  s.addText('Pharos Consultoria  |  Uso Interno  |  Confidencial', { x:0.28, y:5.37, w:5.5, h:0.255, fontSize:7, color:'6FA882', fontFace:'Trebuchet MS', valign:'middle', margin:0 });
  s.addText(`Cliente: ${data.cliente}  |  Semana: ${data.semana}`, { x:5.6, y:5.37, w:4.1, h:0.255, fontSize:7, color:C.gold, fontFace:'Trebuchet MS', align:'right', valign:'middle', margin:0 });

  // ─── Retorna como Buffer (diferença do browser: write() em vez de writeFile()) ─
  return (await pres.write({ outputType: 'nodebuffer' })) as Buffer;
}

export function buildOprFileName(cliente: string, semana: string): string {
  const clean = (s: string) => s.replace(/[\s/\\]/g, '-').replace(/[<>:"|?*]/g, '').trim();
  return `OPR_${clean(cliente)}_${clean(semana)}.pptx`;
}
