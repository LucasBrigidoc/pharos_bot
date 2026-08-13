import * as fs from 'fs';
import * as path from 'path';
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HorizontalPositionRelativeFrom,
  ImageRun,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  TextWrappingType,
  VerticalPositionRelativeFrom,
  WidthType,
} from 'docx';

// ─── Medidas ─────────────────────────────────────────────────────────────────
// A4: 11906 × 16838 twips | margens: top 2268, right 1417, bottom 1531, left 1417
// Largura útil = 11906 - 1417 - 1417 = 9072 twips
// Usando WidthType.DXA (twips) em vez de PERCENTAGE para evitar bug do docx
// onde PERCENTAGE usa "fiftieths of a percent" e size:32 vira 0,64% na prática.

const W = 9072;

const INFO_COL  = [2900, W - 2900]                              as const; // 32% / 68%
const PRES_COL  = [3629, 3175, W - 3629 - 3175]                as const; // 40% / 35% / 25%
const ENC_COL   = [726, 4264, 2268, W - 726 - 4264 - 2268]    as const; // 8%/47%/25%/20%

// ─── Paleta e fontes ─────────────────────────────────────────────────────────

const C = {
  darkGreen: '072E22',
  medGreen:  '14663D',
  gold:      'AF7932',
  offWhite:  'F0F4F0',
  white:     'FFFFFF',
  body:      '111111',
};
const F = { title: 'Playfair Display', body: 'Helvetica Now' };

const noBorder  = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = {
  top: noBorder, bottom: noBorder,
  left: noBorder, right: noBorder,
  insideHorizontal: noBorder, insideVertical: noBorder,
};

// ─── Schema ──────────────────────────────────────────────────────────────────

export interface AtaJSON {
  projeto: string;
  assunto: string;
  data: string;
  duracao: string;
  modalidade: string;
  presentes: { nome: string; empresa: string; papel: string }[];
  observacao_participantes: string | null;
  secoes: { titulo: string; corpo: string }[];
  decisoes_alinhamentos: string[];
  encaminhamentos: {
    numero: number;
    encaminhamento: string;
    responsavel: string;
    situacao: string;
  }[];
  pontos_a_confirmar: string[];
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function sectionTitle(text: string): Paragraph {
  return new Paragraph({
    spacing: { before: 400, after: 140 },
    border: { bottom: { color: C.gold, style: BorderStyle.SINGLE, size: 12, space: 6 } },
    children: [new TextRun({
      text: text.toUpperCase(),
      font: F.title,
      color: C.medGreen,
      bold: true,
      size: 28,
    })],
  });
}

function bodyParagraphs(corpo: string): Paragraph[] {
  return corpo.split('\n').map(l => l.trim()).filter(Boolean).map(line => {
    const isSubtitle = /^\d+\.\d+/.test(line);
    return new Paragraph({
      spacing: { before: isSubtitle ? 200 : 100, after: 80 },
      children: [new TextRun({
        text: line,
        font: F.body,
        size: 24,
        color: isSubtitle ? C.gold : C.body,
        italics: isSubtitle,
        bold: isSubtitle,
      })],
    });
  });
}

function hCell(text: string, width: number): TableCell {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen },
    margins: { top: 120, bottom: 120, left: 180, right: 180 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: F.body, color: C.white, bold: true, size: 22 })],
    })],
  });
}

function dCell(
  text: string,
  width: number,
  rowIdx: number,
  color = C.body,
  bold = false,
): TableCell {
  const fill = rowIdx % 2 === 0 ? C.white : C.offWhite;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.SOLID, color: fill },
    margins: { top: 120, bottom: 120, left: 180, right: 180 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: F.body, size: 22, color, bold })],
    })],
  });
}

function infoRow(label: string, value: string): TableRow {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: INFO_COL[0], type: WidthType.DXA },
        borders: noBorders,
        margins: { top: 80, bottom: 80, left: 0, right: 240 },
        children: [new Paragraph({
          children: [new TextRun({ text: label, font: F.body, size: 24, bold: true, color: C.medGreen })],
        })],
      }),
      new TableCell({
        width: { size: INFO_COL[1], type: WidthType.DXA },
        borders: noBorders,
        margins: { top: 80, bottom: 80, left: 0, right: 0 },
        children: [new Paragraph({
          children: [new TextRun({ text: value, font: F.body, size: 24, color: C.body })],
        })],
      }),
    ],
  });
}

// ─── Header com papel timbrado ────────────────────────────────────────────────

function buildHeader(): Header {
  const bgPath = path.join(process.cwd(), 'assets', 'timbrado-bg.png');
  if (!fs.existsSync(bgPath)) {
    return new Header({ children: [new Paragraph({ children: [] })] });
  }
  return new Header({
    children: [
      new Paragraph({
        children: [
          new ImageRun({
            data: fs.readFileSync(bgPath),
            transformation: { width: 794, height: 1123 },
            floating: {
              horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
              verticalPosition:   { relative: VerticalPositionRelativeFrom.PAGE,   offset: 0 },
              behindDocument: true,
              allowOverlap: true,
              lockAnchor: false,
              wrap: { type: TextWrappingType.NONE },
            },
          }),
        ],
      }),
    ],
  });
}

// ─── Gerador principal ────────────────────────────────────────────────────────

export async function generateAtaDocx(ata: AtaJSON): Promise<Buffer> {
  const children: (Paragraph | Table)[] = [];

  children.push(new Paragraph({ spacing: { before: 0, after: 480 }, children: [] }));

  // Título
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 100 },
      children: [new TextRun({ text: 'ATA DE REUNIÃO', font: F.title, color: C.darkGreen, bold: true, size: 44 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 400 },
      border: { bottom: { color: C.gold, style: BorderStyle.SINGLE, size: 18, space: 6 } },
      children: [new TextRun({ text: ata.assunto, font: F.title, color: C.medGreen, italics: true, size: 28 })],
    }),
  );

  // Bloco de informações
  children.push(
    new Table({
      width: { size: W, type: WidthType.DXA },
      borders: noBorders,
      rows: [
        infoRow('Projeto / Cliente:', ata.projeto),
        infoRow('Data:', ata.data),
        infoRow('Duração:', ata.duracao),
        infoRow('Modalidade:', ata.modalidade),
      ],
    }),
    new Paragraph({ spacing: { before: 0, after: 200 }, children: [] }),
  );

  // Participantes
  children.push(sectionTitle('Participantes'));
  children.push(
    new Table({
      width: { size: W, type: WidthType.DXA },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            hCell('Nome',    PRES_COL[0]),
            hCell('Empresa', PRES_COL[1]),
            hCell('Papel',   PRES_COL[2]),
          ],
        }),
        ...ata.presentes.map((p, i) => new TableRow({
          children: [
            dCell(p.nome,    PRES_COL[0], i),
            dCell(p.empresa, PRES_COL[1], i),
            dCell(p.papel,   PRES_COL[2], i),
          ],
        })),
      ],
    }),
  );

  if (ata.observacao_participantes) {
    children.push(new Paragraph({
      spacing: { before: 100, after: 0 },
      children: [new TextRun({
        text: `Obs.: ${ata.observacao_participantes}`,
        font: F.body, size: 22, italics: true, color: C.body,
      })],
    }));
  }

  // Seções de conteúdo
  const secoes = ata.secoes.filter(s =>
    !s.titulo.toLowerCase().includes('decis') &&
    !s.titulo.toLowerCase().includes('alinhamento'),
  );
  secoes.forEach((sec, idx) => {
    children.push(sectionTitle(`${idx + 1}. ${sec.titulo}`));
    children.push(...bodyParagraphs(sec.corpo));
  });

  // Decisões e Alinhamentos
  if (ata.decisoes_alinhamentos.length > 0) {
    children.push(sectionTitle('Decisões e Alinhamentos'));
    ata.decisoes_alinhamentos.forEach(d => {
      children.push(new Paragraph({
        spacing: { before: 100, after: 60 },
        bullet: { level: 0 },
        indent: { left: 360 },
        children: [new TextRun({ text: d, font: F.body, size: 24, color: C.body })],
      }));
    });
  }

  // Encaminhamentos
  if (ata.encaminhamentos.length > 0) {
    children.push(sectionTitle('Encaminhamentos'));
    children.push(
      new Table({
        width: { size: W, type: WidthType.DXA },
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              hCell('#',              ENC_COL[0]),
              hCell('Encaminhamento', ENC_COL[1]),
              hCell('Responsável',    ENC_COL[2]),
              hCell('Situação',       ENC_COL[3]),
            ],
          }),
          ...ata.encaminhamentos.map((e, i) => new TableRow({
            children: [
              dCell(String(e.numero), ENC_COL[0], i, C.gold, true),
              dCell(e.encaminhamento, ENC_COL[1], i),
              dCell(e.responsavel,    ENC_COL[2], i),
              dCell(e.situacao,       ENC_COL[3], i),
            ],
          })),
        ],
      }),
    );
  }

  // Pontos a confirmar
  if (ata.pontos_a_confirmar.length > 0) {
    children.push(sectionTitle('Pontos a Confirmar'));
    ata.pontos_a_confirmar.forEach(p => {
      children.push(new Paragraph({
        spacing: { before: 100, after: 60 },
        bullet: { level: 0 },
        indent: { left: 360 },
        children: [new TextRun({ text: p, font: F.body, size: 24, color: C.gold, italics: true })],
      }));
    });
  }

  // ─── Documento ──────────────────────────────────────────────────────────────

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 2268, right: 1417, bottom: 1531, left: 1417 },
        },
      },
      headers: { default: buildHeader() },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({
                text: 'Ata elaborada por Pharos Consultoria.',
                font: F.body, size: 18, color: C.darkGreen, italics: true,
              })],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({
                text: 'Você define o destino, nós mostramos o caminho.',
                font: F.title, size: 18, color: C.gold, italics: true,
              })],
            }),
          ],
        }),
      },
      children,
    }],
  });

  return Buffer.from(await Packer.toBuffer(doc));
}

export function buildAtaFileName(projeto: string, data: string, assunto: string): string {
  const clean = (s: string) => s.replace(/[<>:"/\\|?*\n\r]/g, '').trim().slice(0, 50);
  return `Ata ${clean(projeto)} - ${data.replace(/\//g, '-')} - ${clean(assunto)}.docx`;
}
