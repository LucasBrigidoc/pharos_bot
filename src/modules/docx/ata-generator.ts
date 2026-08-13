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

// ─── Paleta e fontes ────────────────────────────────────────────────────────

const C = {
  darkGreen: '072E22',
  medGreen:  '14663D',
  gold:      'AF7932',
  offWhite:  'F0F4F0',
  white:     'FFFFFF',
  body:      '111111',
};
const F = { title: 'Playfair Display', body: 'Helvetica Now' };

// Bordas invisíveis reutilizáveis
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder };

// ─── Schema ─────────────────────────────────────────────────────────────────

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
  encaminhamentos: { numero: number; encaminhamento: string; responsavel: string; situacao: string }[];
  pontos_a_confirmar: string[];
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function sectionTitle(text: string): Paragraph {
  return new Paragraph({
    spacing: { before: 400, after: 120 },
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
      indent: { left: 0 },
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

function infoRow(label: string, value: string): TableRow {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 32, type: WidthType.PERCENTAGE },
        borders: noBorders,
        margins: { top: 60, bottom: 60, left: 0, right: 200 },
        children: [new Paragraph({
          children: [new TextRun({ text: label, font: F.body, size: 24, bold: true, color: C.medGreen })],
        })],
      }),
      new TableCell({
        width: { size: 68, type: WidthType.PERCENTAGE },
        borders: noBorders,
        margins: { top: 60, bottom: 60, left: 0, right: 0 },
        children: [new Paragraph({
          children: [new TextRun({ text: value, font: F.body, size: 24, color: C.body })],
        })],
      }),
    ],
  });
}

function headerCell(text: string): TableCell {
  return new TableCell({
    shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen },
    margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: F.body, color: C.white, bold: true, size: 22 })],
    })],
  });
}

function dataCell(text: string, rowIndex: number, color = C.body, bold = false): TableCell {
  const fill = rowIndex % 2 === 0 ? C.white : C.offWhite;
  return new TableCell({
    shading: { fill, type: ShadingType.SOLID, color: fill },
    margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: F.body, size: 22, color, bold })],
    })],
  });
}

// ─── Header com papel timbrado ───────────────────────────────────────────────

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

// ─── Generator principal ─────────────────────────────────────────────────────

export async function generateAtaDocx(ata: AtaJSON): Promise<Buffer> {
  const children: (Paragraph | Table)[] = [];

  // — Espaço superior para não sobrepor o timbrado
  children.push(new Paragraph({ spacing: { before: 0, after: 480 }, children: [] }));

  // — Título principal
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

  // — Bloco de informações
  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
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

  // — Participantes
  children.push(sectionTitle('Participantes'));
  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            { text: 'Nome',    width: 40 },
            { text: 'Empresa', width: 35 },
            { text: 'Papel',   width: 25 },
          ].map(h => {
            const cell = headerCell(h.text);
            (cell as any).options = { ...((cell as any).options ?? {}), width: { size: h.width, type: WidthType.PERCENTAGE } };
            return new TableCell({
              width: { size: h.width, type: WidthType.PERCENTAGE },
              shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen },
              margins: { top: 120, bottom: 120, left: 160, right: 160 },
              children: [new Paragraph({ children: [new TextRun({ text: h.text, font: F.body, color: C.white, bold: true, size: 22 })] })],
            });
          }),
        }),
        ...ata.presentes.map((p, i) => new TableRow({
          children: [
            new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: p.nome,    font: F.body, size: 22, color: C.body })] })] }),
            new TableCell({ width: { size: 35, type: WidthType.PERCENTAGE }, shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: p.empresa, font: F.body, size: 22, color: C.body })] })] }),
            new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: p.papel,   font: F.body, size: 22, color: C.body })] })] }),
          ],
        })),
      ],
    }),
  );

  if (ata.observacao_participantes) {
    children.push(new Paragraph({
      spacing: { before: 100, after: 0 },
      children: [new TextRun({ text: `Obs.: ${ata.observacao_participantes}`, font: F.body, size: 22, italics: true, color: C.body })],
    }));
  }

  // — Seções de conteúdo
  const secoes = ata.secoes.filter(s =>
    !s.titulo.toLowerCase().includes('decis') && !s.titulo.toLowerCase().includes('alinhamento'),
  );
  secoes.forEach((sec, idx) => {
    children.push(sectionTitle(`${idx + 1}. ${sec.titulo}`));
    children.push(...bodyParagraphs(sec.corpo));
  });

  // — Decisões e Alinhamentos
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

  // — Encaminhamentos
  if (ata.encaminhamentos.length > 0) {
    children.push(sectionTitle('Encaminhamentos'));
    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({ width: { size: 8,  type: WidthType.PERCENTAGE }, shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: '#',            font: F.body, color: C.white, bold: true, size: 22 })] })] }),
              new TableCell({ width: { size: 47, type: WidthType.PERCENTAGE }, shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: 'Encaminhamento', font: F.body, color: C.white, bold: true, size: 22 })] })] }),
              new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: 'Responsável',   font: F.body, color: C.white, bold: true, size: 22 })] })] }),
              new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: 'Situação',      font: F.body, color: C.white, bold: true, size: 22 })] })] }),
            ],
          }),
          ...ata.encaminhamentos.map((e, i) => new TableRow({
            children: [
              new TableCell({ width: { size: 8,  type: WidthType.PERCENTAGE }, shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: String(e.numero), font: F.body, bold: true, color: C.gold, size: 22 })] })] }),
              new TableCell({ width: { size: 47, type: WidthType.PERCENTAGE }, shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: e.encaminhamento, font: F.body, size: 22, color: C.body })] })] }),
              new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: e.responsavel,    font: F.body, size: 22, color: C.body })] })] }),
              new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: e.situacao,       font: F.body, size: 22, color: C.body })] })] }),
            ],
          })),
        ],
      }),
    );
  }

  // — Pontos a confirmar
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

  // ─── Documento ───────────────────────────────────────────────────────────

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
              children: [new TextRun({ text: 'Ata elaborada por Pharos Consultoria.', font: F.body, size: 18, color: C.darkGreen, italics: true })],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: 'Você define o destino, nós mostramos o caminho.', font: F.title, size: 18, color: C.gold, italics: true })],
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
