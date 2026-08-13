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
    spacing: { before: 320, after: 80 },
    border: { bottom: { color: C.gold, style: BorderStyle.SINGLE, size: 12, space: 4 } },
    children: [new TextRun({ text: text.toUpperCase(), font: F.title, color: C.medGreen, bold: true, size: 26 })],
  });
}

function bodyParagraphs(corpo: string): Paragraph[] {
  return corpo.split('\n').map(l => l.trim()).filter(Boolean).map(line => {
    const isSubtitle = /^\d+\.\d+/.test(line);
    return new Paragraph({
      spacing: { before: isSubtitle ? 160 : 80, after: 60 },
      children: [new TextRun({ text: line, font: F.body, size: 22, color: isSubtitle ? C.gold : C.body, italics: isSubtitle })],
    });
  });
}

function infoRow(label: string, value: string): TableRow {
  const cell = (text: string, bold = false) => new TableCell({
    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
    children: [new Paragraph({ children: [new TextRun({ text, font: F.body, size: 22, bold, color: bold ? C.medGreen : C.body })] })],
  });
  return new TableRow({ children: [cell(label, true), cell(value)] });
}

// ─── Header com papel timbrado ───────────────────────────────────────────────

function buildHeader(): Header {
  const bgPath = path.join(process.cwd(), 'assets', 'timbrado-bg.png');
  if (!fs.existsSync(bgPath)) {
    return new Header({ children: [new Paragraph({ children: [] })] });
  }
  const bgData = fs.readFileSync(bgPath);
  return new Header({
    children: [
      new Paragraph({
        children: [
          new ImageRun({
            data: bgData,
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

  // — Título
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'ATA DE REUNIÃO', font: F.title, color: C.darkGreen, bold: true, size: 40 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 320 },
      border: { bottom: { color: C.gold, style: BorderStyle.SINGLE, size: 18, space: 4 } },
      children: [new TextRun({ text: ata.assunto, font: F.title, color: C.medGreen, italics: true, size: 26 })],
    }),
  );

  // — Informações
  children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      infoRow('Projeto / Cliente:', ata.projeto),
      infoRow('Data:', ata.data),
      infoRow('Duração:', ata.duracao),
      infoRow('Modalidade:', ata.modalidade),
    ],
  }));

  // — Participantes
  children.push(sectionTitle('Participantes'));
  const presHeader = new TableRow({
    tableHeader: true,
    children: ['Nome', 'Empresa', 'Papel'].map(h => new TableCell({
      shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen },
      children: [new Paragraph({ children: [new TextRun({ text: h, font: F.body, color: C.white, bold: true, size: 20 })] })],
    })),
  });
  const presRows = ata.presentes.map((p, i) => new TableRow({
    children: [p.nome, p.empresa, p.papel].map(v => new TableCell({
      shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite },
      children: [new Paragraph({ children: [new TextRun({ text: v, font: F.body, size: 20, color: C.body })] })],
    })),
  }));
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [presHeader, ...presRows] }));

  if (ata.observacao_participantes) {
    children.push(new Paragraph({
      spacing: { before: 80 },
      children: [new TextRun({ text: `Obs.: ${ata.observacao_participantes}`, font: F.body, size: 20, italics: true, color: C.body })],
    }));
  }

  // — Seções de conteúdo (exclui "decisões/alinhamentos" — vai formatada separado)
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
        spacing: { before: 80, after: 40 },
        bullet: { level: 0 },
        children: [new TextRun({ text: d, font: F.body, size: 22, color: C.body })],
      }));
    });
  }

  // — Encaminhamentos
  if (ata.encaminhamentos.length > 0) {
    children.push(sectionTitle('Encaminhamentos'));
    const encHeader = new TableRow({
      tableHeader: true,
      children: ['#', 'Encaminhamento', 'Responsável', 'Situação'].map(h => new TableCell({
        shading: { fill: C.darkGreen, type: ShadingType.SOLID, color: C.darkGreen },
        children: [new Paragraph({ children: [new TextRun({ text: h, font: F.body, color: C.white, bold: true, size: 20 })] })],
      })),
    });
    const encRows = ata.encaminhamentos.map((e, i) => new TableRow({
      children: [
        new TableCell({
          shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite },
          children: [new Paragraph({ children: [new TextRun({ text: String(e.numero), font: F.body, bold: true, color: C.gold, size: 20 })] })],
        }),
        ...[e.encaminhamento, e.responsavel, e.situacao].map(v => new TableCell({
          shading: { fill: i % 2 === 0 ? C.white : C.offWhite, type: ShadingType.SOLID, color: i % 2 === 0 ? C.white : C.offWhite },
          children: [new Paragraph({ children: [new TextRun({ text: v, font: F.body, size: 20, color: C.body })] })],
        })),
      ],
    }));
    children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [encHeader, ...encRows] }));
  }

  // — Pontos a confirmar
  if (ata.pontos_a_confirmar.length > 0) {
    children.push(sectionTitle('Pontos a Confirmar'));
    ata.pontos_a_confirmar.forEach(p => {
      children.push(new Paragraph({
        spacing: { before: 80, after: 40 },
        bullet: { level: 0 },
        children: [new TextRun({ text: p, font: F.body, size: 22, color: C.gold, italics: true })],
      }));
    });
  }

  // ─── Documento ──────────────────────────────────────────────────────────

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
              children: [new TextRun({ text: 'Ata elaborada por Pharos Consultoria.', font: F.body, size: 16, color: C.darkGreen, italics: true })],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: 'Você define o destino, nós mostramos o caminho.', font: F.title, size: 16, color: C.gold, italics: true })],
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
