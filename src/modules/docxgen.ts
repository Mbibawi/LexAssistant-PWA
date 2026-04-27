/**
 * docxgen.ts
 * Generates .docx files in the browser using the docx IIFE bundle
 * loaded from /vendor/docx.umd.js (window.docx global).
 *
 * Packer API is static in docx v9: Packer.toBlob(doc), Packer.toBuffer(doc)
 */

export interface DocxGenerateOptions {
  title: string;
  content: string;
  caseRef: string;
  date?: string;
}

interface DocxLib {
  Document: new (opts: unknown) => unknown;
  Packer: { toBlob(doc: unknown): Promise<Blob>; toBuffer(doc: unknown): Promise<ArrayBuffer> };
  Paragraph: new (opts: unknown) => unknown;
  TextRun: new (opts: unknown) => unknown;
  HeadingLevel: { HEADING_1: string; HEADING_2: string; HEADING_3: string };
  AlignmentType: { JUSTIFIED: string; CENTER: string; LEFT: string; RIGHT: string };
  BorderStyle: { SINGLE: string };
}

declare global {
  interface Window { docx: DocxLib; }
}

let _loadPromise: Promise<DocxLib> | null = null;

function loadDocx(): Promise<DocxLib> {
  if (_loadPromise) return _loadPromise;
  _loadPromise = new Promise((resolve, reject) => {
    if (window.docx) { resolve(window.docx); return; }
    const s = document.createElement('script');
    s.src = '/vendor/docx.umd.js';
    s.onload = () => window.docx ? resolve(window.docx) : reject(new Error('docx IIFE non chargé'));
    s.onerror = () => reject(new Error('Impossible de charger /vendor/docx.umd.js'));
    document.head.appendChild(s);
  });
  return _loadPromise;
}

function buildParagraphs(content: string, d: DocxLib): unknown[] {
  const { Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle } = d;
  const lines = content.split('\n');
  const out: unknown[] = [];

  for (const raw of lines) {
    const line = raw.trimEnd();
    if (line.startsWith('### ')) {
      out.push(new Paragraph({ text: line.slice(4), heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 80 } }));
    } else if (line.startsWith('## ')) {
      out.push(new Paragraph({ text: line.slice(3), heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 120 } }));
    } else if (line.startsWith('# ')) {
      out.push(new Paragraph({ text: line.slice(2), heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 180 } }));
    } else if (line.trim() === '') {
      out.push(new Paragraph({ text: '', spacing: { before: 80 } }));
    } else if (/^---+$/.test(line.trim())) {
      out.push(new Paragraph({
        text: '',
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 1 } },
        spacing: { before: 120, after: 120 }
      }));
    } else {
      const parts = line.split(/(\*\*[^*]+\*\*)/);
      const runs: unknown[] = parts.map(p =>
        p.startsWith('**') && p.endsWith('**')
          ? new TextRun({ text: p.slice(2, -2), bold: true })
          : new TextRun({ text: p })
      );
      out.push(new Paragraph({
        children: runs,
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 100, after: 100 }
      }));
    }
  }
  return out;
}

export async function generateDocx(opts: DocxGenerateOptions): Promise<Blob> {
  const d = await loadDocx();
  const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle } = d;

  const date = opts.date ?? new Date().toLocaleDateString('fr-FR', {
    day: '2-digit', month: 'long', year: 'numeric'
  });

  const header = new Paragraph({
    children: [
      new TextRun({ text: opts.caseRef, bold: true, size: 18, font: 'Arial', color: '0C2340' }),
      new TextRun({ text: '    ' + date, size: 18, font: 'Arial', color: '888888' })
    ],
    alignment: AlignmentType.RIGHT,
    spacing: { after: 360 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: '185FA5', space: 6 } }
  });

  const title = new Paragraph({
    children: [new TextRun({ text: opts.title, bold: true, size: 36, font: 'Arial', color: '0C2340' })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 600 }
  });

  const body = buildParagraphs(opts.content, d);

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 24, color: '1A1A18' } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 30, bold: true, font: 'Arial', color: '0C2340' },
          paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: '185FA5' },
          paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 24, bold: true, font: 'Arial', color: '333333' },
          paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1260, bottom: 1440, left: 1700 }
        }
      },
      children: [header, title, ...body]
    }]
  });

  return Packer.toBlob(doc);
}


