import type { DocKind, CaseDocumentMeta, LibDocumentMeta, LibDomain } from '../types.js';

const SUPPORTED_EXTS: Record<string, string> = {
  pdf:  'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  doc:  'application/msword',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls:  'application/vnd.ms-excel',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  ppt:  'application/vnd.ms-powerpoint',
  txt:  'text/plain',
  md:   'text/markdown',
  rtf:  'application/rtf',
};

export function guessMime(fileName: string, typehint = ''): string {
  if (typehint) return typehint;
  const ext = fileName.split('.').pop()?.toLowerCase() ?? '';
  return SUPPORTED_EXTS[ext] ?? 'application/octet-stream';
}

export function isSupported(name: string): boolean {
  const ext = name.split('.').pop()?.toLowerCase() ?? '';
  return ext in SUPPORTED_EXTS;
}

export function guessKind(name: string): DocKind {
  const l = name.toLowerCase();
  if (/jurisp|arrêt|arret|décision|cass|conseil.d.état/.test(l)) return 'jurisprudence';
  if (/doctrine|article|revue|doctr/.test(l)) return 'doctrine';
  return 'piece';
}

export function makeCaseDocMeta(file: File): CaseDocumentMeta {
  return {
    name:      file.name,
    kind:      guessKind(file.name),
    mimeType:  guessMime(file.name, file.type),
    sizeBytes: file.size,
    addedAt:   Date.now(),
  };
}

export function makeLibDocMeta(file: File): LibDocumentMeta {
  return {
    name:      file.name,
    mimeType:  guessMime(file.name, file.type),
    sizeBytes: file.size,
    addedAt:   Date.now(),
    tags:      [],
  };
}

export function toBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let bin = '';
  const chunk = 8192;
  for (let i = 0; i < bytes.byteLength; i += chunk) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}

export function formatSize(bytes: number): string {
  if (bytes < 1024)        return `${bytes} o`;
  if (bytes < 1024*1024)   return `${(bytes/1024).toFixed(0)} Ko`;
  return `${(bytes/1024/1024).toFixed(1)} Mo`;
}

export function kindLabel(kind: DocKind): string {
  const m: Record<DocKind,string> = { piece:'Pièce', jurisprudence:'Jurisprudence', doctrine:'Doctrine', redige:'Rédigé' };
  return m[kind];
}

export function mimeIcon(mime: string): string {
  if (mime === 'application/pdf')            return '📄';
  if (mime.includes('wordprocessingml') || mime.includes('msword'))  return '📝';
  if (mime.includes('spreadsheetml')    || mime.includes('excel'))   return '📊';
  if (mime.includes('presentationml')   || mime.includes('powerpoint')) return '📑';
  if (mime === 'text/plain' || mime === 'text/markdown') return '📃';
  return '📎';
}

export function mimeLabel(mime: string): string {
  if (mime === 'application/pdf')            return 'PDF';
  if (mime.includes('wordprocessingml') || mime.includes('msword'))  return 'Word';
  if (mime.includes('spreadsheetml')    || mime.includes('excel'))   return 'Excel';
  if (mime.includes('presentationml')   || mime.includes('powerpoint')) return 'PPT';
  if (mime === 'text/plain')  return 'TXT';
  if (mime === 'text/markdown') return 'MD';
  return 'Fichier';
}

export function uid(): string { return crypto.randomUUID(); }
