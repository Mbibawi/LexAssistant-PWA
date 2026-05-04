export function formatSize(bytes: number): string {
  if (bytes < 1024)        return `${bytes} o`;
  if (bytes < 1024*1024)   return `${(bytes/1024).toFixed(0)} Ko`;
  return `${(bytes/1024/1024).toFixed(1)} Mo`;
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
