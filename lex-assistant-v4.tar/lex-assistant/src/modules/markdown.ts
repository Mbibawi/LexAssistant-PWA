/**
 * Minimal markdown → safe HTML renderer.
 * No external dependencies. Handles: headings, bold, italic,
 * inline code, code blocks, bullet lists, numbered lists, paragraphs.
 */

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function inlineFormat(s: string): string {
  return escapeHtml(s)
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/\*([^*]+)\*/g, '<em>$1</em>')
    .replace(/`([^`]+)`/g, '<code>$1</code>');
}

export function renderMarkdown(md: string): string {
  const lines = md.split('\n');
  const out: string[] = [];
  let inCode = false;
  let inUl = false;
  let inOl = false;
  let codeBuffer: string[] = [];

  function closeList() {
    if (inUl) { out.push('</ul>'); inUl = false; }
    if (inOl) { out.push('</ol>'); inOl = false; }
  }

  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];

    // Code fence
    if (raw.startsWith('```')) {
      if (!inCode) {
        closeList();
        inCode = true;
        codeBuffer = [];
      } else {
        out.push(`<pre><code>${escapeHtml(codeBuffer.join('\n'))}</code></pre>`);
        inCode = false;
        codeBuffer = [];
      }
      continue;
    }
    if (inCode) { codeBuffer.push(raw); continue; }

    // Headings
    const h3 = raw.match(/^### (.+)/);
    const h2 = raw.match(/^## (.+)/);
    const h1 = raw.match(/^# (.+)/);
    if (h1) { closeList(); out.push(`<h1>${inlineFormat(h1[1])}</h1>`); continue; }
    if (h2) { closeList(); out.push(`<h2>${inlineFormat(h2[1])}</h2>`); continue; }
    if (h3) { closeList(); out.push(`<h3>${inlineFormat(h3[1])}</h3>`); continue; }

    // Horizontal rule
    if (/^---+$/.test(raw.trim())) { closeList(); out.push('<hr>'); continue; }

    // Unordered list
    const ulMatch = raw.match(/^[-*] (.+)/);
    if (ulMatch) {
      if (!inUl) { if (inOl) { out.push('</ol>'); inOl = false; } out.push('<ul>'); inUl = true; }
      out.push(`<li>${inlineFormat(ulMatch[1])}</li>`);
      continue;
    }

    // Ordered list
    const olMatch = raw.match(/^\d+\. (.+)/);
    if (olMatch) {
      if (!inOl) { if (inUl) { out.push('</ul>'); inUl = false; } out.push('<ol>'); inOl = true; }
      out.push(`<li>${inlineFormat(olMatch[1])}</li>`);
      continue;
    }

    // Blank line
    if (raw.trim() === '') {
      closeList();
      out.push('<p class="spacer"></p>');
      continue;
    }

    // Paragraph
    closeList();
    out.push(`<p>${inlineFormat(raw)}</p>`);
  }

  closeList();
  if (inCode) out.push(`<pre><code>${escapeHtml(codeBuffer.join('\n'))}</code></pre>`);

  return out.join('\n');
}
