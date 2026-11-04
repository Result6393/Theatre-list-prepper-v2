/**
 * Tab-separated text in the form Excel expects on paste: cells containing
 * tabs, line breaks or quotes are quoted so they stay in one cell.
 * Plain text (rather than HTML) keeps the target sheet's own formatting.
 */
export function toTsv(rows: string[][]): string {
  return rows.map((row) => row.map(quoteCell).join('\t')).join('\r\n');
}

function quoteCell(value: string): string {
  const text = value.replace(/\r\n?/g, '\n');
  return /[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

export async function copyText(text: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(text);
    return;
  } catch {
    // Fall back for browsers that block the async clipboard API (e.g. some file:// pages).
  }
  const area = document.createElement('textarea');
  area.value = text;
  area.style.position = 'fixed';
  area.style.opacity = '0';
  document.body.append(area);
  area.select();
  const ok = document.execCommand('copy');
  area.remove();
  if (!ok) throw new Error('Copy was blocked by the browser.');
}

/**
 * Copies formatted text. Apps that understand HTML (like Outlook) paste it with its line structure;
 * everything else gets the plain text. Falls back to a copy event for browsers without the async clipboard.
 */
export async function copyHtml(html: string, text: string): Promise<void> {
  try {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([text], { type: 'text/plain' }),
      }),
    ]);
    return;
  } catch {
    // Not supported or blocked: try the older route below.
  }
  const onCopy = (e: ClipboardEvent) => {
    e.preventDefault();
    e.clipboardData?.setData('text/html', html);
    e.clipboardData?.setData('text/plain', text);
  };
  document.addEventListener('copy', onCopy, { once: true });
  // Some browsers only fire the copy event when something is selected.
  const area = document.createElement('textarea');
  area.value = text;
  area.style.position = 'fixed';
  area.style.opacity = '0';
  document.body.append(area);
  area.select();
  const ok = document.execCommand('copy');
  area.remove();
  document.removeEventListener('copy', onCopy);
  if (!ok) throw new Error('Copy was blocked by the browser.');
}

/**
 * Cells of the first table in HTML copied from Excel (browser only). Used when the plain text
 * didn't come through cleanly. Line breaks in a cell are <br>; other whitespace is collapsed.
 */
export function htmlTableRows(html: string): string[][] | null {
  const table = new DOMParser().parseFromString(html, 'text/html').querySelector('table');
  if (!table) return null;
  const text = (node: Node): string => {
    if (node.nodeType === Node.TEXT_NODE) return (node.textContent ?? '').replace(/\s+/g, ' ');
    if (node.nodeName === 'BR') return '\n';
    return [...node.childNodes].map(text).join('');
  };
  return [...table.querySelectorAll('tr')].map((tr) =>
    [...tr.querySelectorAll('td, th')].flatMap((cell) => {
      const value = text(cell)
        .split('\n')
        .map((line) => line.trim())
        .join('\n')
        .trim();
      const span = Math.max(1, Number((cell as HTMLTableCellElement).colSpan) || 1);
      return [value, ...Array.from({ length: span - 1 }, () => '')];
    }),
  );
}
