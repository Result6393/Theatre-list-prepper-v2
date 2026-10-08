import Papa from 'papaparse';

export interface Table {
  headers: string[];
  rows: Record<string, string>[];
}

export class ParseError extends Error {}

// Columns that identify a SurgiNet theatre run sheet (SN006A export).
const REQUIRED_HEADERS = ['THEATRE', 'PT_NAME'];

/**
 * Parses a run sheet given as CSV (from a file) or tab-separated text
 * (pasted from Excel). Headers are found by name, so column order and any
 * lines above the header row don't matter.
 */
export function parseTable(text: string): Table {
  const delimiter = countOf(text, '\t') > countOf(text, ',') ? '\t' : ',';
  const result = Papa.parse<string[]>(text.replace(/^\uFEFF/, ''), {
    delimiter,
    skipEmptyLines: 'greedy',
  });
  const data = result.data;

  const headerIndex = data.findIndex((row) => {
    const names = row.map(normaliseHeader);
    return REQUIRED_HEADERS.every((h) => names.includes(h));
  });
  if (headerIndex === -1) {
    throw new ParseError(
      'Couldn’t find the run sheet headings (e.g. THEATRE, PT_NAME). Paste the whole table including its heading row.',
    );
  }

  const headers = data[headerIndex].map(normaliseHeader);
  const rows = data.slice(headerIndex + 1).map((cells) => {
    const row: Record<string, string> = {};
    headers.forEach((h, i) => {
      if (h) row[h] = (cells[i] ?? '').trim();
    });
    return row;
  });
  return { headers: headers.filter(Boolean), rows };
}

function normaliseHeader(h: string): string {
  return h.trim().toUpperCase().replace(/\s+/g, '_');
}

function countOf(text: string, char: string): number {
  return text.split(char).length - 1;
}
