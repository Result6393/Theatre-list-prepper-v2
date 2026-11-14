import { describe, expect, it } from 'vitest';
import { summaryHtml, summaryPage, summaryText, type Summary } from '../src/summary';

const summary: Summary = {
  title: 'OR 4 · AM · Smith · Tuesday 14/10',
  headings: ['Eye', 'Patient', 'Procedure', 'Diagnosis', 'Ordered', 'Implants'],
  nameColumns: [0, 1],
  rows: [
    { time: '08:00', cells: ['RIGHT', '3000001 ALPHA, Ann 70F', 'CE/IOL', '1st eye: NS2+ C1\nPXF', '', 'DCB00 +21.5D\nTarget plano'], check: '✓ Right eye: DCB00 +21.50 → predicted −1.03' },
    { time: '08:30', cells: ['LEFT', '3000002 BRAVO, Ben 65M', 'CE/IOL', '', '', ''], check: '' },
  ],
};

describe('summary of the list to send on', () => {
  it('as text: a block per patient, empty columns left out', () => {
    expect(summaryText(summary)).toBe(
      [
        'OR 4 · AM · Smith · Tuesday 14/10: 2 patients',
        '',
        '1. 08:00 · RIGHT · 3000001 ALPHA, Ann 70F',
        '   Procedure: CE/IOL',
        '   Diagnosis: 1st eye: NS2+ C1 / PXF',
        '   Implants: DCB00 +21.5D / Target plano',
        '   IOL report: ✓ Right eye: DCB00 +21.50 → predicted −1.03',
        '',
        '2. 08:30 · LEFT · 3000002 BRAVO, Ben 65M',
        '   Procedure: CE/IOL',
      ].join('\n'),
    );
  });

  it('as a table for an email, and a page to print', () => {
    const html = summaryHtml(summary);
    const header = [...html.matchAll(/<th\b[^>]*>(.*?)<\/th>/g)].map((m) => m[1]);
    expect(header).toEqual(['#', 'Time', 'Eye', 'Patient', 'Procedure', 'Diagnosis', 'Implants', 'IOL report']);
    expect(html).toContain('1st eye: NS2+ C1<br>PXF');
    const page = summaryPage(summary, 'Prepared today');
    expect(page).toContain('@page { size: A4 landscape');
    expect(page).toContain('<h1>OR 4 · AM · Smith · Tuesday 14/10</h1>');
    expect(page).toContain('2 patients · Prepared today');
  });

  it('escapes what’s typed in the cells', () => {
    const odd = { ...summary, rows: [{ ...summary.rows[0], cells: ['RIGHT', 'A <b>& B', '', '', '', ''] }] };
    expect(summaryHtml(odd)).toContain('A &lt;b&gt;&amp; B');
  });
});
