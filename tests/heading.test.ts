import { describe, expect, it } from 'vitest';
import { headingCells } from '../src/heading';
import { defaultLayout } from '../src/layout';

const layout = defaultLayout();
const values = { consultant: 'Willow', date: '15/10/2026', or: 'OR 4', ampm: 'AM' };
const columns = ['', ...layout.columns.map((c) => c.heading)];
const sheet = (rows: string[][]) => ({ name: 'S', rows });
const simple = (cells: ReturnType<typeof headingCells>) => cells.map((c) => [c.label, c.cell, c.value]);

describe('filling in a sheet’s heading', () => {
  it('writes each value in the cell to the right of its label, keeping a real Excel date a date', () => {
    const rows = [
      ['', 'Consultant', 'Cedar', '', 'Date', '46309'],
      ['', 'OR', '2', '', 'AM/PM', 'PM'],
      [],
      columns,
    ];
    expect(simple(headingCells(sheet(rows), 4, values, layout))).toEqual([
      ['Consultant', 'C1', 'Willow'],
      ['Date', 'F1', '46310'],
      ['OR', 'C2', '4'],
      ['AM/PM', 'F2', 'AM'],
    ]);
  });

  it('fills empty cells in a blank template, with labels ending in a colon', () => {
    const rows = [['Consultant:', '', 'Date:', '', 'Theatre:', '', 'AM/PM:', ''], [], columns];
    expect(simple(headingCells(sheet(rows), 3, values, layout))).toEqual([
      ['Consultant', 'B1', 'Willow'],
      ['Date', 'D1', '15/10/2026'],
      ['Theatre', 'F1', '4'],
      ['AM/PM', 'H1', 'AM'],
    ]);
  });

  it('rewrites a label and value written in one cell', () => {
    const rows = [['', 'Consultant: Cedar', 'Date: 14/10/2026', 'OR: OR 2'], columns];
    expect(simple(headingCells(sheet(rows), 2, values, layout))).toEqual([
      ['Consultant', 'B1', 'Consultant: Willow'],
      ['Date', 'C1', 'Date: 15/10/2026'],
      ['OR', 'D1', 'OR: OR 4'],
    ]);
  });

  it('keeps “OR 4” where the sheet writes the OR that way', () => {
    const rows = [['OR', 'OR 2'], columns];
    expect(simple(headingCells(sheet(rows), 2, values, layout))).toEqual([['OR', 'B1', 'OR 4']]);
  });

  it('never writes over the list’s column headings, other labels, or anything below the list', () => {
    const rows = [['Date', 'Consultant', ''], columns, ['', 'Consultant', 'x']];
    const cells = headingCells(sheet(rows), 2, values, layout);
    expect(simple(cells)).toEqual([['Consultant', 'C1', 'Willow']]);
    expect(headingCells(sheet([columns]), 1, values, layout)).toEqual([]);
  });

  it('puts the value after a label merged across cells', () => {
    const rows = [['', 'Consultant', '', 'Cedar', 'Date', '', ''], columns];
    const merged = { name: 'S', rows, merges: ['B1:C1', 'E1:F1'] };
    expect(simple(headingCells(merged, 2, values, layout))).toEqual([
      ['Consultant', 'D1', 'Willow'],
      ['Date', 'G1', '15/10/2026'],
    ]);
  });

  it('fills the theatre-list template: Hospital / OR, and Session with its times', () => {
    // Laid out like the user's template: labels merged across two cells, values in the yellow cell after.
    const rows = [['Consultant:', '', '', '', 'Date:', '', 'Hospital / OR:', '', 'Session:', '', ''], [], [], columns];
    const template = { name: 'List', rows, merges: ['A1:B1', 'I1:J1', 'A2:B2', 'F2:G2', 'H2:K2'] };
    const full = { ...values, hospital: 'TEH', times: '8:00–12:00' };
    expect(simple(headingCells(template, 4, full, layout))).toEqual([
      ['Consultant', 'C1', 'Willow'],
      ['Date', 'F1', '15/10/2026'],
      ['Hospital / OR', 'H1', 'TEH OR 4'],
      ['Session', 'K1', 'AM 8:00–12:00'],
    ]);
  });

  it('leaves out values the list doesn’t have', () => {
    const rows = [['Consultant', '', 'Date', ''], columns];
    expect(simple(headingCells(sheet(rows), 2, { ...values, consultant: '' }, layout))).toEqual([['Date', 'D1', '15/10/2026']]);
  });
});
