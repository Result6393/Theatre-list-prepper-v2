import { describe, expect, it } from 'vitest';
import { ParseError, parseTable } from '../src/parse';
import { ROWS, runSheetCsv, runSheetTsv } from './fixtures';

describe('parseTable', () => {
  it('reads a CSV run sheet', () => {
    const table = parseTable(runSheetCsv());
    expect(table.headers).toContain('PT_NAME');
    expect(table.rows).toHaveLength(ROWS.length);
    expect(table.rows[1].PT_NAME).toBe('TESTER, Alice Jane');
  });

  it('reads tab-separated text pasted from Excel, trimming cells', () => {
    const table = parseTable(runSheetTsv());
    expect(table.rows).toHaveLength(ROWS.length);
    expect(table.rows[1].SCHED_TM).toBe('09:20');
    expect(table.rows[1].PT_NAME).toBe('TESTER, Alice Jane');
  });

  it('finds the heading row below other lines and ignores column order', () => {
    const text = 'Theatre report\n\nPT_NAME,THEATRE,TYPE\n"SMITH, Ann",TEH OR 01,booking\n';
    const table = parseTable(text);
    expect(table.rows).toEqual([{ PT_NAME: 'SMITH, Ann', THEATRE: 'TEH OR 01', TYPE: 'booking' }]);
  });

  it('keeps quoted line breaks inside a cell', () => {
    const table = parseTable('TYPE\tTHEATRE\tPT_NAME\tGENERAL_COMMENT\nbooking\tOR 1\tX, Y\t"line one\nline two"\n');
    expect(table.rows[0].GENERAL_COMMENT).toBe('line one\nline two');
  });

  it('explains when the headings are missing', () => {
    expect(() => parseTable('just some text\nwithout headings')).toThrow(ParseError);
  });
});
