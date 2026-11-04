import { describe, expect, it } from 'vitest';
import { classifyFile } from '../src/files';

describe('classifyFile', () => {
  it('treats JSON as settings', () => {
    expect(classifyFile('theatre-list-prepper-settings.json')).toBe('settings');
    expect(classifyFile('SETTINGS.JSON')).toBe('settings');
    expect(classifyFile('export', 'application/json')).toBe('settings');
  });
  it('treats Excel files as workbooks', () => {
    expect(classifyFile('Cheat sheet.xlsx')).toBe('workbook');
    expect(classifyFile('list.XLSM')).toBe('workbook');
    expect(classifyFile('old.xls')).toBe('workbook');
    expect(classifyFile('download', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')).toBe('workbook');
  });
  it('treats everything else as a run sheet', () => {
    expect(classifyFile('3019_SN006A_OUTPUT.csv', 'text/csv')).toBe('runsheet');
    expect(classifyFile('list.txt')).toBe('runsheet');
    expect(classifyFile('list')).toBe('runsheet');
    expect(classifyFile('json.csv')).toBe('runsheet');
  });
});
