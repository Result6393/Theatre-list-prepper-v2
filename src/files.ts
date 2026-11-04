export type FileKind = 'settings' | 'workbook' | 'runsheet';

/** Settings files are JSON and Excel files are workbooks; anything else is treated as a run sheet (CSV). */
export function classifyFile(name: string, type = ''): FileKind {
  if (/\.json$/i.test(name) || /json/i.test(type)) return 'settings';
  if (/\.(xlsx|xlsm|xls)$/i.test(name) || /spreadsheetml|ms-excel/i.test(type)) return 'workbook';
  return 'runsheet';
}
