export type FileKind = 'settings' | 'runsheet';

/** Settings files are JSON; anything else dropped or chosen is treated as a run sheet. */
export function classifyFile(name: string, type = ''): FileKind {
  return /\.json$/i.test(name) || /json/i.test(type) ? 'settings' : 'runsheet';
}
