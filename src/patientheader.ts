// Picks the patient's details out of the text read from a screenshot of the patient banner
// in the electronic record, which looks like:
//
//   Surname, Firstname        MRN:1234567        DOB:01-02-1950
//   Allergies/ADR: ...                           Resus Status: ...

export interface HeaderDetails {
  first: string;
  surname: string;
  /** dd/mm/yyyy */
  dob: string;
  mrn: string;
}

/** SMITH -> Smith, O'BRIEN -> O'Brien, SMITH-JONES -> Smith-Jones. Mixed case is left as typed. */
export function tidyName(name: string): string {
  const trimmed = name.replace(/\s+/g, ' ').trim();
  if (!trimmed || trimmed !== trimmed.toUpperCase()) return trimmed;
  return trimmed.toLowerCase().replace(/(^|[\s'’-])([a-z])/g, (_, before: string, letter: string) => before + letter.toUpperCase());
}

const pad = (v: string) => v.padStart(2, '0');

/** Reads whatever it can; anything not found comes back as an empty string. */
export function parsePatientHeader(text: string): HeaderDetails {
  const out: HeaderDetails = { first: '', surname: '', dob: '', mrn: '' };
  const lines = text.split(/\r?\n/).map((l) => l.trim());

  // "MRN:" is sometimes read with a stray space, a semicolon, or a lower-case letter.
  const mrnLabel = /\bM\s?R\s?N\s*[:;.]?\s*/i;
  const mrnLine = lines.find((l) => mrnLabel.test(l)) ?? '';
  const mrn = new RegExp(mrnLabel.source + '([A-Za-z0-9]{4,})', 'i').exec(mrnLine)?.[1];
  if (mrn) out.mrn = mrn.toUpperCase();

  const dob = /\bD\s?[O0]\s?B\s*[:;.]?\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{4})/i.exec(text);
  if (dob) out.dob = `${pad(dob[1])}/${pad(dob[2])}/${dob[3]}`;

  // The name is on the same line as the MRN, before it; otherwise the first "Surname, First" line.
  const nameRe = /([A-Za-z][A-Za-z'’ -]*?)\s*,\s*([A-Za-z][A-Za-z0-9'’ -]*)/;
  const beforeMrn = mrnLine ? mrnLine.split(mrnLabel)[0] : '';
  const name = nameRe.exec(beforeMrn) ?? lines.map((l) => nameRe.exec(l.split(/\b(?:MRN|DOB|Allerg)/i)[0])).find(Boolean) ?? null;
  if (name) {
    // Drop junk read from the picture icon before the surname, e.g. "@ Smith" or "a SMITH".
    const surname = name[1].trim().split(/\s+/).filter((w, i, all) => !(i === 0 && all.length > 1 && w.length === 1)).join(' ');
    out.surname = tidyName(surname);
    out.first = tidyName(name[2]);
  }
  return out;
}

export function foundAnything(d: HeaderDetails): boolean {
  return Boolean(d.first || d.surname || d.dob || d.mrn);
}
