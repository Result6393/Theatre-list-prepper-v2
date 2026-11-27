import { describe, expect, it } from 'vitest';
import { foundAnything, parsePatientHeader, tidyName } from '../src/patientheader';

// Invented banner text, in the layouts the record system shows.
describe('patient banner', () => {
  it('reads name, MRN and DOB from a clean banner', () => {
    const text = 'Zzexample, Patient        MRN:TEMP1234567        DOB:01-02-1950\nAllergies/ADR: Allergies Not Recorded        Resus Status: INACTIVE';
    expect(parsePatientHeader(text)).toEqual({ first: 'Patient', surname: 'Zzexample', dob: '01/02/1950', mrn: 'TEMP1234567' });
  });

  it('copes with a tab title line above the banner', () => {
    const text = 'SAMPLE, PERSON ×\nSAMPLE, PERSON  MRN:7654321  DOB:12-03-1984\nAllergies/ADR: penicillin  Resus Status: INACTIVE';
    expect(parsePatientHeader(text)).toEqual({ first: 'Person', surname: 'Sample', dob: '12/03/1984', mrn: '7654321' });
  });

  it('keeps digits in a first name and tidies capitals', () => {
    const text = 'EXAMPLEPATIENT, TEST101 MRN:1112223 DOB:30-08-1991';
    expect(parsePatientHeader(text)).toEqual({ first: 'Test101', surname: 'Examplepatient', dob: '30/08/1991', mrn: '1112223' });
  });

  it('tolerates common misreads', () => {
    const text = '@ Sample-Jones, Mary Ann   MRN; 0012345   D0B: 1-2-1960\nAllergies/ADR: No known allergies, No Known Medication Allergies';
    expect(parsePatientHeader(text)).toEqual({ first: 'Mary Ann', surname: 'Sample-Jones', dob: '01/02/1960', mrn: '0012345' });
  });

  it('does not take the allergy list for a name', () => {
    const text = 'MRN:1234567 DOB:01-01-1900\nAllergies/ADR: Oranges, Apples';
    expect(parsePatientHeader(text)).toEqual({ first: '', surname: '', dob: '01/01/1900', mrn: '1234567' });
  });

  it('returns blanks for unrelated text', () => {
    const d = parsePatientHeader('nothing useful here');
    expect(d).toEqual({ first: '', surname: '', dob: '', mrn: '' });
    expect(foundAnything(d)).toBe(false);
  });

  it('tidies names', () => {
    expect(tidyName("O'EXAMPLE")).toBe("O'Example");
    expect(tidyName('SAMPLE-JONES')).toBe('Sample-Jones');
    expect(tidyName('van der Example')).toBe('van der Example');
    expect(tidyName('  MARY   ANN ')).toBe('Mary Ann');
  });
});
