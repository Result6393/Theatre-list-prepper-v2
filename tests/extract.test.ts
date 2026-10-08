import { describe, expect, it } from 'vitest';
import {
  abbreviateProcedure, detectEye, extractSessions, formatAgeSex, formatAlerts, formatAllergy, parseComment,
  theatreToOr, type Session,
} from '../src/extract';
import { parseTable } from '../src/parse';
import { runSheetCsv } from './fixtures';

const OPTIONS = { flagHospitals: ['POWH', 'STG'], hospitalCodes: ['TEH', 'SSEH', 'POWH', 'STG'] };

function sessions(): Session[] {
  return extractSessions(parseTable(runSheetCsv()), OPTIONS);
}

function find(or: string, ampm: string): Session {
  const s = sessions().find((x) => x.or === or && x.ampm === ampm);
  if (!s) throw new Error(`no session ${or} ${ampm}`);
  return s;
}

describe('extractSessions', () => {
  it('groups bookings under their theatre session', () => {
    const list = sessions();
    expect(list.map((s) => `${s.or} ${s.ampm} ${s.consultant} ${s.patients.length}`)).toEqual([
      'OR 1 AM Cedar 2',
      'OR 1 PM Hazel 2',
      'OR 2 AM Alder 1',
      'OR 2 PM Birch 2',
      'OR 3 PM  0',
      'OR 5 AM Rowan 1',
    ]);
  });

  it('labels the session with date, times and slot owner', () => {
    const s = find('OR 2', 'AM');
    expect(s).toMatchObject({ date: '14/10/2026', start: '08:00', end: '12:00', slotOwner: 'Alder' });
  });

  it('sorts patients by scheduled time', () => {
    expect(find('OR 1', 'AM').patients.map((p) => p.time)).toEqual(['08:00', '09:20']);
  });

  it('builds the cheat-sheet fields for a patient', () => {
    const p = find('OR 1', 'AM').patients[1];
    expect(p).toMatchObject({
      patient: '1000001 TESTER, Alice Jane 69F',
      eye: 'LEFT',
      procedure: 'LE CE/IOL',
      history: 'Cantonese interpreter',
      allergy: 'NKA',
      alerts: '',
      planNotes: '',
      admissionType: 'Day Only',
      arrival: '07:00',
      consultantSurname: 'Cedar',
    });
  });

  it('flags POWH/STG patients and keeps other comments in Plan & notes', () => {
    expect(find('OR 2', 'AM').patients[0].planNotes).toBe('POWH patient');
    expect(find('OR 1', 'PM').patients[0]).toMatchObject({
      planNotes: 'STG patient',
      history: 'Diabetic',
      procedure: 'LE trabeculectomy + MMC',
      allergy: 'Allergy not recorded',
    });
    expect(find('OR 2', 'PM').patients[0]).toMatchObject({
      history: 'Diabetic',
      planNotes: 'need US marking morning of surgery',
      procedure: 'RE ptosis repair',
    });
  });

  it('splits multiple clinical alerts onto lines', () => {
    expect(find('OR 1', 'AM').patients[0]).toMatchObject({
      alerts: 'Appointment of Enduring Guardian\nRisk of pressure area',
      allergy: 'Allergy: Y',
      history: 'Diabetic (Ozempic)',
      ageSex: '70M',
    });
  });

  it('leaves non-eye procedures as listed', () => {
    expect(find('OR 5', 'AM').patients[0]).toMatchObject({ eye: 'LEFT', procedure: 'removal of kwires left thumb' });
  });
});

describe('field rules', () => {
  it.each([
    ['left cataract extraction and insertion of an intra ocular lens implant', 'LEFT', 'LE CE/IOL'],
    ['Right eye cataract extraction and insertion of an intraocular lens implant', 'RIGHT', 'RE CE/IOL'],
    ['right cataract extraction/insertion intra ocular lens', 'RIGHT', 'RE CE/IOL'],
    ['left eye cataract extraction + intraocular lens insertion', 'LEFT', 'LE CE/IOL'],
    ['left trabeculectomy + mitomicyn c', 'LEFT', 'LE trabeculectomy + MMC'],
    ['left eye trabeculectomy', 'LEFT', 'LE trabeculectomy'],
    ['Right dacryocystorhinostomy', 'RIGHT', 'RE DCR'],
    ['excision pterygium and autograft', '', 'pterygium excision + autograft'],
    ['bilateral ptosis repair', 'BOTH', 'BE ptosis repair'],
    ['RE phaco + trab', 'RIGHT', 'RE CE + trabeculectomy'],
    ['Right cataract extraction, insertion of intraocular lens and iStent inject', 'RIGHT', 'RE CE/IOL + iStent inject'],
    ['left phaco/IOL + istent insertion', 'LEFT', 'LE CE/IOL + iStent'],
    ['right eye cataract extraction + i-stent', 'RIGHT', 'RE CE + iStent'],
    ['left phacoemulsification + intraocular lens + Hydrus microstent', 'LEFT', 'LE CE/IOL + Hydrus'],
    ['right cataract extraction and insertion of toric intraocular lens', 'RIGHT', 'RE CE/toric IOL'],
    ['left cataract extraction and insertion of intraocular lens under GA', 'LEFT', 'LE CE/IOL + under GA'],
    ['right cataract extraction + IOL + capsular tension ring', 'RIGHT', 'RE CE/IOL + capsular tension ring'],
    ['left eye cataract extraction + intraocular lens + vitrectomy', 'LEFT', 'LE CE/IOL + PPV'],
  ])('%s', (procedure, eye, abbreviated) => {
    expect(detectEye(procedure)).toBe(eye);
    expect(abbreviateProcedure(procedure, detectEye(procedure))).toBe(abbreviated);
  });

  it.each([
    ['Cantonese interpreter 0730 (Booking ID: RS-1)', 'Cantonese interpreter', '', []],
    ['cantonese intep @11 am', 'Cantonese interpreter', '', []],
    ['mandarin interp @ 07.30 am', 'Mandarin interpreter', '', []],
    ['DIABETIC, Mandarin interpreter 0800 (Booking ID: RS-2)', 'Mandarin interpreter', 'Diabetic', []],
    ['STG, Croation Interpreter', 'Croation interpreter', '', []],
    ['needs interpreter', 'Interpreter', '', []],
    ['diabetic/ OZEMPIC', '', 'Diabetic (Ozempic)', []],
    ['daibetes type2', '', 'Diabetic', []],
    ['STG. T2DM', '', 'Diabetic', []],
    ['SSEH', '', '', []],
    ['POWH pt, GA', '', '', ['GA']],
    ['need US marking morning of surgery', '', '', ['need US marking morning of surgery']],
  ])('comment %s', (comment, interpreter, diabetic, leftover) => {
    expect(parseComment(comment, OPTIONS.hospitalCodes)).toEqual({ interpreter, diabetic, leftover });
  });

  it('formats allergy, alerts, age/sex and theatre', () => {
    expect(['NKA', 'Y', 'Not Recorded', '', 'Penicillin'].map(formatAllergy)).toEqual([
      'NKA', 'Allergy: Y', 'Allergy not recorded', 'Allergy not recorded', 'Allergy: Penicillin',
    ]);
    expect(formatAlerts('No')).toBe('');
    expect(formatAlerts('Risk of falls')).toBe('Risk of falls');
    expect(formatAgeSex(' 69 Y', 'F')).toBe('69F');
    expect(formatAgeSex(' 8 M', 'M')).toBe('8moM');
    expect(theatreToOr('SSEH OR 03')).toBe('OR 3');
  });
});
