import Papa from 'papaparse';

// Entirely invented patients and surgeons, in the shape of a SurgiNet SN006A export.
export const HEADERS = [
  'TYPE', 'SCHED_DATE', 'FACILITY', 'AREA', 'THEATRE', 'SLOT_NAME', 'SLOT_BEG', 'SLOT_END', 'SLOT_COMMENT',
  'CONSULTANT', 'MRN', 'PT_NAME', 'PT_AGE', 'PT_SEX', 'CURRENT_WARD', 'WARD_FROM', 'SCHED_TM', 'ADM_DT_TM',
  'PLANNED_ADM_TM', 'TCI_TM', 'FIN_CLASS', 'CLINICAL_ALERT', 'ALLERGY', 'LISTING_DATE', 'CLINICAL_PRIORITY',
  'EMER_PRIORITY', 'ADDITIONAL_REQ', 'OTHER_RESOURCES', 'INTEND_ADM_TYPE', 'PLANNED_PROCEDURE', 'ANAESTHETIST',
  'GENERAL_COMMENT', 'CONF_COMMENTS',
];

type Row = Record<string, string>;

export function slot(theatre: string, name: string, beg: string, end: string): Row {
  return {
    TYPE: 'slot', SCHED_DATE: '  14/10/26', FACILITY: 'Test Eye Hospital', AREA: 'OpSuite TEH',
    THEATRE: theatre, SLOT_NAME: name, SLOT_BEG: `14/10/2026 ${beg}`, SLOT_END: `14/10/2026 ${end}`,
  };
}

export function booking(s: Row, fields: Row): Row {
  return {
    ...s, TYPE: 'booking', PT_AGE: ' 70 Y', PT_SEX: 'F', CURRENT_WARD: 'DPU TEH', WARD_FROM: 'DPU TEH',
    PLANNED_ADM_TM: '14/10/2026 7:00', TCI_TM: '14/10/2026 7:00', FIN_CLASS: 'MNW', CLINICAL_ALERT: 'No',
    ALLERGY: 'NKA', CLINICAL_PRIORITY: 'C', INTEND_ADM_TYPE: 'Day Only ( same calendar day)', ...fields,
  };
}

const or1am = slot('TEH OR 01', 'Open Unallocated', '8:00', '12:00');
const or1pm = slot('TEH OR 01', 'Waitlist', '13:00', '17:00');
const or2am = slot('TEH OR 02', 'Alder, P (Max)  Eye - General', '8:00', '12:00');
const or2pm = slot('TEH OR 02', 'Birch, Q Eye - University', '13:00', '17:00');
const or3pm = slot('TEH OR 03', 'Emerg Eye', '17:00', '19:00');
const or5am = slot('TEH OR 05', 'Rowan, K - Hand', '8:00', '12:00');

export const ROWS: Row[] = [
  or1am,
  booking(or1am, {
    CONSULTANT: 'Cedar, M', MRN: '1000001', PT_NAME: 'TESTER, Alice Jane', PT_AGE: ' 69 Y', SCHED_TM: '   09:20',
    PLANNED_PROCEDURE: 'left cataract extraction and insertion of an intra ocular lens implant',
    GENERAL_COMMENT: '[G]: Cantonese interpreter 0730 (Booking ID: RS-11111111)',
  }),
  booking(or1am, {
    CONSULTANT: 'Cedar, M', MRN: '1000002', PT_NAME: "O'SAMPLE, Bob", PT_SEX: 'M', SCHED_TM: '   08:00',
    ALLERGY: 'Y', CLINICAL_ALERT: 'Appointment of Enduring Guardian;Risk of pressure area',
    PLANNED_PROCEDURE: 'Right eye cataract extraction/insertion intra ocular lens',
    GENERAL_COMMENT: '[G]: DIABETIC/ OZEMPIC',
  }),
  or1pm,
  booking(or1pm, {
    CONSULTANT: 'Hazel, T', MRN: '1000003', PT_NAME: 'EXAMPLE, Carol', SCHED_TM: '   13:00',
    INTEND_ADM_TYPE: 'DOSA - 1 Overnight Stay', ALLERGY: 'Not Recorded', CLINICAL_PRIORITY: 'B',
    PLANNED_PROCEDURE: 'left trabeculectomy + mitomicyn c', GENERAL_COMMENT: '[G]: STG. T2DM',
  }),
  booking(or1pm, {
    CONSULTANT: 'Birch, Q', MRN: '1000004', PT_NAME: 'DUMMY, Dan', SCHED_TM: '   14:40',
    PLANNED_PROCEDURE: 'Right dacryocystorhinostomy', GENERAL_COMMENT: '[G]: mandarin interp @ 07.30 am',
  }),
  or2am,
  booking(or2am, {
    CONSULTANT: 'Alder, P', MRN: '1000005', PT_NAME: 'PLACEHOLDER, Eve', SCHED_TM: '   08:00',
    PLANNED_PROCEDURE: 'excision pterygium and autograft', GENERAL_COMMENT: '[G]: POWH',
  }),
  or2pm,
  booking(or2pm, {
    CONSULTANT: 'Hazel, T', MRN: '1000006', PT_NAME: 'MOCK, Frank', SCHED_TM: '   13:41',
    CLINICAL_ALERT: 'Risk of falls',
    PLANNED_PROCEDURE: 'left eye cataract extraction + lens insertion',
    GENERAL_COMMENT: '[G]: STG, Croation Interpreter',
  }),
  booking(or2pm, {
    CONSULTANT: 'Birch, Q', MRN: '1000007', PT_NAME: 'SPECIMEN, Gwen', SCHED_TM: '   13:00',
    PLANNED_PROCEDURE: 'right ptosis repair', GENERAL_COMMENT: '[G]: daibetes type2, need US marking morning of surgery',
  }),
  or3pm,
  or5am,
  booking(or5am, {
    CONSULTANT: 'Rowan, K', MRN: '1000008', PT_NAME: 'HANDY, Hugo', PT_AGE: ' 21 Y', SCHED_TM: '   08:00',
    PLANNED_PROCEDURE: 'removal of kwires left thumb',
  }),
];

export function runSheetCsv(rows: Row[] = ROWS): string {
  return Papa.unparse({ fields: HEADERS, data: rows.map((r) => HEADERS.map((h) => r[h] ?? '')) });
}

export function runSheetTsv(rows: Row[] = ROWS): string {
  return Papa.unparse({ fields: HEADERS, data: rows.map((r) => HEADERS.map((h) => r[h] ?? '')) }, { delimiter: '\t' });
}

/** The same OR 1 AM list again the next day, plus one patient, to exercise multi-day run sheets. */
export function nextDayRows(): Row[] {
  const day2 = { ...slot('TEH OR 01', 'Open Unallocated', '8:00', '12:00'), SCHED_DATE: '  15/10/26' };
  const shift = (r: Row): Row => ({ ...r, SLOT_BEG: r.SLOT_BEG.replace('14/10', '15/10'), SLOT_END: r.SLOT_END.replace('14/10', '15/10') });
  return [
    shift(day2),
    shift(
      booking(day2, {
        CONSULTANT: 'Cedar, M', MRN: '2000001', PT_NAME: 'NEXTDAY, Nina', SCHED_TM: '   08:00',
        PLANNED_PROCEDURE: 'right cataract extraction and insertion of intraocular lens implant',
      }),
    ),
  ];
}
