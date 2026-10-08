import type { Patient } from './extract';

/**
 * Where an output column gets its value:
 * - `blank` — left empty for the registrar to fill in
 * - `field:<id>` — a cleaned-up field from FIELDS
 * - `raw:<HEADER>` — a run-sheet column exactly as listed
 */
export type ColumnSource = string;

export interface Column {
  heading: string;
  source: ColumnSource;
}

export interface Layout {
  columns: Column[];
  /** Put the column headings on the first copied row. */
  includeHeadings: boolean;
  /** Cell the user pastes into; shown as a hint only. */
  pasteStart: string;
}

export interface Field {
  id: string;
  label: string;
  get: (p: Patient) => string;
}

export const FIELDS: Field[] = [
  { id: 'eye', label: 'Eye (RIGHT / LEFT / BOTH)', get: (p) => p.eye },
  { id: 'patient', label: 'Patient (MRN NAME AgeSex)', get: (p) => p.patient },
  { id: 'mrn', label: 'MRN', get: (p) => p.mrn },
  { id: 'name', label: 'Name', get: (p) => p.name },
  { id: 'ageSex', label: 'Age/Sex', get: (p) => p.ageSex },
  { id: 'procedure', label: 'Procedure (abbreviated)', get: (p) => p.procedure },
  { id: 'procedureRaw', label: 'Procedure (as listed)', get: (p) => p.procedureRaw },
  { id: 'history', label: 'Diagnosis & history prefill (interpreter, diabetic)', get: (p) => p.history },
  { id: 'allergy', label: 'Allergy', get: (p) => p.allergy },
  { id: 'alerts', label: 'Alerts', get: (p) => p.alerts },
  { id: 'planNotes', label: 'Plan & notes (hospital flag + other comments)', get: (p) => p.planNotes },
  { id: 'interpreter', label: 'Interpreter', get: (p) => p.interpreter },
  { id: 'diabetic', label: 'Diabetic', get: (p) => p.diabetic },
  { id: 'comment', label: 'Comment (as listed)', get: (p) => p.comment },
  { id: 'time', label: 'Scheduled time', get: (p) => p.time },
  { id: 'consultant', label: 'Consultant', get: (p) => p.consultant },
  { id: 'ward', label: 'Ward', get: (p) => p.ward },
  { id: 'admissionType', label: 'Admission type', get: (p) => p.admissionType },
  { id: 'arrival', label: 'Arrival (TCI) time', get: (p) => p.arrival },
  { id: 'priority', label: 'Clinical priority', get: (p) => p.priority },
];

const FIELD_BY_ID = new Map(FIELDS.map((f) => [f.id, f]));

/** Matches the consultant cheat sheet: paste at B5, columns B–K (A is a row-number formula). */
export const DEFAULT_LAYOUT: Layout = {
  columns: [
    { heading: 'Eye', source: 'field:eye' },
    { heading: 'Patient', source: 'field:patient' },
    { heading: 'Procedure', source: 'field:procedure' },
    { heading: 'Diagnosis & relevant history', source: 'field:history' },
    { heading: 'Exam (VA / IOP)', source: 'blank' },
    { heading: 'Drops / allergies', source: 'field:allergy' },
    { heading: 'Implants / to order', source: 'blank' },
    { heading: 'Ordered', source: 'blank' },
    { heading: 'Alerts', source: 'field:alerts' },
    { heading: 'Plan & notes', source: 'field:planNotes' },
  ],
  includeHeadings: false,
  pasteStart: 'B5',
};

export function defaultLayout(): Layout {
  return structuredClone(DEFAULT_LAYOUT);
}

export function cellValue(patient: Patient, source: ColumnSource): string {
  if (source.startsWith('field:')) return FIELD_BY_ID.get(source.slice(6))?.get(patient) ?? '';
  if (source.startsWith('raw:')) return patient.raw[source.slice(4)] ?? '';
  return '';
}

export function buildRow(patient: Patient, layout: Layout): string[] {
  return layout.columns.map((c) => cellValue(patient, c.source));
}

/** Columns of the SurgiNet SN006A run sheet, offered as raw sources even before one is loaded. */
export const RUN_SHEET_COLUMNS = [
  'SCHED_DATE', 'FACILITY', 'AREA', 'THEATRE', 'SLOT_NAME', 'SLOT_BEG', 'SLOT_END', 'SLOT_COMMENT', 'CONSULTANT',
  'MRN', 'PT_NAME', 'PT_AGE', 'PT_SEX', 'CURRENT_WARD', 'WARD_FROM', 'SCHED_TM', 'ADM_DT_TM', 'PLANNED_ADM_TM',
  'TCI_TM', 'FIN_CLASS', 'CLINICAL_ALERT', 'ALLERGY', 'LISTING_DATE', 'CLINICAL_PRIORITY', 'EMER_PRIORITY',
  'ADDITIONAL_REQ', 'OTHER_RESOURCES', 'INTEND_ADM_TYPE', 'PLANNED_PROCEDURE', 'ANAESTHETIST', 'GENERAL_COMMENT',
  'CONF_COMMENTS',
];
