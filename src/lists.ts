import type { Patient, Session } from './extract';

/**
 * Lists that run on into each other without a break can be prepped as one. The combined list keeps
 * each patient's own list, so the Medical Records email can still name each one.
 */
export function combineSessions(sessions: Session[]): Session {
  if (sessions.length === 1) return sessions[0];
  const join = (values: string[], separator: string) => [...new Set(values.filter(Boolean))].join(separator);
  return {
    id: sessions.map((s) => s.id).join('+'),
    theatre: join(sessions.map((s) => s.theatre), ' + '),
    or: join(sessions.map((s) => s.or), ' + '),
    date: join(sessions.map((s) => s.date), ' + '),
    start: sessions[0].start,
    end: sessions[sessions.length - 1].end,
    ampm: join(sessions.map((s) => s.ampm), ' + '),
    slotName: join(sessions.map((s) => s.slotName), ' + '),
    slotOwner: join(sessions.map((s) => s.slotOwner), ' / '),
    consultant: join(sessions.map((s) => s.consultant), ' / '),
    patients: sessions.flatMap((s) => s.patients),
  };
}

/** Which of the lists each patient is booked on. */
export function listOfPatient(sessions: Session[]): Map<Patient, Session> {
  return new Map(sessions.flatMap((s) => s.patients.map((p) => [p, s] as const)));
}

/** Keeps the ticked lists in run-sheet order, so combined lists always read the same way. */
export function orderedSelection(all: Session[], ids: Iterable<string>): string[] {
  const wanted = new Set(ids);
  return all.filter((s) => wanted.has(s.id)).map((s) => s.id);
}
