import type { Layout } from './layout';
import type { GridRow } from './merge';

/** Which column of your sheet holds the allergy, or -1 if it isn't on your sheet. */
export function allergyColumn(layout: Layout): number {
  return layout.columns.findIndex((c) => c.source === 'field:allergy');
}

/** "", "Allergy: Y", "Allergy Yes", "Y": the run sheet said there's an allergy, but nobody has said what. */
export function lacksAllergyDetails(text: string): boolean {
  return /^(allergy[:\s]*)?(y|yes)?\.?$/i.test(text.trim());
}

/**
 * A patient the run sheet says has an allergy whose allergy cell still has no details, and who you
 * haven't acknowledged. It clears when you type the details, or acknowledge it.
 */
export function allergyPending(row: GridRow, col: number): boolean {
  return col >= 0 && row.status !== 'cancelled' && Boolean(row.patient?.allergyPending) && !row.allergyAck && lacksAllergyDetails(row.cells[col] ?? '');
}

/** Ticked patients still to check, i.e. the ones that would go out in a copy or an email. */
export function pendingAllergies(rows: GridRow[], col: number): GridRow[] {
  return rows.filter((r) => r.included && allergyPending(r, col));
}
