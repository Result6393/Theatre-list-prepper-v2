/**
 * The SRK/T formula (Retzlaff, Sanders and Kraff, J Cataract Refract Surg 1990, with the 1990 erratum),
 * which is published. Used only for the difference between two lenses' constants in one eye: the
 * report's own (Barrett) numbers stay the reference.
 */

const NA = 1.336; // aqueous and vitreous
const NC_1 = 0.333; // cornea (1.333) minus 1
const VERTEX = 12; // spectacle vertex distance, mm

interface Eye {
  r: number;
  lopt: number;
  elp: number;
}

function eye(axialLength: number, k: number, aConstant: number): Eye {
  const r = 337.5 / k;
  const lcor = axialLength > 24.2 ? -3.446 + 1.716 * axialLength - 0.0237 * axialLength * axialLength : axialLength;
  const width = -5.41 + 0.58412 * lcor + 0.098 * k;
  const h = r * r - (width * width) / 4 > 0 ? r - Math.sqrt(r * r - (width * width) / 4) : r;
  const offset = 0.62467 * aConstant - 68.747 - 3.336;
  return { r, lopt: axialLength + 0.65696 - 0.02029 * axialLength, elp: h + offset };
}

/** IOL power for a target refraction (spectacle plane). */
export function srktPower(axialLength: number, k: number, aConstant: number, target: number): number {
  const { r, lopt, elp } = eye(axialLength, k, aConstant);
  const top = 1000 * NA * (NA * r - NC_1 * lopt - 0.001 * target * (VERTEX * (NA * r - NC_1 * lopt) + lopt * r));
  const bottom = (lopt - elp) * (NA * r - NC_1 * elp - 0.001 * target * (VERTEX * (NA * r - NC_1 * elp) + elp * r));
  return top / bottom;
}

/** Refraction (spectacle plane) expected with a given IOL power. */
export function srktRefraction(axialLength: number, k: number, aConstant: number, power: number): number {
  const { r, lopt, elp } = eye(axialLength, k, aConstant);
  const top = 1000 * NA * (NA * r - NC_1 * lopt) - power * (lopt - elp) * (NA * r - NC_1 * elp);
  const bottom = NA * (VERTEX * (NA * r - NC_1 * lopt) + lopt * r) - 0.001 * power * (lopt - elp) * (VERTEX * (NA * r - NC_1 * elp) + elp * r);
  return top / bottom;
}
