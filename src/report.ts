/**
 * Screenshots of a patient's IOL report (e.g. the IOLMaster 700 IOL calculation), kept with their row
 * and saved as pictures on that row in your Excel file, to check the lens against at timeout.
 */

export interface ReportImage {
  bytes: Uint8Array;
  /** e.g. "image/png". */
  type: string;
  /** Size in pixels, as the image was captured. */
  width: number;
  height: number;
}

const TYPES: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif' };

/** The image's type from its first bytes (PNG, JPEG or GIF), or null for anything else. */
export function sniffImageType(bytes: Uint8Array): string | null {
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'image/png';
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return 'image/jpeg';
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'image/gif';
  return null;
}

/** File extension for an image type, as Excel names its media ("png", "jpeg"). */
export function imageExtension(type: string): string {
  return TYPES[type] ?? 'png';
}

/** Width and height read from a PNG, JPEG or GIF header, or null if it can't be read. */
export function imageSize(bytes: Uint8Array): { width: number; height: number } | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const type = sniffImageType(bytes);
  try {
    if (type === 'image/png') return { width: view.getUint32(16), height: view.getUint32(20) };
    if (type === 'image/gif') return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
    if (type === 'image/jpeg') {
      let at = 2;
      while (at + 9 < bytes.length) {
        if (bytes[at] !== 0xff) return null;
        const marker = bytes[at + 1];
        const length = view.getUint16(at + 2);
        // Start-of-frame markers (not DHT, JPG or DAC) carry the size.
        if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
          return { height: view.getUint16(at + 5), width: view.getUint16(at + 7) };
        }
        at += 2 + length;
      }
    }
  } catch {
    // Truncated header.
  }
  return null;
}

/** A report image from a file's bytes, or null if it isn't a PNG, JPEG or GIF picture. */
export function reportImage(bytes: Uint8Array, fallbackType = ''): ReportImage | null {
  const type = sniffImageType(bytes) ?? (TYPES[fallbackType] ? fallbackType : null);
  if (!type) return null;
  const size = imageSize(bytes);
  if (!size || !size.width || !size.height) return null;
  return { bytes, type, ...size };
}

/** Same picture (same bytes), e.g. one read back from the file it was saved in. */
export function sameImage(a: ReportImage, b: ReportImage): boolean {
  if (a === b || a.bytes === b.bytes) return true;
  if (a.bytes.length !== b.bytes.length) return false;
  for (let i = 0; i < a.bytes.length; i++) if (a.bytes[i] !== b.bytes[i]) return false;
  return true;
}

export function sameImages(a: ReportImage[], b: ReportImage[]): boolean {
  return a.length === b.length && a.every((image, i) => sameImage(image, b[i]));
}
