/**
 * Reads the text of an IOL report screenshot with Tesseract, entirely in this browser and offline:
 * the OCR engine, its worker and its English data are built into the page, nothing is downloaded,
 * and the picture never leaves the computer.
 */
import Tesseract from 'tesseract.js/dist/tesseract.esm.min.js';
import workerSource from 'tesseract.js/dist/worker.min.js?raw';
import coreSource from 'tesseract.js-core/tesseract-core-lstm.wasm.js?raw';
import languageBase64 from '@tesseract.js-data/eng/4.0.0_best_int/eng.traineddata.gz?base64';
import { headingGaps, readReport, withReread, type OcrWord } from './iolreport';

/** Identifies how a picture was read: OCR results saved by another engine or settings are read again. */
export const OCR_ENGINE = 'tesseract.js 6, eng best_int, LSTM, grey 1900px, lens headings re-read';

/** Where the worker asks for the English data; it never goes to the network (see below). */
const LANG_PATH = 'https://offline.invalid/tessdata';

let engine: Promise<Tesseract.Worker> | null = null;

async function start(): Promise<Tesseract.Worker> {
  // A worker started from a page opened from disk can't load any other file, not even one made in
  // the page, so it gets everything in one script: the engine (which Tesseract then doesn't load),
  // the English data (handed over when the worker asks for "<langPath>/eng.traineddata.gz"), and
  // Tesseract's worker code.
  const script = [
    coreSource,
    'self.TesseractCore = TesseractCore;',
    `const LANGUAGE = ${JSON.stringify(languageBase64)};`,
    `const LANG_PATH = ${JSON.stringify(LANG_PATH)};`,
    'const realFetch = self.fetch.bind(self);',
    'self.fetch = (url, options) => String(url).startsWith(LANG_PATH)',
    '  ? Promise.resolve(new Response(Uint8Array.from(atob(LANGUAGE), (c) => c.charCodeAt(0))))',
    '  : realFetch(url, options);',
    workerSource,
  ].join('\n');
  return Tesseract.createWorker('eng', Tesseract.OEM.LSTM_ONLY, {
    workerPath: URL.createObjectURL(new Blob([script], { type: 'text/javascript' })),
    workerBlobURL: false,
    langPath: LANG_PATH,
    gzip: true,
    cacheMethod: 'none',
  });
}

/** Screenshots are enlarged to about this width first: at screen size OCR drops decimal points and minus signs. */
const READ_WIDTH = 1900;

async function enlarged(image: Blob): Promise<HTMLCanvasElement> {
  const bitmap = await createImageBitmap(image);
  const scale = Math.min(3, Math.max(1, READ_WIDTH / bitmap.width));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  const context = canvas.getContext('2d')!;
  context.imageSmoothingQuality = 'high';
  context.filter = 'grayscale(1)';
  context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  return canvas;
}

/** The words in a picture, with where each one is (in a copy of the picture `width` × `height` pixels). */
export async function readWords(image: Blob): Promise<{ words: OcrWord[]; width: number; height: number }> {
  engine ??= start().catch((e) => {
    engine = null;
    throw e;
  });
  const worker = await engine;
  const canvas = await enlarged(image);
  const { data } = await worker.recognize(canvas, {}, { blocks: true, text: false });
  let words = wordsOf(data);
  // Reading the whole page, OCR can drop the second line of a lens's heading (the line with its name,
  // e.g. "nsar AR40 3-Piece"); read on their own, those headings come out whole.
  for (const area of headingGaps(readReport(words))) {
    const left = Math.max(0, Math.floor(area.x0));
    const top = Math.max(0, Math.floor(area.y0));
    const width = Math.min(canvas.width, Math.ceil(area.x1)) - left;
    const height = Math.min(canvas.height, Math.ceil(area.y1)) - top;
    if (width < 10 || height < 10) continue;
    const again = await worker.recognize(canvas, { rectangle: { left, top, width, height } }, { blocks: true, text: false });
    words = withReread(words, area, wordsOf(again.data));
  }
  return { words, width: canvas.width, height: canvas.height };
}

function wordsOf(data: Tesseract.Page): OcrWord[] {
  const words: OcrWord[] = [];
  for (const block of data.blocks ?? [])
    for (const paragraph of block.paragraphs)
      for (const line of paragraph.lines)
        for (const word of line.words) words.push({ text: word.text, x0: word.bbox.x0, y0: word.bbox.y0, x1: word.bbox.x1, y1: word.bbox.y1 });
  return words;
}

/**
 * Enlarged, grey copy of a screenshot of the patient banner (from the Medical Letter Prepper app).
 * The record system's banner is white text on dark blue, which reads far better as dark text on a
 * light background, so dark pictures are inverted.
 */
async function preparedBanner(image: Blob): Promise<HTMLCanvasElement> {
  const bitmap = await createImageBitmap(image);
  const scale = Math.min(4, Math.max(1, READ_WIDTH / bitmap.width));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  const context = canvas.getContext('2d', { willReadFrequently: true })!;
  context.imageSmoothingQuality = 'high';
  context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();

  const pixels = context.getImageData(0, 0, canvas.width, canvas.height);
  const d = pixels.data;
  let total = 0;
  for (let i = 0; i < d.length; i += 4) {
    const grey = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
    d[i] = d[i + 1] = d[i + 2] = grey;
    total += grey;
  }
  if (total / (d.length / 4) < 128) for (let i = 0; i < d.length; i += 4) d[i] = d[i + 1] = d[i + 2] = 255 - d[i];
  context.putImageData(pixels, 0, 0);
  return canvas;
}

/** The text in a screenshot of the patient banner, line by line (see patientheader.ts). Nothing is kept. */
export async function readText(image: Blob): Promise<string> {
  engine ??= start().catch((e) => {
    engine = null;
    throw e;
  });
  const worker = await engine;
  const { data } = await worker.recognize(await preparedBanner(image));
  return data.text ?? '';
}
