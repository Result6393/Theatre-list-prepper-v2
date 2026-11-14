/// <reference types="vite/client" />

/** The package.json version, filled in at build time. */
declare const __APP_VERSION__: string;

declare module 'tesseract.js/dist/tesseract.esm.min.js' {
  const Tesseract: typeof import('tesseract.js');
  export default Tesseract;
}

declare module '*?base64' {
  const base64: string;
  export default base64;
}
