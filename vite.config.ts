import type { Plugin } from 'vite';
import { readFileSync } from 'node:fs';
import { defineConfig } from 'vitest/config';

const { version } = JSON.parse(readFileSync('package.json', 'utf8')) as { version: string };

// Inline the built JS and CSS into index.html so the whole app is one file
// that can be opened straight from disk, with no network access.
function inlineIntoHtml(): Plugin {
  return {
    name: 'inline-into-html',
    apply: 'build',
    enforce: 'post',
    generateBundle(_options, bundle) {
      const html = bundle['index.html'];
      if (!html || html.type !== 'asset') return;
      let source = String(html.source);
      for (const [fileName, item] of Object.entries(bundle)) {
        if (item.type === 'chunk' && item.isEntry) {
          const tag = new RegExp(`<script[^>]*src="[^"]*${fileName}"[^>]*></script>`);
          const code = item.code.replace(/<\/script/gi, '<\\/script');
          source = source.replace(tag, () => `<script type="module">${code}</script>`);
          delete bundle[fileName];
        } else if (item.type === 'asset' && fileName.endsWith('.css')) {
          const tag = new RegExp(`<link[^>]*href="[^"]*${fileName}"[^>]*>`);
          source = source.replace(tag, () => `<style>${String(item.source)}</style>`);
          delete bundle[fileName];
        }
      }
      html.source = source;
    },
  };
}

// `import data from 'file?base64'` gives the file's bytes as base64 text, so binary data (like the
// OCR engine's language data) is built into the single offline file.
function binaryAsBase64(): Plugin {
  return {
    name: 'binary-as-base64',
    enforce: 'pre',
    async resolveId(id, importer) {
      if (!id.endsWith('?base64')) return null;
      const resolved = await this.resolve(id.slice(0, -'?base64'.length), importer, { skipSelf: true });
      return resolved ? `${resolved.id}?base64` : null;
    },
    load(id) {
      if (!id.endsWith('?base64')) return null;
      return `export default ${JSON.stringify(readFileSync(id.slice(0, -'?base64'.length)).toString('base64'))};`;
    },
  };
}

// Puts the package.json version in the page title.
function versionInTitle(): Plugin {
  return {
    name: 'version-in-title',
    transformIndexHtml: (html) => html.replace('<title>Theatre List Prepper</title>', `<title>Theatre List Prepper v${version}</title>`),
  };
}

export default defineConfig({
  base: './',
  define: { __APP_VERSION__: JSON.stringify(version) },
  plugins: [binaryAsBase64(), versionInTitle(), inlineIntoHtml()],
  build: {
    cssCodeSplit: false,
    assetsInlineLimit: Number.MAX_SAFE_INTEGER,
    modulePreload: false,
  },
  test: {
    environment: 'node',
  },
});
