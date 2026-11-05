import archiver from 'archiver';
import { createWriteStream, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PNG } from 'pngjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const required = ['TEAMS_APP_ID', 'TEAMS_DOMAIN', 'ENTRA_CLIENT_ID'];
const missing = required.filter((key) => !process.env[key]);
if (missing.length) throw new Error(`Missing Teams package settings: ${missing.join(', ')}`);

const domain = process.env.TEAMS_DOMAIN.replace(/^https?:\/\//, '').replace(/\/$/, '');
const clientId = process.env.ENTRA_CLIENT_ID;
const appIdUri = process.env.ENTRA_APP_ID_URI || `api://${domain}/${clientId}`;
const manifestText = readFileSync(join(root, 'teams', 'manifest.template.json'), 'utf8')
  .replaceAll('__TEAMS_APP_ID__', process.env.TEAMS_APP_ID)
  .replaceAll('__TEAMS_DOMAIN__', domain)
  .replaceAll('__ENTRA_CLIENT_ID__', clientId)
  .replaceAll('__ENTRA_APP_ID_URI__', appIdUri);
const manifest = JSON.parse(manifestText);
const packageDir = join(root, '.teams-build');
mkdirSync(packageDir, { recursive: true });
writeFileSync(join(packageDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);

function makeIcon(size, outline = false) {
  const image = new PNG({ width: size, height: size });
  const scale = size / 192;
  const setPixel = (x, y, color) => {
    if (x < 0 || y < 0 || x >= size || y >= size) return;
    const index = (size * y + x) << 2;
    image.data[index] = color[0];
    image.data[index + 1] = color[1];
    image.data[index + 2] = color[2];
    image.data[index + 3] = color[3];
  };
  const rect = (x, y, width, height, color) => {
    for (let py = Math.round(y * scale); py < Math.round((y + height) * scale); py += 1) {
      for (let px = Math.round(x * scale); px < Math.round((x + width) * scale); px += 1) setPixel(px, py, color);
    }
  };
  if (!outline) {
    const background = [23, 60, 54, 255];
    for (let y = 0; y < size; y += 1) for (let x = 0; x < size; x += 1) setPixel(x, y, background);
  }
  const color = outline ? [255, 255, 255, 255] : [220, 187, 94, 255];
  rect(48, 96, 22, 42, color);
  rect(85, 55, 22, 83, color);
  rect(122, 76, 22, 62, color);
  return PNG.sync.write(image);
}

const iconPath = join(packageDir, 'color.png');
const outlinePath = join(packageDir, 'outline.png');
writeFileSync(iconPath, makeIcon(192));
writeFileSync(outlinePath, makeIcon(32, true));

const outputPath = join(root, 'dist', 'fieldnotes-teams.zip');
mkdirSync(dirname(outputPath), { recursive: true });
const output = createWriteStream(outputPath);
const archive = archiver('zip', { zlib: { level: 9 } });
archive.on('error', (error) => { throw error; });
archive.pipe(output);
archive.file(join(packageDir, 'manifest.json'), { name: 'manifest.json' });
archive.file(iconPath, { name: 'color.png' });
archive.file(outlinePath, { name: 'outline.png' });
await archive.finalize();
await new Promise((resolvePromise, reject) => {
  output.on('close', resolvePromise);
  output.on('error', reject);
});
console.log(`Created ${outputPath}`);