import { spawn } from 'child_process';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const userDataDir = path.join(__dirname, '.edge-profile');

const edgePaths = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'
];

const edgeExe = edgePaths.find(p => fs.existsSync(p));

if (!edgeExe) {
  console.error('Error: Microsoft Edge not found in standard paths.');
  process.exit(1);
}

const args = [
  '--remote-debugging-port=9222',
  `--user-data-dir=${userDataDir}`,
  '--no-first-run',
  '--no-default-browser-check',
  'https://x.com'
];

const child = spawn(edgeExe, args, {
  detached: true,
  stdio: 'ignore'
});

child.unref();

console.log('Edge started.');
console.log('Remote debugging: http://127.0.0.1:9222');
console.log(`Profile: ${userDataDir}`);
process.exit(0);
