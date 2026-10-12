import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const changed = execFileSync('git', ['diff', '--name-only'], { encoding: 'utf8' })
    .split(/\r?\n/);
const untracked = execFileSync('git', ['ls-files', '--others', '--exclude-standard'],
    { encoding: 'utf8' }).split(/\r?\n/);
const files = [...new Set([...changed, ...untracked])].filter(path =>
    /\.(cs|json|md|mjs)$/.test(path)
    && !path.endsWith('Bing.Offices.EntityProviderContracts.xml'));
const decoder = new TextDecoder('utf-8', { fatal: true });
const failures = [];
for (const path of files) {
    const bytes = readFileSync(path);
    const bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
    let content;
    try {
        content = decoder.decode(bytes);
    } catch {
        failures.push(`${path}: invalid UTF-8`);
        continue;
    }
    if (bom !== path.endsWith('.cs'))
        failures.push(`${path}: unexpected BOM state`);
    if (content.includes('\r'))
        failures.push(`${path}: CR or mixed line endings`);
    if (!content.endsWith('\n') || content.endsWith('\n\n'))
        failures.push(`${path}: missing or extra terminal newline`);
}
process.stdout.write(`Checked ${files.length} changed text files.\n`);
for (const failure of failures)
    process.stderr.write(`${failure}\n`);
if (failures.length)
    process.exitCode = 1;
