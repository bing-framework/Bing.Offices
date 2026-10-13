import { readFileSync, writeFileSync } from 'node:fs';

const task = 'ai_docs/tasks/BING-OFFICES-BATCH-IMPORT-20260929-001';
const baseline = JSON.parse(readFileSync('build/api-snapshot-baseline.json', 'utf8'));
const candidate = JSON.parse(readFileSync(`${task}/artifacts/api-candidate.json`, 'utf8'));
const differences = {};

for (const tfm of ['net6.0', 'net8.0']) {
    differences[tfm] = {};
    for (const [assembly, snapshot] of Object.entries(candidate.assemblies[tfm])) {
        const previous = new Set(baseline.assemblies[tfm][assembly].lines);
        const current = new Set(snapshot.lines);
        const added = [...current].filter(line => !previous.has(line));
        const removed = [...previous].filter(line => !current.has(line));
        if (added.length || removed.length) {
            differences[tfm][assembly] = { added, removed };
            process.stdout.write(`${tfm}/${assembly}: +${added.length} -${removed.length}\n`);
        }
    }
}
writeFileSync(`${task}/api-diff.json`, JSON.stringify(differences, null, 2) + '\n', 'utf8');
