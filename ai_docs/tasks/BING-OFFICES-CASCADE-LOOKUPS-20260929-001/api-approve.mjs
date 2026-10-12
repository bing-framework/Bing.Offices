import { readFileSync, writeFileSync } from 'node:fs';

const task = 'ai_docs/tasks/BING-OFFICES-CASCADE-LOOKUPS-20260929-001';
const differences = JSON.parse(readFileSync(`${task}/api-diff.json`, 'utf8'));
for (const tfm of ['net6.0', 'net8.0']) {
    const entries = Object.entries(differences[tfm]);
    if (entries.length !== 1 || entries[0][0] !== 'Bing.Offices.Abstractions'
        || entries[0][1].added.length !== 13 || entries[0][1].removed.length !== 0)
        throw new Error(`API approval does not match the reviewed ${tfm} diff`);
}

const candidate = JSON.parse(readFileSync(`${task}/api-candidate.json`, 'utf8').replace(/^\uFEFF/, ''));
candidate.approvedBy = 'user:cascade-lookups-2026-09-29';
candidate.approvedAt = '2026-09-29';
writeFileSync('build/api-snapshot-baseline.json', JSON.stringify(candidate, null, 2) + '\n', 'utf8');
