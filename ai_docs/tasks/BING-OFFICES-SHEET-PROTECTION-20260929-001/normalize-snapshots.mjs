import { readFileSync, writeFileSync } from 'node:fs';

const task = 'ai_docs/tasks/BING-OFFICES-SHEET-PROTECTION-20260929-001';
for (const name of ['api-candidate.json', 'api-snapshot-net6.0.json', 'api-snapshot-net8.0.json']) {
    const path = `${task}/${name}`;
    const value = readFileSync(path, 'utf8').replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
    writeFileSync(path, value.replace(/\n*$/, '') + '\n', 'utf8');
}
