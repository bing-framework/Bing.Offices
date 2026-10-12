import { readFileSync, writeFileSync } from 'node:fs';

const directory = 'ai_docs/tasks/BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001';
for (const name of ['api-candidate.json', 'api-diff.json', 'api-snapshot-net6.0.json',
    'api-snapshot-net8.0.json']) {
    const path = `${directory}/${name}`;
    const content = readFileSync(path, 'utf8').replace(/^\uFEFF/, '')
        .replace(/\r\n?/g, '\n').replace(/\n*$/, '\n');
    writeFileSync(path, content, 'utf8');
}
