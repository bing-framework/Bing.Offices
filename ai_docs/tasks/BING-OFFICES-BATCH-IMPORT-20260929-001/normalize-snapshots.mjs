import { readFileSync, writeFileSync } from 'node:fs';

const directory = 'ai_docs/tasks/BING-OFFICES-BATCH-IMPORT-20260929-001/artifacts';
for (const name of ['api-candidate.json', 'api-snapshot-net6.0.json', 'api-snapshot-net8.0.json', 'api-diff.json']) {
    const path = `${directory}/${name}`;
    try {
        const value = JSON.parse(readFileSync(path, 'utf8').replace(/^\uFEFF/, ''));
        writeFileSync(path, JSON.stringify(value, null, 2) + '\n', 'utf8');
    } catch (error) {
        if (error.code !== 'ENOENT') throw error;
    }
}
