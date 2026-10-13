import { readFileSync, writeFileSync } from 'node:fs';

const task = 'ai_docs/tasks/BING-OFFICES-BATCH-IMPORT-20260929-001';
const differences = JSON.parse(readFileSync(`${task}/api-diff.json`, 'utf8'));
for (const tfm of ['net6.0', 'net8.0']) {
    const expected = differences[tfm];
    const assemblies = Object.keys(expected).sort();
    if (assemblies.join('|') !== 'Bing.Offices.Abstractions|Bing.Offices.ExcelDataReader')
        throw new Error(`Unexpected assemblies in ${tfm}: ${assemblies.join(', ')}`);
    if (expected['Bing.Offices.Abstractions'].removed.length !== 0
        || expected['Bing.Offices.Abstractions'].added.length !== 31)
        throw new Error(`Unexpected Abstractions diff in ${tfm}`);
    if (expected['Bing.Offices.ExcelDataReader'].removed.length !== 1
        || expected['Bing.Offices.ExcelDataReader'].added.length !== 3)
        throw new Error(`Unexpected ExcelDataReader diff in ${tfm}`);
}

const candidate = JSON.parse(readFileSync(`${task}/artifacts/api-candidate.json`, 'utf8'));
candidate.approvedBy = 'user:batch-import-2026-09-29';
candidate.approvedAt = '2026-09-29';
writeFileSync('build/api-snapshot-baseline.json', JSON.stringify(candidate, null, 2) + '\n', 'utf8');
