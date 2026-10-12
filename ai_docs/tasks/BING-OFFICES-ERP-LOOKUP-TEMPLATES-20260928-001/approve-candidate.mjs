import { readFileSync, writeFileSync } from 'node:fs';

const directory = 'ai_docs/tasks/BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001';
const path = 'build/api-snapshot-baseline.json';
const original = readFileSync(path);
const candidate = JSON.parse(readFileSync(`${directory}/api-candidate.json`, 'utf8')
    .replace(/^\uFEFF/, ''));
const difference = JSON.parse(readFileSync(`${directory}/api-diff.json`, 'utf8'));
for (const tfm of ['net6.0', 'net8.0']) {
    const assemblies = Object.entries(difference[tfm]);
    if (assemblies.length !== 1 || assemblies[0][0] !== 'Bing.Offices.Abstractions'
        || assemblies[0][1].added.length !== 14 || assemblies[0][1].removed.length !== 0)
        throw new Error(`Unexpected API change in ${tfm}`);
}
if (!candidate.candidateIdentity?.candidateSourceManifestSha256
    || !candidate.candidateIdentity?.nupkgFiles
    || !candidate.candidateIdentity?.breakingApprovalSha256)
    throw new Error('Candidate identity is incomplete');

candidate.approvedBy = 'user:erp-lookup-templates-2026-09-29';
candidate.approvedAt = '2026-09-29';
const hasBom = original[0] === 0xef && original[1] === 0xbb && original[2] === 0xbf;
const serialized = JSON.stringify(candidate, null, 2) + '\n';
writeFileSync(path, (hasBom ? '\uFEFF' : '') + serialized, 'utf8');
