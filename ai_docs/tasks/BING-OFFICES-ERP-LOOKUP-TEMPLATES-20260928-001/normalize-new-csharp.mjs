import { readFileSync, writeFileSync } from 'node:fs';

const files = [
    'src/Bing.Offices.Abstractions/Bing/Offices/Exports/Content/ExcelLookupListDefinition.cs',
    'src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiLookupListWriter.cs',
    'src/Bing.Offices.ClosedXml/Bing/Offices/Exports/ClosedXmlLookupListWriter.cs',
    'tests/Bing.Offices.Npoi.Tests/EntityNativeValidationTest.cs',
    'tests/Bing.Offices.ClosedXml.Tests/EntityNativeValidationTest.cs',
    'tests/Bing.Offices.Npoi.Tests/LookupListModelTest.cs'
];
for (const path of files) {
    const body = readFileSync(path, 'utf8').replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
    writeFileSync(path, '\uFEFF' + body.replace(/\n*$/, '\n'), 'utf8');
}
