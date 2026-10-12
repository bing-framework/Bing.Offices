import { spawnSync } from 'node:child_process';

for (const tfm of ['Net6', 'Net8']) {
    const project = `tests/Bing.Offices.Consumer.${tfm}/Bing.Offices.Consumer.${tfm}.csproj`;
    const result = spawnSync('dotnet', ['run', '--project', project, '-c', 'Release',
        '--no-build', '--no-restore', '-p:BingOfficesPackageVersion=2.0.0'], {
        env: { ...process.env, BING_OFFICES_PACKAGE_VERSION: '2.0.0' },
        stdio: 'inherit',
        shell: false
    });
    if (result.error)
        throw result.error;
    if (result.status !== 0)
        process.exit(result.status ?? 1);
}
