import fs from 'node:fs/promises';
import path from 'node:path';

const assetsDir = path.resolve('dist', 'assets');
const legacyAliases = [
  'index-BFoxlFun.js',
  'index-B27fRKpq.js',
  'index-CIBoO_Gq.js'
];

const isIndexBundle = (fileName) => /^index-[A-Za-z0-9_-]+\.js$/.test(fileName);

async function main() {
  const files = await fs.readdir(assetsDir);
  const candidates = files.filter((f) => isIndexBundle(f) && !legacyAliases.includes(f));

  if (candidates.length === 0) {
    throw new Error(`Nenhum bundle index encontrado em ${assetsDir}`);
  }

  const withStat = await Promise.all(
    candidates.map(async (fileName) => ({
      fileName,
      stat: await fs.stat(path.join(assetsDir, fileName))
    }))
  );

  withStat.sort((a, b) => b.stat.mtimeMs - a.stat.mtimeMs);
  const latestBundle = withStat[0].fileName;

  const compatContent = (target) =>
    `// Compat alias generated post-build\nimport './${target}';\n`;

  for (const legacyName of legacyAliases) {
    if (legacyName === latestBundle) continue;

    const legacyPath = path.join(assetsDir, legacyName);
    await fs.writeFile(legacyPath, compatContent(latestBundle), 'utf8');
    console.log(`[postbuild:compat] Alias gerado: ${legacyName} -> ${latestBundle}`);
  }
}

main().catch((err) => {
  console.error('[postbuild:compat] Erro:', err.message);
  process.exit(1);
});
