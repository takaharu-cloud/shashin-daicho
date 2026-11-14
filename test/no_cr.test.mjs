import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

// 置き先に写すファイル（web/ と src/）に、改行の CR（0x0D）が混ざっていない事。
// Windows の字のモードで書き直すと CRLF になり、置き先のリポにそのまま入ってしまう。vendor/ は元のまま持つ決まりなので見ない。
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

for (const dir of ['web', 'src']) {
  test(`${dir}/ の全部のファイルに CR（0x0D）が 1 バイトも無い`, () => {
    const files = readdirSync(path.join(ROOT, dir), { withFileTypes: true }).filter((d) => d.isFile()).map((d) => d.name);
    assert.ok(files.length > 0, `${dir}/ にファイルが無い`);
    const withCr = files.filter((f) => readFileSync(path.join(ROOT, dir, f)).includes(0x0d));
    assert.deepEqual(withCr, [], `CR が入っているファイル：${withCr.join('、')}`);
  });
}
