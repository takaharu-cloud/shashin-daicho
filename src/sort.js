// 写真の一覧の並べ替え。DOM に触らない。

const collator = new Intl.Collator('ja', { numeric: true, sensitivity: 'base' });
const key = (it) => it.relativePath ?? it.name ?? '';
const byName = (a, b) => collator.compare(key(a), key(b));

/**
 * @param {Array<{name:string, relativePath?:string, dateTimeOriginal:string|null}>} items
 * @param {'date'|'name'} mode 知らない値は 'name' と同じ
 * @returns 新しい配列（元の配列は変えない）
 */
export function sortPhotos(items, mode) {
  const list = [...items];
  if (mode !== 'date') return list.sort(byName);
  return list.sort((a, b) => {
    const da = a.dateTimeOriginal ?? null;
    const db = b.dateTimeOriginal ?? null;
    if (da === null && db === null) return byName(a, b);
    if (da === null) return 1;
    if (db === null) return -1;
    if (da !== db) return da < db ? -1 : 1;
    return byName(a, b);
  });
}
