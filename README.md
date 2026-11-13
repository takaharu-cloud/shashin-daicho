# shashin-daicho（工事写真台帳を作る道具）

写真のフォルダを選ぶと、Excel の写真台帳（A4 縦・1 ページ 3 枚）を作るページ。ブラウザの中だけで動き、写真は PC の外に出ない。

## 決まり

- **原本は読むだけ。** 台帳に貼る写しだけを縮小する（国土交通省「写真管理基準（案）」令和8年3月改定 2-5 は、写真の編集を認めていない）。
- 撮影年月日は EXIF の DateTimeOriginal から入れる。無い写真は空欄にする（ファイルの更新日時を撮影日として入れない）。
- 開けない写真は、理由をつけて一覧に出す。黙って飛ばさない。
- 置き先のサイトの CSP（`script-src 'self'`・`style-src 'self'`・`img-src 'self'`）の下で動かす。インラインのスクリプトとスタイル、`eval`、`blob:` の画像は使わない。
- 試す写真は `tools/make_fixtures.py` で作る架空の絵だけ。実物の現場写真は使わない。
- 国の工事の電子納品（写真の原本を電子媒体で出す）の代わりにはならない。

## 並び

| 場所 | 何 |
|---|---|
| `vendor/` | ExcelJS 4.4.0（MIT）。版を固定して同梱する |
| `src/` | `exif.js`・`layout.js`・`sort.js`・`ledger.js`・`pipeline.js` |
| `web/` | ページ（`index.html`・`app.js`・`app.css`） |
| `test/` | `node --test` の単体テスト |
| `tools/` | 架空の写真を作る・作った台帳を Excel で開いて検査する |
| `scripts/` | 置き先と同じ CSP を付けた試験用のサーバー |

## 試し方

単体テスト（36 件）:

```
npm test
```

架空の写真を作る（`fixtures/` の下に作る。git の対象外。Python 3 と Pillow だけを使う）:

```
python -X utf8 tools/make_fixtures.py --set smoke|mix60|bulk300|big|all
```

組ごとに `fixtures/<組>/`（写真）・`fixtures/<組>.manifest.json`（正解）・`fixtures/<組>.copies/`（台帳に貼る写しの見本）の 3 つができる。作り直すときは、その組の 3 つだけを消して作る。作った後、道具が四隅の色と番号の帯を自分で確かめる。

作った写真が書き換わっていないかを確かめる（md5 と重さを manifest と比べる）:

```
python -X utf8 tools/make_fixtures.py --verify smoke|mix60|bulk300|big|all
```

`src/exif.js` の読み取りを manifest と突き合わせる:

```
node tools/check_fixtures.mjs smoke|mix60|bulk300|big|all
```
