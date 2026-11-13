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
| `scripts/` | 置き先に置く形を組む・置き先と同じ CSP を付けた試験用のサーバー・ブラウザを通す試験 |
| `e2e/` | ブラウザを通す試験のページ（試験用のサーバーの `--e2e` の時だけ配る。本番のページではない） |

## 試し方

単体テスト（94 件）:

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

出来た台帳を本物の Excel で開いて検査する（Windows・Excel・Python 3・pywin32・PyMuPDF・Pillow が要る）。まず台帳を出し、次に検査する:

```
node tools/build_ledger.mjs smoke|mix60|bulk300 [--sort date|name]
python -X utf8 tools/inspect_xlsx.py out/<組>_<sort>.expect.json [--out out/inspect/<名前>] [--dpi 110] [--no-pictures]
```

専用の Excel を起こして読み取り専用で開く（保存しない・終わったら自分の Excel だけ閉じる）。A＝修復なしで開く・シート・印刷設定・改ページ・列の幅と行の高さ・絵の位置と縦横比・セル、B＝PDF にして絵・文字・題・ページ番号・名前の欄を測る、C＝画像にして四隅の色と番号の帯を読む。外れた判定は 1 行ずつ出て、終了コードは 1。`--out` の下に `result.json`・`<ブック>.pdf`・`<ブック>_pNN.png`（全ページ）・`<ブック>_sheet_NN.png`（6 ページを 3×2 に並べた一覧。人が目で見る用）を出す。`expect.json` に任意の `"title"`（切った後の頭の題）が有れば PDF の頭と比べる。

Excel は行の高さを画面の画素の刻み（100% は 0.75pt・125% は 0.6pt）に丸めるので、刻みに乗らない高さ（83.25pt・7.5pt など）は、画面の倍率によって 82.8pt・7.2pt のように読める。検査は、画面の DPI（レジストリの `AppliedDPI`。読めなければ 96）から刻みを出し、刻みに乗らない行だけ 1 画素まで許す（ファイルの中の値は `test/` が確かめている）。許す幅の計算だけは Excel なしで `python -X utf8 tools/inspect_xlsx.py --selftest` で確かめられる。

置き先に置く形を組む（`dist/shashin-daicho/` を作り直す。`web/`・`src/` → `lib/`・`vendor/` の ExcelJS と LICENSE を写し、SHA-256 が元と同じ事を 1 行ずつ出す）:

```
node scripts/build_site.mjs
```

置き先と同じ CSP を付けた試験用のサーバー（`127.0.0.1` だけで待つ。`/shashin-daicho/` を配る。`--assets` を渡した時だけ `/assets/` をそのフォルダから読むだけで配る）:

```
node scripts/serve_csp.mjs [--port 8733] [--assets <フォルダ>]
```

ブラウザ（Edge・Chrome の headless）を通す試験。置き先と同じ CSP の下で、架空の写真を 1 枚ずつ写しにして台帳を作り、写しの縦横・四隅の色・番号の帯・撮影日時・向き・CSP の違反・外への要求を確かめる。前後で原本の照合も回す。ブラウザは CDP で操り、止めるのは自分が起こした PID だけ:

```
node scripts/run_e2e.mjs <edge|chrome> <smoke|mix60|bulk300|big> [--sort date|name]
python -X utf8 tools/inspect_xlsx.py out/e2e/<ブラウザ>_<組>_<sort>.expect.json
```

`out/e2e/` に、台帳（xlsx）・`expect.json`・ページの報告（`report.json`）・終わりの画面（png）・サーバーが受けた要求の一覧が残る。外れが有れば 1 行ずつ出て、終了コードは 1。

本番のページ（`web/index.html`）を、人が使うのと同じ順に動かす試験。`dist/` を組み直し、置き先と同じ CSP の下で開き、`#input-files` に組の写真を渡して、一覧（並び・まとめ・開けなかった写真・小さい絵）を確かめ、保存のボタンを押して落ちた台帳を `tools/inspect_xlsx.py` で検査する。続けて、中止・`#input-folder`（フォルダの道が入るか）・300 枚を超えた時・写真が 0 の時・最初からやり直す、を回す。ページの見た目は置き先のサイトの `inner.css` に頼るので、`--assets` にその assets のフォルダを渡す:

```
node scripts/run_page.mjs <edge|chrome> <smoke|mix60|bulk300> --assets <置き先のサイトの assets> [--shots]
```

`out/page/<ブラウザ>_<組>/` に落ちた台帳・`expect.json`・要求の一覧、`out/inspect/<ブラウザ>_<組>/` に検査の結果が残る。`--shots` を付けると、幅 1280 と 390 の画面を `out/page/shots/` に撮る（人が目で見る用）。前後で原本の照合も回す。外れが有れば 1 行ずつ出て、終了コードは 1。
