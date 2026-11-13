# vendor（同梱している外の部品）

CDN に頼らず、版を固定して同梱する。差し替えるときは、この表のハッシュも直す。

## ExcelJS 4.4.0（MIT）

- 出どころ＝npm の公式の包み https://registry.npmjs.org/exceljs/-/exceljs-4.4.0.tgz （2026-11-13 に取得）
- 包みの integrity（npm の公表値と、手元で測った値が一致）
  `sha512-XctvKaEMaj1Ii9oDOqbW/6e1gXknSY4g/aLCDicOXqBE4M0nRWkUu0PTp++UPNzoFY12BNHMfs/VadKIS6llvg==`
- ライセンス＝`LICENSE.exceljs.txt`（包みの `LICENSE` をそのまま置いた）

| ファイル | 包みの中の場所 | 大きさ | SHA-256 |
|---|---|---|---|
| `exceljs.min.js` | `dist/exceljs.min.js` | 947,702 | `7E49DA68588E250DBB8BBA190D2CAA8AB3787CC0284BDA1D8B2F805C4DF742C9` |
| `exceljs.bare.min.js` | `dist/exceljs.bare.min.js` | 862,631 | `B039E32BFB7164DADD84A6973B30200DC07E1B58E1F21C546378FCA6920A1F94` |
| `LICENSE.exceljs.txt` | `LICENSE` | 1,084 | `2ED526B7E7C7681DC2347995B5036DDF73DC27D0539CDEA59B6D4A72B63A57A0` |

中身には手を入れていない。
