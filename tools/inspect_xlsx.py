# 出来た台帳（xlsx）を本物の Excel で開いて検査する。Excel で開く(A)・PDF にして測る(B)・画像にして絵の中身を見る(C)。
# 使い方: python -X utf8 tools/inspect_xlsx.py out/<名前>.expect.json [--out out/inspect/<名前>] [--dpi 110] [--no-pictures]
# expect.json は tools/build_ledger.mjs が書く。任意の "title"（切った後の頭の題）が有れば、PDF の頭と比べる。
# 専用の Excel（DispatchEx）・Visible=False・DisplayAlerts=False・ReadOnly で開く・保存しない。名前で強制終了しない。
# 終了コード: 0＝全部 OK／1＝FAIL が有る／2＝使い方の誤り
import argparse, json, os, re, sys, time

import fitz  # PyMuPDF
import pythoncom
import win32com.client
from PIL import Image, ImageDraw

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---- 台帳の設計（src/layout.js の SHEET と同じ数字。試作 1a で本物の Excel に開いて確かめた物）----
SHEET_NAME = '写真台帳'
COL_PX = [420, 8, 64, 166]
ROW_PT = [18, 18, 45, 18, 18, 18, 18, 83.25, 7.5]
LABELS = ['写真番号', '撮影年月日', 'ファイル名', '工種', '種別', '細別', '撮影箇所', '備考']
ROWS_PER_BLOCK = 9
BLOCKS_PER_PAGE = 3
FRAME_PX = (420, 315)
FONT_NAME = '游ゴシック'
DEFAULT_TITLE = '工事写真台帳'

# ---- 許す幅（決めた数と理由は task-3b-report.md）----
TOL_COL_PX = 0.5      # 列の幅（指定）
TOL_ROW_PT = 0.1      # 行の高さ（指定）
TOL_SHAPE_IN = 0.5    # Excel 上の絵が枠の中（pt）
TOL_AR = 0.01         # 縦横比（指定）
TOL_PDF_IN = 0.3      # 紙の上の絵が枠の中（pt・絵は枠より 9pt 内側に置く設計）
FRAME_AR_REL = 0.06   # 紙の上の枠の縦横比＝設計(420/315)の ±6%（この PC は列 +1.9%・行 −1.9% で約 +3.9%）
TOL_TEXT_PT = 2.0     # 名前の欄の字の箱が、欄の線を越えてよい量（pt）。4 行に溢れると上下に 7.5pt 以上越える
CORNER_TOL = 60
COLORS = {'tl': (255, 0, 0), 'tr': (0, 255, 0), 'bl': (0, 0, 255), 'br': (255, 255, 0)}

NOW = time.time


def ws_strip(s):
    return re.sub(r'\s+', '', s or '')


class Book:
    def __init__(self, name):
        self.name = name
        self.fails = []
        self.measured = {}

    def fail(self, code, msg):
        self.fails.append(f'{code}: {msg}')


def screen_dpi():
    """Excel は行の高さを画面の画素の刻みに丸める。その刻みを知るために画面の DPI を読む（読めなければ 96＝倍率 100%）。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Control Panel\Desktop\WindowMetrics') as k:
            v = int(winreg.QueryValueEx(k, 'AppliedDPI')[0])
            if 96 <= v <= 480:
                return v
    except (OSError, ValueError, TypeError):
        pass
    return 96


def row_tolerance(want_pt, dpi):
    """指定の高さが画素の刻みに乗る時は TOL_ROW_PT、乗らない時は 1 画素ぶん（Excel が丸めるため）。"""
    step = 72.0 / dpi
    on_grid = abs(want_pt / step - round(want_pt / step)) < 1e-6
    return TOL_ROW_PT if on_grid else step + 0.01


def selftest():
    """row_tolerance の確かめ（Excel を起こさない・ファイルも要らない）。"""
    cases = [(83.25, 96, 0.1), (7.5, 96, 0.1), (18, 120, 0.1), (83.25, 120, 0.61), (7.5, 120, 0.61), (7.5, 144, 0.1), (83.25, 144, 0.51)]
    bad = [(h, d, w, row_tolerance(h, d)) for h, d, w in cases if not abs(row_tolerance(h, d) - w) < 1e-9]
    for h, d, w, got in bad:
        print(f'selftest 外れ：高さ {h}pt・DPI {d} → 許す幅 {got}（{w} のはず）')
    if bad:
        print(f'selftest: {len(cases) - len(bad)}/{len(cases)}')
        return 1
    print(f'selftest: {len(cases)}/{len(cases)}')
    return 0


def rect_in(inner, outer, tol):
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[2] <= outer[2] + tol and inner[3] <= outer[3] + tol)


def rects_overlap(a, b, eps=0.01):
    return min(a[2], b[2]) - max(a[0], b[0]) > eps and min(a[3], b[3]) - max(a[1], b[1]) > eps


# =========================================================== A. Excel で開く
def check_excel(xl, path, bk, exp):
    photos = exp['photos']
    n = len(photos)
    pages = exp['pages']
    pdf = os.path.join(bk.outdir, os.path.splitext(os.path.basename(path))[0] + '.pdf')
    wb = None
    try:
        try:
            wb = xl.Workbooks.Open(path, 0, True, None, None, None, True, None, None, False, False, None, False, None, 0)
        except Exception as e:
            bk.fail('修復', f'Excel が普通に開けない（修復が要る）：{e!r}')
            return None
        ws = None
        if wb.Worksheets.Count != 1:
            bk.fail('A シート', f'シートが {wb.Worksheets.Count} 枚（1 枚のはず）')
        ws = wb.Worksheets(1)
        if ws.Name != SHEET_NAME:
            bk.fail('A シート', f'シート名が「{ws.Name}」（「{SHEET_NAME}」のはず）')

        ps = ws.PageSetup
        if ps.PaperSize != 9:
            bk.fail('A 印刷設定', f'用紙 {ps.PaperSize}（A4＝9 のはず）')
        if ps.Orientation != 1:
            bk.fail('A 印刷設定', f'向き {ps.Orientation}（縦＝1 のはず）')
        if ps.FitToPagesWide != 1:
            bk.fail('A 印刷設定', f'FitToPagesWide {ps.FitToPagesWide}（1 のはず）')
        want_area = f'$A$1:$D${n * ROWS_PER_BLOCK}'
        if ps.PrintArea != want_area:
            bk.fail('A 印刷範囲', f'PrintArea {ps.PrintArea}（{want_area} のはず）')

        # 手動の改ページ（ExcelJS は読み返せないので Excel で見る）
        rows = [ws.HPageBreaks(i).Location.Row for i in range(1, ws.HPageBreaks.Count + 1)]
        types = [ws.HPageBreaks(i).Type for i in range(1, ws.HPageBreaks.Count + 1)]
        want_rows = [BLOCKS_PER_PAGE * ROWS_PER_BLOCK * k + 1 for k in range(1, pages)]
        bk.measured['hpagebreaks'] = rows
        if len(rows) != pages - 1:
            bk.fail('A 改ページ', f'改ページが {len(rows)} 個（ページ数 {pages}−1＝{pages - 1} 個のはず）')
        elif rows != want_rows:
            bk.fail('A 改ページ', f'改ページの行 {rows[:6]}…（{want_rows[:6]}… のはず）')
        if any(t != -4135 for t in types):
            bk.fail('A 改ページ', '手動でない改ページ（自動）が混じる')

        # 列の幅・行の高さ
        cols_px = [ws.Columns(i).Width * 96 / 72 for i in range(1, 5)]
        bk.measured['col_px'] = [round(v, 2) for v in cols_px]
        for i, (got, want) in enumerate(zip(cols_px, COL_PX)):
            if abs(got - want) > TOL_COL_PX:
                bk.fail('A 列の幅', f'{"ABCD"[i]} 列 {got:.2f}px（{want}px のはず）')
        bad_rows = 0
        dpi = bk.dpi
        bk.measured['dpi'] = dpi
        bk.measured['row_step_pt'] = round(72.0 / dpi, 4)
        rounded = []
        for r in range(1, n * ROWS_PER_BLOCK + 1):
            got = ws.Rows(r).Height
            want = ROW_PT[(r - 1) % ROWS_PER_BLOCK]
            tol = row_tolerance(want, dpi)
            if tol != TOL_ROW_PT and abs(got - want) > TOL_ROW_PT and abs(got - want) <= tol:
                pair = f'{want:g}→{got:g}'
                if pair not in rounded:
                    rounded.append(pair)
            if abs(got - want) > tol:
                bad_rows += 1
                if bad_rows <= 5:
                    bk.fail('A 行の高さ', f'{r} 行目 {got}pt（{want}pt のはず）')
        if bad_rows > 5:
            bk.fail('A 行の高さ', f'ほか {bad_rows - 5} 行')
        bk.measured['row_rounded'] = rounded

        # 標準の字体・保護
        font = wb.Styles('Normal').Font.Name
        bk.measured['normal_font'] = font
        if font != FONT_NAME:
            bk.fail('A 字体', f'標準の字体が「{font}」（「{FONT_NAME}」のはず）')
        if ws.ProtectContents:
            bk.fail('A 保護', 'シートに保護がかかっている')

        # 枠（A 列の結合セル）
        frames = []
        for i in range(n):
            first = i * ROWS_PER_BLOCK + 1
            a = ws.Cells(first, 1)
            area = a.MergeArea
            if area.Address != f'$A${first}:$A${first + 7}':
                bk.fail('A 結合', f'{i + 1} 段目の A 列の結合が {area.Address}')
            frames.append((area.Left, area.Top, area.Left + area.Width, area.Top + area.Height))

        # 絵
        cnt = ws.Shapes.Count
        if cnt != n:
            bk.fail('A 絵の数', f'絵が {cnt} 個（写真 {n} 枚のはず）')
        shapes = []
        for i in range(1, cnt + 1):
            s = ws.Shapes(i)
            shapes.append((s.Left, s.Top, s.Left + s.Width, s.Top + s.Height))
        owner = {}
        for si, sr in enumerate(shapes):
            cx, cy = (sr[0] + sr[2]) / 2, (sr[1] + sr[3]) / 2
            hit = [i for i, f in enumerate(frames) if f[0] <= cx <= f[2] and f[1] <= cy <= f[3]]
            if len(hit) != 1:
                bk.fail('A 絵の位置', f'絵 {si + 1} 個目が、どの枠の中にも無い')
            else:
                owner.setdefault(hit[0], []).append(si)
        for i, p in enumerate(photos):
            lst = owner.get(i, [])
            tag = f'写真番号 {p["number"]}（{p["name"]}）'
            if len(lst) != 1:
                if cnt == n:
                    bk.fail('A 絵の数', f'{tag} の枠に絵が {len(lst)} 個')
                continue
            sr = shapes[lst[0]]
            if not rect_in(sr, frames[i], TOL_SHAPE_IN):
                bk.fail('A 絵が枠の中', f'{tag} の絵が枠からはみ出す')
            ar = (sr[2] - sr[0]) / (sr[3] - sr[1])
            err = ar / (p['imageW'] / p['imageH']) - 1
            if abs(err) > TOL_AR:
                bk.fail('A 縦横比', f'{tag} の縦横比が {err * 100:.2f}% ずれる')
        for a in range(len(shapes)):
            for b in range(a + 1, len(shapes)):
                if rects_overlap(shapes[a], shapes[b]):
                    bk.fail('A 絵の重なり', f'絵 {a + 1} と {b + 1} が重なる')

        # セルの中身（まとめて読む）
        vals = ws.Range(ws.Cells(1, 1), ws.Cells(n * ROWS_PER_BLOCK, 4)).Value
        for i, p in enumerate(photos):
            first = i * ROWS_PER_BLOCK
            tag = f'写真番号 {p["number"]}（{p["name"]}）'
            for k, lab in enumerate(LABELS):
                got = vals[first + k][2]
                if got != lab:
                    bk.fail('A セル', f'{tag} の C 列 {k + 1} 行目が「{got}」（「{lab}」のはず）')
            d1, d2, d3 = (vals[first + k][3] for k in range(3))
            if not (isinstance(d1, (int, float)) and int(d1) == p['number'] and d1 == int(d1)):
                bk.fail('A セル', f'{tag} の D 列 1 行目（写真番号）が「{d1}」')
            if (d2 or '') != p['dateText']:
                bk.fail('A セル', f'{tag} の撮影年月日が「{d2 or ""}」（「{p["dateText"]}」のはず）')
            if (d3 or '') != p['shownName']:
                bk.fail('A セル', f'{tag} のファイル名が「{d3 or ""}」（「{p["shownName"]}」のはず）')

        # PDF
        if os.path.exists(pdf):
            os.remove(pdf)
        t = NOW()
        wb.ExportAsFixedFormat(0, pdf)
        bk.measured['pdf_sec'] = round(NOW() - t, 2)
        return pdf
    except Exception as e:
        bk.fail('検査中の例外', repr(e))
        return None
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass


# =========================================================== B・C. PDF と画像
def lines_of(page):
    vs, hs = [], []
    for d in page.get_drawings():
        for it in d['items']:
            if it[0] == 'l':
                p, q = it[1], it[2]
                if abs(p.x - q.x) < 0.5:
                    vs.append((p.x, min(p.y, q.y), max(p.y, q.y)))
                elif abs(p.y - q.y) < 0.5:
                    hs.append((p.y, min(p.x, q.x), max(p.x, q.x)))
            elif it[0] == 're':
                r = it[1]
                if r.width < 2 and r.height > 2:
                    vs.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
                elif r.height < 2 and r.width > 2:
                    hs.append(((r.y0 + r.y1) / 2, r.x0, r.x1))
    return vs, hs


def frame_of(img, vs, hs):
    cx, cy = (img.x0 + img.x1) / 2, (img.y0 + img.y1) / 2
    left = max([x for x, y0, y1 in vs if x <= img.x0 + 0.5 and y0 <= cy <= y1], default=None)
    right = min([x for x, y0, y1 in vs if x >= img.x1 - 0.5 and y0 <= cy <= y1], default=None)
    top = max([y for y, x0, x1 in hs if y <= img.y0 + 0.5 and x0 <= cx <= x1], default=None)
    bot = min([y for y, x0, x1 in hs if y >= img.y1 - 0.5 and x0 <= cx <= x1], default=None)
    if None in (left, right, top, bot):
        return None
    return (left, top, right, bot)


def row_bounds(hs, frame, d_x):
    """D 列の横線を集めて、枠の上から下までの 9 本（8 行の境）にする。取れなければ None。"""
    ys = sorted(y for y, x0, x1 in hs if x0 - 1 <= d_x <= x1 + 1 and frame[1] - 1 <= y <= frame[3] + 1)
    out = []
    for y in ys:
        if not out or y - out[-1] > 0.5:
            out.append(y)
    return out if len(out) == 9 else None


def span_rows(page):
    # Excel の PDF は「見出し＋値」を 1 つの文字列にまとめる事が有るので、1 字ずつ（箱つき）で取る。順は PDF の描いた順。
    out = []
    for b in page.get_text('rawdict')['blocks']:
        for ln in b.get('lines', []):
            for sp in ln['spans']:
                for ch in sp['chars']:
                    if ch['c'].strip():
                        out.append((ch['bbox'], ch['c']))
    return out


def pix_to_image(pix):
    return Image.frombytes('RGB', (pix.width, pix.height), pix.samples)


def patch_mean(im, x, y, half, gray=False):
    px = []
    for yy in range(int(y) - half, int(y) + half + 1):
        for xx in range(int(x) - half, int(x) + half + 1):
            if 0 <= xx < im.width and 0 <= yy < im.height:
                px.append(im.getpixel((xx, yy)))
    if not px:
        return None
    if gray:
        return sum(0.299 * r + 0.587 * g + 0.114 * b for r, g, b in px) / len(px)
    return tuple(sorted(c[i] for c in px)[len(px) // 2] for i in range(3))


def read_band(im, box, dpi):
    s = dpi / 72
    x0, y0, x1, y1 = (v * s for v in box)
    w, h = x1 - x0, y1 - y0
    cy = y0 + 0.935 * h
    bits = []
    for i in range(12):
        cx = x0 + 0.25 * w + 0.5 * w * (i + 0.5) / 12
        m = patch_mean(im, cx, cy, 2, gray=True)
        bits.append(1 if m is not None and m < 128 else 0)
    n = 0
    for b in bits[1:11]:
        n = n * 2 + b
    return bits[0] == 1 and bits[11] == 1, n


def corner_problem(im, box, dpi):
    s = dpi / 72
    x0, y0, x1, y1 = (v * s for v in box)
    w, h = x1 - x0, y1 - y0
    pts = {'tl': (x0 + 0.04 * w, y0 + 0.04 * h), 'tr': (x1 - 0.04 * w, y0 + 0.04 * h),
           'bl': (x0 + 0.04 * w, y1 - 0.04 * h), 'br': (x1 - 0.04 * w, y1 - 0.04 * h)}
    bad = []
    for k, (x, y) in pts.items():
        got = patch_mean(im, x, y, 1)
        if got is None or any(abs(a - b) > CORNER_TOL for a, b in zip(got, COLORS[k])):
            bad.append(f'{k}={got}（{COLORS[k]} のはず）')
    return bad


def check_pdf(pdf, bk, exp, title, dpi, pictures, stem):
    photos = exp['photos']
    n = len(photos)
    pages = exp['pages']
    doc = fitz.open(pdf)
    bk.measured['pdf_pages'] = doc.page_count
    if doc.page_count != pages:
        bk.fail('B ページ数', f'PDF が {doc.page_count} ページ（{pages} ページのはず）')
    want_title = ws_strip(title)
    pngs = []
    ar_nominal = FRAME_PX[0] / FRAME_PX[1]
    for pi in range(doc.page_count):
        page = doc[pi]
        pix = page.get_pixmap(dpi=dpi)
        png_path = os.path.join(bk.outdir, f'{stem}_p{pi + 1:02d}.png')
        pix.save(png_path)
        pngs.append(png_path)
        im = pix_to_image(pix) if pictures else None
        infos = sorted(page.get_image_info(), key=lambda i: (i['bbox'][1], i['bbox'][0]))
        want_n = max(0, min(BLOCKS_PER_PAGE, n - pi * BLOCKS_PER_PAGE))
        if len(infos) != want_n:
            bk.fail('B 絵の数', f'p{pi + 1} の絵が {len(infos)} 個（{want_n} 個のはず）')
        vs, hs = lines_of(page)
        spans = span_rows(page)

        # 頭の題と足のページ番号
        H = page.rect.height
        head = ws_strip(''.join(t for b, t in spans if b[3] < 40))
        foot = ws_strip(''.join(t for b, t in spans if b[1] > H - 40))
        if head != want_title:
            bk.fail('B 題', f'p{pi + 1} の頭が「{head}」（「{want_title}」のはず）')
        if foot != f'{pi + 1}/{pages}':
            bk.fail('B ページ番号', f'p{pi + 1} の足が「{foot}」（「{pi + 1}/{pages}」のはず）')

        boxes = []
        for j, inf in enumerate(infos):
            gi = pi * BLOCKS_PER_PAGE + j
            if gi >= n:
                break
            p = photos[gi]
            tag = f'p{pi + 1} 写真番号 {p["number"]}（{p["name"]}）'
            r = fitz.Rect(inf['bbox'])
            boxes.append((r, p, tag))
            fr = frame_of(r, vs, hs)
            if fr is None:
                bk.fail('B 枠', f'{tag} の枠の罫線が見つからない')
                continue
            fw, fh = fr[2] - fr[0], fr[3] - fr[1]
            far = fw / fh
            if abs(far / ar_nominal - 1) > FRAME_AR_REL:
                bk.fail('B 枠', f'{tag} の枠の縦横比 {far:.3f}（設計 {ar_nominal:.3f} の ±{FRAME_AR_REL * 100:.0f}% の外）')
            if not rect_in((r.x0, r.y0, r.x1, r.y1), fr, TOL_PDF_IN):
                bk.fail('B 絵が枠の中', f'{tag} の絵が枠の外')
            err = (r.width / r.height) / (p['imageW'] / p['imageH']) - 1
            if abs(err) > TOL_AR:
                bk.fail('B 縦横比', f'{tag} の縦横比が {err * 100:.2f}% ずれる')

            # 文字（D 列の行ごと・C 列の見出し）
            sx = fw / FRAME_PX[0]
            d_x0 = fr[2] + (COL_PX[1] + COL_PX[2]) * sx
            d_x1 = d_x0 + COL_PX[3] * sx
            c_x0 = fr[2] + COL_PX[1] * sx
            rb = row_bounds(hs, fr, (d_x0 + d_x1) / 2)
            if rb is None:
                bk.fail('B 罫線', f'{tag} の D 列の横線が 9 本取れない')
                continue
            cell = {c: [''] * 8 for c in 'CD'}
            over = []
            for bb, t in spans:
                cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
                if not (fr[1] - 1 <= cy <= fr[3] + 1):
                    continue
                col = 'D' if d_x0 - 1 <= cx <= d_x1 + 1 else 'C' if c_x0 - 1 <= cx < d_x0 - 1 else None
                if col is None:
                    continue
                row = next((k for k in range(8) if rb[k] <= cy < rb[k + 1]), None)
                if row is None:
                    continue
                cell[col][row] += t
                if col == 'D' and row == 2 and (bb[3] > rb[3] + TOL_TEXT_PT or bb[1] < rb[2] - TOL_TEXT_PT):
                    over.append(t)
            for k, lab in enumerate(LABELS):
                if ws_strip(cell['C'][k]) != lab:
                    bk.fail('B 見出し', f'{tag} の見出し {k + 1} 行目が「{ws_strip(cell["C"][k])}」')
            if ws_strip(cell['D'][0]) != str(p['number']):
                bk.fail('B 文字', f'{tag} の写真番号の欄が「{ws_strip(cell["D"][0])}」')
            if ws_strip(cell['D'][1]) != ws_strip(p['dateText']):
                bk.fail('B 文字', f'{tag} の撮影年月日の欄が「{ws_strip(cell["D"][1])}」（「{p["dateText"]}」のはず）')
            if ws_strip(cell['D'][2]) != ws_strip(p['shownName']):
                bk.fail('B 文字', f'{tag} のファイル名の欄が「{ws_strip(cell["D"][2])}」（「{ws_strip(p["shownName"])}」のはず）')
            if over:
                bk.fail('B 名前の欄', f'{tag} の字が欄の線を越える：{"".join(over)[:40]}')
            if any(ws_strip(cell['D'][k]) for k in range(3, 8)):
                bk.fail('B 文字', f'{tag} の工種〜備考の欄に字が有る（名前の溢れの疑い）：{ws_strip("".join(cell["D"][3:]))[:40]}')
        for a in range(len(infos)):
            for b in range(a + 1, len(infos)):
                if fitz.Rect(infos[a]['bbox']).intersects(fitz.Rect(infos[b]['bbox'])):
                    bk.fail('B 絵の重なり', f'p{pi + 1} の絵 {a + 1} と {b + 1} が重なる')

        # C. 画像にして絵の中身を見る
        if pictures:
            for r, p, tag in boxes:
                if p.get('label') is None:
                    continue
                bk.measured['c_checked'] = bk.measured.get('c_checked', 0) + 1
                box = (r.x0, r.y0, r.x1, r.y1)
                bad = corner_problem(im, box, dpi)
                if bad:
                    bk.fail('C 四隅の色', f'{tag}：{" ".join(bad)}')
                ends, num = read_band(im, box, dpi)
                if not ends:
                    bk.fail('C 番号の帯', f'{tag}：帯の両端が黒でない')
                elif num != p['label']:
                    bk.fail('C 番号の帯', f'{tag}：帯が {num}（期待 {p["label"]}）＝絵と隣のファイル名が別の写真')
    doc.close()
    # 見るための一覧（6 ページを 3×2・幅 2400px）
    for g in range(0, len(pngs), 6):
        group = pngs[g:g + 6]
        cw = 800
        thumbs = []
        for pth in group:
            t = Image.open(pth).convert('RGB')
            t = t.resize((cw, round(t.height * cw / t.width)), Image.LANCZOS)
            thumbs.append(t)
        ch = max(t.height for t in thumbs)
        sheet = Image.new('RGB', (cw * 3, ch * 2), (255, 255, 255))
        d = ImageDraw.Draw(sheet)
        for k, t in enumerate(thumbs):
            ox, oy = (k % 3) * cw, (k // 3) * ch
            sheet.paste(t, (ox, oy))
            d.rectangle([ox, oy, ox + cw - 1, oy + t.height - 1], outline=(150, 150, 150))
        sheet.save(os.path.join(bk.outdir, f'{stem}_sheet_{g // 6 + 1:02d}.png'))
    return pngs


# =========================================================== 本体
def main():
    ap = argparse.ArgumentParser(description='台帳（xlsx）を本物の Excel で開いて検査する')
    ap.add_argument('expect', nargs='?')
    ap.add_argument('--selftest', action='store_true')
    ap.add_argument('--out', default=None)
    ap.add_argument('--dpi', type=int, default=110)
    ap.add_argument('--no-pictures', action='store_true')
    a = ap.parse_args()
    if a.selftest:
        return selftest()
    if not a.expect:
        ap.error('expect.json の道が要る（または --selftest）')
    if not os.path.isfile(a.expect):
        print(f'expect.json が無い：{a.expect}', file=sys.stderr)
        return 2
    exp_path = os.path.abspath(a.expect)
    exp = json.load(open(exp_path, encoding='utf-8'))
    base = os.path.basename(exp_path)
    name = base[:-len('.expect.json')] if base.endswith('.expect.json') else os.path.splitext(base)[0]
    outdir = os.path.abspath(a.out or os.path.join(ROOT, 'out', 'inspect', name))
    os.makedirs(outdir, exist_ok=True)
    title = exp.get('title', DEFAULT_TITLE)
    results = []
    if not exp.get('books'):
        print('FAIL  expect.json に books が無い（空）＝検査するブックが 0 冊', file=sys.stderr)
        return 1

    dpi = screen_dpi()
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx('Excel.Application')
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        try:
            xl.AutomationSecurity = 3
        except Exception:
            pass
        for b in exp['books']:
            path = os.path.join(os.path.dirname(exp_path), b['file'])
            bk = Book(b['file'])
            bk.outdir = outdir
            bk.dpi = dpi
            t0 = NOW()
            stem = os.path.splitext(b['file'])[0]
            if not b.get('photos'):
                bk.fail('入力', 'photos が空（検査する写真が 0 枚）')
            elif not os.path.isfile(path):
                bk.fail('入力', f'xlsx が無い：{path}')
            else:
                bk.measured['bytes'] = os.path.getsize(path)
                pdf = check_excel(xl, path, bk, b)
                if pdf and os.path.isfile(pdf):
                    pictures = not a.no_pictures
                    try:
                        check_pdf(pdf, bk, b, b.get('title', title), a.dpi, pictures, stem)
                    except Exception as e:
                        bk.fail('検査中の例外', f'PDF の検査：{e!r}')
                elif pdf or not bk.fails:
                    # PDF が出来ていないのに B・C を黙って飛ばして OK にしない
                    bk.fail('B PDF', f'PDF が出来ていない：{pdf}')
            bk.measured['sec'] = round(NOW() - t0, 1)
            n_ph = len(b.get('photos', []))
            c_done = bk.measured.get('c_checked', 0)
            bk.measured['c_checked'] = c_done
            bk.measured['c_skipped'] = n_ph - c_done
            results.append({'file': b['file'], 'ok': not bk.fails, 'fails': bk.fails, 'measured': bk.measured})
            head = 'OK  ' if not bk.fails else f'FAIL {len(bk.fails)} 件  '
            print(f'{head}{b["file"]}  {len(b.get("photos", []))} 枚  {b.get("pages")} ページ  {bk.measured.get("bytes", 0)} B  {bk.measured["sec"]} 秒  '
                  + (f'C {c_done}/{n_ph}' if c_done == n_ph else f'C {c_done}/{n_ph}（飛ばした {n_ph - c_done}）'))
            if bk.measured.get('row_rounded'):
                pairs = '、'.join(p + 'pt' for p in bk.measured['row_rounded'])
                print(f'    行の高さ：画面の倍率 {round(dpi / 96 * 100)}%（DPI {dpi}・1 画素 {72.0 / dpi:g}pt）。画素の刻みに乗らない高さは Excel が丸めるので、1 画素まで許した（{pairs}）')
            for f in bk.fails[:40]:
                print('   ', f)
            if len(bk.fails) > 40:
                print(f'    …ほか {len(bk.fails) - 40} 件（result.json に全部）')
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        del xl
        pythoncom.CoUninitialize()

    ok = all(r['ok'] for r in results)
    with open(os.path.join(outdir, 'result.json'), 'w', encoding='utf-8') as fp:
        json.dump({'expect': exp_path, 'dpi': a.dpi, 'pictures': not a.no_pictures, 'ok': ok, 'books': results},
                  fp, ensure_ascii=False, indent=2)
    print(f'{"全部 OK" if ok else "FAIL あり"}  → {outdir}')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
