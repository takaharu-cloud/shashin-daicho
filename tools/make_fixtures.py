"""架空の写真を作る（試験の材料。ページには入らない）。

正しい向きで見て 左上＝赤・右上＝緑・左下＝青・右下＝黄、真ん中に黒い番号、上の辺の真ん中に上向きの黒い三角。
使い方:
  python -X utf8 tools/make_fixtures.py --set smoke|mix60|bulk300|big|all [--out fixtures]
  python -X utf8 tools/make_fixtures.py --verify smoke|mix60|bulk300|big|all [--out fixtures]
部品は Pillow だけ。実物の写真・実在の名前は使わない。乱数は種を決めて作る（同じ環境なら同じ md5）。
"""
import argparse
import datetime
import hashlib
import io
import json
import os
import random
import shutil
import sys
import time

from PIL import Image, ImageDraw, ImageFont, ImageOps

# 108MP の絵（12,000×9,000）は Pillow の既定の上限（約 8,900 万画素）を超えて警告が出る。
# 自分で作った絵だけを開くので、上限だけを 2 億画素に上げる（それを超える絵は今も例外になる）。
Image.MAX_IMAGE_PIXELS = 200_000_000

REPO = os.path.realpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
FIXTURES_ROOT = os.path.join(REPO, 'fixtures')
SEED = 20261113
SETS = ['smoke', 'mix60', 'bulk300', 'big']
BASE = datetime.datetime(2026, 4, 1, 9, 0, 0)

NOISE_ALPHA = 0.20   # ノイズの濃さ（12MP の JPEG が約 2.5 MB になる値）
JPEG_QUALITY = 80
COPY_LONG = 1200
COPY_QUALITY = 80
COLORS = {'tl': (255, 0, 0), 'tr': (0, 255, 0), 'bl': (0, 0, 255), 'br': (255, 255, 0)}

FONT_PATH = 'C:/Windows/Fonts/arialbd.ttf'

T = Image.Transpose
# 保存する画素 ＝ 正しい向きの絵をこう変えた物。Orientation で直すと元に戻る。
TO_STORED = {
    2: T.FLIP_LEFT_RIGHT, 3: T.ROTATE_180, 4: T.FLIP_TOP_BOTTOM,
    5: T.TRANSPOSE, 6: T.ROTATE_90, 7: T.TRANSVERSE, 8: T.ROTATE_270,
}


def fmt_dt(dt):
    return dt.strftime('%Y:%m:%d %H:%M:%S')


# ---------------------------------------------------------------- 絵を描く

_noise_row = {}   # 幅 → ノイズの 1 帯（幅 w・高さ 1024）
_bg_cache = {}


def _tile():
    rng = random.Random(SEED * 7919 + 1)
    return Image.frombytes('RGB', (1024, 1024), rng.randbytes(1024 * 1024 * 3))


_TILE = None


def noise_band(w):
    global _TILE
    if _TILE is None:
        _TILE = _tile()
    if w not in _noise_row:
        band = Image.new('RGB', (w, 1024))
        for x in range(0, w, 1024):
            band.paste(_TILE, (x, 0))
        _noise_row.clear()   # 幅ごとに 1 つだけ持つ（メモリ節約）
        _noise_row[w] = band
    return _noise_row[w]


def build_bg(w, h, noise):
    """明るい灰色の地に薄い縞。noise なら薄いノイズを混ぜる。帯ごとに作ってメモリを抑える。"""
    step = max(w // 20, 2)
    bg = Image.new('RGB', (w, h), (200, 200, 200))
    nb = noise_band(w) if noise else None
    for y0 in range(0, h, 1024):
        bh = min(1024, h - y0)
        band = Image.new('RGB', (w, bh), (200, 200, 200))
        d = ImageDraw.Draw(band)
        for x in range(0, w, step):
            d.rectangle([x, 0, x + max(step // 2, 1) - 1, bh - 1], fill=(190, 195, 205))
        if nb is not None:
            band = Image.blend(band, nb.crop((0, 0, w, bh)), NOISE_ALPHA)
        bg.paste(band, (0, y0))
    return bg


def get_bg(w, h, noise):
    if w * h > 20_000_000:
        return build_bg(w, h, noise)   # 大きい絵は 1 回しか使わない＝持たない
    key = (w, h, noise)
    if key not in _bg_cache:
        _bg_cache[key] = build_bg(w, h, noise)
    return _bg_cache[key].copy()


def make_upright(w, h, label, noise=True):
    im = get_bg(w, h, noise)
    d = ImageDraw.Draw(im)
    s = int(max(w, h) * 0.10)
    d.rectangle([0, 0, s - 1, s - 1], fill=COLORS['tl'])
    d.rectangle([w - s, 0, w - 1, s - 1], fill=COLORS['tr'])
    d.rectangle([0, h - s, s - 1, h - 1], fill=COLORS['bl'])
    d.rectangle([w - s, h - s, w - 1, h - 1], fill=COLORS['br'])
    d.polygon([(w // 2 - s // 2, s), (w // 2 + s // 2, s), (w // 2, 0)], fill=(0, 0, 0))
    fs = int(min(w, h) * 0.4)
    try:
        font = ImageFont.truetype(FONT_PATH, fs)
    except OSError:
        font = ImageFont.load_default(size=fs)
    d.text((w // 2, h // 2), str(label), fill=(0, 0, 0), font=font, anchor='mm')
    draw_band(d, w, h, label)   # ノイズを混ぜた後の地に描くので、濃さは負けない
    return im


def band_bits(label):
    """12 マス: 両端は黒(1)、間の 10 マスが label の 2 進数（左が上の桁）。"""
    return [1] + [(label >> (9 - i)) & 1 for i in range(10)] + [1]


def draw_band(d, w, h, label):
    d.rectangle([round(0.23 * w), round(0.885 * h), round(0.77 * w), round(0.985 * h)], fill=(255, 255, 255))
    x0, x1 = 0.25 * w, 0.75 * w
    y0, y1 = round(0.90 * h), round(0.97 * h)
    for i, bit in enumerate(band_bits(label)):
        if bit:
            d.rectangle([round(x0 + (x1 - x0) * i / 12), y0, round(x0 + (x1 - x0) * (i + 1) / 12) - 1, y1],
                        fill=(0, 0, 0))


def read_band(im):
    """絵（正しい向き）から帯を読む。(両端が黒か, 読んだ数)。マスの真ん中が 128 より下なら 1。"""
    g = im.convert('L')
    w, h = g.size
    y = round(0.935 * h)
    bits = []
    for i in range(12):
        x = round(0.25 * w + 0.5 * w * (i + 0.5) / 12)
        bits.append(1 if g.getpixel((x, y)) < 128 else 0)
    n = 0
    for b in bits[1:11]:
        n = n * 2 + b
    return bits[0] == 1 and bits[11] == 1, n


def to_stored(up, orient):
    if orient in TO_STORED:
        return up.transpose(TO_STORED[orient])
    return up


def build_exif(mode, orient, dt):
    """mode: full（向き＋撮影日時）／dtonly（DateTime と DateTimeDigitized だけ）／zero（0000 の撮影日時）"""
    ex = Image.Exif()
    ex[0x010F] = 'FakeCam'
    if orient is not None:
        ex[0x0112] = orient
    sub = ex.get_ifd(0x8769)
    if mode == 'full':
        sub[0x9003] = fmt_dt(dt)
    elif mode == 'zero':
        sub[0x9003] = '0000:00:00 00:00:00'
    elif mode == 'dtonly':
        ex[0x0132] = fmt_dt(dt)
        sub[0x9004] = fmt_dt(dt)
    return ex.tobytes()


def jpeg_bytes(stored, exif_bytes):
    buf = io.BytesIO()
    kw = {'quality': JPEG_QUALITY}
    if exif_bytes is not None:
        kw['exif'] = exif_bytes
    stored.save(buf, 'JPEG', **kw)
    return buf.getvalue()


def png_bytes(im):
    buf = io.BytesIO()
    im.save(buf, 'PNG')
    return buf.getvalue()


def copy_bytes(up):
    w, h = up.size
    k = COPY_LONG / max(w, h)
    size = (max(1, round(w * k)), max(1, round(h * k)))
    c = up.resize(size, Image.LANCZOS, reducing_gap=2.0)
    buf = io.BytesIO()
    c.save(buf, 'JPEG', quality=COPY_QUALITY)
    return buf.getvalue(), size


def sniff(b):
    """ファイルの頭のバイトで決まる種類（拡張子は見ない）。"""
    if b[:3] == b'\xff\xd8\xff':
        return 'jpeg'
    if b[:8] == b'\x89PNG\r\n\x1a\n':
        return 'png'
    if b[4:8] == b'ftyp' and b[8:12] in (b'heic', b'heix', b'hevc', b'hevx', b'heim', b'heis', b'mif1', b'msf1'):
        return 'heic'
    return 'other'


# ---------------------------------------------------------------- 組の中身（仕様）

def S(folder, name, typ, **kw):
    d = {'folder': folder, 'name': name, 'type': typ, 'up': None, 'orient': None, 'exif': 'none',
         'dt': None, 'note': '', 'noise': True, 'expect': 'ok'}
    d.update(kw)
    return d


def specs_smoke():
    return [
        S('', 'IMG_0001.jpg', 'jpeg', up=(1600, 1200), orient=1, exif='full', dt=BASE + datetime.timedelta(minutes=30)),
        S('', 'IMG_0002.jpg', 'jpeg', up=(1600, 1200), orient=6, exif='full', dt=BASE + datetime.timedelta(minutes=10)),
        S('', 'IMG_0003.JPG', 'jpeg', up=(1600, 1200), orient=8, exif='full', dt=BASE + datetime.timedelta(minutes=20)),
        S('', '全景 (2).jpg', 'jpeg', up=(1200, 1600), orient=None, exif='full', dt=BASE,
          note='縦で保存・Orientation のタグなし'),
        S('', 'IMG_0005.jpeg', 'jpeg', up=(1600, 1200), exif='none', note='EXIF なし'),
        S('', 'IMG_0006.png', 'png', up=(1600, 1200), noise=False, note='PNG'),
        S('', 'IMG_0007.HEIC', 'heic', expect='unsupported', note='にせの HEIC'),
    ]


FOLDERS = ['', '01_着工前', '02_施工中', '03_完成']
SPECIAL_NAMES = {5: '全景 (2).jpg', 12: '状況 (1).jpg', 21: 'IMG 0021 コピー.JPG', 30: '近景 (3).jpeg'}
EXT_OVERRIDE = {3: '.JPG', 9: '.jpeg', 14: '.JPG', 27: '.JPG', 40: '.jpeg'}


def specs_mix60():
    ok = []
    L = (4000, 3000)
    for _ in range(28):
        ok.append(S('', '', 'jpeg', up=L, orient=1, exif='full', dt=True))
    for o in [6] * 6 + [8] * 4 + [3] * 2:
        ok.append(S('', '', 'jpeg', up=L, orient=o, exif='full', dt=True, note='回した物'))
    for o in (2, 4, 5, 7):
        ok.append(S('', '', 'jpeg', up=L, orient=o, exif='full', dt=True, note='裏返しの物'))
    for _ in range(2):
        ok.append(S('', '', 'jpeg', up=(3000, 4000), orient=None, exif='full', dt=True,
                    note='縦で保存・Orientation のタグなし'))
    for up in [(4000, 2250), (3000, 3000), (6000, 2000)]:
        ok.append(S('', '', 'jpeg', up=up, orient=1, exif='full', dt=True, note='形ちがい'))
    for _ in range(2):
        ok.append(S('', '', 'jpeg', up=L, exif='none', note='EXIF なし'))
    ok.append(S('', '', 'jpeg', up=L, orient=None, exif='dtonly', dt=BASE + datetime.timedelta(days=3),
                note='DateTime と DateTimeDigitized だけ（DateTimeOriginal なし）'))
    ok.append(S('', '', 'jpeg', up=L, orient=1, exif='zero',
                note='DateTimeOriginal に 0000:00:00 00:00:00 を書いた'))
    ok.append(S('', '', 'png', up=(2000, 1500), noise=False, note='PNG'))
    assert len(ok) == 54
    random.Random(SEED).shuffle(ok)

    # 撮影日時（名前の順・作った順と食い違うように振る）。同じ日時の組を 3 つ。
    dated = [s for s in ok if s['dt'] is True]
    n = len(dated)
    for k, s in enumerate(dated):
        s['dt'] = BASE + datetime.timedelta(minutes=((k * 11) % n) * 17)
    for a, b in [(5, 20), (30, 2), (41, 9)]:
        dated[a]['dt'] = dated[b]['dt']

    for i, s in enumerate(ok, 1):
        s['folder'] = FOLDERS[i % 4]
        ext = '.png' if s['type'] == 'png' else EXT_OVERRIDE.get(i, '.jpg')
        s['name'] = SPECIAL_NAMES.get(i, 'IMG_%04d%s' % (1 + (i * 11) % 211, ext))
        if i in SPECIAL_NAMES and s['type'] == 'png':
            s['name'] = 'IMG_%04d.png' % (1 + (i * 11) % 211)

    bad = [
        S('03_完成', 'IMG_9001.HEIC', 'heic', expect='unsupported', note='にせの HEIC'),
        S('02_施工中', 'IMG_9002.jpg', 'trunc', up=(4000, 3000), orient=1, exif='full',
          dt=BASE + datetime.timedelta(minutes=45), expect='broken', note='正しい JPEG の先頭 40% だけ'),
        S('01_着工前', 'IMG_9003.jpg', 'empty', expect='broken', note='0 バイト'),
        S('03_完成', 'IMG_9004.jpg', 'text', expect='broken', note='中身が文字の .jpg'),
        S('', 'memo.txt', 'memo', expect='not-image', note='文字のファイル'),
        S('02_施工中', 'Thumbs.db', 'thumbs', expect='not-image', note='でたらめなバイト列'),
    ]
    out = ok + bad
    seen = set()
    for s in out:
        key = (s['folder'] + '/' + s['name']).lower()
        assert key not in seen, key
        seen.add(key)
    return out


def specs_bulk300():
    out = []
    for i in range(300):
        out.append(S('', 'P%07d.JPG' % (i + 1), 'jpeg', up=(4000, 3000), orient=6 if i % 10 == 0 else 1,
                     exif='full', dt=BASE + datetime.timedelta(minutes=i)))
    return out


def specs_big():
    return [
        S('', 'big_24mp.jpg', 'jpeg', up=(6000, 4000), orient=1, exif='full', dt=BASE),
        S('', 'big_48mp.jpg', 'jpeg', up=(8000, 6000), orient=1, exif='full', dt=BASE + datetime.timedelta(minutes=1)),
        S('', 'big_108mp.jpg', 'jpeg', up=(12000, 9000), orient=1, exif='full', dt=BASE + datetime.timedelta(minutes=2)),
    ]


SPEC_FUNCS = {'smoke': specs_smoke, 'mix60': specs_mix60, 'bulk300': specs_bulk300, 'big': specs_big}


# ---------------------------------------------------------------- 作る

def safe_remove(path, root):
    """root（＝<リポ>/fixtures）の下にある物だけ消す。解決した絶対の道で確かめ、外なら止まる。"""
    real = os.path.realpath(path)
    r = os.path.realpath(root)
    if not (real != r and real.startswith(r + os.sep)):
        sys.exit('止めた: 消す道が %s の下にありません: %s' % (r, real))
    if os.path.islink(path):
        sys.exit('止めた: リンクは消しません: %s' % path)
    if os.path.isdir(path):
        shutil.rmtree(path)
    elif os.path.exists(path):
        os.remove(path)


def make_set(name, out):
    t0 = time.time()
    specs = SPEC_FUNCS[name]()
    set_dir = os.path.join(out, name)
    manifest_path = os.path.join(out, name + '.manifest.json')
    copies_dir = os.path.join(out, name + '.copies')
    for p in (set_dir, manifest_path, copies_dir):
        safe_remove(p, out)
    os.makedirs(set_dir)
    os.makedirs(copies_dir)
    rng = random.Random(SEED + 17)

    files = []
    label = 0
    for s in specs:
        typ = s['type']
        entry = {'path': (s['folder'] + '/' if s['folder'] else '') + s['name'], 'bytes': 0, 'md5': '',
                 'kind': '', 'expect': s['expect'], 'label': None, 'stored': None, 'orientation': None,
                 'upright': None, 'dateTimeOriginal': None, 'note': s['note'], 'copy': None}
        up = None
        if typ in ('jpeg', 'png', 'trunc'):
            label += 1
            w, h = s['up']
            up = make_upright(w, h, label if typ != 'trunc' else 0, s['noise'])
        if typ == 'jpeg':
            ex = None if s['exif'] == 'none' else build_exif(s['exif'], s['orient'], s['dt'])
            data = jpeg_bytes(to_stored(up, s['orient']), ex)
        elif typ == 'trunc':
            full = jpeg_bytes(up, build_exif('full', s['orient'], s['dt']))
            data = full[:int(len(full) * 0.4)]
            label -= 1
        elif typ == 'png':
            data = png_bytes(up)
        elif typ == 'heic':
            data = b'\x00\x00\x00\x18ftypheic' + rng.randbytes(3000)
        elif typ == 'empty':
            data = b''
        elif typ == 'text':
            data = 'これは写真ではありません。文字だけのファイルです。\n'.encode('utf-8')
        elif typ == 'memo':
            data = 'メモ（架空）\n撮影の順番を確かめる\n'.encode('utf-8')
        elif typ == 'thumbs':
            data = b'\xd0\xcf\x11\xe0' + rng.randbytes(2500)
        else:
            raise AssertionError(typ)

        full_path = os.path.join(set_dir, *entry['path'].split('/'))
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        with open(full_path, 'wb') as f:
            f.write(data)
        entry['bytes'] = len(data)
        entry['md5'] = hashlib.md5(data).hexdigest()
        entry['kind'] = sniff(data)
        if typ in ('jpeg', 'trunc') and s['exif'] != 'none':
            entry['orientation'] = s['orient']
            if s['exif'] == 'full':
                entry['dateTimeOriginal'] = fmt_dt(s['dt'])
        if typ == 'jpeg':
            entry['label'] = label
            entry['stored'] = {'w': to_stored(up, s['orient']).size[0], 'h': to_stored(up, s['orient']).size[1]}
            entry['upright'] = {'w': up.size[0], 'h': up.size[1]}
        elif typ == 'png':
            entry['label'] = label
            entry['stored'] = {'w': up.size[0], 'h': up.size[1]}
            entry['upright'] = dict(entry['stored'])
        if entry['label'] is not None:
            cb, (cw, ch) = copy_bytes(up)
            rel = '%s.copies/%03d.jpg' % (name, label)
            with open(os.path.join(out, *rel.split('/')), 'wb') as f:
                f.write(cb)
            entry['copy'] = {'path': rel, 'w': cw, 'h': ch}
        files.append(entry)
        del up

    with open(manifest_path, 'w', encoding='utf-8', newline='\n') as f:
        json.dump({'set': name, 'seed': SEED, 'files': files}, f, ensure_ascii=False, indent=2)
        f.write('\n')
    return files, time.time() - t0


# ---------------------------------------------------------------- 自己点検

def corner_problem(im):
    w, h = im.size
    ins = max(2, int(max(w, h) * 0.04))
    pts = {'tl': (ins, ins), 'tr': (w - 1 - ins, ins), 'bl': (ins, h - 1 - ins), 'br': (w - 1 - ins, h - 1 - ins)}
    for k, xy in pts.items():
        got = im.getpixel(xy)[:3]
        if any(abs(a - b) > 40 for a, b in zip(got, COLORS[k])):
            return '%s が %s（正しくは %s）' % (k, got, COLORS[k])
    return None


def upright_check(path, label=None):
    """開いて exif_transpose で直し、四隅の色を見る。大きい絵は縮めて読む。"""
    try:
        with Image.open(path) as im:
            if im.format == 'JPEG':
                im.draft('RGB', (im.size[0] // 4, im.size[1] // 4))
            fixed = ImageOps.exif_transpose(im).convert('RGB')
        r = corner_problem(fixed)
        if r or label is None:
            return r
        ends, n = read_band(fixed)
        if not ends:
            return '番号の帯の両端が黒でない'
        if n != label:
            return '番号の帯が %d（正しくは %d）' % (n, label)
        return None
    except Exception as e:  # noqa: BLE001  壊れた絵は理由をつけて返す
        return '開けない: %s' % e


def self_check(name, out, files):
    problems = []
    for i, e in enumerate(files):
        if e['expect'] != 'ok':
            continue
        if name == 'bulk300' and i % 10 != 0:
            continue
        for p, lb in ((os.path.join(out, name, *e['path'].split('/')), None),
                      (os.path.join(out, *e['copy']['path'].split('/')), e['label'])):
            r = upright_check(p, lb)
            if r:
                problems.append('%s: %s' % (os.path.relpath(p, out).replace(os.sep, '/'), r))
    return problems


# ---------------------------------------------------------------- 照合

def verify_set(name, out):
    mp = os.path.join(out, name + '.manifest.json')
    if not os.path.exists(mp):
        return ['%s: manifest が無い' % name], 0
    with open(mp, encoding='utf-8') as f:
        m = json.load(f)
    diffs = []
    want = set()
    for e in m['files']:
        want.add(e['path'])
        p = os.path.join(out, name, *e['path'].split('/'))
        if not os.path.isfile(p):
            diffs.append('足りない: %s/%s' % (name, e['path']))
            continue
        with open(p, 'rb') as f:
            data = f.read()
        if len(data) != e['bytes']:
            diffs.append('重さがちがう: %s/%s（%d → %d）' % (name, e['path'], e['bytes'], len(data)))
        elif hashlib.md5(data).hexdigest() != e['md5']:
            diffs.append('md5 がちがう: %s/%s' % (name, e['path']))
        if e['copy'] and not os.path.isfile(os.path.join(out, *e['copy']['path'].split('/'))):
            diffs.append('写しが足りない: %s' % e['copy']['path'])
    root = os.path.join(out, name)
    for dp, _dn, fn in os.walk(root):
        for f in fn:
            rel = os.path.relpath(os.path.join(dp, f), root).replace(os.sep, '/')
            if rel not in want:
                diffs.append('余計: %s/%s' % (name, rel))
    return diffs, len(m['files'])


# ---------------------------------------------------------------- 入り口

def resolve_out(arg):
    out = arg if os.path.isabs(arg) else os.path.join(REPO, arg)
    out = os.path.realpath(out)
    r = os.path.realpath(FIXTURES_ROOT)
    if not (out == r or out.startswith(r + os.sep)):
        sys.exit('止めた: --out は %s の中だけです（%s）' % (r, out))
    return out


def main():
    ap = argparse.ArgumentParser(description='架空の写真を作る')
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument('--set', dest='make', choices=SETS + ['all'])
    g.add_argument('--verify', dest='verify', choices=SETS + ['all'])
    ap.add_argument('--out', default='fixtures')
    a = ap.parse_args()
    out = resolve_out(a.out)
    os.makedirs(out, exist_ok=True)

    if a.verify:
        total, bad = 0, 0
        for n in (SETS if a.verify == 'all' else [a.verify]):
            diffs, cnt = verify_set(n, out)
            for d in diffs:
                print(d)
            total += cnt
            bad += len(diffs)
            print('%s: %d 件・ちがい %d' % (n, cnt, len(diffs)))
        if a.verify == 'all':
            print('合計: %d 件・ちがい %d' % (total, bad))
        sys.exit(1 if bad else 0)

    failed = False
    for n in (SETS if a.make == 'all' else [a.make]):
        files, sec = make_set(n, out)
        problems = self_check(n, out, files)
        mb = sum(e['bytes'] for e in files) / 1048576
        print('%s: %d ファイル・合計 %.1f MB・%.1f 秒・自己点検 %s' % (
            n, len(files), mb, sec, '合格' if not problems else 'ちがい %d' % len(problems)))
        for p in problems:
            print('  ' + p)
            failed = True
    sys.exit(1 if failed else 0)


if __name__ == '__main__':
    main()
