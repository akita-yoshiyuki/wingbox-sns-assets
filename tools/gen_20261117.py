# tools/gen_20261012.py（承認済みひな形）のコピー。POSTS のデータとラベルの編名だけを変更（2026-11-17、対象週 11/23〜11/27 探しもの編 #04〜#06）。
# 使い方: python3 tools/gen_20261117.py <キャラクター素材フォルダ> <出力フォルダ>
import asyncio, os, sys, html
from playwright.async_api import async_playwright

ASSET_DIR = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
OUT_DIR = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
A = "file://" + ASSET_DIR + "/"

POSTS = [
  dict(key="20261123_sagashimono04", no="#04", title="自動更新の契約、どれだった？",
       crumbs=["共有ドライブ","総務部","契約書","取引先別"],
       files=[
         ("業務委託契約書_ひかり設計.pdf","2026/10/01 09:12","210 KB",False,False),
         ("賃貸借契約書_本社.pdf","2022/03/25 15:40","1.4 MB",False,False),
         ("複合機リース契約.pdf","2023/07/10 11:02","540 KB",True,False),
         ("清掃業務契約_2024.pdf","2024/03/28 16:15","290 KB",False,False),
         ("ソフトウェア利用契約_スキャン.pdf","2021/11/05 10:30","3.2 MB",False,False),
         ("保守契約書_2025.pdf","2025/04/01 09:30","318 KB",False,False),
       ],
       tip="『自動で更新』と書いてあるのはどれ？ ファイル名には書いていない…",
       chars=[("michio_side.png","解約の連絡が要る契約は、どれだったかな？"),
              ("hanako_side.png","条文は契約書の中です。1件ずつ開いて確かめます。")]),
  dict(key="20261125_sagashimono05", no="#05", title="去年の見積、番号しか分からない",
       crumbs=["共有ドライブ","営業部","見積書","2025年度"],
       files=[
         ("Q-2025-0412.xlsx","2025/06/10 14:20","36 KB",False,False),
         ("Q-2025-0413.xlsx","2025/06/11 10:05","41 KB",False,False),
         ("Q-2025-0413_再.xlsx","2025/06/18 17:42","41 KB",True,False),
         ("Q-2025-0414.xlsx","2025/06/12 09:58","33 KB",False,False),
         ("Q-2025-0415.xlsx","2025/06/16 15:31","38 KB",False,False),
         ("見積番号台帳.xlsx","2025/05/30 18:02","22 KB",False,False),
       ],
       tip="みなと商事様に出したのは、どの番号？ 台帳は5月で止まっている…",
       chars=[("taro_side.png","去年みなと商事さんに、いくらで出しましたっけ？"),
              ("hanako_side.png","見積書は番号だけの名前です。台帳も5月から止まっていますね。")]),
  dict(key="20261127_sagashimono06", no="#06", title="去年の休業案内、どこ？",
       crumbs=["共有ドライブ","総務部","お知らせ・掲示"],
       files=[
         ("お知らせ.docx","2023/12/08 10:12","24 KB",False,False),
         ("年末年始休業のお知らせ_取引先.docx","2022/12/05 16:40","27 KB",False,False),
         ("休業案内_2024.docx","2024/12/10 11:25","25 KB",False,False),
         ("年末.docx","2025/12/09 13:48","26 KB",True,False),
         ("夏季休業のお知らせ.docx","2026/08/01 09:30","23 KB",False,False),
         ("掲示用_最終.pdf","2025/12/12 17:05","180 KB",False,False),
       ],
       tip="去年のはどれ？ 毎年作るのに、毎年ファイル名が違う…",
       chars=[("taro_side.png","年末年始のお知らせ、去年のを下敷きにしたいんですが……"),
              ("michio_side.png","毎年作っているはずだが、毎年名前が違うな。")]),
]

CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Noto Sans CJK JP','Noto Sans JP',sans-serif;background:#FEF1E0;color:#32353D}
.canvas{position:relative;overflow:hidden;background:#FEF1E0}
.label{position:absolute;font-weight:700;color:#2960AD;background:#fff;border:3px solid #2960AD;border-radius:999px}
.no{position:absolute;font-weight:900;color:#E78536;font-family:Roboto,'Noto Sans CJK JP',sans-serif}
.title{position:absolute;font-weight:900;color:#32353D;letter-spacing:.02em}
.win{position:absolute;background:#fff;border:1px solid #b9c3cf;border-radius:10px;box-shadow:0 10px 30px rgba(50,53,61,.18);overflow:hidden}
.tabs{background:#e8edf3;display:flex;align-items:flex-end;padding:8px 10px 0;gap:4px}
.tab{background:#fff;border-radius:8px 8px 0 0;padding:6px 14px;font-size:13px;display:flex;gap:8px;align-items:center}
.tab .x{color:#888}
.ctrl{margin-left:auto;display:flex;gap:18px;padding:0 6px 8px;color:#555;font-size:14px}
.bar{display:flex;align-items:center;gap:10px;padding:8px 12px;border-bottom:1px solid #e3e7ec;font-size:13px}
.nav{color:#777;letter-spacing:6px}
.addr{flex:1;border:1px solid #d5dbe2;border-radius:6px;padding:5px 10px;white-space:nowrap;overflow:hidden}
.addr span.sep{color:#999;margin:0 6px}
.search{width:180px;border:1px solid #d5dbe2;border-radius:6px;padding:5px 10px;color:#999}
.cmd{display:flex;gap:16px;padding:7px 14px;border-bottom:1px solid #e3e7ec;font-size:12.5px;color:#444}
table{border-collapse:collapse;width:100%}
th{font-weight:400;font-size:12.5px;color:#555;text-align:left;padding:6px 10px;border-bottom:1px solid #e3e7ec}
td{padding:var(--rp,6px) 10px;white-space:nowrap}
tr.sel td{background:#cce8ff}
tr.tmp td{color:#8a8f96}
.ic{display:inline-block;width:16px;height:19px;border-radius:2px;margin-right:8px;vertical-align:-4px;position:relative}
.ic.xls{background:#3f9b62}.ic.doc{background:#3b6fc4}.ic.pdf{background:#d0473f}.ic.tmp{background:#c7ccd2}
.ic::after{content:"";position:absolute;right:0;top:0;border-left:5px solid rgba(255,255,255,.55);border-bottom:5px solid rgba(255,255,255,.55);}
.status{border-top:1px solid #e3e7ec;font-size:12px;color:#555;padding:5px 12px}
.tip{position:absolute;background:#fffbe6;border:1px solid #c9b86b;padding:5px 9px;box-shadow:0 2px 6px rgba(0,0,0,.15);white-space:nowrap}
.char{position:absolute}
.bub{position:absolute;background:#fff;border:3px solid #32353D;border-radius:22px;font-weight:700;line-height:1.5}
.bub i{position:absolute;width:22px;height:22px;background:#fff;border-right:3px solid #32353D;border-bottom:3px solid #32353D}
.cta{position:absolute;left:0;right:0;bottom:0;background:#2960AD;color:#fff;font-weight:700;text-align:center}
.cta b{color:#FFD340}
"""

def ext_icon(name):
    if name.startswith("~$"): return "tmp"
    if name.endswith(".pdf"): return "pdf"
    if name.endswith(".docx"): return "doc"
    return "xls"

def sel_index(p):
    return [f[3] for f in p["files"]].index(True)

def window(p, left, top, width, rows_font, rp):
    crumbs = '<span class="sep">›</span>'.join(html.escape(c) for c in p["crumbs"])
    rows = ""
    for (n, d, s, sel, tmp) in p["files"]:
        cls = "sel" if sel else ("tmp" if tmp else "")
        rows += f'<tr class="{cls}"><td><span class="ic {ext_icon(n)}"></span>{html.escape(n)}</td><td>{d}</td><td style="text-align:right">{s}</td></tr>'
    return f'''<div class="win" style="left:{left}px;top:{top}px;width:{width}px">
<div class="tabs"><div class="tab">📁 {html.escape(p["crumbs"][-1])} <span class="x">✕</span></div><div class="tab" style="background:transparent">＋</div><div class="ctrl"><span>—</span><span>☐</span><span>✕</span></div></div>
<div class="bar"><span class="nav">← → ↑ ⟳</span><div class="addr">{crumbs}</div><div class="search">検索</div></div>
<div class="cmd"><span>＋ 新規作成</span><span>✂</span><span>⧉</span><span>並べ替え ▾</span><span>表示 ▾</span><span>…</span></div>
<table style="font-size:{rows_font}px;--rp:{rp}px"><tr><th>名前</th><th>更新日時</th><th style="text-align:right">サイズ</th></tr>{rows}</table>
<div class="status">{len(p["files"])} 個の項目　　1 個の項目を選択　{p["files"][sel_index(p)][2]}</div>
</div>'''

def bubble(text, left, top, width, fs, tail_css):
    return f'<div class="bub" style="left:{left}px;top:{top}px;width:{width}px;font-size:{fs}px;padding:{int(fs*0.7)}px {int(fs*0.9)}px">{html.escape(text)}<i style="{tail_css}"></i></div>'

def ig(p):
    c1, c2 = p["chars"]
    return 1080, 1350, f'''
<div class="label" style="left:44px;top:40px;font-size:24px;padding:8px 22px">バックオフィスあるある｜探しもの編</div>
<div class="no" style="right:48px;top:30px;font-size:58px">{p["no"]}</div>
<div class="title" style="left:48px;top:118px;font-size:56px;width:990px">{html.escape(p["title"])}</div>
{window(p, 40, 220, 1000, 19, 11)}
<div class="tip" style="left:330px;top:{220+178+49*sel_index(p)+44}px;font-size:17px;z-index:3">{html.escape(p["tip"])}</div>
<img class="char" src="{A}{c1[0]}" style="left:60px;bottom:90px;height:540px">
<img class="char" src="{A}{c2[0]}" style="right:70px;bottom:90px;height:540px">
{bubble(c1[1], 300, 770, 330, 27, "left:-14px;top:40px;transform:rotate(135deg)")}
{bubble(c2[1], 470, 960, 360, 27, "left:348px;top:50px;transform:rotate(-45deg)")}
<div class="cta" style="height:80px;font-size:30px;line-height:80px">この続きは、<b>プロフィールのリンク</b>から</div>'''

def wide(p):
    c1, c2 = p["chars"]
    return 1200, 630, f'''
<div class="label" style="left:28px;top:22px;font-size:17px;padding:5px 16px">バックオフィスあるある｜探しもの編</div>
<div class="no" style="left:370px;top:12px;font-size:38px">{p["no"]}</div>
<div class="title" style="left:30px;top:70px;font-size:36px;width:760px">{html.escape(p["title"])}</div>
{window(p, 24, 128, 720, 12.5, 6)}
<img class="char" src="{A}{c1[0]}" style="left:770px;bottom:62px;height:400px">
<img class="char" src="{A}{c2[0]}" style="left:990px;bottom:62px;height:400px">
{bubble(c1[1], 760, 14, 210, 17, "left:80px;bottom:-13px;transform:rotate(45deg)")}
{bubble(c2[1], 975, 14, 210, 17, "left:90px;bottom:-13px;transform:rotate(45deg)")}
<div class="cta" style="height:54px;font-size:22px;line-height:54px">この続きは、<b>投稿文のリンク</b>から</div>'''

async def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    async with async_playwright() as pw:
        b = await pw.chromium.launch()
        pg = await b.new_page()
        for p in POSTS:
            for kind, fn in (("ig", ig), ("wide", wide)):
                W, H, body = fn(p)
                doc = f"<!doctype html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body><div class='canvas' style='width:{W}px;height:{H}px'>{body}</div></body></html>"
                path = os.path.join(OUT_DIR, f"{p['key']}_{kind}.html")
                open(path, "w").write(doc)
                await pg.set_viewport_size({"width": W, "height": H})
                await pg.goto("file://" + path)
                await pg.wait_for_timeout(300)
                await pg.screenshot(path=os.path.join(OUT_DIR, f"{p['key']}_{kind}.png"), clip={"x":0,"y":0,"width":W,"height":H})
                os.remove(path)
        await b.close()

asyncio.run(main())
