# tools/gen_20261012.py（承認済みひな形）のコピー。POSTS のデータとラベルの編名だけを変更（2026-11-03、対象週 11/9〜11/13 古い書式・規程編。横長画像で番号と重ならないようラベルは「古い書式編」）。
# 使い方: python3 tools/gen_20261103.py <キャラクター素材フォルダ> <出力フォルダ>
import asyncio, os, sys, html
from playwright.async_api import async_playwright

ASSET_DIR = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
OUT_DIR = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
A = "file://" + ASSET_DIR + "/"

POSTS = [
  dict(key="20261109_furuishoshiki01", no="#01", title="その規程、いつの版ですか？",
       crumbs=["共有ドライブ","総務部","社内規程","人事・労務"],
       files=[
         ("在宅勤務規程_2021.pdf","2021/04/01 09:00","310 KB",False,False),
         ("在宅勤務規程_改定案.docx","2023/11/20 18:42","64 KB",False,False),
         ("在宅勤務規程_改定案_コメント反映.docx","2024/01/15 11:08","66 KB",False,False),
         ("在宅勤務規程_2024年4月改定.pdf","2024/03/29 16:30","325 KB",False,False),
         ("在宅勤務規程.pdf","2025/06/10 10:21","310 KB",True,False),
         ("~$在宅勤務規程_改定案.docx","2026/11/02 13:05","1 KB",False,True),
       ],
       tip="名前に年がない『在宅勤務規程.pdf』、中身は2021年版のコピー…",
       chars=[("taro_side.png","社内ポータルから開いた規程、これが今のですよね？"),
              ("hanako_side.png","そのリンク先、2021年版のままです。今は2024年4月改定版です。")]),
  dict(key="20261111_furuishoshiki02", no="#02", title="「いつものひな形」が3種類",
       crumbs=["PC","デスクトップ","鈴木","よく使う"],
       files=[
         ("秘密保持契約書_ひな形.docx","2022/05/18 14:02","41 KB",False,False),
         ("秘密保持契約書_ひな形(旧住所).docx","2023/02/07 10:36","41 KB",False,False),
         ("秘密保持契約書_ひな形_法務確認済.docx","2024/08/21 17:50","43 KB",True,False),
         ("秘密保持契約書_ひな形_先方版.docx","2025/03/04 09:14","45 KB",False,False),
         ("NDA_英文_ひな形.docx","2022/05/18 14:05","39 KB",False,False),
         ("~$秘密保持契約書_ひな形.docx","2026/11/02 15:40","1 KB",False,True),
       ],
       tip="法務が2025年に改訂した版は、共有ドライブ › 法務 にだけある…",
       chars=[("michio_side.png","先方とのNDA、いつものひな形で出しておいて。"),
              ("taro_side.png","『いつもの』が、手元に3種類あるんですけど……")]),
  dict(key="20261113_furuishoshiki03", no="#03", title="その請求書、移転前の住所です",
       crumbs=["共有ドライブ","経理部","書式","請求書"],
       files=[
         ("請求書_テンプレート.xlsx","2022/04/01 09:30","35 KB",False,False),
         ("請求書_テンプレート_インボイス対応.xlsx","2023/09/25 15:12","37 KB",True,False),
         ("請求書_テンプレート_新住所.xlsx","2025/07/01 10:00","36 KB",False,False),
         ("請求書_テンプレート_新住所_修正.xlsx","2025/07/03 11:47","36 KB",False,False),
         ("請求書の書き方_2023.pdf","2023/09/25 15:20","520 KB",False,False),
         ("~$請求書_テンプレート_インボイス対応.xlsx","2026/11/02 16:18","1 KB",False,True),
       ],
       tip="『インボイス対応』と『新住所』、両方そろった版はどれ…？",
       chars=[("hanako_side.png","この請求書、会社の住所が移転前のままです。"),
              ("taro_side.png","インボイス対応版を使ったのに……新住所版は別だったんですね。")]),
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
<div class="label" style="left:44px;top:40px;font-size:24px;padding:8px 22px">バックオフィスあるある｜古い書式編</div>
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
<div class="label" style="left:28px;top:22px;font-size:17px;padding:5px 16px">バックオフィスあるある｜古い書式編</div>
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
