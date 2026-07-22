#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GANYANRADAR — üyelik kapılı analiz sitesi (Faz 1: bedava ay, üye toplama)
=========================================================================
- Kayıt/Giriş: telefon + tek kullanımlık kod (ŞİFRESİZ). Cihaz 180 gün hatırlanır.
- Kod gönderimi: SMS_MODU="test" iken kod ekranda gösterilir (SMS sağlayıcı
  bağlanana kadar). Netgsm anlaşınca SMS_MODU="netgsm" yapılır, tek yerden.
- Analiz sayfası: analiz/tjk_analiz_prototip.html (pipeline her gün üstüne yazar).
- Admin paneli: /panel (sadece ADMIN_TEL numarası girince görünür).
Çalıştırma:  python3 app.py   ->  http://sunucu:8000
"""
import os, re, json, sqlite3, secrets, random, datetime
from flask import (Flask, request, redirect, session, send_file,
                   render_template_string, abort)

# ------------------------------------------------------------------ AYARLAR
VT = os.path.join(os.path.dirname(__file__), "uyeler.db")
ANALIZ_HTML = os.path.join(os.path.dirname(__file__), "analiz", "tjk_analiz_prototip.html")
ADMIN_TEL = "5512297710"          # Emrah — admin paneli sadece bu numaraya açılır
SMS_MODU = "test"
NETGSM_KULLANICI = ""             # Netgsm abone no (KENDIN DOLDUR)
NETGSM_SIFRE = ""                 # Netgsm API sifresi (KENDIN DOLDUR)
NETGSM_BASLIK = ""                # Onayli baslik, orn. SEVNUR KIRANSAL (KENDIN DOLDUR)                 # "test" = kod ekranda | "netgsm" = gerçek SMS
KOD_OMRU_DK = 5                   # kodun geçerlilik süresi (dakika)
OTURUM_GUN = 1825
SITE_AD = "GanyanRadar"

app = Flask(__name__)
app.secret_key = os.environ.get("GIZLI_ANAHTAR") or "degistir-" + "sabit-gelistirme-anahtari"
app.permanent_session_lifetime = datetime.timedelta(days=OTURUM_GUN)


# ------------------------------------------------------------------ VERİTABANI
def vt():
    c = sqlite3.connect(VT)
    c.row_factory = sqlite3.Row
    return c


def vt_kur():
    with vt() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS uyeler(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ad TEXT, telefon TEXT UNIQUE, kayit TEXT, son_giris TEXT,
            aktif INTEGER DEFAULT 1)""")
        c.execute("""CREATE TABLE IF NOT EXISTS kodlar(
            telefon TEXT PRIMARY KEY, kod TEXT, biter TEXT, deneme INTEGER DEFAULT 0)""")
        try:
            c.execute("ALTER TABLE uyeler ADD COLUMN son_gorulme TEXT")
        except Exception:
            pass


def tel_temizle(t):
    t = re.sub(r"\D", "", str(t or ""))
    if t.startswith("90") and len(t) == 12: t = t[2:]
    if t.startswith("0") and len(t) == 11: t = t[1:]
    return t if re.match(r"^5\d{9}$", t) else ""


# ------------------------------------------------------------------ KOD GÖNDERİMİ
def kod_gonder(tel, kod):
    """SMS_MODU: test -> kod ekranda | netgsm -> gercek SMS"""
    if SMS_MODU == "netgsm":
        import urllib.request, urllib.parse
        mesaj = SITE_AD + " giris kodunuz: " + str(kod)
        prm = urllib.parse.urlencode({"usercode": NETGSM_KULLANICI, "password": NETGSM_SIFRE,
            "gsmno": tel, "message": mesaj, "msgheader": NETGSM_BASLIK, "dil": "TR"})
        try:
            r = urllib.request.urlopen("https://api.netgsm.com.tr/sms/send/get?" + prm, timeout=10)
            cevap = r.read().decode("utf-8", "ignore").strip()
            print("[NETGSM]", tel, cevap)
        except Exception as e:
            print("[NETGSM] HATA:", e)
        return None
    return kod if SMS_MODU == "test" else None


def kod_uret(tel):
    kod = f"{random.SystemRandom().randint(0, 999999):06d}"
    biter = str(int(datetime.datetime.now().timestamp()) + KOD_OMRU_DK*60)
    with vt() as c:
        c.execute("INSERT OR REPLACE INTO kodlar(telefon,kod,biter,deneme) VALUES(?,?,?,0)",
                  (tel, kod, biter))
    return kod_gonder(tel, kod)


def kod_dogru_mu(tel, girilen):
    with vt() as c:
        r = c.execute("SELECT * FROM kodlar WHERE telefon=?", (tel,)).fetchone()
        if not r: return False
        if r["deneme"] >= 5: return False
        if int(datetime.datetime.now().timestamp()) > int(float(r["biter"] or 0)): return False
        if r["kod"] != re.sub(r"\D", "", girilen or ""):
            c.execute("UPDATE kodlar SET deneme=deneme+1 WHERE telefon=?", (tel,))
            return False
        c.execute("DELETE FROM kodlar WHERE telefon=?", (tel,))
        return True


# ------------------------------------------------------------------ ŞABLON
SAYFA = """<!DOCTYPE html><html lang="tr"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>GanyanRadar — Tamamen Matematiksel At Yarışı Analizi</title>
<meta name="description" content="Atları filtreleyin, birbiriyle yarıştırın: orijin analizi, aynılaştırılmış toplam derece, yarıştaki son 800 ve 30 günlük galop. Tamamen ücretsiz, tamamen matematiksel at yarışı analizi.">
<link rel="canonical" href="https://ganyanradar.com/">
<link rel="icon" type="image/png" href="/logo.png">
<link rel="icon" href="/favicon.ico" sizes="any">
<link rel="apple-touch-icon" sizes="180x180" href="/logo.png">
<script type="application/ld+json">{"@context":"https://schema.org","@type":"Organization","name":"GanyanRadar","url":"https://ganyanradar.com/","logo":"https://ganyanradar.com/logo.png","sameAs":["https://x.com/Ganyanradar","https://instagram.com/ganyanradar"]}</script>
<meta property="og:type" content="website">
<meta property="og:site_name" content="GanyanRadar">
<meta property="og:title" content="GanyanRadar — Tamamen Matematiksel At Yarışı Analizi">
<meta property="og:description" content="Orijin analizi, aynılaştırılmış toplam derece, yarıştaki son 800, 30 günlük galop — tamamen ücretsiz.">
<meta property="og:url" content="https://ganyanradar.com/">
<meta property="og:image" content="https://ganyanradar.com/og.png">
<meta name="twitter:card" content="summary_large_image"><style>
 body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#f2f5f4;margin:0;
      display:flex;min-height:100vh;align-items:center;justify-content:center}
 .kut{background:#fff;border-radius:18px;box-shadow:0 6px 30px rgba(20,40,30,.12);
      padding:34px 30px;width:340px;text-align:center}
 h1{font-size:22px;margin:0 0 4px;color:#1e6f5c} .alt{color:#7a8388;font-size:13px;margin-bottom:22px}
 input{width:100%;box-sizing:border-box;padding:13px 14px;font-size:17px;border:1.5px solid #d8dee2;
       border-radius:11px;margin-bottom:12px;text-align:center}
 input:focus{outline:none;border-color:#1e6f5c}
 button{width:100%;padding:13px;font-size:16px;font-weight:700;color:#fff;background:#1e6f5c;
        border:none;border-radius:11px;cursor:pointer} button:active{opacity:.85}
 .hata{color:#c23a3a;font-size:13.5px;margin-bottom:10px}
 .bilgi{background:#e8f3f0;color:#1e6f5c;border-radius:10px;padding:10px;font-size:14px;margin-bottom:12px}
 .kodkutu{font-size:26px;letter-spacing:8px} .kodkart{background:#eafaf3;border:2px solid #1e6f5c;border-radius:16px;padding:18px 14px;margin:16px 0;text-align:center} .kodkart .ust{font-size:15px;color:#1e6f5c;font-weight:bold} .kodkart .kod{font-size:46px;font-weight:bold;letter-spacing:10px;color:#12463a;margin:10px 0} .kodkart .alt2{font-size:13px;color:#3a6b5e}
 .ufak{font-size:12px;color:#9aa4a9;margin-top:16px}
</style></head><body><div class="kut">
 <h1>🏇 {{site}}</h1><div class="alt">Günlük at yarışı analiz platformu</div>
 {% if hata %}<div class="hata">{{hata}}</div>{% endif %}
 {% if test_kod %}<div class="kodkart"><div class="ust">&#128071; Giri&#351; kodun burada</div><div class="kod">{{test_kod}}</div><div class="alt2">SMS beklemene gerek yok &#8212; bu kodu a&#351;a&#287;&#305;ya yaz</div></div>{% endif %}
 {{ govde|safe }}
 <div class="ufak">Ücretsiz Kayıt Ol</div>
</div></body></html>"""

GIRIS_FORM = """<form method="post" action="/kod-iste">
 <input name="ad" placeholder="Adınız Soyadınız" value="{ad}" autocomplete="name">
 <input name="telefon" placeholder="Telefon (5xx xxx xx xx)" value="{tel}"
        inputmode="tel" autocomplete="tel">
 <button>Giriş kodumu göster</button></form>"""

KOD_FORM = """<form method="post" action="/kod-dogrula">
 <div class="alt">{tel} için ekranda görünen 6 haneli kodu gir</div>
 <input name="kod" class="kodkutu" placeholder="______" inputmode="numeric"
        autocomplete="one-time-code" maxlength="6" autofocus>
 <button>Giriş yap</button></form>
 <form method="post" action="/kod-iste" style="margin-top:8px">
 <input type="hidden" name="ad" value="{ad}"><input type="hidden" name="telefon" value="{tel}">
 <button style="background:#fff;color:#1e6f5c;border:1.5px solid #1e6f5c">Yeni kod göster</button></form>"""


# ------------------------------------------------------------------ ROTALAR
@app.route("/google67e102f6f60c1b1f.html")
def google_dogrula():
    return send_file("/opt/ganyanradar/google67e102f6f60c1b1f.html")


@app.route("/robots.txt")
def robots_txt():
    return ("User-agent: *\nAllow: /\nDisallow: /panel\nDisallow: /analiz\n"
            "Sitemap: https://ganyanradar.com/sitemap.xml\n"), 200, {"Content-Type": "text/plain; charset=utf-8"}


@app.route("/sitemap.xml")
def sitemap_xml():
    x = ('<?xml version="1.0" encoding="UTF-8"?>'
         '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'
         '<url><loc>https://ganyanradar.com/</loc><changefreq>daily</changefreq></url>'
         '</urlset>')
    return x, 200, {"Content-Type": "application/xml; charset=utf-8"}


@app.route("/logo.png")
def logo_png():
    y = os.path.join(os.path.dirname(__file__), "logo.png")
    if not os.path.exists(y):
        abort(404)
    return send_file(y, mimetype="image/png")


@app.route("/og.png")
def og_png():
    y = os.path.join(os.path.dirname(__file__), "og.png")
    if not os.path.exists(y):
        abort(404)
    return send_file(y, mimetype="image/png")


@app.route("/")
def ana():
    if session.get("tel"):
        return redirect("/analiz")
    return render_template_string(SAYFA, site=SITE_AD, hata=None, test_kod=None,
                                  govde=GIRIS_FORM.format(ad="", tel=""))


@app.route("/kod-iste", methods=["POST"])
def kod_iste():
    ad = (request.form.get("ad") or "").strip()[:60]
    tel = tel_temizle(request.form.get("telefon"))
    if not tel:
        return render_template_string(SAYFA, site=SITE_AD, test_kod=None,
            hata="Telefon numarası 5 ile başlamalı ve 10 haneli olmalı.",
            govde=GIRIS_FORM.format(ad=ad, tel=request.form.get("telefon") or ""))
    session["bekleyen_tel"], session["bekleyen_ad"] = tel, ad
    test_kod = kod_uret(tel)
    return render_template_string(SAYFA, site=SITE_AD, hata=None, test_kod=test_kod,
                                  govde=KOD_FORM.format(tel="0" + tel, ad=ad))


@app.route("/kod-dogrula", methods=["POST"])
def kod_dogrula():
    tel = session.get("bekleyen_tel", "")
    ad = session.get("bekleyen_ad", "")
    if not tel:
        return redirect("/")
    if not kod_dogru_mu(tel, request.form.get("kod")):
        return render_template_string(SAYFA, site=SITE_AD, test_kod=None,
            hata="Kod yanlış ya da süresi doldu.", govde=KOD_FORM.format(tel="0"+tel, ad=ad))
    simdi = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    with vt() as c:
        r = c.execute("SELECT id, aktif FROM uyeler WHERE telefon=?", (tel,)).fetchone()
        if r is None:
            c.execute("INSERT INTO uyeler(ad,telefon,kayit,son_giris) VALUES(?,?,?,?)",
                      (ad or "İsimsiz", tel, simdi, simdi))
        else:
            if not r["aktif"]:
                return render_template_string(SAYFA, site=SITE_AD, test_kod=None,
                    hata="Üyeliğiniz pasif durumda. İletişime geçin.",
                    govde=GIRIS_FORM.format(ad=ad, tel=tel))
            c.execute("UPDATE uyeler SET son_giris=?, ad=CASE WHEN ad='İsimsiz' AND ?<>'' THEN ? ELSE ad END WHERE telefon=?",
                      (simdi, ad, ad, tel))
    session.clear()
    session.permanent = True
    session["tel"] = tel
    return redirect("/analiz")


@app.route("/analiz")
def analiz():
    tel = session.get("tel")
    if not tel:
        return redirect("/")
    with vt() as c:
        r = c.execute("SELECT aktif FROM uyeler WHERE telefon=?", (tel,)).fetchone()
        if r and r["aktif"]:
            c.execute("UPDATE uyeler SET son_gorulme=? WHERE telefon=?",
                      (datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"), tel))
    if not r or not r["aktif"]:
        session.clear()
        return redirect("/")
    if not os.path.exists(ANALIZ_HTML):
        return "Analiz sayfası henüz yüklenmedi.", 503
    _t = request.args.get("t", "")
    if re.match(r"^\d{4}-\d{2}-\d{2}$", _t):
        _ars = os.path.join(os.path.dirname(__file__), "analiz", "arsiv", _t + ".html")
        if os.path.exists(_ars):
            return send_file(_ars)
    # VARSAYILAN: o anki gun (yeni gune gece 01:00'de gecer)
    _bugun = (datetime.datetime.now() - datetime.timedelta(hours=1)).strftime("%Y-%m-%d")
    _ars = os.path.join(os.path.dirname(__file__), "analiz", "arsiv", _bugun + ".html")
    if os.path.exists(_ars):
        return send_file(_ars)
    return send_file(ANALIZ_HTML)


@app.route("/sonuclar.json")
def sonuclar_json():
    if not session.get("tel"):
        abort(404)
    yol = os.path.join(os.path.dirname(__file__), "analiz", "sonuclar.json")
    if not os.path.exists(yol):
        return "{}", 200, {"Content-Type": "application/json"}
    r = send_file(yol)
    r.headers["Cache-Control"] = "no-store"
    return r


@app.route("/analiz-tarihler")
def analiz_tarihler():
    if not session.get("tel"):
        abort(404)
    kl = os.path.join(os.path.dirname(__file__), "analiz", "arsiv")
    ts = []
    if os.path.isdir(kl):
        ts = sorted(x[:-5] for x in os.listdir(kl) if x.endswith(".html"))
    # OK SINIRI: en fazla yarin gosterilir; obur gun ihtiyat olarak gizli kalir
    _sinir = (datetime.datetime.now() + datetime.timedelta(hours=23)).strftime("%Y-%m-%d")
    ts = [x for x in ts if x <= _sinir]
    import json as _j
    return _j.dumps(ts), 200, {"Content-Type": "application/json", "Cache-Control": "no-store"}


@app.route("/sonuclar-arsiv/<t>.json")
def sonuclar_arsiv(t):
    if not session.get("tel"):
        abort(404)
    if len(t) != 10 or t[4] != "-" or t[7] != "-" or not (t[:4] + t[5:7] + t[8:]).isdigit():
        abort(404)
    yol = os.path.join(os.path.dirname(__file__), "analiz", "sonucarsiv", t + ".json")
    if not os.path.exists(yol):
        return "{}", 200, {"Content-Type": "application/json"}
    r = send_file(yol)
    r.headers["Cache-Control"] = "no-store"
    return r


@app.route("/tanitim.mp4")
def tanitim_video():
    yol = os.path.join(os.path.dirname(__file__), "analiz", "tanitim.mp4")
    if not os.path.exists(yol):
        abort(404)
    return send_file(yol, conditional=True)


@app.route("/cikis")
def cikis():
    session.clear()
    return redirect("/")


@app.route("/kanit")
def kanit_liste():
    if session.get("tel") != ADMIN_TEL:
        abort(404)
    kl = os.path.join(os.path.dirname(__file__), "kanit")
    try:
        ds = sorted((x for x in os.listdir(kl) if x.endswith(".png")), reverse=True)
    except Exception:
        ds = []
    parca = []
    for x in ds:
        parca.append('<div style="margin:22px auto;max-width:540px">'
                     '<img src="/kanit/%s" style="width:100%%;border-radius:14px;display:block">'
                     '<a href="/kanit/%s" download style="display:block;text-align:center;'
                     'margin-top:10px;background:#f4b63f;color:#123c32;padding:14px;'
                     'border-radius:12px;text-decoration:none;font:bold 17px sans-serif">Indir</a>'
                     '</div>' % (x, x))
    g = "".join(parca) or "<p style='text-align:center'>Bugun kart yok.</p>"
    return ("<html><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width,initial-scale=1'>"
            "<title>Kanit Kartlari</title><style>body{background:#123c32;color:#fff;"
            "font-family:sans-serif;padding:16px;margin:0}h2{text-align:center}</style>"
            "</head><body><h2>Gunun Kanit Kartlari</h2>" + g + "</body></html>")


@app.route("/kanit/<ad>")
def kanit_dosya(ad):
    if session.get("tel") != ADMIN_TEL:
        abort(404)
    if (not ad.endswith(".png")) or "/" in ad or ".." in ad:
        abort(404)
    y = os.path.join(os.path.dirname(__file__), "kanit", ad)
    if not os.path.exists(y):
        abort(404)
    return send_file(y, mimetype="image/png")


@app.route("/panel")
def panel():
    if session.get("tel") != ADMIN_TEL:
        abort(404)
    with vt() as c:
        uyeler = c.execute("SELECT * FROM uyeler ORDER BY kayit DESC").fetchall()
    sat = "".join(
        f"<tr><td>{u['id']}</td><td>{u['ad']}</td><td>0{u['telefon']}</td>"
        f"<td>{u['kayit'][:16].replace('T',' ')}</td><td>{(u['son_giris'] or '')[:16].replace('T',' ')}</td>"
        f"<td>{(u['son_gorulme'] or '')[:16].replace('T',' ')}</td>"
        f"<td><a href='/panel/degistir/{u['id']}'>{'✅ aktif' if u['aktif'] else '⛔ pasif'}</a></td></tr>"
        for u in uyeler)
    return f"""<html><head><meta charset=utf8><title>Panel</title><style>
      body{{font-family:sans-serif;padding:24px}} table{{border-collapse:collapse;width:100%}}
      td,th{{border:1px solid #ccc;padding:6px 10px;font-size:14px}} th{{background:#eef}}
      </style></head><body><h2>👥 Üyeler ({len(uyeler)})</h2>
      <table><tr><th>#</th><th>Ad</th><th>Telefon</th><th>Kayıt</th><th>Son Giriş</th><th>Son Görülme</th><th>Durum</th></tr>
      {sat}</table></body></html>"""


@app.route("/panel/degistir/<int:uid>")
def panel_degistir(uid):
    if session.get("tel") != ADMIN_TEL:
        abort(404)
    with vt() as c:
        c.execute("UPDATE uyeler SET aktif=1-aktif WHERE id=?", (uid,))
    return redirect("/panel")


vt_kur()

@app.route("/favicon.ico")
def favicon_ico():
    from flask import send_file
    return send_file(os.path.join(os.path.dirname(__file__), "favicon.ico"),
                     mimetype="image/x-icon")

if __name__ == "__main__":
    os.makedirs(os.path.join(os.path.dirname(__file__), "analiz"), exist_ok=True)
    app.run(host="0.0.0.0", port=8000)
