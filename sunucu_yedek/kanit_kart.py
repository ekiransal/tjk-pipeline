#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GANYANRADAR KANIT KARTI v3 — kazananin EN IYI oldugu metrik (TD/Son800/Orijin/Galop)."""
import sys, os, re, json, io, datetime
from PIL import Image, ImageDraw, ImageFont

KOK = os.path.dirname(os.path.abspath(__file__))
ARS_JSON = os.path.join(KOK, "analiz", "sonucarsiv")
ARS_HTML = os.path.join(KOK, "analiz", "arsiv")
CIK = os.path.join(KOK, "kanit")
LOGO = os.path.join(KOK, "logo.png")
FONT_ADAYLARI = [
    "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf",
]
FYOL = next((a for a in FONT_ADAYLARI if os.path.exists(a)), None)
AYLAR = ["", "OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN",
         "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]
ETIKET = {"td": "Aynılaştırılmış toplam derecede", "s8": "Aynılaştırılmış son 800'de",
          "orj": "Orijin analizinde", "gl": "30 günlük aynılaştırılmış galopta"}


def ff(s):
    return ImageFont.truetype(FYOL, s) if FYOL else ImageFont.load_default()


def tw(f, t):
    try:
        b = f.getbbox(t); return b[2]-b[0]
    except Exception:
        try:
            return f.getsize(t)[0]
        except Exception:
            return len(t)*10


def rrect(d, box, rad, fill):
    x0, y0, x1, y1 = [int(v) for v in box]
    if rad*2 > x1-x0: rad = (x1-x0)//2
    if rad*2 > y1-y0: rad = (y1-y0)//2
    d.rectangle([x0+rad, y0, x1-rad, y1], fill=fill)
    d.rectangle([x0, y0+rad, x1, y1-rad], fill=fill)
    d.pieslice([x0, y0, x0+2*rad, y0+2*rad], 180, 270, fill=fill)
    d.pieslice([x1-2*rad, y0, x1, y0+2*rad], 270, 360, fill=fill)
    d.pieslice([x0, y1-2*rad, x0+2*rad, y1], 90, 180, fill=fill)
    d.pieslice([x1-2*rad, y1-2*rad, x1, y1], 0, 90, fill=fill)


def tr_up(s):
    return str(s or "").replace("i", "İ").replace("ı", "I").upper().strip()


def data_cek(yol):
    h = io.open(yol, encoding="utf-8").read()
    m = re.search(r"DATA\s*=\s*(\{)", h)
    if not m: return {}
    i = m.start(1); dep = 0; j = i
    for j in range(i, len(h)):
        if h[j] == "{": dep += 1
        elif h[j] == "}":
            dep -= 1
            if dep == 0: break
    return json.loads(h[i:j+1])


def blok(bloklar, il, k):
    for b in bloklar:
        hd = b.get("header") or {}
        if tr_up(hd.get("İl")) == tr_up(il) and str(hd.get("Koşu No")) == str(k):
            return b
    return None


def detay_gor(b):
    gor = []
    if not b: return gor
    for r in b.get("detay", []):
        if len(r) > 41 and str(r[41]) == "GECMIS_YOK": continue
        n = str(r[0] or "").strip()
        a = re.sub(r"\d+$", "", str(r[1] or "").strip()).strip()
        if n and n not in [g[0] for g in gor]:
            gor.append((n, a))
    return gor


def detay_rank(gor, no):
    for i, (n, a) in enumerate(gor, 1):
        if str(n) == str(no): return i, a
    return None, None


def detay_adlar(b):
    """GECMIS_YOK dahil TUM detay satirlarindan no->ad haritasi (isim kurtarma)."""
    adlar = {}
    if not b: return adlar
    for r in b.get("detay", []):
        try:
            n = str(r[0] or "").strip()
            a = re.sub(r"\d+$", "", str(r[1] or "").strip()).strip()
        except Exception:
            continue
        if n and a and n not in adlar:
            adlar[n] = a
    return adlar


def orijin_rank(b, no):
    if not b: return None
    best = None
    for t in b.get("tables", []):
        rows = t.get("rows", [])
        for i, row in enumerate(rows, 1):
            if row and str(row[0]).strip() == str(no):
                if best is None or i < best: best = i
                break
    return best


def galop_rank(b, no):
    if not b: return None
    en = {}
    for g in b.get("galops", []):
        for row in g.get("rows", []):
            cell = row[0] if isinstance(row, (list, tuple)) else row
            p = [x.strip() for x in str(cell).split("/")]
            if len(p) >= 2:
                hn = p[0]
                try:
                    v = float(p[1].replace(",", "."))
                except Exception:
                    continue
                if hn and (hn not in en or v < en[hn]): en[hn] = v
    if str(no) not in en or len(en) < 3: return None
    for i, (hn, v) in enumerate(sorted(en.items(), key=lambda kv: kv[1]), 1):
        if str(hn) == str(no): return i
    return None


def en_iyi(b1, gor1, b2, gor2, no):
    """kazananin (metrik_key, sira, at_adi) en iyi (en kucuk sira) degeri."""
    td, ad = detay_rank(gor1, no)
    s8, ad2 = detay_rank(gor2, no)
    ad = ad or ad2 or "?"
    aday = [("td", td), ("s8", s8), ("orj", orijin_rank(b1, no)), ("gl", galop_rank(b1, no))]
    aday = [(k, v) for k, v in aday if v]
    if not aday:
        return None, None, ad
    aday.sort(key=lambda kv: kv[1])
    return aday[0][0], aday[0][1], ad


def gy_num(g):
    try:
        return float(str(g).replace(".", "").replace(",", "."))
    except Exception:
        return -1.0


def kart_ciz(il, tarih, gec):
    gg, aa = tarih.split("-")[2], tarih.split("-")[1]
    tarih_txt = "%s %s" % (gg.lstrip("0"), AYLAR[int(aa)])
    W, H = 1080, 1350
    BG = (18, 60, 50); GOLD = (244, 182, 63); CREAM = (207, 232, 223)
    CARD = (26, 74, 62); WHT = (255, 255, 255)
    im = Image.new("RGB", (W, H), BG); d = ImageDraw.Draw(im)

    def ctr(y, t, f, fill):
        d.text(((W-tw(f, t))/2, y), t, font=f, fill=fill)

    if os.path.exists(LOGO):
        lg = Image.open(LOGO).convert("RGBA").resize((84, 84))
        im.paste(lg, (int(W/2-42), 36), lg)
    ctr(132, "GANYANRADAR", ff(40), WHT)

    et = "DÜNÜN SONUÇLARI · %s · %s" % (tr_up(il), tarih_txt)
    f = ff(28); w = tw(f, et)
    rrect(d, ((W-w)/2-24, 196, (W+w)/2+24, 254), 16, GOLD)
    d.text(((W-w)/2, 210), et, font=f, fill=BG)

    ilk3 = [x for x in gec if x["r"] <= 3]
    ctr(298, "%d koşunun %d tanesinde kazanan," % (len(gec), len(ilk3)), ff(42), WHT)
    ctr(354, "radarın İLK 3'ündeydi.", ff(42), GOLD)

    hl = sorted(ilk3, key=lambda x: -gy_num(x["gy"]))[:3]
    y = 448
    for x in hl:
        rrect(d, (70, y, W-70, y+120), 20, CARD)
        d.text((100, y+18), str(x["ad"])[:20], font=ff(38), fill=WHT)
        gt = "ganyan " + str(x["gy"]); f = ff(30); w = tw(f, gt)
        rrect(d, (W-100-w-36, y+16, W-100, y+66), 12, GOLD)
        d.text((W-100-w-18, y+24), gt, font=f, fill=BG)
        ac = "%s %d. sıradaydı → kazandı" % (ETIKET[x["mk"]], x["r"])
        fa = ff(25)
        while tw(fa, ac) > 880 and fa.size > 17:
            fa = ff(fa.size - 1)
        d.text((100, y+74), ac, font=fa, fill=CREAM)
        y += 142

    ctr(902, "Bu şehirde kazananın en iyi analiz sırası:", ff(25), CREAM)
    seri = " · ".join(str(x["r"]) for x in gec)
    f = ff(40)
    while tw(f, seri) > 900 and f.size > 20:
        f = ff(f.size-2)
    w = tw(f, seri)
    rrect(d, ((W-w)/2-30, 946, (W+w)/2+30, 1016), 16, (14, 46, 38))
    d.text(((W-w)/2, 958), seri, font=f, fill=GOLD)
    ctr(1036, "Derece, son 800, orijin ve galop ayrı ölçülür;", ff(24), CREAM)
    ctr(1070, "kazananın en güçlü olduğu sıra alınmıştır — sitede işaretli.", ff(24), CREAM)

    d.line((240, 1132, 840, 1132), fill=(60, 110, 95), width=3)
    ctr(1158, "Tahmin değil, matematiksel kıyaslama.", ff(28), WHT)
    ctr(1203, "Tamamen ücretsiz", ff(25), CREAM)
    u = "ganyanradar.com"; f = ff(42); w = tw(f, u)
    rrect(d, ((W-w)/2-32, 1250, (W+w)/2+32, 1314), 16, GOLD)
    d.text(((W-w)/2, 1260), u, font=f, fill=BG)

    cev = {"İ": "I", "I": "I", "Ş": "S", "Ç": "C", "Ğ": "G", "Ü": "U", "Ö": "O"}
    slug = "".join(cev.get(c, c) for c in tr_up(il))
    slug = re.sub(r"[^A-Za-z0-9]", "", slug) or "SEHIR"
    im.save(os.path.join(CIK, "kanit_%s_%s.png" % (slug, tarih)))
    return slug


def uret(tarih):
    if not os.path.isdir(CIK):
        os.makedirs(CIK)
    jyol = os.path.join(ARS_JSON, tarih + ".json")
    hyol = os.path.join(ARS_HTML, tarih + ".html")
    if not os.path.exists(jyol) or not os.path.exists(hyol):
        print("VERI YOK:", jyol, hyol); return []
    son = json.load(io.open(jyol, encoding="utf-8"))
    dd = data_cek(hyol)
    s1, s2 = dd.get("Sayfa1", []), dd.get("Sayfa2", [])
    ciktilar = []
    for il, kosular in (son.get("iller") or {}).items():
        gec = []
        for kno in sorted(kosular, key=lambda x: int(x)):
            kz = kosular[kno].get("kazanan")
            b1 = blok(s1, il, kno); b2 = blok(s2, il, kno)
            gor1 = detay_gor(b1); gor2 = detay_gor(b2)
            mk, r, ad = en_iyi(b1, gor1, b2, gor2, kz)
            if not ad or ad == "?":
                ad = (detay_adlar(b1).get(str(kz))
                      or detay_adlar(b2).get(str(kz)) or "?")
            if r:
                gec.append(dict(kno=kno, gy=kosular[kno].get("ganyan", "?"),
                                ad=ad, r=r, mk=mk))
        if len(gec) < 4:
            print("ATLANDI (az kosu):", il, len(gec)); continue
        slug = kart_ciz(il, tarih, gec)
        print("KART: %s (%d kosu, %d ilk3) -> kanit_%s_%s.png"
              % (il, len(gec), sum(1 for x in gec if x["r"] <= 3), slug, tarih))
        ciktilar.append(il)
    return ciktilar


if __name__ == "__main__":
    t = sys.argv[1] if len(sys.argv) > 1 else \
        (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    print("KANIT KARTI v3 | tarih:", t, "| font:", FYOL)
    print("URETILEN:", uret(t))
