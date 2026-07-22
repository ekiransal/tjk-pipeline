#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""TJK sonuc CSV -> analiz/sonuclar.json (canli) + analiz/sonucarsiv/YYYY-MM-DD.json (KALICI).
Kullanim: python3 sonuc_cek.py             -> bugun: canli json + arsiv
          python3 sonuc_cek.py 10.07.2026  -> SADECE o gunun arsivini yazar (canliya dokunmaz)
Kazanan = GANYAN(n) satiri."""
import json, os, re, sys, datetime, urllib.request, urllib.parse

SEHIRLER = ["\u0130stanbul", "Ankara", "\u0130zmir", "Bursa", "Adana", "Kocaeli",
            "Elaz\u0131\u011f", "\u015eanl\u0131urfa", "Diyarbak\u0131r", "Antalya"]
KOK = os.path.dirname(os.path.abspath(__file__))
HEDEF = os.path.join(KOK, "analiz", "sonuclar.json")
ARSIV = os.path.join(KOK, "analiz", "sonucarsiv")

def bugun():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=3)   # Istanbul gunu

def csv_url(t, sehir):
    dosya = "{0}-{1}-GunlukYarisSonuclari-TR.csv".format(t.strftime("%d.%m.%Y"), sehir)
    return "https://medya-cdn.tjk.org/raporftp/TJKPDF/{0}/{1}/CSV/GunlukYarisSonuclari/{2}".format(
        t.strftime("%Y"), t.strftime("%Y-%m-%d"), urllib.parse.quote(dosya))

def indir(url):
    try:
        r = urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=20)
        return r.read().decode("utf-8", "ignore")
    except Exception:
        return None

def _cell(p, i):
    return p[i].strip() if len(p) > i and p[i].strip() else ""


def derece_ms(s):
    """1:58.53 -> karsilastirilabilir sayi (kucuk=hizli). Kosmaz/bos -> None."""
    m = re.match(r"^\s*(\d+):(\d+)[.,](\d+)", str(s or "").strip())
    if m:
        return int(m.group(1)) * 6000 + int(m.group(2)) * 100 + int(m.group(3))
    return None


def boy_say(s):
    """Boy farkini yaklasik uzunluga cevirir. 'Burun'->0.05, 'Boyun'->0.3,
    'Bas'->0.2, '3,5 Boy'->3.5, 'Farksiz'->0. Bilinmiyorsa None."""
    t = str(s or "").strip().lower()
    if not t:
        return None
    if "farks" in t:
        return 0.0
    if "burun" in t:
        return 0.05
    if "k" in t and "ba" in t and "boy" not in t:   # kisa bas
        return 0.1
    m = re.search(r"([\d]+(?:[.,]\d+)?)\s*boy", t)
    if m:
        try:
            return float(m.group(1).replace(",", "."))
        except Exception:
            pass
    if "yar" in t and "boy" in t:
        return 0.5
    if "boyun" in t:
        return 0.3
    if "ba" in t:                                    # bas (head)
        return 0.2
    return None


def _at_kaydi(p):
    at = {
        "no": int(p[0].strip()),
        "ad": _cell(p, 1),
        "yas": _cell(p, 2),
        "baba": _cell(p, 3),
        "anne": _cell(p, 4),
        "kilo": _cell(p, 5),
        "jokey": _cell(p, 6),
        "sahip": _cell(p, 7),
        "antrenor": _cell(p, 8),
        "st": _cell(p, 9),
        "agf": _cell(p, 10),
        "hp": _cell(p, 11),
        "derece": _cell(p, 12),
        "ganyan": _cell(p, 13),
        "fark": _cell(p, 14),
    }
    at["_ms"] = derece_ms(at["derece"])
    b = boy_say(at["fark"])
    if b is not None:
        at["boy"] = b                                # onundeki ata gore boy farki
    am = re.search(r"%?\s*([\d.,]+)\s*\((\d+)\)", at["agf"])
    if am:
        try:
            at["agf_pct"] = float(am.group(1).replace(",", "."))
        except Exception:
            pass
        at["agf_sira"] = int(am.group(2))
    return at


def _baslik(satir):
    """'1. Kosu :  17.15;Maiden; 3 Yasli...; 58kg; 1800m; Kum' -> dict."""
    p = [x.strip() for x in satir.split(";")]
    h = {}
    m = re.search(r"(\d{1,2}[.:]\d{2})", p[0])
    if m:
        h["saat"] = m.group(1).replace(".", ":")
    if len(p) > 1:
        h["cins"] = p[1]
    if len(p) > 2:
        h["sart"] = p[2]
    for x in p:
        mm = re.match(r"^(\d{3,4})m$", x)
        if mm:
            h["mesafe"] = mm.group(1)
        if x in ("Kum", "Çim", "Sentetik"):
            h["zemin"] = x
    return h


def parse(metin):
    out = {}
    kosu = None
    atlar = []
    basladi = False
    baslik = {}

    def kapat(k, ats, bh):
        kosanlar = [a for a in ats if a.get("_ms") is not None]
        kosmazlar = [a for a in ats if a.get("_ms") is None]
        kosanlar.sort(key=lambda a: a["_ms"])
        sirali = kosanlar + kosmazlar
        kazanan_ms = kosanlar[0]["_ms"] if kosanlar else None
        # RESMI BOY FARKI: CSV 'Fark' = atin BIR ARKASINDAKI ata attigi mesafe.
        # k. atin onundekine mesafesi = (k-1). atin fark'i; kazanana mesafe = kumulatif toplam.
        kumulatif = 0.0
        kirik = False
        for i, a in enumerate(kosanlar):
            a["sira"] = i + 1
            if i == 0:
                a["onundeki_boy"] = 0.0
                a["kazanana_boy"] = 0.0
            else:
                ob = boy_say(kosanlar[i - 1].get("fark"))   # onundeki atin arkaya attigi mesafe
                a["onundeki_boy"] = ob
                if ob is None or kirik:
                    kirik = True
                    a["kazanana_boy"] = None
                else:
                    kumulatif += ob
                    a["kazanana_boy"] = round(kumulatif, 2)
            if a.get("_ms") is not None and kazanan_ms is not None:
                sn = (a["_ms"] - kazanan_ms) / 100.0        # kazanandan kac saniye geride
                a["kazanana_sn"] = round(sn, 2)
                a["kazanana_boy_tahmin"] = round(sn / 0.16, 1)  # zaman-tabanli YEDEK (~0.16 sn/boy)
            a.pop("_ms", None)
        for a in kosmazlar:
            a["sira"] = None
            a.pop("_ms", None)
        if not kosanlar:
            return None
        kayit = {"kazanan": kosanlar[0]["no"]}
        if bh:
            kayit["kosu"] = bh
        gy = kosanlar[0].get("ganyan")
        if gy:
            kayit["ganyan"] = gy
        kayit["ilk3"] = [a["no"] for a in kosanlar[:3]]
        kayit["ilk4"] = [a["no"] for a in kosanlar[:4]]
        kayit["atlar"] = sirali
        return kayit

    for satir in metin.splitlines():
        m = re.match(r"\s*(\d+)\.\s*Kosu", satir)
        if m:
            if kosu is not None and atlar:
                kk = kapat(kosu, atlar, baslik)
                if kk:
                    out[kosu] = kk
            kosu = m.group(1); atlar = []; basladi = False
            baslik = _baslik(satir)
            continue
        if kosu is None:
            continue
        p = satir.split(";")
        if len(p) > 1 and p[0].strip() == "At No":
            basladi = True
            continue
        if basladi and len(p) >= 13 and p[0].strip().isdigit() and p[1].strip():
            atlar.append(_at_kaydi(p))
    if kosu is not None and atlar:
        kk = kapat(kosu, atlar, baslik)
        if kk:
            out[kosu] = kk
    return out


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    simdi = bugun()
    t = datetime.datetime.strptime(arg, "%d.%m.%Y") if arg else simdi
    sonuc = {}
    for s in SEHIRLER:
        m = indir(csv_url(t, s))
        if not m:
            continue
        k = parse(m)
        if k:
            sonuc[s] = k
    veri = {"tarih": t.strftime("%d.%m.%Y"), "guncelleme": simdi.strftime("%H:%M"), "iller": sonuc}
    os.makedirs(ARSIV, exist_ok=True)
    arsiv_yol = os.path.join(ARSIV, t.strftime("%Y-%m-%d") + ".json")
    if sonuc or not os.path.exists(arsiv_yol):     # dolu arsivi bos sonucla EZME
        json.dump(veri, open(arsiv_yol, "w", encoding="utf-8"), ensure_ascii=False)
    if not arg:                                    # tarih verildiyse canli json'a dokunma
        json.dump(veri, open(HEDEF, "w", encoding="utf-8"), ensure_ascii=False)
    print("[SONUC]", ("ARSIV " if arg else "") + veri["tarih"], veri["guncelleme"],
          dict((a, len(b)) for a, b in sonuc.items()))

if __name__ == "__main__":
    main()
