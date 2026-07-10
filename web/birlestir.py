#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SONUC ESLEYICI (Katman 2)
web/ klasorunden calistirilir:  python3 birlestir.py
- ../veri/ozellikler_*.csv dosyalarini tarar
- Her gun+il icin TJK resmi sonuc CSV'sini indirir (once indirilmisse tekrar indirmez)
- Atlari ISIMLE esler; kazandi / ilk3 / kacinci / final AGF / ganyan ekler
- Hepsini ../veri/gecmis_veri.csv'ye yazar ve kisa istatistik basar
"""
import csv, glob, os, re, sys, urllib.request, urllib.parse

VERI = os.path.join("..", "veri")
os.makedirs(VERI, exist_ok=True)

# at adindaki donanim kisaltmalari (sonuc CSV'sinde ada yapisik gelir)
EKLER = {"K","KG","DB","SK","SKG","G","GK","HC","HÇ","B","D","S","KGDB","DBSK"}

def temiz_ad(s):
    s = str(s or "").strip().upper()
    s = re.sub(r"\s+", " ", s)
    p = s.split(" ")
    while len(p) > 1 and p[-1] in EKLER:
        p.pop()
    return " ".join(p)

def indir(url, hedef):
    if os.path.exists(hedef) and os.path.getsize(hedef) > 200:
        return True
    try:
        r = urllib.request.urlopen(urllib.request.Request(
            url, headers={"User-Agent": "Mozilla/5.0"}), timeout=25)
        veri = r.read()
        if len(veri) < 200:
            return False
        open(hedef, "wb").write(veri)
        return True
    except Exception:
        return False

def sonuc_url(tarih, il):   # tarih GG.AA.YYYY
    g, a, y = tarih.split(".")
    dosya = "%s-%s-GunlukYarisSonuclari-TR.csv" % (tarih, il)
    return ("https://medya-cdn.tjk.org/raporftp/TJKPDF/%s/%s-%s-%s/CSV/"
            "GunlukYarisSonuclari/%s" % (y, y, a, g, urllib.parse.quote(dosya)))

def sonuc_parse(metin):
    """{kosu_no: {"atlar": {temiz_ad: {...}}, "kazanan_no": int|None, "ilk3_no": [..]}}"""
    out = {}
    kosu = None; sira = 0
    for satir in metin.splitlines():
        m = re.match(r"\s*(\d+)\.\s*Kosu", satir)
        if m:
            kosu = m.group(1); sira = 0
            out[kosu] = {"atlar": {}, "kazanan_no": None, "ilk3_no": []}
            continue
        if kosu is None:
            continue
        g = re.search(r"GANYAN\((\d+)\)", satir)
        if g:
            out[kosu]["kazanan_no"] = int(g.group(1))
            u = re.search(r"BAH[İI]S\((\d+)/(\d+)/(\d+)\)", satir)
            if u:
                out[kosu]["ilk3_no"] = [int(u.group(i)) for i in (1, 2, 3)]
            kosu = None
            continue
        p = satir.split(";")
        # at satiri: ilk alan bitis sirasi (sayi), 2. alan at ismi, >=13 alan
        if len(p) >= 13 and p[0].strip().isdigit() and p[1].strip():
            sira += 1
            agf = None; agf_sira = None
            ma = re.search(r"%([\d.,]+)\((\d+)\)", p[10])
            if ma:
                agf = float(ma.group(1).replace(",", "."))
                agf_sira = int(ma.group(2))
            out[kosu]["atlar"][temiz_ad(p[1])] = {
                "kacinci": sira,
                "derece": p[12].strip(),
                "agf": agf, "agf_sira": agf_sira,
                "ganyan": p[13].strip() if len(p) > 13 else "",
            }
    return out

# ---- ozellik dosyalarini oku ----
oz_dosyalar = sorted(glob.glob(os.path.join(VERI, "ozellikler_*.csv")))
if not oz_dosyalar:
    sys.exit("veri/ klasorunde ozellikler_*.csv yok — once ozellik_kaydet.py calistir.")

tum = []
for yol in oz_dosyalar:
    with open(yol, encoding="utf-8") as fh:
        tum.extend(list(csv.DictReader(fh)))
print("[BIRLESTIR] ozellik satiri: %d (%d gun)" % (len(tum), len(oz_dosyalar)))

# ---- gerekli sonuc dosyalarini indir + parse ----
sonuclar = {}   # (tarih, il) -> parse ciktisi
for (tarih, il) in sorted(set((r["tarih"], r["il"]) for r in tum)):
    iso = "-".join(reversed(tarih.split(".")))
    hedef = os.path.join(VERI, "sonuc_%s_%s.csv" % (iso, il))
    if indir(sonuc_url(tarih, il), hedef):
        metin = open(hedef, encoding="utf-8", errors="ignore").read()
        sonuclar[(tarih, il)] = sonuc_parse(metin)
        n = sum(len(v["atlar"]) for v in sonuclar[(tarih, il)].values())
        print("  %s %-10s sonuc: %d kosu, %d at" % (tarih, il,
              len(sonuclar[(tarih, il)]), n))
    else:
        print("  %s %-10s sonuc YOK (henuz kosulmadi ya da indirilemedi)" % (tarih, il))

# ---- esle ve yaz ----
EKLE = ["kacinci", "kazandi", "ilk3", "agf_final", "agf_sira", "derece_sonuc", "ganyan", "durum"]
cikti = os.path.join(VERI, "gecmis_veri.csv")
esk = 0; kos = 0
with open(cikti, "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    basliklar = list(tum[0].keys()) + EKLE
    w.writerow(basliklar)
    for r in tum:
        sn = sonuclar.get((r["tarih"], r["il"]), {}).get(str(int(r["kosu"])) if r["kosu"].isdigit() else r["kosu"])
        ek = ["", "", "", "", "", "", "", "sonuc_bekliyor"]
        if sn:
            at = sn["atlar"].get(temiz_ad(r["at"]))
            if at:
                # bitis sirasi (isim eslesmesinden) ASIL gercek: 1. geldiyse kazandi
                kazandi = 1 if at["kacinci"] == 1 else 0
                ilk3 = 1 if at["kacinci"] <= 3 else 0
                ek = [at["kacinci"], kazandi, ilk3, at["agf"], at["agf_sira"],
                      at["derece"], at["ganyan"], "tamam"]
                esk += 1
            else:
                ek = ["", 0, 0, "", "", "", "", "kosmadi_veya_eslesmedi"]
                kos += 1
        w.writerow([r[k] for k in tum[0].keys()] + ek)
print("[BIRLESTIR] eslesen: %d | eslesmeyen/kosmayan: %d -> %s" % (esk, kos, cikti))

# ---- kisa istatistik (yeterince veri birikince anlamlanir) ----
try:
    rows = [x for x in csv.DictReader(open(cikti, encoding="utf-8")) if x["durum"] == "tamam"]
    n = len(rows)
    if n >= 50:
        fav = [x for x in rows if x["agf_sira"] == "1"]
        fav_k = sum(1 for x in fav if x["kazandi"] == "1")
        print("[STAT] n=%d | AGF favorisi kazanma: %d/%d (%%%.0f)" % (
            n, fav_k, len(fav), 100.0 * fav_k / max(1, len(fav))))
        cb = [x for x in rows if "çok belirgin" in str(x.get("kosu_final_kademe", "")) ]
    else:
        print("[STAT] su an %d eslesmis satir var; anlamli istatistik icin biriktirmeye devam." % n)
except Exception as e:
    print("[STAT] atlandi:", e)
