#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""yeni_yer_SONUC.xlsx icindeki gecmis kosu satirlarina VIDEO LINKI (kolon 43) ekler.

Kaynak: derece cekiminin yazdigi '%45 MP4 URL' / '%45 Video URL' kolonlari
(TUM_ILLER_..._STIL_EKLI.xlsx). Anahtar: (at adi, tarih, sehir).
Link bulunamayan satira DOKUNULMAZ (tarih duz yazi kalir) - hicbir sey bozulmaz.
GUNLUK_CALISTIR.sh icinden tjk_yeni_yer.py'den SONRA calisir.
"""
import os, re, sys, json, io

try:
    import pandas as pd
    from openpyxl import load_workbook
except Exception as e:
    print("UYARI: pandas/openpyxl yok (%s) - video linki adimi atlandi" % e)
    sys.exit(0)

KOK = os.path.dirname(os.path.abspath(__file__))
HEDEF = os.path.join(KOK, "yeni_yer_SONUC.xlsx")
KAYNAKLAR = [
    os.path.join(KOK, "TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_SON6AY_ERTESI_GUN_STIL_EKLI.xlsx"),
    os.path.join(KOK, "SON800_BUGUN_ILK_AT_STIL_FINAL_TEST_STIL_EKLI.xlsx"),
]
LINK_KOLON = 43           # sheet'te 43. kolon -> DATA detay index 42 (video url)
ATNO_KOLON = 44           # sheet'te 44. kolon -> DATA detay index 43 (o kosudaki at no)
TARIH_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")


def tr_up(s):
    return str(s or "").replace("i", "İ").replace("ı", "I").upper().strip()


def at_norm(s):
    t = str(s or "").strip()
    t = re.sub(r"\d+$", "", t)             # BOARDING PASS1 -> BOARDING PASS
    t = re.sub(r"\([^)]*\)", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return tr_up(t)


def tjk_url_mu(u):
    return bool(re.match(r"^https://[a-z0-9.-]*tjk\.org/", str(u or "").strip(), re.I))


def atno_temiz(v):
    s = str(v if v is not None else "").strip()
    if s.lower() in ("", "nan", "none"):
        return ""
    s = re.sub(r"\.0+$", "", s)              # 5.0 -> 5
    return s if re.match(r"^\d{1,2}$", s) else ""



SEHIRLER = ["İstanbul", "Ankara", "İzmir", "Bursa", "Adana", "Kocaeli",
            "Elazığ", "Şanlıurfa", "Diyarbakır", "Antalya", "Kayseri", "Malatya"]
VIDEO_KALIP = "https://www.tjk.org/TR/YarisSever/Info/YarisVideoAt/At?AtKodu={atk}&KosuKodu={kk}"
DUMP = os.path.join(KOK, "kosu_link_dump.jsonl")


def kosu_kodu_bul(hrefler):
    for h in (hrefler or []):
        m = re.search(r"#(\d+)\b", str(h)) or re.search(r"KosuKodu=(\d+)", str(h), re.I)
        if m:
            return m.group(1)
    return ""


def sehir_bul(text):
    t = str(text or "")
    for s in SEHIRLER:
        if s in t:
            return tr_up(s)
    return ""


def atkodu_haritasi():
    """at (norm) -> AtKodu  (dogrulanmis: son-6-ay video url'lerinden)."""
    hmap = {}
    for yol in KAYNAKLAR:
        if not os.path.exists(yol):
            continue
        try:
            df = pd.read_excel(yol)
        except Exception:
            continue
        kolonlar = {str(c): c for c in df.columns}
        vidk = next((kolonlar[c] for c in kolonlar if "VIDEO URL" in c.upper()), None)
        if vidk is None or "At Adı" not in kolonlar:
            continue
        for _, r in df.iterrows():
            m = re.search(r"AtKodu=(\d+)", str(r.get(vidk) or ""))
            if m:
                a = at_norm(r.get("At Adı"))
                if a and a not in hmap:
                    hmap[a] = m.group(1)
    return hmap


def dump_lut(atkodu_map):
    """Eski kosular: dump'tan (at,tarih,sehir) -> (kurulan video url, '')."""
    lut = {}
    if not os.path.exists(DUMP):
        print("NOT: kosu_link_dump.jsonl yok (eski video adimi atlandi)")
        return lut
    say = 0
    for satir in io.open(DUMP, encoding="utf-8"):
        try:
            d = json.loads(satir)
        except Exception:
            continue
        an = at_norm(d.get("at"))
        atk = atkodu_map.get(an)
        if not atk:
            _aid = str(d.get("at_id") or "").strip()
            atk = _aid if _aid.isdigit() else ""
        if not an or not atk:
            continue
        for k in d.get("kayitlar", []):
            tarih = str(k.get("tarih") or "").strip()
            if not TARIH_RE.match(tarih):
                continue
            kk = kosu_kodu_bul(k.get("hrefler"))
            if not kk:
                continue
            sehir = sehir_bul(k.get("text"))
            if not sehir:
                continue
            key = (an, tarih, sehir)
            if key not in lut:
                lut[key] = (VIDEO_KALIP.format(atk=atk, kk=kk), "")
                say += 1
    print("DUMP (eski kosular): %d link kuruldu" % say)
    return lut


def lut_kur():
    """(at, tarih, sehir) -> (video url, o kosudaki at no). MP4 oncelikli."""
    lut = {}
    for yol in KAYNAKLAR:
        if not os.path.exists(yol):
            continue
        try:
            df = pd.read_excel(yol)
        except Exception as e:
            print("UYARI: okunamadi %s (%s)" % (os.path.basename(yol), e))
            continue
        kolonlar = {str(c): c for c in df.columns}
        if "At Adı" not in kolonlar or "Tarih" not in kolonlar or "Şehir" not in kolonlar:
            continue
        mp4k = next((kolonlar[c] for c in kolonlar if "MP4" in c.upper()), None)
        vidk = next((kolonlar[c] for c in kolonlar if "VIDEO URL" in c.upper()), None)
        nok = next((kolonlar[c] for c in kolonlar if "AT NO" in c.upper()), None)
        if mp4k is None and vidk is None:
            continue
        say = 0
        for _, r in df.iterrows():
            url = ""
            if mp4k is not None and tjk_url_mu(r.get(mp4k)):
                url = str(r.get(mp4k)).strip()
            elif vidk is not None and tjk_url_mu(r.get(vidk)):
                url = str(r.get(vidk)).strip()
            if not url:
                continue
            no = atno_temiz(r.get(nok)) if nok is not None else ""
            k = (at_norm(r.get("At Adı")), str(r.get("Tarih") or "").strip(),
                 tr_up(r.get("Şehir")))
            if k[0] and k[1] and k not in lut:
                lut[k] = (url, no)
                say += 1
        print("KAYNAK %s: %d link" % (os.path.basename(yol), say))
    # ESKI KOSULAR: dump'tan kur, sadece eksik anahtarlara ekle (yakinlar oncelikli)
    try:
        dl = dump_lut(atkodu_haritasi())
        for k, v in dl.items():
            if k not in lut:
                lut[k] = v
    except Exception as e:
        print("UYARI: eski video adimi atlandi (%s)" % e)
    return lut


def isle(ws, at_c, tarih_c, sehir_c, lut):
    """Bir sayfadaki detay satirlarina linki yazar. Yazilan satir sayisini dondurur."""
    n = 0
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        try:
            at = row[at_c - 1].value
            tarih = str(row[tarih_c - 1].value or "").strip()
            sehir = row[sehir_c - 1].value
        except IndexError:
            continue
        if not at or not TARIH_RE.match(tarih):
            continue
        k = (at_norm(at), tarih, tr_up(sehir))
        kayit = lut.get(k)
        if not kayit:
            continue
        url, no = kayit
        hucre = ws.cell(row=row[0].row, column=LINK_KOLON)
        if not hucre.value:
            hucre.value = url
            n += 1
        if no:
            h2 = ws.cell(row=row[0].row, column=ATNO_KOLON)
            if not h2.value:
                h2.value = no
    return n


def main():
    if not os.path.exists(HEDEF):
        print("UYARI: yeni_yer_SONUC.xlsx yok - video linki adimi atlandi")
        return
    lut = lut_kur()
    if not lut:
        print("UYARI: hic video linki bulunamadi - adim atlandi")
        return
    wb = load_workbook(HEDEF)
    toplam = 0
    # 'yapilacak yer'  : at=kolon2, tarih=kolon9, sehir=kolon10
    # 'yapilacak yer 800': at=kolon2, tarih=kolon3, sehir=kolon4
    for ad, atc, trc, shc in (("yapılacak yer", 2, 9, 10),
                              ("yapılacak yer 800", 2, 3, 4)):
        if ad in wb.sheetnames:
            n = isle(wb[ad], atc, trc, shc, lut)
            print("%s: %d satira video linki yazildi" % (ad, n))
            toplam += n
    if toplam:
        wb.save(HEDEF)
        print("VIDEO LINKLERI EKLENDI: %d satir" % toplam)
    else:
        print("NOT: eslesen satir yok (dosyalar bos olabilir)")


if __name__ == "__main__":
    main()
