#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GANYANRADAR NUMARA HIZALA
Sonuc CSV'sinin at numarasi ile sitenin (analiz) at numarasi FARKLI olabiliyor.
Bu script kazanan/ilk3/ilk4 (ve atlar 'no') degerlerini ISIMLE eslestirip
sitenin numaralandirmasina cevirir. Boylece tik dogru ata oturur.
EK/GUVENLI: yedek alir (.bak_numara), eslesmeyen kosuya DOKUNMAZ, idempotent ('hizali').
Kullanim: python3 numara_hizala.py            -> canli sonuclar.json + bugunku arsiv
          python3 numara_hizala.py --hepsi    -> tum sonucarsiv/*.json (sinyal motoru icin)
"""
import json, io, os, re, sys, glob, shutil

KOK = os.environ.get("GR_KOK", "/opt/ganyanradar")
ANALIZ = os.path.join(KOK, "analiz")
ARS_HTML = os.path.join(ANALIZ, "arsiv")


def tr_up(s):
    return str(s or "").replace("i", "İ").replace("ı", "I").upper().strip()


def site_norm(ad):
    """site adi: sondaki rakamlari at, buyuk harf. 'MY RAINY HEART1' -> 'MY RAINY HEART'."""
    t = re.sub(r"\d+$", "", str(ad or "").strip())
    t = re.sub(r"\s+", " ", t).strip()
    return tr_up(t)


def csv_core(ad):
    return re.sub(r"\s+", " ", str(ad or "").strip()) and tr_up(re.sub(r"\s+", " ", str(ad).strip()))


def data_cek(yol):
    try:
        h = io.open(yol, encoding="utf-8").read()
    except Exception:
        return {}
    m = re.search(r"DATA\s*=\s*(\{)", h)
    if not m:
        return {}
    i = m.start(1); dep = 0; j = i
    for j in range(i, len(h)):
        if h[j] == "{":
            dep += 1
        elif h[j] == "}":
            dep -= 1
            if dep == 0:
                break
    try:
        return json.loads(h[i:j + 1])
    except Exception:
        return {}


def site_haritasi(iso):
    """(il_up, kosu_no) -> {site_norm_ad: site_no(int)}."""
    yol = os.path.join(ARS_HTML, iso + ".html")
    dd = data_cek(yol)
    s1 = dd.get("Sayfa1", []) if isinstance(dd, dict) else []
    out = {}
    for b in s1:
        hd = b.get("header") or {}
        il = tr_up(hd.get("İl"))
        kno = str(hd.get("Koşu No"))
        if not il or not kno:
            continue
        ad_no = {}
        for row in (b.get("detay") or []):
            if row and str(row[0]).strip().isdigit() and len(row) > 1:
                sn = site_norm(row[1])
                if sn and sn not in ad_no:
                    ad_no[sn] = int(str(row[0]).strip())
        if ad_no:
            out[(il, kno)] = ad_no
    return out


def eslesme(ad_no, csv_ad):
    """csv adi (kodlu) -> site no. En UZUN prefix eslesmesini secer."""
    cn = csv_core(csv_ad)
    if not cn:
        return None
    aday = None
    for snorm, sno in ad_no.items():
        if cn == snorm or (cn.startswith(snorm) and (len(cn) == len(snorm) or cn[len(snorm)] == " ")):
            if aday is None or len(snorm) > aday[0]:
                aday = (len(snorm), sno)
    return aday[1] if aday else None


def hizala_dosya(sonuc_yol, iso):
    if not os.path.exists(sonuc_yol):
        return "yok"
    try:
        veri = json.load(io.open(sonuc_yol, encoding="utf-8"))
    except Exception as e:
        return "okunamadi: %s" % e
    smap = site_haritasi(iso)
    if not smap:
        return "analiz haritasi yok (%s) - DOKUNULMADI" % iso
    degisti = 0
    atlandi = 0
    rapor = []
    for il, ks in (veri.get("iller") or {}).items():
        ilu = tr_up(il)
        for kno, r in (ks or {}).items():
            if r.get("hizali"):
                continue
            ad_no = smap.get((ilu, str(kno)))
            if not ad_no:
                atlandi += 1
                continue
            atlar = r.get("atlar") or []
            csvno_ad = {str(a.get("no")): a.get("ad") for a in atlar}

            def cev(cno):
                ad = csvno_ad.get(str(cno))
                return eslesme(ad_no, ad) if ad else None

            yeni_kaz = cev(r.get("kazanan"))
            if yeni_kaz is None:
                atlandi += 1
                continue
            eski_kaz = r.get("kazanan")
            r["kazanan"] = yeni_kaz
            if isinstance(r.get("ilk3"), list):
                r["ilk3"] = [cev(x) or x for x in r["ilk3"]]
            if isinstance(r.get("ilk4"), list):
                r["ilk4"] = [cev(x) or x for x in r["ilk4"]]
            for a in atlar:
                sno = eslesme(ad_no, a.get("ad"))
                if sno is not None:
                    a["no"] = sno
            r["hizali"] = True
            degisti += 1
            if eski_kaz != yeni_kaz:
                rapor.append("%s K%s: kazanan %s -> %s" % (ilu, kno, eski_kaz, yeni_kaz))
    if degisti:
        shutil.copy(sonuc_yol, sonuc_yol + ".bak_numara")
        json.dump(veri, io.open(sonuc_yol, "w", encoding="utf-8"), ensure_ascii=False)
    return "degisti=%d atlandi=%d | %s" % (degisti, atlandi, "; ".join(rapor[:8]))


def main():
    hepsi = "--hepsi" in sys.argv
    if hepsi:
        for jf in sorted(glob.glob(os.path.join(ANALIZ, "sonucarsiv", "*.json"))):
            iso = os.path.basename(jf)[:-5]
            print("[ARSIV]", iso, "->", hizala_dosya(jf, iso))
    # her zaman canli dosyayi da hizala
    canli = os.path.join(ANALIZ, "sonuclar.json")
    if os.path.exists(canli):
        try:
            t = json.load(io.open(canli, encoding="utf-8")).get("tarih", "")
            iso = "-".join(t.split(".")[::-1]) if t else ""
        except Exception:
            iso = ""
        print("[CANLI]", iso, "->", hizala_dosya(canli, iso))
        # ayni gunun arsivini de
        if iso:
            af = os.path.join(ANALIZ, "sonucarsiv", iso + ".json")
            if os.path.exists(af):
                print("[ARSIV-bugun]", iso, "->", hizala_dosya(af, iso))
    print("NUMARA_HIZALA_BITTI")


if __name__ == "__main__":
    main()
