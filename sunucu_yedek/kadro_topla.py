#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GANYANRADAR KADRO TOPLAYICI - her kosu icin TEK birlesik anlik goruntu.
Sonuc arsivi (tum kadro: hp, kilo, ganyan, agf, bitis sirasi, kazanana boy, zemin, mesafe)
+ analiz arsivi (orijin rank, galop, varsa stil) BIRLESTIRILIR.
Cikti: analiz/kadro_arsiv/YYYY-MM-DD.json  (KALICI, EK - mevcut hicbir seyi bozmaz)
Ilk calismada analiz yapisi hakkinda kadro_tani.json yazar (stil/sicil nerede diye bakariz).
At ADI ana anahtar: haftadan haftaya ayni atlari baglar -> sicil/gecis SONRADAN buradan hesaplanir.
"""
import os, re, json, io, glob, datetime
from collections import defaultdict

KOK = os.environ.get("GR_KOK", "/opt/ganyanradar")
ARS_HTML = os.path.join(KOK, "analiz", "arsiv")
ARS_JSON = os.path.join(KOK, "analiz", "sonucarsiv")
CIKTI = os.path.join(KOK, "analiz", "kadro_arsiv")
TANI = os.path.join(KOK, "analiz", "kadro_tani.json")

DOGU = {"ELAZIĞ", "ŞANLIURFA", "URFA", "DIYARBAKIR", "DİYARBAKIR"}
STIL_KELIME = ["KAÇAK", "ERKEN", "TAKİP", "GERİDEN", "ÖNDE", "LİDER", "SONDAN", "ORTADAN"]


def tr_up(s):
    return str(s or "").replace("i", "İ").replace("ı", "I").upper().strip()


def at_norm(s):
    t = str(s or "").strip()
    t = re.sub(r"\(.*?\)", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return tr_up(t)


_TRMAP = str.maketrans("İIıiŞşÇçĞğÖöÜü", "IIIISSCCGGOOUU")


def at_key(s):
    """Turkce harfleri ASCII'ye indir + parantez + sondaki rakamlari at + BUYUK."""
    t = re.sub(r"\(.*?\)", " ", str(s or ""))
    t = re.sub(r"\d+$", "", t.strip())
    t = re.sub(r"\s+", " ", t).strip()
    return t.translate(_TRMAP).upper()


def derece_yukle():
    """Aynilastirilmis toplam derece + son 800 (varsa). derece_bugun.json ->
    {(SEHIR,AD):rec} + {AD:rec}. Isimler ASCII-fold + kod-siz TEMIZ ad."""
    yol = os.path.join(KOK, "analiz", "derece_bugun.json")
    smap, amap = {}, {}
    try:
        d = json.load(io.open(yol, encoding="utf-8"))
    except Exception:
        return smap, amap
    for r in d.get("kayitlar", []):
        adk = at_key(r.get("at"))
        if not adk:
            continue
        rec = {"td": r.get("toplam_derece"), "td_sira": r.get("toplam_derece_sira"),
               "s800": r.get("son800"), "s800_sira": r.get("son800_sira")}
        sh = at_key(r.get("sehir"))
        if sh:
            smap[(sh, adk)] = rec
        amap.setdefault(adk, rec)
    return smap, amap


def derece_bul(smap, amap, il, ad):
    """Arsiv adi 'GERCEK AD + malzeme kodlari' (MY RAINY HEART KG DB SK), derece adi
    TEMIZ. En UZUN prefix derece adini bul (kodlar derece adi degildir)."""
    ilu = at_key(il)
    toks = at_key(ad).split()
    for n in range(len(toks), 0, -1):
        key = " ".join(toks[:n])
        r = smap.get((ilu, key)) or amap.get(key)
        if r:
            return r
    return None


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


def blok(bloklar, il, k):
    for b in bloklar:
        hd = b.get("header") or {}
        if tr_up(hd.get("İl")) == tr_up(il) and str(hd.get("Koşu No")) == str(k):
            return b
    return None


def orijin_ranklar(b1):
    """at no -> (ort rank, ort ORNEK SAYISI). Ornek = tablo 3. sutundaki '154/1201' -> 1201.
    Kucuk ornek (2/4) gurultu, buyuk ornek (225/2109) saglam - sinyalde agirliklandirilir."""
    if not b1:
        return {}, {}
    siralar = defaultdict(list)
    ornekler = defaultdict(list)
    for t in b1.get("tables", []):
        for idx, row in enumerate(t.get("rows", []), 1):
            no = str(row[0]).strip() if row else ""
            if no.isdigit():
                siralar[no].append(idx)
                if len(row) > 2:
                    m = re.search(r"/\s*(\d+)", str(row[2]))
                    if m:
                        ornekler[no].append(int(m.group(1)))
    rank = {}
    ornek = {}
    for no, s in siralar.items():
        if len(s) >= 2:
            rank[no] = round(sum(s) / len(s), 2)
            if ornekler.get(no):
                ornek[no] = int(sum(ornekler[no]) / len(ornekler[no]))
    return rank, ornek


def galop_map(b1):
    """at NO -> {tipler:{tip:deger}, iyi:tum tiplerin en iyisi, g800:800+400 degeri}.
    Gercek format: satir[0] = 'atno / deger / zemin'  (deger kucuk/negatif = hizli)."""
    if not b1:
        return {}
    by = defaultdict(dict)
    for g in b1.get("galops", []):
        tip = g.get("name") if isinstance(g, dict) else None
        rows = g.get("rows", []) if isinstance(g, dict) else (g or [])
        for row in rows:
            cell = row[0] if isinstance(row, (list, tuple)) else row
            p = [x.strip() for x in str(cell).split("/")]
            if len(p) >= 2 and p[0].isdigit():
                no = p[0]
                try:
                    v = float(p[1].replace(",", "."))
                except Exception:
                    continue
                if tip not in by[no] or v < by[no][tip]:
                    by[no][tip] = v
    out = {}
    for no, d in by.items():
        out[no] = {"tipler": d, "iyi": (min(d.values()) if d else None),
                   "g800": d.get("800+400")}
    return out


def stil_bul(b1):
    """at no -> stil (varsa). Tablo hucrelerinde stil kelimesi arar (temkinli)."""
    out = {}
    if not b1:
        return out
    for t in b1.get("tables", []):
        for row in t.get("rows", []):
            if not row:
                continue
            no = str(row[0]).strip()
            if not no.isdigit():
                continue
            for cell in row[1:]:
                cu = tr_up(cell)
                if len(cu) <= 22:
                    for kel in STIL_KELIME:
                        if kel in cu:
                            out.setdefault(no, cu)
                            break
    return out


def sn_cevir(t):
    """'1:09,47' -> 69.47 saniye."""
    m = re.match(r"(\d+):(\d+)[.,](\d+)", str(t or "").strip())
    if m:
        return int(m.group(1)) * 60 + int(m.group(2)) + int(m.group(3)) / 100.0
    return None


def stil_iz_coz(iz):
    """'▷▶▷▷' izinde dolu ucgen(▶) konumu: erken->KAÇAK, orta->TAKİP, gec->GERİDEN."""
    s = str(iz or "")
    n = len(s)
    dolu = [k for k, ch in enumerate(s) if ch == "▶"]
    if not dolu or n < 2:
        return None
    pos = (sum(dolu) / len(dolu)) / (n - 1)      # 0..1
    if pos <= 0.34:
        return "KAÇAK"
    if pos <= 0.67:
        return "TAKİP"
    return "GERİDEN"


def detay_coz(dt):
    """detay(list) -> at_no -> gecmis yaris listesi [{zemin,mesafe,sn,hiz,stil,...}]."""
    by = defaultdict(list)
    for row in (dt or []):
        if not row or len(row) < 15:
            continue
        no = str(row[0]).strip()
        if not no.isdigit():
            continue
        try:
            mes = int(re.sub(r"[^0-9]", "", str(row[13])))
        except Exception:
            mes = None
        sn = sn_cevir(row[14])
        if not mes or not sn or mes <= 0 or sn <= 0:
            continue
        iz = row[34] if len(row) > 34 else ""
        by[no].append({
            "tarih": row[8] if len(row) > 8 else "",
            "sehir": row[9] if len(row) > 9 else "",
            "zemin": tr_up(row[10]) if len(row) > 10 else "",
            "mesafe": mes, "sn": round(sn, 2), "hiz": round(mes / sn, 3),
            "stil": stil_iz_coz(iz),
        })
    return by


def yas_no(s):
    """'4y a  a' -> 4 (atin yasi)."""
    m = re.search(r"(\d+)", str(s or ""))
    return int(m.group(1)) if m else None


def cinsiyet(yas):
    """'4y d  a' -> ('DİŞİ','d'). Kod: d/k=disi, a/e/i/t=erkek. (ham kod da donuyor - DOGRULANACAK)"""
    toks = str(yas or "").split()
    kod = None
    for t in toks[1:]:
        t = t.strip().lower()
        if t and t[0].isalpha():
            kod = t[0]
            break
    if kod in ("d", "k"):
        return "DİŞİ", kod
    if kod in ("a", "e", "i", "t"):
        return "ERKEK", kod
    return None, kod


def yas_sart_coz(sart):
    """'3 ve Yukarı İngilizler' -> '3+', '2 Yaşlı' -> '2yo'. Kosu yas sinifi."""
    t = tr_up(sart)
    m = re.search(r"(\d+)", t)
    if not m:
        return "?"
    n = m.group(1)
    if "YUKAR" in t:
        return n + "+"
    if "YAŞ" in t or "YASL" in t:
        return n + "yo"
    return n


def gun_fark(tarih_str, bugun_iso):
    """'12.06.2026' ve '2026-07-14' -> aradaki gun (pozitif=gecmis). Cozemezse None."""
    try:
        g = datetime.datetime.strptime(str(tarih_str).strip(), "%d.%m.%Y")
        b = datetime.datetime.strptime(str(bugun_iso).strip(), "%Y-%m-%d")
        return (b - g).days
    except Exception:
        return None


def sicil_ozet(sic, bugun_zemin, bugun_mesafe, bugun_tarih=None):
    """Bir atin gecmisinden: toplam derece hizi + zemin bazli (AYRI) + bugune uygun hiz + stil
    + SON 2 AY formu (yarisa gore son 60 gun) + trend (son2ay vs genel).
    Sentetik = kopru (kum/cim ikisi de gider). Kum<->cim gecisi VARSAYILMAZ, ayri saklanir."""
    if not sic:
        return {}
    hepsi = [x["hiz"] for x in sic]
    o = {"sicil_n": len(sic),
         "toplam_derece_hiz": round(sum(hepsi) / len(hepsi), 3),
         "en_iyi_hiz": max(hepsi)}
    # SON 2 AY formu (yarisa gore son 60 gun icindeki kosular)
    if bugun_tarih:
        yakin = [x["hiz"] for x in sic
                 if (gun_fark(x.get("tarih"), bugun_tarih) is not None
                     and 0 <= gun_fark(x.get("tarih"), bugun_tarih) <= 60)]
        o["son2ay_n"] = len(yakin)
        o["son2ay_hiz"] = round(sum(yakin) / len(yakin), 3) if yakin else None
        # trend: son2ay hizi genel ortalamanin uzerinde mi (yukselen form)
        if yakin and o["toplam_derece_hiz"]:
            o["form_trend"] = round(o["son2ay_hiz"] - o["toplam_derece_hiz"], 3)
    # her zemini AYRI sakla (transfer'i olcmek icin)
    for et, ad in (("KUM", "kum"), ("ÇİM", "cim"), ("SENTETİK", "sentetik")):
        zz = [x["hiz"] for x in sic if x["zemin"] == et]
        o[ad + "_n"] = len(zz)
        o[ad + "_hiz"] = round(sum(zz) / len(zz), 3) if zz else None
    # bugune UYGUN hiz: sentetikse kopru (hepsi), kum/cimde sentetik yardimci; karsi zemin haric
    zt = tr_up(bugun_zemin)
    if zt == "SENTETİK":
        uyg = hepsi
    elif zt == "KUM":
        uyg = [x["hiz"] for x in sic if x["zemin"] in ("KUM", "SENTETİK")]
    elif zt in ("ÇİM", "ÇIM"):
        uyg = [x["hiz"] for x in sic if x["zemin"] in ("ÇİM", "ÇIM", "SENTETİK")]
    else:
        uyg = hepsi
    o["uygun_n"] = len(uyg)
    o["uygun_hiz"] = round(sum(uyg) / len(uyg), 3) if uyg else None
    try:
        bm = int(re.sub(r"[^0-9]", "", str(bugun_mesafe or "")))
    except Exception:
        bm = 0
    if bm:
        mh = [x["hiz"] for x in sic if abs(x["mesafe"] - bm) <= 300]
        o["mesafe_n"] = len(mh)
        o["mesafe_hiz"] = round(sum(mh) / len(mh), 3) if mh else None
    stiller = [x["stil"] for x in sic if x["stil"]]
    o["stil"] = max(set(stiller), key=stiller.count) if stiller else None
    return o


def tempo(stil_map, atlar):
    if not stil_map:
        return None, None
    kacak = 0
    for a in atlar:
        s = tr_up(stil_map.get(str(a["no"]), ""))
        if "KAÇAK" in s or "ERKEN" in s or "ÖNDE" in s:
            kacak += 1
    if kacak >= 3:
        return kacak, "hizli"
    if kacak <= 1:
        return kacak, "yavas"
    return kacak, "normal"


def main():
    os.makedirs(CIKTI, exist_ok=True)
    json_g = sorted(os.path.basename(x)[:-5] for x in glob.glob(os.path.join(ARS_JSON, "*.json")))
    html_g = set(os.path.basename(x)[:-5] for x in glob.glob(os.path.join(ARS_HTML, "*.html")))

    top_gun = top_kosu = top_at = 0
    e_orijin = e_galop = e_stil = e_sicil = 0
    derece_smap, derece_amap = derece_yukle()
    e_derece = 0
    tani = {"analiz_gun": len(html_g), "sonuc_gun": len(json_g), "ornek": None}

    for g in json_g:
        try:
            son = json.load(io.open(os.path.join(ARS_JSON, g + ".json"), encoding="utf-8"))
        except Exception:
            continue
        dd = data_cek(os.path.join(ARS_HTML, g + ".html")) if g in html_g else {}
        s1 = dd.get("Sayfa1", []) if isinstance(dd, dict) else []
        gun_kosular = []
        for il, ks in (son.get("iller") or {}).items():
            ilu = tr_up(il)
            for kno, r in (ks or {}).items():
                atlar = r.get("atlar") or []
                if not atlar:
                    continue
                b1 = blok(s1, il, kno)
                orij, orij_ornek = orijin_ranklar(b1)
                gal = galop_map(b1)
                stil_map = stil_bul(b1)
                det = detay_coz(b1.get("detay") if b1 else [])
                hd = (b1.get("header") if b1 else {}) or {}
                kosu = r.get("kosu") or {}
                zemin_bugun = tr_up(hd.get("Zemin") or kosu.get("zemin"))
                mesafe_bugun = hd.get("Mesafe") or kosu.get("mesafe")
                kadro = []
                stil_bugun = {}
                for a in atlar:
                    no = str(a.get("no"))
                    ad = at_norm(a.get("ad"))
                    gm = gal.get(no) or {}
                    so = sicil_ozet(det.get(no, []), zemin_bugun, mesafe_bugun, bugun_tarih=g)
                    stil = so.get("stil") or stil_map.get(no)
                    stil_bugun[no] = stil
                    rec = {
                        "no": a.get("no"), "ad": a.get("ad"), "ad_norm": ad,
                        "sira": a.get("sira"),
                        "hp": a.get("hp"), "kilo": a.get("kilo"),
                        "agf_pct": a.get("agf_pct"), "agf_sira": a.get("agf_sira"),
                        "ganyan": a.get("ganyan"),
                        # RESMI boy farki (CSV Fark'tan) oncelikli, yoksa zaman-tahmini
                        "kazanana_boy": (a.get("kazanana_boy") if a.get("kazanana_boy") is not None
                                         else a.get("kazanana_boy_tahmin")),
                        "onundeki_boy": a.get("onundeki_boy"),
                        "fark": a.get("fark"),
                        "jokey": a.get("jokey"), "baba": a.get("baba"), "anne": a.get("anne"),
                        "antrenor": a.get("antrenor"), "yas": a.get("yas"),
                        "yas_no": yas_no(a.get("yas")),
                        "cinsiyet": cinsiyet(a.get("yas"))[0],
                        "cins_kod": cinsiyet(a.get("yas"))[1],
                        "orijin_rank": orij.get(no),
                        "orijin_ornek": orij_ornek.get(no),
                        # AYNILASTIRILMIS TOPLAM DERECE + SON 800 (varsa; yoksa None -> kaba proxy)
                        "resmi_toplam_derece": None, "resmi_toplam_derece_sira": None,
                        "resmi_son800": None, "resmi_son800_sira": None,
                        "galop_iyi": gm.get("iyi"),
                        "galop_800": gm.get("g800"),
                        "galop_tipler": gm.get("tipler"),
                        "stil": stil,
                        # --- SICIL (gecmis derece / zemin / mesafe uygunlugu) ---
                        "toplam_derece_hiz": so.get("toplam_derece_hiz"),
                        "uygun_hiz": so.get("uygun_hiz"),
                        "en_iyi_hiz": so.get("en_iyi_hiz"),
                        "kum_hiz": so.get("kum_hiz"), "kum_n": so.get("kum_n"),
                        "cim_hiz": so.get("cim_hiz"), "cim_n": so.get("cim_n"),
                        "sentetik_hiz": so.get("sentetik_hiz"), "sentetik_n": so.get("sentetik_n"),
                        "mesafe_hiz": so.get("mesafe_hiz"), "mesafe_n": so.get("mesafe_n"),
                        "sicil_n": so.get("sicil_n"),
                        # SON 2 AY formu + trend
                        "son2ay_hiz": so.get("son2ay_hiz"), "son2ay_n": so.get("son2ay_n"),
                        "form_trend": so.get("form_trend"),
                    }
                    if rec["orijin_rank"] is not None:
                        e_orijin += 1
                    if rec["galop_iyi"] is not None:
                        e_galop += 1
                    if rec["stil"] is not None:
                        e_stil += 1
                    if rec["toplam_derece_hiz"] is not None:
                        e_sicil += 1
                    _dr = derece_bul(derece_smap, derece_amap, ilu, a.get("ad"))
                    if _dr:
                        rec["resmi_toplam_derece"] = _dr.get("td")
                        rec["resmi_toplam_derece_sira"] = _dr.get("td_sira")
                        rec["resmi_son800"] = _dr.get("s800")
                        rec["resmi_son800_sira"] = _dr.get("s800_sira")
                        e_derece += 1
                    kadro.append(rec)
                    top_at += 1
                # tempo: artik gercek per-at stil (sicilden) ile kacak sayilir
                kacak = sum(1 for no, st in stil_bugun.items()
                            if st and ("KAÇAK" in st or "ERKEN" in st))
                if any(stil_bugun.values()):
                    tmp = "hizli" if kacak >= 3 else ("yavas" if kacak <= 1 else "normal")
                    ks_say = kacak
                else:
                    ks_say, tmp = tempo(stil_map, atlar)
                gun_kosular.append({
                    "tarih": g, "il": ilu, "dogu": ilu in DOGU,
                    "kosu_no": kno,
                    "zemin": zemin_bugun,
                    "mesafe": mesafe_bugun,
                    "irk": hd.get("Irk"),
                    "yorum": (hd.get("Yorum") or "").strip() or None,
                    "cins": kosu.get("cins"), "sart": kosu.get("sart"),
                    "yas_sart": yas_sart_coz(kosu.get("sart")),
                    "kazanan": r.get("kazanan"), "ilk3": r.get("ilk3"), "ilk4": r.get("ilk4"),
                    "kacak_sayisi": ks_say, "tempo": tmp,
                    "atlar": kadro,
                })
                top_kosu += 1
        if gun_kosular:
            io.open(os.path.join(CIKTI, g + ".json"), "w", encoding="utf-8").write(
                json.dumps({"tarih": g, "kosular": gun_kosular}, ensure_ascii=False))
            top_gun += 1
            if tani["ornek"] is None and s1:
                b = s1[0]
                tani["ornek"] = {
                    "gun": g,
                    "blok_anahtarlari": sorted(list(b.keys())),
                    "tablo_sayisi": len(b.get("tables", [])),
                    "galop_grup": len(b.get("galops", [])),
                    "header_alanlari": sorted(list((b.get("header") or {}).keys())),
                }

    tani["ozet"] = {"gun": top_gun, "kosu": top_kosu, "at": top_at,
                    "orijinli_at": e_orijin, "galoplu_at": e_galop, "stilli_at": e_stil,
                    "sicilli_at": e_sicil}
    io.open(TANI, "w", encoding="utf-8").write(json.dumps(tani, ensure_ascii=False, indent=1))

    print("KADRO TOPLANDI ->", CIKTI)
    print("  gun:", top_gun, "| kosu:", top_kosu, "| at:", top_at)
    print("  zenginlestirme (at basina): orijin", e_orijin, "| galop", e_galop,
          "| stil", e_stil, "| sicil(derece)", e_sicil)
    print("  AYNILASTIRILMIS derece eslesen at:", e_derece, "(resmi toplam derece + son 800)")
    print("  TANI dosyasi:", TANI)
    if tani.get("ornek"):
        print("  analiz blok anahtarlari:", tani["ornek"]["blok_anahtarlari"])
        print("  header alanlari       :", tani["ornek"]["header_alanlari"])
        print("  tablo/galop grubu     :", tani["ornek"]["tablo_sayisi"], "/", tani["ornek"]["galop_grup"])
    if e_stil == 0:
        print("  NOT: stil arsivde bulunamadi -> stil'i ileri-capture ile (Mac scraper) eklememiz gerekebilir.")
    print("KADRO_TOPLAYICI_BITTI")


if __name__ == "__main__":
    main()

