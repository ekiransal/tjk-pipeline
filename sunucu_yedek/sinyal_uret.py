#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GANYANRADAR SINYAL URETICI - kadro_arsiv anlik goruntulerinden sinyal cikarir.
"Illa birinciye bakma": her secim icin kazanma% + tabela(ilk3/ilk4)% + ROI + ort. kazanana boy.
DOGU illeri (Elazig/Urfa/Diyarbakir) AYRI kova. CINS (kosu sinifi) ve HP-DOMINANS (rakiplerle
kiyas) da degerlendirilir. Kazananin HP-ilk3'te cikma orani ayrica raporlanir.
Guven esikleri: <40 yetersiz | 40-99 izlemede | 100-249 orta | 250+ yuksek.
Cikti: analiz/sinyal_kutuphane_v2.json (+ tarihli anlik goruntu). EK - eski dosyaya dokunmaz.
"""
import os, re, json, io, glob
from collections import defaultdict

KOK = os.environ.get("GR_KOK", "/opt/ganyanradar")
KADRO = os.path.join(KOK, "analiz", "kadro_arsiv")
KUTUP = os.path.join(KOK, "analiz", "sinyal_kutuphane_v2.json")
GECMIS = os.path.join(KOK, "analiz", "sinyal_gecmis_v2")


def tr_up(s):
    return str(s or "").replace("i", "İ").replace("ı", "I").upper().strip()


def gy_num(g):
    try:
        return float(str(g).replace(".", "").replace(",", "."))
    except Exception:
        return None


def sayi(x):
    try:
        return float(str(x).replace(",", "."))
    except Exception:
        return None


def guven(n):
    if n < 40:
        return "yetersiz"
    if n < 100:
        return "izlemede"
    if n < 250:
        return "orta"
    return "yuksek"


def cins_norm(c):
    t = tr_up(c)
    if not t:
        return "?"
    if "MAIDEN" in t or "MAİDEN" in t:
        return "MAIDEN"
    if "HANDIKAP" in t or "HANDİKAP" in t or t.startswith("HK") or "HANDIK" in t:
        return "HANDIKAP"
    if "KV" in t:
        return "KV(kazanmamis)"
    if "ŞARTLI" in t or "SARTLI" in t or t.startswith("Ş"):
        return "SARTLI"
    if "SATIS" in t or "SATIŞ" in t:
        return "SATIS"
    if "DERECE" in t or t.startswith("D"):
        return "DERECELI"
    return t[:14]


def mesafe_kova(m):
    try:
        mm = int(re.sub(r"[^0-9]", "", str(m or "")))
    except Exception:
        return "?"
    if mm == 0:
        return "?"
    if mm <= 1300:
        return "sprint<=1300"
    if mm <= 1700:
        return "orta1400-1700"
    return "uzun>=1800"


class P:
    __slots__ = ("n", "kaz", "ilk3", "ilk4", "stake", "donen", "boy_top", "boy_n", "puan")

    def __init__(s):
        s.n = s.kaz = s.ilk3 = s.ilk4 = 0
        s.stake = s.donen = s.boy_top = 0.0
        s.boy_n = 0
        s.puan = 0.0


def ekle(p, secno, kosu, atlar_by_no):
    if secno is None:
        return
    a = atlar_by_no.get(str(secno))
    if not a:
        return
    p.n += 1
    p.stake += 1
    sira = a.get("sira")
    ilk3 = kosu.get("ilk3") or []
    ilk4 = kosu.get("ilk4") or []
    kazandi = (sira == 1)
    if kazandi:
        p.kaz += 1
        g = gy_num(a.get("ganyan"))
        if g:
            p.donen += g
    if a.get("no") in ilk3:
        p.ilk3 += 1
    if a.get("no") in ilk4:
        p.ilk4 += 1
    boy = a.get("kazanana_boy")
    if not kazandi and boy is not None:
        p.boy_top += boy
        p.boy_n += 1
    if kazandi:
        p.puan += 1.0
    elif a.get("no") in ilk3:
        p.puan += 0.6
    elif a.get("no") in ilk4:
        p.puan += 0.4
    elif boy is not None:
        p.puan += max(0.0, 0.3 - 0.02 * boy)


def ozet(p):
    if not p.n:
        return {"n": 0, "kazanma": None, "ilk3": None, "ilk4": None,
                "roi": None, "ort_boy": None, "puan": None, "guven": "yetersiz"}
    return {
        "n": p.n,
        "kazanma": round(100.0 * p.kaz / p.n, 1),
        "ilk3": round(100.0 * p.ilk3 / p.n, 1),
        "ilk4": round(100.0 * p.ilk4 / p.n, 1),
        "roi": round(100.0 * (p.donen - p.stake) / p.stake, 1) if p.stake else None,
        "ort_boy": round(p.boy_top / p.boy_n, 2) if p.boy_n else None,
        "puan": round(p.puan / p.n, 3),
        "guven": guven(p.n),
    }


def hp_ozellik(kosu):
    """HP'yi rakiplerle kiyasla: lider, 2.'ye acik ara (gap), kazanan HP-ilk3'te mi."""
    atlar = kosu.get("atlar") or []
    hps = [(sayi(a.get("hp")), a["no"]) for a in atlar if sayi(a.get("hp")) is not None]
    if len(hps) < 2:
        return None
    hps.sort(reverse=True)               # yuksek HP once
    hp_max = hps[0][0]
    hp_2 = hps[1][0]
    top3 = set(n for _, n in hps[:3])
    return {"lider": hps[0][1], "gap": round(hp_max - hp_2, 2),
            "hp_top3": top3, "kazanan_hp_top3": kosu.get("kazanan") in top3}


def secimler(kosu):
    atlar = kosu.get("atlar") or []
    s = {}
    gec = [(gy_num(a.get("ganyan")), a["no"]) for a in atlar if gy_num(a.get("ganyan"))]
    s["favori"] = min(gec)[1] if gec else None
    s["agf"] = next((a["no"] for a in atlar if a.get("agf_sira") == 1), None)
    oj = [(a.get("orijin_rank"), a["no"]) for a in atlar if a.get("orijin_rank") is not None]
    s["orijin"] = min(oj)[1] if oj else None
    # orijin_saglam: sadece yeterli ORNEK sayisi olan orijinler (kucuk-ornek gurultuyu dis birak)
    ojs = [(a.get("orijin_rank"), a["no"]) for a in atlar
           if a.get("orijin_rank") is not None and (a.get("orijin_ornek") or 0) >= 50]
    s["orijin_saglam"] = min(ojs)[1] if ojs else None
    # form: son 2 ayin en iyi (yuksek) hizi - son form lideri
    fm = [(a.get("son2ay_hiz"), a["no"]) for a in atlar if a.get("son2ay_hiz") is not None]
    s["form"] = max(fm)[1] if fm else None
    gl = []
    for a in atlar:
        v = a.get("galop_800")
        if v is None:
            v = a.get("galop_iyi")
        if v is not None:
            gl.append((v, a["no"]))
    s["galop"] = min(gl)[1] if gl else None
    hp = [(sayi(a.get("hp")), a["no"]) for a in atlar if sayi(a.get("hp")) is not None]
    s["hp_lider"] = max(hp)[1] if hp else None
    # genc: sahadaki en genc at (gelisen/kilo avantajli olabilir)
    yn = [(a.get("yas_no"), a["no"]) for a in atlar if a.get("yas_no")]
    s["genc"] = min(yn)[1] if yn else None
    # toplam derece: bugune UYGUN hiz (sentetik kopru, karsi zemin haric) - yuksek=hizli
    td = []
    for a in atlar:
        v = a.get("uygun_hiz")
        if v is None:
            v = a.get("toplam_derece_hiz")
        if v is not None:
            td.append((v, a["no"]))
    s["toplam_derece"] = max(td)[1] if td else None
    # AYNILASTIRILMIS resmi toplam derece: sira KUCUK = daha iyi (1. sirada)
    rd = [(a.get("resmi_toplam_derece_sira"), a["no"]) for a in atlar
          if a.get("resmi_toplam_derece_sira") is not None]
    s["resmi_derece"] = min(rd)[1] if rd else None
    # AYNILASTIRILMIS son 800: sira KUCUK = daha iyi
    s8 = [(a.get("resmi_son800_sira"), a["no"]) for a in atlar
          if a.get("resmi_son800_sira") is not None]
    s["son800"] = min(s8)[1] if s8 else None
    return s


STRATEJILER = ["favori", "agf", "orijin", "orijin_saglam", "form",
               "galop", "hp_lider", "toplam_derece", "resmi_derece", "son800", "genc"]


def yeni_kova():
    return {st: P() for st in STRATEJILER}


def kova_ozet(kv):
    return {st: ozet(kv[st]) for st in STRATEJILER}


def main():
    gunler = sorted(glob.glob(os.path.join(KADRO, "*.json")))

    # --- 1. gecis: tum kosulari yukle + HP gap dagilimini ol ---
    tum = []
    gaplar = []
    son_gun = None
    for yol in gunler:
        try:
            d = json.load(io.open(yol, encoding="utf-8"))
        except Exception:
            continue
        son_gun = d.get("tarih") or son_gun
        for kosu in d.get("kosular", []):
            atlar = kosu.get("atlar") or []
            if not atlar:
                continue
            hpz = hp_ozellik(kosu)
            tum.append((kosu, {str(a["no"]): a for a in atlar}, secimler(kosu), hpz))
            if hpz:
                gaplar.append(hpz["gap"])

    gaplar.sort()
    if gaplar:
        t1 = gaplar[len(gaplar) // 3]
        t2 = gaplar[2 * len(gaplar) // 3]
    else:
        t1 = t2 = 0

    def gap_kova_ad(g):
        if g <= t1:
            return "dar(<=%.1f)" % t1
        if g <= t2:
            return "orta(%.1f-%.1f)" % (t1, t2)
        return "genis(>%.1f)" % t2

    # --- 2. gecis: kovalari doldur ---
    genel = yeni_kova(); dogu = yeni_kova(); bati = yeni_kova()
    zemin = defaultdict(yeni_kova); yorum_k = defaultdict(yeni_kova)
    pist_k = defaultdict(yeni_kova)   # PIST baglami (yeni harita, pist_yon.json)
    def _tr_up_pist(x):
        return str(x or "").replace("i", "İ").replace("ı", "I").upper().strip()
    try:
        _pmap_ham = json.load(io.open("/opt/ganyanradar/analiz/pist_yon.json", encoding="utf-8"))
        _pmap = {}
        for _kk, _vv in _pmap_ham.items():
            _pp = _kk.split("|")
            if len(_pp) == 4:
                _pmap["%s|%s|%s|%s" % (_tr_up_pist(_pp[0]), _tr_up_pist(_pp[1]),
                                       _tr_up_pist(_pp[2]), _pp[3])] = _vv
    except Exception:
        _pmap = {}

    def _pist_tipi(kosu):
        il = _tr_up_pist(kosu.get("il") or kosu.get("sehir"))
        z = _tr_up_pist(kosu.get("zemin"))
        ir = _tr_up_pist(kosu.get("irk"))
        try:
            mes = int(float(kosu.get("mesafe")))
        except Exception:
            return "?"
        kod = _pmap.get("%s|%s|%s|%d" % (il, z, ir, mes))
        return {"S": "finis", "K": "cokus", "N": "notr"}.get(kod, "?")
    irk_k = defaultdict(yeni_kova); mesafe_k = defaultdict(yeni_kova)
    cins_k = defaultdict(yeni_kova); hpgap_k = defaultdict(yeni_kova)
    yas_k = defaultdict(yeni_kova)         # kosu yas sinifi (3+, 4+, 2yo...) -> kova
    hp_kapsar = [0, 0]                      # [kazanan HP-ilk3'te, toplam]
    hp_kapsar_cins = defaultdict(lambda: [0, 0])
    komb = {k: P() for k in (
        "favori_destek0", "favori_destek1", "favori_destek2", "favori_destek3",
        "temiz_favori", "kotu_baglam_favori", "iyi_baglam_favori", "tam_mutabakat")}
    # KACAK-EZISMESI SENARYOSU ("oyle olsa boyle olur"): tempoya gore KAZANANIN STILI
    # + hizli tempoda son-form liderine (geriden gelebilir) oynama performansi
    kacak_senaryo = defaultdict(lambda: defaultdict(int))   # tempo -> kazanan_stili -> sayi
    kacak_yorum = defaultdict(lambda: defaultdict(int))     # TJK Yorum etiketi -> kazanan_stili -> sayi
    kacak_pist = defaultdict(lambda: defaultdict(int))   # pist|tempo -> kazanan stili
    saha_pist = defaultdict(lambda: defaultdict(int))    # pist|tempo -> kosan stiller (taban)
    takip_cokus_k = P()   # YENI SINYAL: cokus+hizli -> TAKIP/GERIDEN formlusu (eski sinyalle ayni tip)
    revans_k = P()      # AT-BASINA: bugunku rakiplerinden en cok yenmis at
    yeniat_k = P()      # AT-BASINA: arsivde ilk kez gorulen atlar
    tekrarci_k = P()    # AT-BASINA: son 15 gunde kosup ilk3 yapmis at
    at_gecmis = {}      # ad_norm -> {yaris_id: sira}
    at_songun = {}      # ad_norm -> (tarih, ilk3_mu)
    _gorulen_at = set()

    def _trh(kosu):
        import datetime as _dt
        t = str(kosu.get("tarih") or "")
        for f in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
            try:
                return _dt.datetime.strptime(t[:10], f).date()
            except Exception:
                pass
        return None
    form_tempo = defaultdict(P)                             # tempo -> 'form' stratejisi
    # KACAK SAVASI -> TAKIPCI sinyali: Yorum'da KACAK gecen kosuda en formlu TAKIP-stilli at
    takip_kacakta = P()
    # CINSIYET x MESAFE: disiler uzunda geliyor mu (taban oranli) + ham kod dogrulama
    kazanan_cins_mesafe = defaultdict(lambda: defaultdict(int))
    kosan_cins_mesafe = defaultdict(lambda: defaultdict(int))
    cins_kod_say = defaultdict(int)
    toplam_kosu = 0

    for kosu, abn, sec, hpz in tum:
        toplam_kosu += 1
        cn = cins_norm(kosu.get("cins"))
        hedefler = [genel, dogu if kosu.get("dogu") else bati,
                    zemin[kosu.get("zemin") or "?"],
                    pist_k[_pist_tipi(kosu)],
                    yorum_k[kosu.get("yorum") or "?"],
                    irk_k[kosu.get("irk") or "?"],
                    mesafe_k[mesafe_kova(kosu.get("mesafe"))],
                    cins_k[cn],
                    yas_k[kosu.get("yas_sart") or "?"]]
        if hpz:
            hedefler.append(hpgap_k[gap_kova_ad(hpz["gap"])])
            hp_kapsar[1] += 1
            hp_kapsar_cins[cn][1] += 1
            if hpz["kazanan_hp_top3"]:
                hp_kapsar[0] += 1
                hp_kapsar_cins[cn][0] += 1
        for kv in hedefler:
            for st in STRATEJILER:
                ekle(kv[st], sec[st], kosu, abn)

        # --- KOMBINASYON: favoriyi bagimsiz sinyaller + baglamla birlestir ---
        favn = sec["favori"]
        if favn is not None:
            destek = sum(1 for k in ("orijin", "galop", "hp_lider") if sec.get(k) == favn)
            ekle(komb["favori_destek%d" % destek], favn, kosu, abn)
            zem = kosu.get("zemin") or "?"
            kotu = bool(kosu.get("dogu")) or zem == "KUM" or cn == "HANDIKAP"
            ekle(komb["kotu_baglam_favori" if kotu else "temiz_favori"], favn, kosu, abn)
            if (not kosu.get("dogu")) and zem in ("ÇİM", "SENTETİK") and cn != "HANDIKAP":
                ekle(komb["iyi_baglam_favori"], favn, kosu, abn)
            if sec.get("orijin") == favn and sec.get("galop") == favn and sec.get("hp_lider") == favn:
                ekle(komb["tam_mutabakat"], favn, kosu, abn)

        # --- KACAK-EZISMESI: tempoya gore kazananin stili + form lideri performansi ---
        tmp = kosu.get("tempo") or "?"
        kz = kosu.get("kazanan")
        kzr = abn.get(str(kz)) if kz is not None else None
        kz_stil = (kzr or {}).get("stil") or "?"
        kacak_senaryo[tmp][kz_stil] += 1
        kacak_yorum[kosu.get("yorum") or "?"][kz_stil] += 1
        try:
            _ptk = "%s|%s" % (_pist_tipi(kosu), tmp)
            kacak_pist[_ptk][kz_stil] += 1
            for _sa in (kosu.get("atlar") or []):
                saha_pist[_ptk][(_sa.get("stil") or "?")] += 1
        except Exception:
            pass
        ekle(form_tempo[tmp], sec.get("form"), kosu, abn)
        # --- AT-BASINA SINYALLER (once SADECE gecmisle degerlendir, sonra guncelle) ---
        try:
            _yid = "%s|%s|%s" % (kosu.get("tarih"), kosu.get("il") or kosu.get("sehir"), kosu.get("kosu_no"))
            _bugun_t = _trh(kosu)
            _atl3 = kosu.get("atlar") or []
            _adlar = {}
            for _a3 in _atl3:
                _nm = str(_a3.get("ad_norm") or "").strip()
                if _nm:
                    _adlar[_nm] = _a3
            # 1) REVANS-LIDERI
            _en, _en_no = 0, None
            for _nm, _a3 in _adlar.items():
                _gA = at_gecmis.get(_nm)
                if not _gA:
                    continue
                _puan = 0
                for _nm2 in _adlar:
                    if _nm2 == _nm:
                        continue
                    _gB = at_gecmis.get(_nm2)
                    if not _gB:
                        continue
                    _w = 0; _l = 0
                    for _rid, _sA in _gA.items():
                        _sB = _gB.get(_rid)
                        if _sA is None or _sB is None:
                            continue
                        if _sA < _sB:
                            _w += 1
                        elif _sB < _sA:
                            _l += 1
                    if _w > _l:
                        _puan += 1
                    elif _l > _w:
                        _puan -= 1
                if _puan > _en:
                    _en = _puan; _en_no = _a3.get("no")
            if _en_no is not None and _en >= 2:
                ekle(revans_k, _en_no, kosu, abn)
            # 2) GRUBA-YENI-AT (arsiv isinana kadar sayma)
            if len(_gorulen_at) >= 400:
                for _nm, _a3 in _adlar.items():
                    if _nm not in _gorulen_at:
                        ekle(yeniat_k, _a3.get("no"), kosu, abn)
            # 3) TAZE-TEKRARCI
            if _bugun_t is not None:
                _aday3 = []
                for _nm, _a3 in _adlar.items():
                    _sg = at_songun.get(_nm)
                    if not _sg or _sg[0] is None:
                        continue
                    _frk = (_bugun_t - _sg[0]).days
                    if 1 <= _frk <= 15 and _sg[1]:
                        _aday3.append(((_a3.get("son2ay_hiz") or 0), _a3.get("no")))
                if _aday3:
                    ekle(tekrarci_k, max(_aday3)[1], kosu, abn)
            # --- gecmisi guncelle (degerlendirmeden SONRA) ---
            for _nm, _a3 in _adlar.items():
                _s3 = _a3.get("sira")
                _s3 = _s3 if isinstance(_s3, int) else None
                at_gecmis.setdefault(_nm, {})[_yid] = _s3
                at_songun[_nm] = (_bugun_t, bool(_s3 is not None and _s3 <= 3))
                _gorulen_at.add(_nm)
        except Exception:
            pass
        # KACAK SAVASI -> en formlu TAKIPCI (kaçaklar eziserse arkadaki kapar)
        # YENI SINYAL: COKUS PISTI + HIZLI TEMPO -> en formlu TAKIP/GERIDEN
        try:
            if tmp == "hizli" and _pist_tipi(kosu) == "cokus":
                _atl2 = kosu.get("atlar") or []
                _arka = [a for a in _atl2 if a.get("stil") in ("TAKİP", "GERİDEN")]
                _sec2 = None
                _fl2 = [(a.get("son2ay_hiz"), a["no"]) for a in _arka if a.get("son2ay_hiz") is not None]
                if _fl2:
                    _sec2 = max(_fl2)[1]
                else:
                    _gc2 = [(gy_num(a.get("ganyan")), a["no"]) for a in _arka if gy_num(a.get("ganyan"))]
                    if _gc2:
                        _sec2 = min(_gc2)[1]
                if _sec2 is not None:
                    ekle(takip_cokus_k, _sec2, kosu, abn)
        except Exception:
            pass
        if "KAÇAK" in tr_up(kosu.get("yorum") or ""):
            atl = kosu.get("atlar") or []
            takipciler = [a for a in atl if a.get("stil") == "TAKİP"]
            secim = None
            fl = [(a.get("son2ay_hiz"), a["no"]) for a in takipciler if a.get("son2ay_hiz") is not None]
            if fl:
                secim = max(fl)[1]
            else:
                gc = [(gy_num(a.get("ganyan")), a["no"]) for a in takipciler if gy_num(a.get("ganyan"))]
                secim = min(gc)[1] if gc else None
            ekle(takip_kacakta, secim, kosu, abn)

        # CINSIYET x MESAFE
        mb = mesafe_kova(kosu.get("mesafe"))
        for a in (kosu.get("atlar") or []):
            kc = a.get("cinsiyet")
            if kc:
                kosan_cins_mesafe[mb][kc] += 1
            cins_kod_say[a.get("cins_kod") or "?"] += 1
        kzr2 = abn.get(str(kosu.get("kazanan")))
        if kzr2 and kzr2.get("cinsiyet"):
            kazanan_cins_mesafe[mb][kzr2["cinsiyet"]] += 1

    kutup = {
        "guncelleme": son_gun,
        "toplam_kosu": toplam_kosu,
        "aciklama": "illa birinci degil (kazanma+tabela+ROI+boy); dogu ayri; cins + HP-dominans dahil",
        "genel": kova_ozet(genel),
        "dogu": kova_ozet(dogu),
        "bati": kova_ozet(bati),
        "zemin": {z: kova_ozet(zemin[z]) for z in zemin},
        "pist": {p: kova_ozet(pist_k[p]) for p in pist_k},
        "yorum": {y: kova_ozet(yorum_k[y]) for y in yorum_k},
        "irk": {ir: kova_ozet(irk_k[ir]) for ir in irk_k},
        "mesafe": {mm: kova_ozet(mesafe_k[mm]) for mm in mesafe_k},
        "cins": {c: kova_ozet(cins_k[c]) for c in cins_k},
        "yas_sart": {y: kova_ozet(yas_k[y]) for y in yas_k},
        "hp_gap": {g: kova_ozet(hpgap_k[g]) for g in hpgap_k},
        "hp_kazanan_ilk3_kapsama": {
            "oran": round(100.0 * hp_kapsar[0] / hp_kapsar[1], 1) if hp_kapsar[1] else None,
            "n": hp_kapsar[1],
            "cins": {c: round(100.0 * v[0] / v[1], 1) for c, v in hp_kapsar_cins.items() if v[1]},
        },
        "kombinasyon": {k: ozet(v) for k, v in komb.items()},
        "kacak_senaryo": {t: dict(d) for t, d in kacak_senaryo.items()},
        "kacak_pist": {k: dict(v) for k, v in kacak_pist.items()},
        "saha_pist": {k: dict(v) for k, v in saha_pist.items()},
        "kacak_yorum": {t: dict(d) for t, d in kacak_yorum.items()},
        "form_tempo": {t: ozet(form_tempo[t]) for t in form_tempo},
        "takip_kacakta": ozet(takip_kacakta),
        "takip_cokus": ozet(takip_cokus_k),
        "revans": (ozet(revans_k) if revans_k.n else None),
        "yeni_at": (ozet(yeniat_k) if yeniat_k.n else None),
        "tekrarci": (ozet(tekrarci_k) if tekrarci_k.n else None),
        "cins_kod_say": dict(cins_kod_say),
        "kazanan_cins_mesafe": {m: dict(d) for m, d in kazanan_cins_mesafe.items()},
        "kosan_cins_mesafe": {m: dict(d) for m, d in kosan_cins_mesafe.items()},
    }

    # aktif sinyal: guven>=orta VE favori tabanini ROI'da ya da ilk3'te acik gecen
    tb = kutup["genel"]["favori"]
    tb_roi = tb["roi"] if tb["roi"] is not None else -100
    tb_ilk3 = tb["ilk3"] if tb["ilk3"] is not None else 0
    aktif = []

    def tara(etiket, kova):
        for st in STRATEJILER:
            o = kova[st]
            if o["guven"] not in ("orta", "yuksek") or o["n"] == 0:
                continue
            roi_iyi = o["roi"] is not None and o["roi"] > tb_roi + 5
            tab_iyi = o["ilk3"] is not None and o["ilk3"] > tb_ilk3 + 8
            if roi_iyi or tab_iyi:
                aktif.append("%s | %s: ROI %s%% ilk3 %s%% (n=%d, %s)" %
                             (etiket, st, o["roi"], o["ilk3"], o["n"], o["guven"]))

    tara("GENEL", kutup["genel"])
    tara("DOGU", kutup["dogu"]); tara("BATI", kutup["bati"])
    for z, kv in kutup["zemin"].items():
        tara("ZEMIN:" + z, kv)
    for p, kv in kutup["pist"].items():
        if p != "?":
            tara("PIST:" + p, kv)
    for y, kv in kutup["yorum"].items():
        if y != "?":
            tara("YORUM:" + y, kv)
    for ir, kv in kutup["irk"].items():
        if ir != "?":
            tara("IRK:" + ir, kv)
    for mm, kv in kutup["mesafe"].items():
        if mm != "?":
            tara("MESAFE:" + mm, kv)
    for c, kv in kutup["cins"].items():
        if c != "?":
            tara("CINS:" + c, kv)
    for gg, kv in kutup["hp_gap"].items():
        tara("HPGAP:" + gg, kv)
    for k, o in kutup["kombinasyon"].items():
        if o["guven"] in ("orta", "yuksek") and o["n"]:
            roi_iyi = o["roi"] is not None and o["roi"] > tb_roi + 5
            tab_iyi = o["ilk3"] is not None and o["ilk3"] > tb_ilk3 + 8
            if roi_iyi or tab_iyi:
                aktif.append("KOMBO:%s ROI %s%% ilk3 %s%% (n=%d, %s)" %
                             (k, o["roi"], o["ilk3"], o["n"], o["guven"]))
    kutup["aktif_sinyaller"] = aktif

    os.makedirs(GECMIS, exist_ok=True)
    io.open(KUTUP, "w", encoding="utf-8").write(json.dumps(kutup, ensure_ascii=False, indent=1))
    if son_gun:
        io.open(os.path.join(GECMIS, son_gun.replace(".", "-") + ".json"), "w",
                encoding="utf-8").write(json.dumps(kutup, ensure_ascii=False))

    def satir(etiket, o):
        return "%-22s n=%-4s kaz=%-5s ilk3=%-5s ilk4=%-5s ROI=%-7s (%s)" % (
            etiket[:22], o["n"], o["kazanma"], o["ilk3"], o["ilk4"], o["roi"], o["guven"])

    print("SINYAL KUTUPHANE ->", KUTUP)
    print("  toplam kosu:", toplam_kosu, "| son gun:", son_gun)
    print("  -- GENEL (tum stratejiler) --")
    for st in STRATEJILER:
        print("   ", satir(st, kutup["genel"][st]))
    print("  -- PIST TIPI (yeni harita: finis/cokus/notr) --")
    for p in ("finis", "cokus", "notr"):
        kv = kutup["pist"].get(p)
        if not kv:
            continue
        for st in ("favori", "toplam_derece", "son800", "form", "orijin"):
            if st in kv:
                print("   ", satir("P:%s %s" % (p, st), kv[st]))
    print("  -- DOGU vs BATI (favori) --")
    print("   ", satir("dogu", kutup["dogu"]["favori"]))
    print("   ", satir("bati", kutup["bati"]["favori"]))
    print("  -- ZEMIN (favori) --")
    for z, kv in kutup["zemin"].items():
        print("   ", satir("zemin:" + str(z), kv["favori"]))
    print("  -- CINS/kosu sinifi (favori) --")
    for c, kv in kutup["cins"].items():
        print("   ", satir("cins:" + str(c), kv["favori"]))

    def en_iyi_metrik(kova, olcut):
        aday = [(kova[st][olcut], st, kova[st]["n"]) for st in STRATEJILER
                if kova[st]["n"] >= 8 and kova[st].get(olcut) is not None]
        if not aday:
            return None
        aday.sort(reverse=True)
        return aday[0]

    print("  -- HANGI CINSTE HANGI METRIK ONDE (on izlenim, n>=10) --")
    for c, kv in sorted(kutup["cins"].items(), key=lambda x: -x[1]["favori"]["n"]):
        if kv["favori"]["n"] < 10:
            continue
        br = en_iyi_metrik(kv, "roi")
        bt = en_iyi_metrik(kv, "ilk3")
        rtxt = ("%s %+.0f%%" % (br[1], br[0])) if br else "-"
        ttxt = ("%s %s%%" % (bt[1], bt[0])) if bt else "-"
        print("    cins:%-14s ROI-lider: %-22s tabela-lider: %s (n=%d)" %
              (str(c), rtxt, ttxt, kv["favori"]["n"]))
    print("  -- HANGI ZEMINDE HANGI METRIK ONDE --")
    for z, kv in sorted(kutup["zemin"].items(), key=lambda x: -x[1]["favori"]["n"]):
        if kv["favori"]["n"] < 10:
            continue
        br = en_iyi_metrik(kv, "roi")
        bt = en_iyi_metrik(kv, "ilk3")
        rtxt = ("%s %+.0f%%" % (br[1], br[0])) if br else "-"
        ttxt = ("%s %s%%" % (bt[1], bt[0])) if bt else "-"
        print("    zemin:%-12s ROI-lider: %-22s tabela-lider: %s (n=%d)" %
              (str(z), rtxt, ttxt, kv["favori"]["n"]))
    print("  -- YAS SINIFI (favori) --")
    for y, kv in sorted(kutup["yas_sart"].items(), key=lambda x: -x[1]["favori"]["n"]):
        print("   ", satir("yas:" + str(y), kv["favori"]))
    print("  -- HANGI YAS SINIFINDA HANGI METRIK ONDE (n>=10) --")
    for y, kv in sorted(kutup["yas_sart"].items(), key=lambda x: -x[1]["favori"]["n"]):
        if kv["favori"]["n"] < 10 or y == "?":
            continue
        br = en_iyi_metrik(kv, "roi")
        bt = en_iyi_metrik(kv, "ilk3")
        rtxt = ("%s %+.0f%%" % (br[1], br[0])) if br else "-"
        ttxt = ("%s %s%%" % (bt[1], bt[0])) if bt else "-"
        print("    yas:%-8s ROI-lider: %-22s tabela-lider: %s (n=%d)" %
              (str(y), rtxt, ttxt, kv["favori"]["n"]))
    print("  -- HP-DOMINANS: HP lideri, rakibe acik araya gore (hp_lider stratejisi) --")
    for gg, kv in kutup["hp_gap"].items():
        print("   ", satir("gap:" + str(gg), kv["hp_lider"]))
    hk = kutup["hp_kazanan_ilk3_kapsama"]
    print("  -- Kazanan, HP'ce ilk-3 icinde mi (kapsama) --")
    print("     genel: %s%% (n=%d)" % (hk["oran"], hk["n"]))
    for c, o in sorted(hk["cins"].items(), key=lambda kv: -kv[1]):
        print("       cins:%-14s %s%%" % (c, o))
    print("  -- KOMBINASYON (birlesik sinyaller) --")
    for k in ("favori_destek3", "favori_destek2", "favori_destek1", "favori_destek0",
              "tam_mutabakat", "iyi_baglam_favori", "temiz_favori", "kotu_baglam_favori"):
        print("   ", satir(k, kutup["kombinasyon"][k]))
    print("  -- KACAK-EZISMESI: tempoya gore KAZANANIN stili (cok kacak -> geriden mi geliyor?) --")
    for t in ("hizli", "normal", "yavas", "?"):
        d = kutup["kacak_senaryo"].get(t)
        if not d:
            continue
        tp = sum(d.values())
        dag = ", ".join("%s %d%%" % (st, round(100.0 * n / tp)) for st, n in
                        sorted(d.items(), key=lambda x: -x[1]))
        print("    tempo:%-7s (n=%d) kazanan stili: %s" % (t, tp, dag))
    print("  -- PIST x TEMPO CAPRAZI: kazanan stil vs kadro payi (taban) --")
    kp_ = kutup.get("kacak_pist") or {}
    sp_ = kutup.get("saha_pist") or {}
    for kk_ in sorted(kp_, key=lambda x: -sum(kp_[x].values())):
        if "?" in kk_:
            continue
        d_ = kp_[kk_]; t_ = sum(d_.values())
        if t_ < 5:
            continue
        sh_ = sp_.get(kk_) or {}
        ts_ = float(sum(sh_.values()) or 1)
        par_ = []
        for st_, n_ in sorted(d_.items(), key=lambda x: -x[1]):
            par_.append("%s %d%%(taban %d%%)" % (st_, round(100.0 * n_ / t_),
                                                round(100.0 * sh_.get(st_, 0) / ts_)))
        print("    %-13s (n=%d) %s" % (kk_, t_, ", ".join(par_)))
    print("  -- KACAK-EZISMESI: TJK YORUM etiketine gore kazananin stili (daha guvenilir) --")
    for y, d in sorted(kutup["kacak_yorum"].items(), key=lambda x: -sum(x[1].values())):
        if y == "?":
            continue
        tp = sum(d.values())
        if tp < 5:
            continue
        dag = ", ".join("%s %d%%" % (st, round(100.0 * n / tp)) for st, n in
                        sorted(d.items(), key=lambda x: -x[1]))
        print("    yorum:%-16s (n=%d) kazanan stili: %s" % (str(y)[:16], tp, dag))
    print("  -- FORM (son 2 ay) lideri, tempoya gore --")
    for t in ("hizli", "normal", "yavas"):
        if t in kutup["form_tempo"]:
            print("   ", satir("form@" + t, kutup["form_tempo"][t]))
    print("  -- SINYAL: KACAK SAVASINDA en formlu TAKIPCI (kaçaklar eziserse arkadaki kapar) --")
    print("   ", satir("takip@kacak-yorumu", kutup["takip_kacakta"]))
    _tc = kutup.get("takip_cokus") or {}
    if _tc.get("n"):
        print("   ", satir("takip@cokus-hizli", _tc))
    else:
        print("    takip@cokus-hizli      (henuz veri toplaniyor - kova bos)")
    print("  -- AT-BASINA SINYALLER (grup dinamikleri) --")
    for _ad4, _kk4 in (("revans-lideri", "revans"), ("gruba-yeni-at", "yeni_at"), ("taze-tekrarci", "tekrarci")):
        _o4 = kutup.get(_kk4) or {}
        if _o4.get("n"):
            print("   ", satir(_ad4, _o4))
        else:
            print("    %-22s (henuz veri toplaniyor)" % _ad4)
    print("  -- CINSIYET x MESAFE (disiler uzunda erkeklerin yaninda geliyor mu?) --")
    print("     ham sex kodu dagilimi (esleme dogrulama):", kutup["cins_kod_say"])
    for mb in ("sprint<=1300", "orta1400-1700", "uzun>=1800"):
        ko = kutup["kosan_cins_mesafe"].get(mb, {})
        kz = kutup["kazanan_cins_mesafe"].get(mb, {})
        tko = sum(ko.values()); tkz = sum(kz.values())
        if tko < 5:
            continue
        dk = round(100.0 * ko.get("DİŞİ", 0) / tko)
        dz = round(100.0 * kz.get("DİŞİ", 0) / tkz) if tkz else 0
        print("    %-16s DISI: kosan %%%d -> kazanan %%%d | ERKEK: kosan %%%d -> kazanan %%%d (kosu n=%d)" %
              (mb, dk, dz, 100 - dk, 100 - dz, tkz))
    print("  -- YORUM/tempo (favori) --")
    for y, kv in kutup["yorum"].items():
        print("   ", satir("yorum:" + str(y), kv["favori"]))
    print("  -- IRK (favori) --")
    for ir, kv in kutup["irk"].items():
        print("   ", satir("irk:" + str(ir), kv["favori"]))
    print("  -- MESAFE (favori) --")
    for mm, kv in kutup["mesafe"].items():
        print("   ", satir("mesafe:" + str(mm), kv["favori"]))
    print("  AKTIF SINYAL (guven>=orta VE favoriyi gecen):", len(aktif))
    for a in aktif:
        print("    +", a)
    if not aktif:
        print("    (henuz yok - veri birikince cikacak)")
    print("SINYAL_URETICI_BITTI")


if __name__ == "__main__":
    main()

