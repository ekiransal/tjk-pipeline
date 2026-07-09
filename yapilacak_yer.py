#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AŞAMA 2 - AnaTabloYeniFinal  (Python karşılığı)
================================================

Excel makrosu 'AnaTabloYeniFinal' işinin Python karşılığı:
  Girdi : 'ana' sayfası (= 1. aşamanın 'yeni yer' çıktısı)
  Çıktı : 'yapılacak yer' ana tablosu (koşu bloklarına bölünmüş, sıkıştırılmış)

Bu modül SADECE ana tabloyu üretir (orjin/dede/galop/son galop yerleştirme
ayrı modüllerde eklenecek). Renk/merge gibi biçim şimdilik uygulanmaz;
amaç doğru DEĞER ve DÜZEN.

Konumlandırma makroyla birebir aynı:
  - İlk koşu başlığı 31. satır (KOSU_UST_BLOK = 30, +1).
  - Veri başlığın 2 altından başlar.
  - Sıkıştırma sonrası bir sonraki başlık: blok_son + 30 + 2.
"""

import re
import datetime

KOSU_UST_BLOK = 30
KOSU_TAMPON = 2

# ana (yeni yer) kolon numaraları -> çıktı kolonları (YazKaynakSatiri eşlemesi)
# (dst_col, ana_col).  ana_col None ise özel hesap.
DST_FROM_ANA = {
    1: 2, 2: 3, 3: 7, 4: 8, 5: 9, 6: 10, 7: 2, 8: 11, 9: 12, 10: 13,
    11: 14, 12: 19, 13: 2, 14: 16, 15: 18, 16: 63, 17: 64, 18: 47,
    20: 2, 21: 28, 22: 29, 23: 30, 24: 31, 25: 48, 26: 2, 27: 50,
    28: 34, 29: 35, 30: 36, 31: 47, 32: 37, 33: 38, 34: 61, 35: 62,
    36: 46, 37: 49, 38: 47, 39: 53, 40: 60, 41: 65, 52: 76, 53: 77,
}


# ---------------- yardımcılar ----------------

_EXCEL_HATA = {"#REF!", "#N/A", "#YOK", "#VALUE!", "#DEĞER!", "#DIV/0!",
               "#SAYI/0!", "#NAME?", "#AD?", "#NULL!", "#BOŞ!", "#NUM!", "#SAYI!"}


def _ham(ana_row, c):
    """ana satırından 1-tabanlı c kolonu; Excel hata hücresi/None -> '' (makro HamDeger gibi)."""
    if c - 1 < len(ana_row):
        v = ana_row[c - 1]
        if v is None:
            return ""
        if isinstance(v, float) and v != v:   # NaN -> boş
            return ""
        if isinstance(v, str) and v.strip().upper() in _EXCEL_HATA:
            return ""
        return v
    return ""


def _meaningful(v):
    if v is None:
        return False
    t = str(v).strip().upper()
    return t not in ("", "#YOK", "#N/A")


def _norm(v):
    if v is None:
        return ""
    t = str(v).strip().upper()
    if t in ("", "#YOK", "#N/A"):
        return ""
    return v


def _sayi(v):
    """SayiyaCevir: -> (ok, float). TR ondalık toleranslı."""
    if v is None:
        return (False, 0.0)
    if isinstance(v, bool):
        return (False, 0.0)
    if isinstance(v, (int, float)):
        f = float(v)
        if f != f or f in (float("inf"), float("-inf")):  # NaN / sonsuz
            return (False, 0.0)
        return (True, f)
    t = str(v).strip()
    if t == "" or t.upper() in ("#YOK", "#N/A"):
        return (False, 0.0)
    t = t.replace("\xa0", "").replace(" ", "")
    if "." in t and "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return (True, float(t))
    except Exception:
        return (False, 0.0)


def _percentile_inc(vals, yuzde):
    """Excel PERCENTILE.INC (linear). vals: sıralı olmayan float listesi."""
    if not vals:
        return None
    arr = sorted(vals)
    n = len(arr)
    if n == 1:
        return arr[0]
    y = max(0.0, min(100.0, yuzde))
    pos = 1 + (n - 1) * (y / 100.0)
    alt = int(pos)            # floor (pos>=1)
    ust = alt + 1
    if ust > n:
        return arr[n - 1]
    oran = pos - alt
    return arr[alt - 1] + (arr[ust - 1] - arr[alt - 1]) * oran


def _liste_percentil(liste_raw, yuzde):
    """'45-50-53-58' / '54-55,5' listesinden percentile.inc. -> float|None"""
    if liste_raw is None:
        return None
    t = str(liste_raw).strip()
    if t == "" or t.upper() in ("#YOK", "#N/A"):
        return None
    for ch in ("–", "‒", ";", "|"):
        t = t.replace(ch, "-")
    vals = []
    for p in t.split("-"):
        ok, d = _sayi(p)
        if ok:
            vals.append(d)
    if not vals:
        return None
    return _percentile_inc(vals, yuzde)


def _tarih_key(v, satir_no):
    """TarihSiralamaAnahtari: tarih -> sıralanabilir sayı; yoksa büyük + satır."""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.toordinal()
    t = str(v).strip() if v is not None else ""
    if t:
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(t[:10], fmt).toordinal()
            except Exception:
                pass
    return 1_000_000_000 + satir_no


def _temiz_hp_metni(v):
    if v is None:
        return ""
    t = str(v).strip()
    if t == "" or t.upper() in ("#YOK", "#N/A"):
        return ""
    ok, d = _sayi(v)
    if ok:
        return str(int(d)) if d == int(d) else str(d)
    return t


# ---------------- ana hesap ----------------

class AnaTablo:
    def __init__(self, ana_rows):
        """ana_rows: list[list]  (başlık DAHİL; satır 1 = başlık). 1-tabanlı kolonlar 0-index."""
        self.ana = ana_rows
        self.son = len(ana_rows)
        # Koşu bazında benzersiz at (ana col3) HP(col54)/Kilo(col29) listeleri
        self._kosu_hp = {}
        self._kosu_kilo = {}
        self._kosu_hazirla()

    def _kosu_hazirla(self):
        hp, kilo = {}, {}
        seen_hp, seen_kilo = {}, {}
        for i in range(1, self.son):  # i=1.. (0=başlık)
            row = self.ana[i]
            kosu = str(_ham(row, 1)).strip()
            if kosu == "":
                continue
            atkey = str(_ham(row, 3)).strip().upper()
            if atkey == "":
                continue
            # HP (col54)
            okh, dh = _sayi(_ham(row, 54))
            if okh:
                s = seen_hp.setdefault(kosu, set())
                if atkey not in s:
                    s.add(atkey)
                    hp.setdefault(kosu, []).append(dh)
            # Kilo (col29)
            okk, dk = _sayi(_ham(row, 29))
            if okk:
                s = seen_kilo.setdefault(kosu, set())
                if atkey not in s:
                    s.add(atkey)
                    kilo.setdefault(kosu, []).append(dk)
        self._kosu_hp = hp
        self._kosu_kilo = kilo

    def _kosu_yeni_hp_p(self, kosu, yuzde):
        return _percentile_inc(self._kosu_hp.get(kosu, []), yuzde)

    def _kosu_yeni_kilo_p(self, kosu, yuzde):
        return _percentile_inc(self._kosu_kilo.get(kosu, []), yuzde)

    def _hp_dominans(self, row, yuzde):
        kosu = str(_ham(row, 1)).strip()
        oky, yeniHp = _sayi(_ham(row, 54))
        if not oky:
            return ""
        yeniP = self._kosu_yeni_hp_p(kosu, yuzde)
        if yeniP is None:
            return ""
        oke, eskiHp = _sayi(_ham(row, 55))
        eskiP = _liste_percentil(_ham(row, 66), yuzde)
        if oke and eskiP is not None:
            return (yeniHp - yeniP) - (eskiHp - eskiP)
        return (yeniHp - yeniP)

    def _kilo_dominans(self, row, yuzde):
        kosu = str(_ham(row, 1)).strip()
        oky, yeniKilo = _sayi(_ham(row, 29))
        if not oky:
            return ""
        yeniP = self._kosu_yeni_kilo_p(kosu, yuzde)
        if yeniP is None:
            return ""
        oke, eskiKilo = _sayi(_ham(row, 30))
        eskiP = _liste_percentil(_ham(row, 67), yuzde)
        if oke and eskiP is not None:
            return (yeniP - yeniKilo) - (eskiP - eskiKilo)
        return (yeniP - yeniKilo)

    def _hp_kronolojik(self, src_i):
        """AP: aynı at (col3) için BC(col55) eski HP'leri tarih(col34) sırasıyla,
        sonuna bu satırın BB(col54) yeni HP'si eklenir."""
        row = self.ana[src_i]
        atk = str(_ham(row, 3)).strip().upper()
        if atk == "":
            return ""
        kayit = []
        for i in range(1, self.son):
            rr = self.ana[i]
            if str(_ham(rr, 3)).strip().upper() != atk:
                continue
            hp = _temiz_hp_metni(_ham(rr, 55))
            if hp != "":
                kayit.append((_tarih_key(_ham(rr, 34), i), i, hp))
        kayit.sort(key=lambda x: (x[0], x[1]))
        gecmis = "-".join(k[2] for k in kayit)
        yeni = _temiz_hp_metni(_ham(row, 54))
        if gecmis != "" and yeni != "":
            return gecmis + "-" + yeni
        return gecmis + yeni

    def _kaynak_satiri(self, src_i):
        """YazKaynakSatiri: 54 kolonluk çıktı satırı (1-index -> liste[0..53])."""
        row = self.ana[src_i]
        out = [""] * 54
        for dst, anac in DST_FROM_ANA.items():
            out[dst - 1] = _ham(row, anac)
        # AP (42) kronolojik HP
        out[41] = self._hp_kronolojik(src_i)
        # AQ (43) = AL(38) değeri (koşu cinsi tekrar)
        out[42] = out[37]
        # AR (44) P50 HP dom, AS(45) P50 kilo dom
        out[43] = self._hp_dominans(row, 50)
        out[44] = self._kilo_dominans(row, 50)
        # AT(46) P66 HP, AU(47) P66 kilo
        out[45] = self._hp_dominans(row, 66.66)
        out[46] = self._kilo_dominans(row, 66.66)
        # AV(48) P75 HP, AW(49) P75 kilo
        out[47] = self._hp_dominans(row, 75)
        out[48] = self._kilo_dominans(row, 75)
        # AX(50) = AV(25) at no tekrar
        out[49] = out[24]
        # AY(51) = AQ(43)
        out[50] = out[42]
        # BB(54) geçici işaret: ana col43 veya col44 sayıysa 1
        m43 = _sayi(_ham(row, 43))[0]
        m44 = _sayi(_ham(row, 44))[0]
        out[53] = 1 if (m43 or m44) else ""
        return out

    # ---- sıkıştırma ----
    @staticmethod
    def _sol_dolu(r):
        for c in (3, 4, 5, 6, 8, 10, 11, 12, 14, 15, 16, 17):
            if _meaningful(r[c - 1]):
                return True
        return False

    @staticmethod
    def _sag_gercek(r):
        return _sayi(r[53])[0]   # çıktı col54 (BB geçici işaret)

    @staticmethod
    def _aj_iki(r):
        aj = str(r[35]).strip().upper() if r[35] is not None else ""  # çıktı col36
        if aj not in ("", "#YOK", "#N/A"):
            return False
        return _sayi(r[53])[0]

    def _sikistir(self, raw):
        sol = []
        for r in raw:
            if self._sol_dolu(r):
                sol.append([_norm(r[j - 1]) for j in range(1, 19)])  # col1..18
        sag = []
        for r in raw:
            if self._sag_gercek(r):
                rr = [_norm(r[j - 1]) for j in range(20, 55)]  # col20..54 (35)
                if self._aj_iki(r):
                    rr[16] = "iki"   # sağ blok col17 = çıktı col36
                sag.append(rr)
        maxs = max(len(sol), len(sag))
        out = []
        for i in range(maxs):
            ro = [""] * 54
            if i < len(sol):
                for j in range(18):
                    ro[j] = sol[i][j]
            ro[18] = ""
            if i < len(sag):
                for j in range(35):
                    ro[19 + j] = sag[i][j]
            ro[53] = ""  # BB temizle
            out.append(ro)
        return out

    def uret(self):
        """Çıktı: dict {(row,col): value} (1-tabanlı) + son satır."""
        G = {}
        def yaz(r, c, v):
            if v != "" and v is not None:
                G[(r, c)] = v

        out_row = KOSU_UST_BLOK + 1   # 31
        onceki = ""
        blok_bas = 0
        raw = []

        def blok_kapat(blok_bas, raw):
            packed = self._sikistir(raw)
            for idx, ro in enumerate(packed):
                rr = blok_bas + idx
                for c in range(1, 55):
                    if ro[c - 1] != "" and ro[c - 1] is not None:
                        G[(rr, c)] = ro[c - 1]
            son = blok_bas + len(packed) - 1
            if son < blok_bas:
                son = blok_bas - 1
            return son

        for i in range(1, self.son):
            row = self.ana[i]
            mevcut = str(_ham(row, 1)).strip()
            baslik = str(_ham(row, 6)).strip()
            if mevcut == "" or baslik == "":
                continue
            if mevcut != onceki:
                if onceki != "":
                    son = blok_kapat(blok_bas, raw)
                    sonraki = (son + 1) if son >= blok_bas else blok_bas
                    out_row = sonraki + KOSU_UST_BLOK + KOSU_TAMPON
                    raw = []
                yaz(out_row, 1, baslik)   # başlık (merge yok)
                blok_bas = out_row + 2
                out_row = blok_bas
                onceki = mevcut
            raw.append(self._kaynak_satiri(i))
            out_row += 1

        if onceki != "":
            blok_kapat(blok_bas, raw)

        son_satir = max((rc[0] for rc in G), default=0)
        return G, son_satir


def hesapla(ana_rows):
    """ana_rows (başlık dahil) -> (G dict {(r,c):val}, son_satir)."""
    return AnaTablo(ana_rows).uret()


# =========================================================================
# ORJIN / DEDE YERLEŞTİRME  (YazOrjinDinamik / YazDedeDinamik karşılığı)
# =========================================================================

_GERCEK_ORJIN_BASLIK = {
    "ORAN", "MESAFE", "YAPISI", "SİTİLİ", "STİLİ", "SITILI", "STILI",
    "GENEL KAÇAK", "GENEL KACAK", "FARK SPRINT", "ORAN SPRINT",
    "KUM", "ÇİM", "CİM", "CIM", "SENTETİK", "SENTETIK",
}
_YASAK_ORJIN_BASLIK = {
    "BABA DOĞU", "BABA DOGU", "EXTREM BABA KUM",
    "EXTREM BABA ÇİM", "EXTREM BABA CIM",
}


def _baslik_kosu_mu(txt):
    t = str(txt or "").strip()
    if t == "":
        return False
    tl = t.lower()
    return (". koşu" in tl) or (". kosu" in tl)


def _baslik_kosu_no(baslik):
    s = ""
    for ch in str(baslik or ""):
        if ch.isdigit():
            s += ch
        elif s != "":
            break
    return int(s) if s else 0


def _gercek_orjin_basligi(txt):
    t = str(txt or "").strip().upper()
    if t in _YASAK_ORJIN_BASLIK:
        return False
    return t in _GERCEK_ORJIN_BASLIK


def _genel_kacak(txt):
    return str(txt or "").strip().upper() in ("GENEL KAÇAK", "GENEL KACAK")


def _formatli_deger(v):
    if v is None:
        return ""
    t = str(v).strip()
    if t == "" or t.upper() in _EXCEL_HATA:
        return ""
    ok, d = _sayi(v)
    if ok:
        from decimal import Decimal, ROUND_HALF_UP
        q = Decimal(str(d)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if q == 0:
            q = abs(q)
        return f"{q}".replace(".", ",")
    return str(v)


def _temiz_hucre_str(v):
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d.%m.%Y")
    t = str(v).strip()
    if t == "" or t.upper() in _EXCEL_HATA:
        return ""
    # tam sayı float ise '11.0' yerine '11'
    if isinstance(v, float):
        if v != v or v in (float("inf"), float("-inf")):  # NaN / sonsuz
            return ""
        if v == int(v):
            return str(int(v))
    return t


def _birlestir_no_deger(no, deger):
    no_t = _temiz_hucre_str(no)
    deger_t = _formatli_deger(deger)
    if no_t == "" and deger_t == "":
        return ""
    if no_t != "" and deger_t != "":
        return f"{no_t} - {deger_t}"
    return no_t or deger_t


def _o(rows, r, c):
    """1-tabanlı kaynak hücre."""
    if 1 <= r <= len(rows):
        row = rows[r - 1]
        if 1 <= c <= len(row):
            return row[c - 1]
    return None


def _kosu_no_satiri_bul(rows, kosu_no):
    for r in range(1, len(rows) + 1):
        t = str(_o(rows, r, 1) or "").strip().upper()
        if t in ("KOŞU NO", "KOSU NO"):
            ok, d = _sayi(_o(rows, r, 2))
            if ok and int(d) == kosu_no:
                return r
    return 0


def _baslik_alt_kolon_bul(rows, alt_satir, baslik_kol, *aranan):
    ar = [a.upper() for a in aranan]
    for c in range(baslik_kol, baslik_kol + 7):
        t = str(_o(rows, alt_satir, c) or "").strip().upper()
        if t in ar:
            return c
    return 0


def _kolondaki_dolu_no(rows, col, start):
    say = 0
    bosluk = 0
    r = start
    while r <= len(rows):
        v = _o(rows, r, col)
        if v is None or str(v).strip().upper() in ({"", } | _EXCEL_HATA):
            bosluk += 1
            if bosluk >= 2:
                break
        else:
            bosluk = 0
            say += 1
        if say > 100:
            break
        r += 1
    return say


def _yaz_orjin_dinamik(G, src_rows, base_row, target_row, start_col=1):
    baslik_satir = base_row + 2
    alt_satir = base_row + 3
    data_start = base_row + 4
    # son başlık kolonu
    son_bk = 1
    for c in range(1, 200):
        if str(_o(src_rows, baslik_satir, c) or "").strip() != "":
            son_bk = c
    bloklar = []
    maxs = 0
    i = 1
    while i <= son_bk:
        bm = str(_o(src_rows, baslik_satir, i) or "").strip()
        if _gercek_orjin_basligi(bm):
            alt = str(_o(src_rows, alt_satir, i + 1) or "").strip().upper()
            if alt == "NO":
                no_kol = i + 1
                deger_kol = _baslik_alt_kolon_bul(src_rows, alt_satir, i, "DEĞER", "DEGER")
                final_kol = 0
                if _genel_kacak(bm):
                    final_kol = _baslik_alt_kolon_bul(src_rows, alt_satir, i, "FİNAL", "FINAL")
                    gen = 2
                else:
                    gen = 1
                if deger_kol > 0:
                    rc = _kolondaki_dolu_no(src_rows, no_kol, data_start)
                    bloklar.append((i, no_kol, deger_kol, final_kol, gen, rc))
                    maxs = max(maxs, rc)
        i += 1
    if not bloklar:
        return
    if maxs == 0:
        maxs = 1
    # başlıklar
    yk = start_col
    for (bk, nok, degk, fink, gen, rc) in bloklar:
        bm = str(_o(src_rows, baslik_satir, bk) or "").strip()
        G[(target_row, yk)] = bm
        G[(target_row + 1, yk)] = "NO - DEĞER"
        if _genel_kacak(bm):
            G[(target_row + 1, yk + 1)] = "FİNAL"
        yk += gen
    # veri
    hedef = target_row + 2
    yk = start_col
    for (bk, nok, degk, fink, gen, rc) in bloklar:
        for k in range(1, maxs + 1):
            if k <= rc:
                nd = _birlestir_no_deger(_o(src_rows, data_start + k - 1, nok),
                                         _o(src_rows, data_start + k - 1, degk))
                if nd != "":
                    G[(hedef + k - 1, yk)] = nd
                if gen == 2 and fink > 0:
                    fv = _formatli_deger(_o(src_rows, data_start + k - 1, fink))
                    if fv != "":
                        G[(hedef + k - 1, yk + 1)] = fv
        yk += gen


def _son_kullanilan_kolon(G, ilk_satir, son_satir):
    son = 0
    for (r, c), v in G.items():
        if ilk_satir <= r <= son_satir and v is not None and str(v).strip() != "":
            if c > son:
                son = c
    return son


def orjin_yerlestir(G, orjin_rows):
    """yapılacak yer grid'ine orjin bloklarını yerleştirir (başlık satırı - 30)."""
    titles = [r for (r, c) in list(G.keys()) if c == 1 and _baslik_kosu_mu(G.get((r, 1)))]
    for r in sorted(set(titles)):
        kno = _baslik_kosu_no(G[(r, 1)])
        if kno <= 0:
            continue
        base = _kosu_no_satiri_bul(orjin_rows, kno)
        if base <= 0:
            continue
        hedef = r - 30
        if hedef >= 1:
            _yaz_orjin_dinamik(G, orjin_rows, base, hedef, start_col=1)


def dede_yerlestir(G, dede_rows):
    """dede bloklarını orjin'in sağına yerleştirir (start_col = son kullanılan + 2)."""
    titles = [r for (r, c) in list(G.keys()) if c == 1 and _baslik_kosu_mu(G.get((r, 1)))]
    for r in sorted(set(titles)):
        kno = _baslik_kosu_no(G[(r, 1)])
        if kno <= 0:
            continue
        base = _kosu_no_satiri_bul(dede_rows, kno)
        if base <= 0:
            continue
        hedef = r - 30
        if hedef < 1:
            continue
        bas_kol = _son_kullanilan_kolon(G, hedef, r - 1) + 2
        if bas_kol < 1:
            bas_kol = 1
        _yaz_orjin_dinamik(G, dede_rows, base, hedef, start_col=bas_kol)


# =========================================================================
# GALOP / SON GALOP YERLEŞTİRME
# =========================================================================

def _galop_baslik_benzeri(v):
    t = str(v or "").strip().upper()
    if t == "" or t in _EXCEL_HATA:
        return False
    return ("+" in t) or ("KENTER" in t) or ("ÇAPRAZ" in t) or ("CAPRAZ" in t) or ("FARK" in t)


def _tr_upper(s):
    return (str(s or "").upper().replace("İ", "I").replace("I", "I")
            .replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U")
            .replace("Ö", "O").replace("Ç", "C"))


def _galop_kosu_satiri(rows, kosu_no, sehir=None, tek_il=False):
    """{kosu_no}. Koşu satırını bulur.
    sehir verilirse SADECE o ili içeren başlık kabul edilir (kesin eşleşme).
    Şehir eşleşmezse:
      - tek_il=True (o gün tek şehir yarışıyor) -> numara tutan blok güvenle kullanılır;
      - tek_il=False (çok-illi gün) -> HİÇBİRİ döndürülmez. Yanlış ilin galopunu
        yapıştırmaktansa (Kocaeli'ye Ankara'nınki) galop boş kalır — bu KASITLI."""
    k1 = f"{kosu_no}. KOŞU"
    k2 = f"{kosu_no}. KOSU"
    sehir_u = _tr_upper(sehir) if sehir else None
    numara_tutan = 0
    for r in range(1, len(rows) + 1):
        traw = str(_o(rows, r, 1) or "").strip()
        t = traw.upper()
        if t and (k1 in t or k2 in t):
            if sehir_u is None:
                return r
            if sehir_u in _tr_upper(traw):
                return r          # doğru il — kesin eşleşme
            if numara_tutan == 0:
                numara_tutan = r   # numara tutuyor ama il farklı
    # şehir eşleşmedi: yalnız tek-illi günde (belirsizlik yok) numara tutanı kullan
    if sehir_u is not None and tek_il:
        return numara_tutan
    return 0                       # çok-illi gün -> yanlış il yapıştırma, boş bırak


def _sonraki_galop_kosu(rows, cur):
    for r in range(cur + 1, len(rows) + 1):
        t = str(_o(rows, r, 1) or "").strip().upper()
        if t and (". KOŞU" in t or ". KOSU" in t):
            return r
    return 0


def _galop_satiri_gecerli(atno, deger):
    if not _meaningful(atno) or not _meaningful(deger):
        return False
    if str(atno).strip().upper() == "AT NO":
        return False
    if str(deger).strip().upper() in ("DEĞER", "DEGER"):
        return False
    return True


def _format_galop_deger(v):
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d.%m.%Y")
    t = str(v).strip()
    if t == "" or t.upper() in _EXCEL_HATA:
        return ""
    ok, d = _sayi(v)
    if ok:
        from decimal import Decimal, ROUND_HALF_UP
        q = Decimal(str(d)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if q == 0:
            q = abs(q)   # -0,00 -> 0,00
        return f"{q}".replace(".", ",")
    return str(v)


def _galop_tek_hucre(atno, deger, sekli):
    a = _temiz_hucre_str(atno)
    dd = _format_galop_deger(deger)
    s = _temiz_hucre_str(sekli)
    return f"{a} / {dd} / {s}"


def _tarih_yaz(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d.%m.%Y")
    t = str(v or "").strip()
    if t == "" or t.upper() in _EXCEL_HATA:
        return ""
    return t


def _header_row_bul(rows, base, bloklar):
    for r in range(base, base + 11):
        for c in bloklar:
            if _galop_baslik_benzeri(_o(rows, r, c)):
                return r
    return 0


def _veri_baslangic_bul(rows, header_row, bloklar):
    for r in range(header_row + 1, header_row + 7):
        for c in (bloklar[0], bloklar[1] if len(bloklar) > 1 else bloklar[0]):
            if str(_o(rows, r, c) or "").strip().upper() == "AT NO":
                return r + 1
    for r in range(header_row + 1, header_row + 7):
        for c in bloklar:
            if _galop_satiri_gecerli(_o(rows, r, c), _o(rows, r, c + 1)):
                return r
    return 0


def _son_veri_satiri(rows, base, bloklar, genis):
    enson = base - 1
    r = base
    while r <= len(rows) + 2:
        bos = _satir_bos(rows, r, bloklar, genis)
        if bos:
            if r > base and _satir_bos(rows, r + 1, bloklar, genis) and _satir_bos(rows, r + 2, bloklar, genis):
                break
        else:
            enson = r
        r += 1
    return enson


def _satir_bos(rows, r, bloklar, genis):
    for c in bloklar:
        for k in range(genis):
            if str(_o(rows, r, c + k) or "").strip() != "":
                return False
    return True


def galop_yerlestir(G, galop_rows, sehir=None, tek_il=False):
    bloklar = [1, 7, 13, 19, 25, 31, 37]
    titles = sorted({r for (r, c) in list(G.keys()) if c == 1 and _baslik_kosu_mu(G.get((r, 1)))})
    for r in titles:
        kno = _baslik_kosu_no(G[(r, 1)])
        if kno <= 0:
            continue
        base = _galop_kosu_satiri(galop_rows, kno, sehir, tek_il)
        if base <= 0:
            continue
        target = r - 18
        arama_bas = r - 22
        if target < 1 or arama_bas < 1:
            continue
        header = _header_row_bul(galop_rows, base, bloklar)
        if header == 0:
            continue
        ds = _veri_baslangic_bul(galop_rows, header, bloklar)
        if ds == 0:
            continue
        nk = _sonraki_galop_kosu(galop_rows, base)
        last = (nk - 1) if nk > 0 else _son_veri_satiri(galop_rows, ds, bloklar, 4)
        if last < ds:
            continue
        outcol = _son_kullanilan_kolon(G, arama_bas, r - 18) + 2
        if outcol < 12:
            outcol = 12
        for srccol in bloklar:
            bb = _temiz_hucre_str(_o(galop_rows, header, srccol))
            if bb == "":
                continue
            dolu = any(_galop_satiri_gecerli(_o(galop_rows, rr, srccol), _o(galop_rows, rr, srccol + 1))
                       for rr in range(ds, last + 1))
            if not dolu:
                continue
            G[(target, outcol)] = bb
            G[(target + 1, outcol)] = "at / değer / şekil"
            G[(target + 1, outcol + 1)] = "tarih"
            od = target + 2
            for rr in range(ds, last + 1):
                atno = _o(galop_rows, rr, srccol)
                deger = _o(galop_rows, rr, srccol + 1)
                sekli = _o(galop_rows, rr, srccol + 2)
                tarih = _o(galop_rows, rr, srccol + 3)
                if _galop_satiri_gecerli(atno, deger):
                    if od <= target + 11:
                        G[(od, outcol)] = _galop_tek_hucre(atno, deger, sekli)
                        tv = _tarih_yaz(tarih)
                        if tv != "":
                            G[(od, outcol + 1)] = tv
                        od += 1
                    else:
                        break
            outcol += 2


def songalop_yerlestir(G, songalop_rows, sehir=None, tek_il=False):
    bloklar = [1, 8, 15, 22, 29, 36, 43]
    titles = sorted({r for (r, c) in list(G.keys()) if c == 1 and _baslik_kosu_mu(G.get((r, 1)))})
    for r in titles:
        kno = _baslik_kosu_no(G[(r, 1)])
        if kno <= 0:
            continue
        base = _galop_kosu_satiri(songalop_rows, kno, sehir, tek_il)
        if base <= 0:
            continue
        target = r - 18
        if target < 1:
            continue
        header = _header_row_bul(songalop_rows, base, bloklar)
        if header == 0:
            continue
        ds = _veri_baslangic_bul(songalop_rows, header, bloklar)
        if ds == 0:
            continue
        nk = _sonraki_galop_kosu(songalop_rows, base)
        last = (nk - 1) if nk > 0 else _son_veri_satiri(songalop_rows, ds, bloklar, 5)
        if last < ds:
            continue
        # son galop, galop'un sağına: targetRow..target+11 son kullanılan + 2
        outcol = _son_kullanilan_kolon(G, target, target + 11) + 2
        if outcol < 12:
            outcol = 12
        for srccol in bloklar:
            bb = _temiz_hucre_str(_o(songalop_rows, header, srccol))
            if bb == "":
                continue
            # KRİTİK: "KENTER (son galop)" bloğunda 'değer' KOLONU YOK.
            #   Mesafe blokları : at no, değer, şekli, tarih, şehir
            #   KENTER (son galop): at no,        şekli, tarih, şehir
            # Eskiden hepsi aynı sanılıp KENTER bloğunda her şey bir kolon
            # kayıyordu (şekli->değer, tarih->şekli, şehir->tarih). Düzeltildi.
            kenter = ("KENTER" in bb.upper()) or ("SON GALOP" in bb.upper())
            if kenter:
                o_deger, o_sekli, o_tarih, o_sehir = None, 1, 2, 3
            else:
                o_deger, o_sekli, o_tarih, o_sehir = 1, 2, 3, 4

            def _spr_gecerli(rr):
                at = _o(songalop_rows, rr, srccol)
                if not _meaningful(at) or str(at).strip().upper() == "AT NO":
                    return False
                ref = _o(songalop_rows, rr, srccol + (o_sekli if kenter else o_deger))
                r = "" if ref is None else str(ref).strip().upper()
                return r not in ("", "DEĞER", "DEGER", "ŞEKLİ", "SEKLI")

            dolu = any(_spr_gecerli(rr) for rr in range(ds, last + 1))
            if not dolu:
                continue
            G[(target, outcol)] = "SON " + bb
            G[(target + 1, outcol)] = ("at / şekil" if kenter else "at / değer / şekil")
            G[(target + 1, outcol + 1)] = "tarih"
            G[(target + 1, outcol + 2)] = "şehir"
            od = target + 2
            for rr in range(ds, last + 1):
                if not _spr_gecerli(rr):
                    continue
                if od > target + 11:
                    break
                atno = _o(songalop_rows, rr, srccol)
                sekli = _o(songalop_rows, rr, srccol + o_sekli)
                tarih = _o(songalop_rows, rr, srccol + o_tarih)
                # DÜZELTME (Elazığ vakası): hücredeki şehir AYRI değişkende tutulur.
                # Eskiden 'sehir' PARAMETRESİ eziliyordu -> ilk İstanbul/Bursa satırından
                # sonra TÜM sonraki koşular yanlış ilin bloğundan veri alıyordu.
                sat_sehir = _o(songalop_rows, rr, srccol + o_sehir)
                if kenter:
                    G[(od, outcol)] = f"{_temiz_hucre_str(atno)} / {_temiz_hucre_str(sekli)}"
                else:
                    deger = _o(songalop_rows, rr, srccol + o_deger)
                    G[(od, outcol)] = _galop_tek_hucre(atno, deger, sekli)
                tv = _tarih_yaz(tarih)
                if tv != "":
                    G[(od, outcol + 1)] = tv
                sv = _temiz_hucre_str(sat_sehir)
                if sv != "":
                    G[(od, outcol + 2)] = sv
                od += 1
            outcol += 3
