#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
'yeni yer' FORMÜL KATMANI  (Excel formüllerinin pure-Python karşılığı)
======================================================================

Excel'deki 'yeni yer' sayfası, Sayfa2'yi üç referans sayfaya VLOOKUP'larla
bağlayıp ~70 kolon hesaplıyor. Bu modül o formül grafiğini birebir Python'da
yeniden kurar.

Girdi sayfaları (hepsi DataFrame veya satır-listesi olarak):
  - Sayfa2 : 10'lu genişletilmiş bugünkü atlar (Kosu_No, At_No, At_Adi_Temiz,
             Siklet, Rep, Anahtar, Kosu_Basligi, Yas, KGS, Handikap_Puani,
             Kosu_HP_Tam_Ortanca, Kosu_Kilo_Tam_Ortanca)  -> A..L
  - sheet_800 : '800' referansı. A=Anahtar, B=rep, C..M veri.
  - derece    : 'derece' referansı. A=Anahtar, B=rep, C..(AG) veri.
  - medyanl   : '800 ve derece çekilecek medyanl'. A=birleşik anahtar, B..O.

Çıktı: 'yeni yer' satırları (A..BY), Excel ile aynı sırada.

VLOOKUP semantiği (Excel, 4. arg = 0 -> tam eşleşme):
  - İlk birebir eşleşen satır.
  - Bulunamazsa  -> NA  (Excel '#N/A')
  - col_index istenen tablo genişliğinden büyükse -> REF ('#REF!')
"""

import re
from openpyxl.utils import column_index_from_string as _col_idx


# 'yeni yer' çıktı başlıkları (A..BY = 77 kolon). Çözümlenen anlamlarıyla.
YENI_YER_BASLIK = [
    "Kosu_No", "At_No", "At_Adi_Temiz", "Siklet", "Anahtar", "Kosu_Basligi",   # A-F
    "Tarih_800", "Sehir_800", "Zemin_800", "PistDurumu_800", "Mesafe_800",     # G-K
    "Siklet_L", "M_800", "N(L-M)", "Irk_800", "P_son800", "Q_800", "R(Q-P)",   # L-R
    "KosuCinsi_800", "AnahtarT(Sehir+Irk+Zemin+Mesafe)", "U_son800_medyan",    # S-U
    "V_son800_orani", "W(P/U)", "X(W*V)", "Y_birlesik", "Z(=R)",               # V-Z
    "At_No_AA", "Anahtar_AB", "Siklet_AC", "Siklet_derece_AD", "AE(AC-AD)",    # AA-AE
    "Irk_derece", "Cinsiyet", "Tarih_derece", "Sehir_derece", "Pist_derece",  # AF-AJ
    "PistDurumu_derece", "Mesafe_derece", "Derece", "AnahtarAN(Sehir+Irk+Pist+Mesafe)",  # AK-AN
    "AO_medyan_derece", "AP_derece_orani", "AQ(Derece/AO)", "AR(AQ*AP)",       # AO-AR
    "AS_birlesik", "AT_lookup", "At_No_AU", "KosuCinsi_derece",               # AS-AV
    "Yas", "KGS", "Irk+Cinsiyet", "Yil_derece", "Sira_derece",                # AW-BA
    "Yeni HP", "Eski HP", "Yeni Ort HP", "Yeni Ort Kilo", "Eski Ort HP",      # BB-BF
    "Eski Ort Kilo", "BH", "Tahmini_Derece(AQ)", "Tahmini_Derece(AR)",        # BG-BJ
    "Tahmini_son800(W)", "Tahmini_son800(X)", "Atin_Ilk3_HP",                 # BK-BM
    "Medyan HP Listesi", "Medyan Kilo Listesi",                               # BN-BO
    "BP", "BQ", "BR", "BS", "BT", "BU",                                       # BP-BU (boş)
    "BV(=Y)", "BW(=BK)", "BX_lookup", "BY_derece_Stil", "BZ_800_Stil",        # BV-BZ
]
assert len(YENI_YER_BASLIK) == 78, len(YENI_YER_BASLIK)


# ----- Excel hata sentinel'leri -----
class _ExcelError(str):
    pass

NA = _ExcelError("#N/A")
REF = _ExcelError("#REF!")
VALUE = _ExcelError("#VALUE!")
DIV0 = _ExcelError("#DIV/0!")


def _is_err(x):
    return isinstance(x, _ExcelError)


# ----------------------------------------------------------------------------
# VLOOKUP
# ----------------------------------------------------------------------------

def build_index(rows, key_col=1):
    """A sütunundaki ilk-görülen anahtarın satırına hızlı erişim indeksi.
    rows: list[list] (1-tabanlı kolonlar 0-index'e map'lenmiş list of lists,
          yani rows[i][0] = A sütunu)."""
    idx = {}
    for row in rows:
        if not row:
            continue
        k = row[key_col - 1]
        if k is None:
            continue
        kk = _key_norm(k)
        if kk not in idx:
            idx[kk] = row
    return idx


def _key_norm(k):
    """VLOOKUP eşleşmesi: Excel metinde büyük/küçük harf duyarsız, baştaki/sondaki
    boşlukları dikkate alır (almaz aslında trim etmez ama TJK verisinde sorun yok).
    Sayı/metin ayrımı: Excel '58' (metin) ile 58 (sayı) eşleşmez. Bunu da koruyoruz."""
    if isinstance(k, bool):
        return ("b", k)
    if isinstance(k, (int, float)):
        # tam sayıysa int gibi davran (Excel 58.0==58)
        f = float(k)
        if f.is_integer():
            return ("n", int(f))
        return ("n", f)
    return ("s", str(k).strip().upper())


def vlookup(key, index_map, col_index, range_width):
    """Excel VLOOKUP(key, table, col_index, 0) tam-eşleşme karşılığı.
    range_width = FORMÜLDE yazılı aralık genişliği (A:M->13, A:W->23, A:AG->33).
    REF kontrolü bu genişliğe göre yapılır (veri genişliğine göre DEĞİL)."""
    if _is_err(key):
        return key
    if col_index > range_width:
        return REF
    row = index_map.get(_key_norm(key))
    if row is None:
        return NA
    if col_index - 1 >= len(row):
        # Satır kısa ama aralık içi: Excel boş hücreyi 0 döndürür.
        return 0
    v = row[col_index - 1]
    return 0 if v is None else v


# ----------------------------------------------------------------------------
# Aritmetik / metin yardımcıları (Excel davranışına yakın)
# ----------------------------------------------------------------------------

def _num(x):
    """Excel'in sayıya zorlaması. Hata -> propagate. Boş -> 0. Metin sayı -> sayı."""
    if _is_err(x):
        return x
    if x is None or x == "":
        return 0
    if isinstance(x, bool):
        return 1 if x else 0
    if isinstance(x, (int, float)):
        # pandas boş hücre -> NaN; Excel'de boş hücre = 0. Sonsuz da -> 0.
        if isinstance(x, float) and (x != x or x in (float("inf"), float("-inf"))):
            return 0
        return x
    s = str(x).strip()
    if s.lower() == "nan":
        return 0
    if s == "":
        return 0
    # Excel ondalık: hücre metniyse nokta ondalık varsayımı
    try:
        return float(s.replace(",", "."))
    except Exception:
        return VALUE


def _sub(a, b):
    a, b = _num(a), _num(b)
    if _is_err(a):
        return a
    if _is_err(b):
        return b
    return a - b


def _div(a, b):
    a, b = _num(a), _num(b)
    if _is_err(a):
        return a
    if _is_err(b):
        return b
    if b == 0:
        return DIV0
    return a / b


def _mul(a, b):
    a, b = _num(a), _num(b)
    if _is_err(a):
        return a
    if _is_err(b):
        return b
    return a * b


def _concat(*args):
    """Excel CONCATENATE: hata varsa propagate; sayılar metne, ondalıkta nokta."""
    out = []
    for a in args:
        if _is_err(a):
            return a
        if a is None:
            out.append("")
        elif isinstance(a, bool):
            out.append("TRUE" if a else "FALSE")
        elif isinstance(a, float):
            if a != a:  # NaN
                out.append("")
            elif a.is_integer():
                out.append(str(int(a)))
            else:
                out.append(str(a))
        else:
            out.append(str(a))
    return "".join(out)


def _right(x, n):
    if _is_err(x):
        return x
    if x is None or (isinstance(x, float) and x != x):  # None / NaN
        s = ""
    elif isinstance(x, float) and x.is_integer():
        s = str(int(x))
    else:
        s = str(x)
    return s[-n:] if n > 0 else ""


def _int(x):
    x = _num(x)
    if _is_err(x):
        return x
    import math
    return math.floor(x)


def _mod(a, b):
    a, b = _num(a), _num(b)
    if _is_err(a):
        return a
    if _is_err(b):
        return b
    if b == 0:
        return DIV0
    return a - b * (a // b)


def _round(x, ndigits):
    x = _num(x)
    if _is_err(x):
        return x
    # Excel round half away from zero
    from decimal import Decimal, ROUND_HALF_UP
    q = Decimal(10) ** (-ndigits)
    d = Decimal(str(x)).quantize(q, rounding=ROUND_HALF_UP)
    f = float(d)
    return int(f) if ndigits <= 0 else f


def _year(x):
    if _is_err(x):
        return x
    import datetime as _dt
    if isinstance(x, (_dt.datetime, _dt.date)):
        return x.year
    # Excel seri tarihi (sayı) ise
    n = _num(x)
    if _is_err(n):
        return n
    if isinstance(n, (int, float)) and n > 0:
        base = _dt.date(1899, 12, 30)
        return (base + _dt.timedelta(days=int(n))).year
    return VALUE


def _excel_out(v):
    """Hücreye yazılacak nihai değer: hata sentinel'i string olarak kalır."""
    return v


# ----------------------------------------------------------------------------
# Süre/derece dönüştürme formülleri (BI, BJ, BK, BL)
# ----------------------------------------------------------------------------

def _time_fmt_6000(ratio):
    """BI/BJ:  INT((r*6684)/6000)&":"&RIGHT("00"&INT(MOD(r*6684,6000)/100),2)
              &","&RIGHT("00"&ROUND(MOD(r*6684,100),0),2)"""
    if _is_err(ratio):
        return ratio
    base = _mul(ratio, 6684)
    if _is_err(base):
        return base
    p1 = _int(_div(base, 6000))
    p2 = _right(_concat("00", _int(_div(_mod(base, 6000), 100))), 2)
    p3 = _right(_concat("00", _round(_mod(base, 100), 0)), 2)
    return _concat(p1, ":", p2, ",", p3)


def _derece_fmt_4691(ratio):
    """BK/BL:  INT(r*4691/100)&"."&RIGHT("00"&ROUND(MOD(r*4691,100),0),2)"""
    if _is_err(ratio):
        return ratio
    base = _mul(ratio, 4691)
    if _is_err(base):
        return base
    p1 = _int(_div(base, 100))
    p2 = _right(_concat("00", _round(_mod(base, 100), 0)), 2)
    return _concat(p1, ".", p2)


# ----------------------------------------------------------------------------
# ANA HESAP
# ----------------------------------------------------------------------------

def hesapla_yeni_yer(sayfa2_rows, rows_800, rows_derece, rows_medyanl):
    """
    Parametreler list[list] (1-tabanlı kolonlar 0-index): her satır bir liste.
      sayfa2_rows : Sayfa2 VERİ satırları (başlık hariç). [A..L]
      rows_800    : '800' satırları (başlık DAHİL, çünkü 800'de başlık yok-veri).
      rows_derece : 'derece' satırları.
      rows_medyanl: medyan tablosu satırları.
    Döndürür: list[list] -> 'yeni yer' veri satırları (A..BY = 77 kolon).
    """
    # Formüllerde yazılı aralık genişlikleri (REF kontrolü buna göre)
    R800 = 13      # '800'!A:M
    R800_STIL = 14 # '800'!A:N  (BZ = 800 Stil)
    RMED = 15      # 'medyanl'!A:O
    RDER_M = 13    # derece!A:M
    RDER_G = 7     # derece!A:G  (AH)
    RDER_N = 14    # derece!A:N  (BA)
    RDER_O = 15    # derece!A:O  (BC, BH)
    RDER_R = 18    # derece!A:R  (BG)
    RDER_V = 22    # derece!A:V  (BM)
    RDER_W = 23    # derece!A:W  (BF, BN, BO)
    RDER_AG = 33   # derece!A:AG (BY)

    idx_800 = build_index(rows_800, key_col=1)
    idx_der = build_index(rows_derece, key_col=1)
    idx_med = build_index(rows_medyanl, key_col=1)

    # Sayfa2'nin F:L aralığı için (BD/BE), anahtar = F kolonu (Anahtar)
    # Sayfa2 kolonları: A..L => index 1..12 ; F=6(Anahtar) ... L=12
    s2_FL_index = {}
    for r in sayfa2_rows:
        if len(r) >= 6 and r[5] is not None:
            k = _key_norm(r[5])
            if k not in s2_FL_index:
                s2_FL_index[k] = r
    W_S2FL = 12 - 6 + 1  # F..L = 7 kolon

    def vlookup_s2fl(key, col_in_FL):
        if _is_err(key):
            return key
        if col_in_FL > W_S2FL:
            return REF
        row = s2_FL_index.get(_key_norm(key))
        if row is None:
            return NA
        abs_col = 6 + (col_in_FL - 1)  # F=6 tabanlı
        if abs_col - 1 >= len(row):
            return ""
        v = row[abs_col - 1]
        return "" if v is None else v

    out_rows = []
    # AS->Y:Z ve AS->BV:BW self-lookup için iki geçiş: önce Y ve BK üret, sonra AT/BX.
    # Excel tek geçişte yapıyor ama Y/BK aynı satırda üretiliyor; AT/BX TÜM sayfada
    # arıyor. Bu yüzden önce tüm satırların ham kolonlarını üretip, sonra AT/BX'i
    # ikinci geçişte dolduruyoruz.

    interim = []  # her eleman: dict(col_letter->value) ham

    for s2 in sayfa2_rows:
        g = {}
        def S2(i):
            # Excel: =Sayfa2!X  -> boş hücre referansı 0 döner
            if i - 1 < len(s2) and s2[i - 1] is not None and s2[i - 1] != "":
                return s2[i - 1]
            return 0

        # --- A..F  Sayfa2'den ---
        g["A"] = S2(1)                  # =Sayfa2!A  Kosu_No
        g["B"] = S2(2)                  # =Sayfa2!B  At_No
        g["C"] = S2(3)                  # =Sayfa2!C  At_Adi_Temiz
        g["D"] = S2(4)                  # =Sayfa2!D  Siklet
        g["E"] = S2(6)                  # =Sayfa2!F  Anahtar
        g["F"] = S2(7)                  # =Sayfa2!G  Kosu_Basligi

        E = g["E"]
        # --- G..S  '800' lookuplari ---
        g["G"] = vlookup(E, idx_800, 3, R800)
        g["H"] = vlookup(E, idx_800, 4, R800)
        g["I"] = vlookup(E, idx_800, 5, R800)
        g["J"] = vlookup(E, idx_800, 6, R800)
        g["K"] = vlookup(E, idx_800, 7, R800)
        g["L"] = g["D"]
        g["M"] = vlookup(E, idx_800, 8, R800)
        g["N"] = _sub(g["L"], g["M"])
        g["O"] = vlookup(E, idx_800, 10, R800)
        g["P"] = vlookup(E, idx_800, 11, R800)
        g["Q"] = vlookup(E, idx_800, 12, R800)
        g["R"] = _sub(g["Q"], g["P"])
        g["S"] = vlookup(E, idx_800, 13, R800)
        g["T"] = _concat(g["H"], g["O"], g["I"], g["K"])
        g["U"] = vlookup(g["T"], idx_med, 8, RMED)
        g["V"] = vlookup(g["T"], idx_med, 9, RMED)
        g["W"] = _div(g["P"], g["U"])
        g["X"] = _mul(g["W"], g["V"])
        g["Y"] = _concat(g["C"], g["G"], g["T"], g["M"])
        g["Z"] = g["R"]
        g["AA"] = g["B"]
        g["AB"] = g["E"]
        g["AC"] = g["D"]
        AB = g["AB"]
        # --- derece lookuplari ---
        g["AD"] = vlookup(AB, idx_der, 4, RDER_M)
        g["AE"] = _sub(g["AC"], g["AD"])
        g["AF"] = vlookup(AB, idx_der, 5, RDER_M)
        # AG: cinsiyet, AW (yas metni) son harfine göre
        g["AW"] = S2(8)                 # =Sayfa2!H  Yas
        g["AX"] = S2(9)                 # =Sayfa2!I  KGS
        aw = g["AW"]
        if _is_err(aw):
            g["AG"] = aw
        else:
            aw_s = "" if aw is None else str(aw)
            if aw_s == "":
                g["AG"] = ""
            else:
                last = aw_s[-1:].lower()
                if last in ("a", "e", "g"):
                    g["AG"] = "Erkek"
                elif last in ("d", "k"):
                    g["AG"] = "Dişi"
                else:
                    g["AG"] = ""
        g["AH"] = vlookup(AB, idx_der, 7, RDER_G)      # Tarih (A:G)
        g["AI"] = vlookup(AB, idx_der, 8, RDER_M)
        g["AJ"] = vlookup(AB, idx_der, 9, RDER_M)
        g["AK"] = vlookup(AB, idx_der, 10, RDER_M)
        g["AL"] = vlookup(AB, idx_der, 11, RDER_M)
        g["AM"] = vlookup(AB, idx_der, 12, RDER_M)
        g["AN"] = _concat(g["AI"], g["AF"], g["AJ"], g["AL"])
        g["AO"] = vlookup(g["AN"], idx_med, 6, RMED)
        g["AP"] = vlookup(g["AN"], idx_med, 7, RMED)
        g["AQ"] = _div(g["AM"], g["AO"])
        g["AR"] = _mul(g["AQ"], g["AP"])
        g["AZ"] = _year(g["AH"])
        g["AS"] = _concat(g["C"], g["AZ"], g["AN"], g["AD"])
        # AT: ikinci geçişte (Y:Z self-lookup)
        g["AU"] = g["B"]
        g["AV"] = vlookup(AB, idx_der, 13, RDER_M)
        g["AY"] = _concat(g["AF"], g["AG"])
        g["BA"] = vlookup(AB, idx_der, 14, RDER_N)
        g["BB"] = S2(10)                # =Sayfa2!J  Handikap_Puani (Yeni HP)
        g["BC"] = vlookup(AB, idx_der, 15, RDER_O)     # Eski HP (A:O)
        g["BD"] = vlookup_s2fl(AB, 6)   # Sayfa2 F:L 6. = K = Kosu_HP_Tam_Ortanca
        g["BE"] = vlookup_s2fl(AB, 7)   # Sayfa2 F:L 7. = L = Kosu_Kilo_Tam_Ortanca
        g["BF"] = vlookup(AB, idx_der, 19, RDER_W)     # Eski Ort HP (A:W -> S = Koşu HP Medyan)
        # BG "Eski Ort Kilo": Excel formülü A:R(18) col20 -> #REF! (BUG).
        # Kullanıcı "formüllere bağlı kalma" dedi -> AMAÇLANAN değer: derece col T(20)
        # = Koşu HP Dolu Atlar Kilo Medyan.
        g["BG"] = vlookup(AB, idx_der, 20, RDER_W)
        # BH: başlıksız, Excel'de bozuk (A:O col16 -> #REF!). Boş bırakılıyor.
        g["BH"] = ""
        g["BI"] = _time_fmt_6000(g["AQ"])
        g["BJ"] = _time_fmt_6000(g["AR"])
        g["BK"] = _derece_fmt_4691(g["W"])
        g["BL"] = _derece_fmt_4691(g["X"])
        g["BM"] = vlookup(AB, idx_der, 16, RDER_V)     # A:V(22) -> 16 = P
        g["BN"] = vlookup(AB, idx_der, 21, RDER_W)     # A:W(23) -> 21 = U
        g["BO"] = vlookup(AB, idx_der, 22, RDER_W)     # A:W(23) -> 22 = V
        g["BV"] = g["Y"]
        g["BW"] = g["BK"]
        g["BY"] = vlookup(AB, idx_der, 28, RDER_AG)    # A:AG(33) -> 28 = AB (derece Stil)
        g["BZ"] = vlookup(E, idx_800, 14, R800_STIL)   # '800' 14. (N) = 800 Stil
        interim.append(g)

    # ----- ikinci geçiş: AT (Y:Z) ve BX (BV:BW) self-lookup -----
    # Y:Z  => anahtar Y, dön Z
    idx_YZ = {}
    for g in interim:
        k = _key_norm(g["Y"])
        if k not in idx_YZ:
            idx_YZ[k] = g["Z"]
    idx_BVBW = {}
    for g in interim:
        k = _key_norm(g["BV"])
        if k not in idx_BVBW:
            idx_BVBW[k] = g["BW"]

    for g in interim:
        AS = g["AS"]
        if _is_err(AS):
            g["AT"] = AS
            g["BX"] = AS
        else:
            g["AT"] = idx_YZ.get(_key_norm(AS), NA)
            g["BX"] = idx_BVBW.get(_key_norm(AS), NA)

    # ----- nihai satır listesine (A..BZ) çevir -----
    last_col = _col_idx("BZ")  # 78
    from openpyxl.utils import get_column_letter as GL
    for g in interim:
        row = []
        for c in range(1, last_col + 1):
            letter = GL(c)
            row.append(_excel_out(g.get(letter, "")))
        out_rows.append(row)

    return out_rows
