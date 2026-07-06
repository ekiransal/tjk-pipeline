#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sayfa6 VLOOKUP katmanı  (orjin/dede analiz verisi)
==================================================
Girdi : Sayfa2 (Kosu_No?/At_No/At_Adi/Baba(orjin)/Kosu_Basligi)
        'ÇEKİLECEK YER'  (baba performans DB)
        'kısa orta uzun' (mesafe DB)
Çıktı : Sayfa6 satırları (AnalizCalistir'ın okuduğu ORAN/MESAFE/YAPISI/... kolonları)

Tüm hesap = baba adını (Sayfa2 D) referanslarda aratan VLOOKUP'lar (tam eşleşme).
"""

NA = "#YOK"


def _key(v):
    if v is None:
        return None
    return str(v).strip().upper()


def _vindex(rows, key_col):
    """key_col (1-tabanlı) -> {anahtar: satır}. İlk görülen kazanır."""
    idx = {}
    for row in rows:
        if key_col - 1 < len(row):
            k = _key(row[key_col - 1])
            if k is not None and k != "" and k not in idx:
                idx[k] = row
    return idx


def _vlookup(idx, key, abs_col):
    """idx üzerinden tam eşleşme; abs_col (1-tabanlı) değeri. Yoksa #YOK."""
    row = idx.get(_key(key))
    if row is None:
        return NA
    if abs_col - 1 < len(row):
        v = row[abs_col - 1]
        return NA if v is None else v
    return NA


def _div(a, b):
    try:
        fa = float(a); fb = float(b)
        if fb == 0:
            return "#SAYI/0!"
        return fa / fb
    except Exception:
        return NA


def hesapla_sayfa6(sayfa2_rows, cekilecek_rows, kou_rows):
    """sayfa2_rows: başlık HARİÇ veri satırları. -> sayfa6 satırları (list[list], A.. )."""
    # ÇEKİLECEK YER aralık indeksleri (start_col = aralığın ilk kolonu)
    CY = cekilecek_rows
    idxA  = _vindex(CY, 1)    # A:E
    idxK  = _vindex(CY, 11)   # K:O
    idxP  = _vindex(CY, 16)   # P:T
    idxU  = _vindex(CY, 21)   # U:X
    idxAB = _vindex(CY, 28)   # AB:AF
    idxAO = _vindex(CY, 41)   # AO:AR
    idxAS = _vindex(CY, 45)   # AS:AZ
    idxBF = _vindex(CY, 58)   # BF:BP
    idxKOU = _vindex(kou_rows, 1)  # kısa orta uzun A:J

    out = []
    for s2 in sayfa2_rows:
        def S2(i):
            return s2[i - 1] if i - 1 < len(s2) else None
        g = [None] * 61   # A..BI (1..61)

        def setg(col, v):
            g[col - 1] = v

        # A..D Sayfa2'den (A = Kosu_Basligi=Sayfa2!E ; B=At_No ; C=At_Adi ; D=Baba)
        setg(1, S2(5))
        setg(2, S2(2))
        setg(3, S2(3))
        setg(4, S2(4))
        D = g[3]
        # E,F = ÇEKİLECEK YER A:E 3,4
        setg(5, _vlookup(idxA, D, 3))   # E
        setg(6, _vlookup(idxA, D, 4))   # F
        setg(7, _div(g[6 - 1], g[5 - 1]))   # G = F/E
        # H=B, I=D
        setg(8, g[2 - 1])
        setg(9, g[4 - 1])
        I = g[9 - 1]
        # J..R = kısa orta uzun A:J 2..10
        for off, col in zip(range(2, 11), range(10, 19)):
            setg(col, _vlookup(idxKOU, I, off))
        # S=H, T=I
        setg(19, g[8 - 1])
        setg(20, g[9 - 1])
        T = g[20 - 1]
        # U,V,W = K:O 3,4,5
        setg(21, _vlookup(idxK, T, 11 + 2))
        setg(22, _vlookup(idxK, T, 11 + 3))
        setg(23, _vlookup(idxK, T, 11 + 4))
        # X=S, Y=T
        setg(24, g[19 - 1])
        setg(25, g[20 - 1])
        Y = g[25 - 1]
        # Z,AA,AB = P:T 3,4,5
        setg(26, _vlookup(idxP, Y, 16 + 2))
        setg(27, _vlookup(idxP, Y, 16 + 3))
        setg(28, _vlookup(idxP, Y, 16 + 4))
        # AC=X, AD=Y
        setg(29, g[24 - 1])
        setg(30, g[25 - 1])
        AD = g[30 - 1]
        # AE,AF,AG = AB:AF 3,4,5
        setg(31, _vlookup(idxAB, AD, 28 + 2))
        setg(32, _vlookup(idxAB, AD, 28 + 3))
        setg(33, _vlookup(idxAB, AD, 28 + 4))
        # AH=AC, AI=AD
        setg(34, g[29 - 1])
        setg(35, g[30 - 1])
        AI = g[35 - 1]
        # AJ..AO = AS:AZ 3..8
        for off, col in zip(range(3, 9), range(36, 42)):
            setg(col, _vlookup(idxAS, AI, 45 + off - 1))
        # AP=AH, AQ=AI
        setg(42, g[34 - 1])
        setg(43, g[35 - 1])
        AQ = g[43 - 1]
        # AR,AT,AU,AW,AX,AZ = BF:BP 3,5,6,8,9,11
        setg(44, _vlookup(idxBF, AQ, 58 + 2))   # AR off3
        setg(46, _vlookup(idxBF, AQ, 58 + 4))   # AT off5
        setg(47, _vlookup(idxBF, AQ, 58 + 5))   # AU off6
        setg(49, _vlookup(idxBF, AQ, 58 + 7))   # AW off8
        setg(50, _vlookup(idxBF, AQ, 58 + 8))   # AX off9
        setg(52, _vlookup(idxBF, AQ, 58 + 10))  # AZ off11
        # BA=AP, BB=AQ
        setg(53, g[42 - 1])
        setg(54, g[43 - 1])
        BB = g[54 - 1]
        # BC,BD,BE = AO:AR 2,3,4
        setg(55, _vlookup(idxAO, BB, 41 + 1))
        setg(56, _vlookup(idxAO, BB, 41 + 2))
        setg(57, _vlookup(idxAO, BB, 41 + 3))
        # BF=BA, BG=BB
        setg(58, g[53 - 1])
        setg(59, g[54 - 1])
        BG = g[59 - 1]
        # BH,BI = U:X 3,4
        setg(60, _vlookup(idxU, BG, 21 + 2))
        setg(61, _vlookup(idxU, BG, 21 + 3))

        out.append(g)
    return out
