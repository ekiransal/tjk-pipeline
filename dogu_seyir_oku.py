#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DOGU SEYIR OKUYUCU - dogu illeri video karesindeki beyaz kutulu
numara seridini okur (soldan saga = o anki siralama, 1-3 kutu, 1-2 hane).
Kutular: karenin x %18-62, y %70-90 bandinda, banner ustunde.
Kullanim (tek kare):  oku_kare(cv2_goruntu) -> [1, 3, 6] gibi liste
"""
import re

import cv2
import numpy as np

try:
    import pytesseract
    OCR_VAR = True
except Exception:
    OCR_VAR = False

# serit arama bandi (kare oranlari)
BAND_X = (0.14, 0.68)
BAND_Y = (0.68, 0.92)


def _kutulari_bul(kare):
    H, W = kare.shape[:2]
    x1, x2 = int(W * BAND_X[0]), int(W * BAND_X[1])
    y1, y2 = int(H * BAND_Y[0]), int(H * BAND_Y[1])
    bolge = kare[y1:y2, x1:x2]
    gri = cv2.cvtColor(bolge, cv2.COLOR_BGR2GRAY)
    _, maske = cv2.threshold(gri, 190, 255, cv2.THRESH_BINARY)
    maske = cv2.morphologyEx(maske, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8))
    konturlar, _ = cv2.findContours(maske, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    kutular = []
    for c in konturlar:
        x, y, w, h = cv2.boundingRect(c)
        if not (W * 0.022 <= w <= W * 0.10):
            continue
        if not (H * 0.035 <= h <= H * 0.12):
            continue
        oran = w / float(h)
        if not (0.55 <= oran <= 2.4):
            continue
        # kutu ici gercekten beyaz agirlikli mi (dolu beyaz blok)
        ic = maske[y:y + h, x:x + w]
        if ic.mean() < 140:
            continue
        kutular.append((x, y, w, h))
    if not kutular:
        return bolge, []
    # ayni satirda duranlari sec (y merkezleri en kalabalik kumede)
    kutular.sort(key=lambda k: k[0])
    ymerk = [y + h / 2.0 for _, y, _, h in [(k[0], k[1], k[2], k[3]) for k in kutular]]
    en_iyi = []
    for i, (x, y, w, h) in enumerate(kutular):
        grup = [k for j, k in enumerate(kutular) if abs(ymerk[j] - ymerk[i]) < h * 0.6]
        if len(grup) > len(en_iyi):
            en_iyi = grup
    en_iyi.sort(key=lambda k: k[0])
    return bolge, en_iyi[:3]


def _rakam_oku(bolge, kutu):
    if not OCR_VAR:
        return None
    x, y, w, h = kutu
    p = max(2, h // 10)
    crop = bolge[max(0, y - p):y + h + p, max(0, x - p):x + w + p]
    crop = cv2.resize(crop, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
    gri = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
    _, ikili = cv2.threshold(gri, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    for psm in (7, 10, 6):
        s = pytesseract.image_to_string(
            ikili, config="--psm %d -c tessedit_char_whitelist=0123456789" % psm)
        m = re.search(r"\d{1,2}", s or "")
        if m:
            n = int(m.group(0))
            if 1 <= n <= 20:
                return n
    return None


def oku_kare(kare):
    """kare (BGR) -> okunan numaralar listesi (soldan saga). Okunamayan kutu atlanir."""
    bolge, kutular = _kutulari_bul(kare)
    sonuc = []
    for k in kutular:
        n = _rakam_oku(bolge, k)
        if n is not None:
            sonuc.append(n)
    return sonuc


if __name__ == "__main__":
    import sys, glob, os
    yollar = sys.argv[1:] or sorted(glob.glob(os.path.expanduser("~/Desktop/dogu_video_kareleri/*.jpg")))
    for yol in yollar:
        img = cv2.imread(yol)
        if img is None:
            print(os.path.basename(yol), ": okunamadi")
            continue
        print("%-28s -> %s" % (os.path.basename(yol), oku_kare(img) or "KUTU BULUNAMADI"))
