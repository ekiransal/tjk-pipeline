#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, re, shutil
KOK = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(KOK, "analiz", "tjk_analiz_prototip.html")
ARS = os.path.join(KOK, "analiz", "arsiv")
os.makedirs(ARS, exist_ok=True)
h = open(SRC, encoding="utf-8").read()
m = re.search(r"\"hedef\":\s*\"(\d{2})\.(\d{2})\.(\d{4})\"", h)
if m:
    ad = "{2}-{1}-{0}.html".format(m.group(1), m.group(2), m.group(3))
    shutil.copy(SRC, os.path.join(ARS, ad))
    print("ARSIVLENDI:", ad)
else:
    print("hedef tarih bulunamadi, arsivlenmedi")
