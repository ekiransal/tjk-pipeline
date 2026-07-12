#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SADECE BUGÜN İÇİN (12.07.2026): dünkü üretim, stil düzeltmesinden ÖNCE yapıldığı
için "aynı geçmiş koşudan iki at" durumlarında seyir yanlış kopyalanmış olabilir
(HOLİGAN/SERDAREFE vakası). Hangisinin doğru olduğu veriden bilinemeyeceğinden bu
çiftlerin seyir üçgenleri BUGÜNLÜK boşaltılır (yanlış bilgi > bilgi yok).
Yarından itibaren gerek yok — gece çekimi ata özgü anahtarla doğru üretecek.
Kullanım: web/ içinde, parse_mockup'tan SONRA, prototip_uret'ten ÖNCE:
    python3 bugun_seyir_duzelt.py
"""
import json, re
from collections import defaultdict

SEYIR = {"Sayfa1": (33, 34), "Sayfa2": (25, 26)}   # 0-tabanlı seyir hücreleri
TARIH_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")

d = json.load(open("mockup_parsed.json", encoding="utf-8"))
bosaltilan = 0
for rol, (sa, sb) in SEYIR.items():
    for b in d.get(rol, []):
        gruplar = defaultdict(set)   # (tarih, şehir) -> {at adları}
        satirlar = defaultdict(list)
        for r in b.get("detay", []):
            if str(r[41] if len(r) > 41 else "") == "GECMIS_YOK":
                continue
            at = re.sub(r"\d+$", "", str(r[1] or "").strip().upper()).strip()
            tarih = next((str(c).strip() for c in r if TARIH_RE.match(str(c or "").strip())), "")
            if not at or not tarih:
                continue
            sehir = ""
            for c in r:
                s = str(c or "").strip()
                if s and not TARIH_RE.match(s) and s.upper() in (
                        "İSTANBUL", "ANKARA", "İZMİR", "BURSA", "ADANA", "KOCAELİ",
                        "ELAZIĞ", "ŞANLIURFA", "DİYARBAKIR", "ANTALYA", "ISTANBUL", "IZMIR",
                        "KOCAELI", "ELAZIG", "SANLIURFA", "DIYARBAKIR"):
                    sehir = s.upper(); break
            anah = (tarih, sehir)
            gruplar[anah].add(at)
            satirlar[anah].append(r)
        for anah, atlar in gruplar.items():
            if len(atlar) >= 2:   # aynı geçmiş koşudan 2+ FARKLI at -> seyir şüpheli
                for r in satirlar[anah]:
                    for c in (sa, sb):
                        if len(r) > c and str(r[c] or "").strip():
                            r[c] = ""
                            bosaltilan += 1

json.dump(d, open("mockup_parsed.json", "w", encoding="utf-8"), ensure_ascii=False)
print(f"BUGUN_SEYIR_DUZELTILDI — {bosaltilan} şüpheli seyir hücresi boşaltıldı")
