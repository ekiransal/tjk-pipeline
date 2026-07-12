#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SEYİR ÖNBELLEK DÜZELTMESİ (HOLİGAN vakası, 26.01.2026):
%45 seyir pozisyonu ATA ÖZGÜ bir değerdir ama önbellek anahtarı sadece KOŞUYU
tanımlıyordu. Aynı geçmiş koşudan iki at bugünkü kadroda olunca ikincisi,
birincisinin pozisyonunu 'kütüphaneden' devralıyordu.
Düzeltme: stil anahtarına AT ADI eklenir (derece + son800, iki taraf birden).
Eski (atsız) kayıtlar artık eşleşmez -> ilk gece videolar yeniden okunur (uzun
sürer), sonraki geceler yine hızlıdır. Medyan önbelleğine DOKUNULMAZ.
Çalıştır (Mac'te): cd ~/Desktop/tjk && python3 stil_at_yama.py
"""
import os, sys

KOK = os.path.dirname(os.path.abspath(__file__))

# ---------- 1) DERECE tarafı ----------
D = os.path.join(KOK, "ERTESI_GUN_MODU_OKA_BASAN_PLUS_FAZ30_STIL_TEK_PY.py")
s = open(D, encoding="utf-8").read()
ESKI = '''    z = [tarih, sehir, mesafe, pist_comb, cins]
    if any(x == "" or x == "NAN" for x in z):
        return ""
    return "|".join(z)'''
YENI = '''    z = [tarih, sehir, mesafe, pist_comb, cins]
    at = _n(_g("At Adı", "At Adi", "At"))
    if any(x == "" or x == "NAN" for x in z) or at in ("", "NAN"):
        return ""
    # AT ADI ANAHTARDA: %45 pozisyonu ata özgüdür; koşu-bazlı anahtar, ayni
    # kosudan gelen IKINCI ata birincinin pozisyonunu kopyaliyordu (HOLIGAN vakasi).
    return "|".join(z) + "|AT:" + at'''
if "|AT:" in s:
    print("1) derece zaten yamalı")
elif ESKI in s:
    open(D, "w", encoding="utf-8").write(s.replace(ESKI, YENI, 1))
    print("1) derece: stil anahtarına AT ADI eklendi")
else:
    print("1) HATA: derece anchor bulunamadı!"); sys.exit(1)

# ---------- 2) SON800 tarafı ----------
S8 = os.path.join(KOK, "son800_stil_ekle.py")
s = open(S8, encoding="utf-8").read()
ESKI8 = '''            kimlik = temiz_str(row.get("Koşu Kimliği", ""))'''
YENI8 = '''            kimlik = temiz_str(row.get("Koşu Kimliği", ""))
            # AT ADI ANAHTARDA (derece tarafıyla aynı format: ...|AT:AD)
            import re as _re_at
            _atk = _re_at.sub(r"\\s+", " ", str(at_adi or "").strip().upper())
            kimlik = (kimlik + "|AT:" + _atk) if (kimlik and _atk) else ""'''
if "|AT:" in s:
    print("2) son800 zaten yamalı")
elif ESKI8 in s:
    open(S8, "w", encoding="utf-8").write(s.replace(ESKI8, YENI8, 1))
    print("2) son800: stil anahtarına AT ADI eklendi")
else:
    print("2) HATA: son800 anchor bulunamadı!"); sys.exit(1)

# ---------- 3) söz dizimi kontrolü ----------
import ast
for yol in (D, S8):
    ast.parse(open(yol, encoding="utf-8").read())
    print("   syntax OK:", os.path.basename(yol))

print("STIL_AT_YAMA_TAMAM — bu geceki çalışmadan itibaren seyirler ata özgü.")
