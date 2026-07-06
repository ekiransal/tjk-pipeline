#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJK — YENİ VERİ EKLE (2026 güncelle)  •  TEK KOMUT
==================================================

Akış (idempotent — ne zaman çalıştırırsan çalıştır, aynı sonuç, ASLA çift saymaz):

  1) Koşu Sorgulama'dan TÜM 2026'yı çek (01.01.2026 -> bugün) -> kosu_2026_tam.xlsx
     (ORİJİN: baba/ana, mesafe, kalite)
  2) Son800Ist'ten TÜM 2026'yı çek -> son800_2026.xlsx
     (800-FARK: kaçak/sprinter için)
  3) fark'ı orijin satırlarına EŞLEŞTİR (at|şehir|mesafe|zemin) -> kosu_2026_tam.xlsx'e
     'fark' kolonu eklenir.
  4) Kütüphanede 2026'yı SIFIRLA, yeni tam 2026'yı (orijin + fark) yükle
     (doğu illeri otomatik elenir).
  5) Panel referans sayfalarını (kısa uzun / KALİTE / siprinter / KAÇAK)
     güncel oranlarla yeniden yaz.

Kullanım:
  python3 tjk_yeni_ekle.py                 # TAM akış (kosu + son800 + fark + yükle + panel)
  python3 tjk_yeni_ekle.py 02.07.2026      # bitiş tarihini elle ver
  python3 tjk_yeni_ekle.py --kosu-atla     # KOŞU'yu tekrar çekme (mevcut kosu_2026_tam.xlsx),
                                           #   sadece son800 çek + fark eşleştir + yükle + panel
  python3 tjk_yeni_ekle.py --sadece-yukle  # HİÇ çekme (mevcut xlsx'leri eşleştir + yükle + panel)
  python3 tjk_yeni_ekle.py --fark-atla     # sadece orijin (son800/fark yok)

ÇİFT SAYMA OLMAZ: her çalıştırmada kütüphaneden 2026 önce SİLİNİR, sonra yüklenir.

NOT: 2007-2025 verisi hiç değişmez. 2026 her seferinde sıfırlanıp baştan kurulur.
"""

import os
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
KOSU_2026 = os.path.join(HERE, "kosu_2026_tam.xlsx")
SON800_2026 = os.path.join(HERE, "son800_2026.xlsx")
PANEL = os.path.join(HERE, "YENI_ORJINLER_BABA_DEDE_PANEL_4LU.xlsx")


def _bugun():
    return datetime.now().strftime("%d.%m.%Y")


def main():
    args = [a for a in sys.argv[1:]]
    sadece_yukle = "--sadece-yukle" in args
    fark_atla = "--fark-atla" in args
    kosu_atla = "--kosu-atla" in args
    args = [a for a in args if not a.startswith("--")]
    bit = args[0] if args else _bugun()
    bas = "01.01.2026"
    yil = "2026"

    import tjk_orjin_kutuphane as K
    import tjk_orjin_panel_yaz as P

    # ------------------------------------------------ 1) KOŞU SORGULAMA (orijin)
    if not sadece_yukle and not kosu_atla:
        import tjk_kosu_sorgulama as S
        print("=" * 60)
        print(f"1) KOŞU SORGULAMA (orijin) ÇEKİLİYOR:  {bas}  ->  {bit}")
        print("=" * 60)
        S.cek(bas, bit, KOSU_2026, headless=False)
    else:
        neden = "--sadece-yukle" if sadece_yukle else "--kosu-atla"
        print(f"1) KOŞU ÇEKME ATLANDI ({neden}). Mevcut kullanılacak: {KOSU_2026}")

    if not os.path.exists(KOSU_2026):
        print(f"HATA: koşu dosyası bulunamadı: {KOSU_2026}")
        raise SystemExit(1)

    # ------------------------------------------------ 2) SON800 (fark) + EŞLEŞTİR
    if not fark_atla:
        if not sadece_yukle:
            import tjk_son800_cek as S8
            print("=" * 60)
            print(f"2) SON800 (800-fark) ÇEKİLİYOR: {yil}")
            print("=" * 60)
            S8.cek(yil, SON800_2026, headless=False)
        if os.path.exists(SON800_2026):
            import tjk_fark_esle as F
            print("=" * 60)
            print("2b) FARK eşleştiriliyor (Son800 -> orijin satırları)")
            print("=" * 60)
            F.fark_ekle(KOSU_2026, SON800_2026, KOSU_2026)
        else:
            print(f"UYARI: son800 dosyası yok ({SON800_2026}); fark eklenmeden devam.")
    else:
        print("2) SON800/FARK ATLANDI (--fark-atla). Sadece orijin yüklenecek.")

    # -------------------------------------------------- 3) 2026 SIFIRLA + YÜKLE
    print("=" * 60)
    print("3) KÜTÜPHANE: 2026 sıfırlanıyor + yeni veri yükleniyor")
    print("=" * 60)
    conn = K.baglan()
    onceki = K.sayac(conn)
    silinen = conn.execute("DELETE FROM kazanan WHERE tarih LIKE '%2026%'").rowcount
    conn.commit()
    print(f"  Silinen eski 2026 satırı: {silinen}")
    K.ham_excel_yukle(conn, KOSU_2026)
    print(f"  Kütüphane toplam: {onceki} -> {K.sayac(conn)}")

    # ------------------------------------------------------------- 4) PANEL YAZ
    print("=" * 60)
    print("4) PANEL referans sayfaları güncelleniyor")
    print("=" * 60)
    if not os.path.exists(PANEL):
        print(f"  HATA: şablon panel yok: {PANEL}")
        print("  (Panel dosyası aynı klasörde olmalı; ilk üretim yapılmışsa vardır.)")
        raise SystemExit(1)
    P.yaz(conn, PANEL, PANEL)

    print("=" * 60)
    print("BİTTİ ✔  Panel güncel:  " + os.path.basename(PANEL))
    print("=" * 60)


if __name__ == "__main__":
    main()
