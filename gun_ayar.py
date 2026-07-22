# =====================================================================
#  TJK SİSTEMİ — HEDEF GÜN AYARI  (TEK YER)
# =====================================================================
#  Bütün scriptler (derece, 800, galop, pipeline) bu dosyadan okur.
#
#  İKİ YOL:
#
#  1) SABİT TARİH (ÖNERİLEN, uzun çekimler için):
#        HEDEF_TARIH = "03.07.2026"
#     -> Hangi script ne zaman çalışırsa çalışsın AYNI güne kilitlenir.
#        derece'yi akşam 17:00 başlatıp son800'ü ertesi sabah çalıştırsan
#        bile ikisi aynı günü hedefler. GÜN KAYMASI OLMAZ.
#
#  2) OFFSET (eski yöntem):  HEDEF_TARIH = ""  bırak.
#        0=bugün, 1=yarın, 2=2 gün sonra ...
#     DİKKAT: OFFSET, her scriptin ÇALIŞTIĞI ANDAKİ "bugün"e göre hesaplar.
#     derece dün 17:00 (bugün+1 = A günü), son800 bugün sabah (bugün+1 = B günü)
#     çalışırsa günler KAYAR ve derece/800 birbirini tutmaz. 14 saat süren
#     çekimlerde MUTLAKA sabit tarih kullan.
# =====================================================================

# Hedef yarış günü (GG.AA.YYYY). Boş "" bırakırsan OFFSET kullanılır.
HEDEF_TARIH = "22.07.2026"

# HEDEF_TARIH boşsa kullanılır:
OFFSET = 1


def hedef_gun():
    """Tüm bileşenlerin ortak hedef günü.
    HEDEF_TARIH doluysa o sabit tarih, değilse bugün+OFFSET. -> datetime (00:00)."""
    from datetime import datetime, timedelta
    t = str(HEDEF_TARIH or "").strip()
    if t:
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(t, fmt)
            except Exception:
                pass
    return datetime.now() + timedelta(days=OFFSET)
