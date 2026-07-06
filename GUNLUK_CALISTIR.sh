#!/bin/bash
# =====================================================================
#  TJK GÜNLÜK ÇALIŞTIRICI  —  her şeyi sırayla yapar
# =====================================================================
#  Kullanım (yarışın tarihini ver):
#      bash GUNLUK_CALISTIR.sh 03.07.2026
#  Tarih vermezsen gun_ayar.py'deki HEDEF_TARIH neyse onu kullanır.
#
#  Sırayla: HEDEF_TARIH'i ayarla -> derece(+stil/üçgen) -> son800 ->
#           galop/son galop -> tjk_yeni_yer (birleştir).
#  Çıktı: yeni_yer_SONUC.xlsx
# =====================================================================

cd "$(dirname "$0")" || exit 1

# 1) Tarihi ayarla (verildiyse)
if [ -n "$1" ]; then
  python3 - "$1" <<'PY'
import sys, re
d = sys.argv[1].strip()
p = "gun_ayar.py"
s = open(p, encoding="utf-8").read()
# Sadece SATIR BAŞINDAKİ gerçek atamayı değiştir (yorumdaki örneği DEĞİL).
s2, n = re.subn(r'(?m)^HEDEF_TARIH\s*=\s*".*?"', f'HEDEF_TARIH = "{d}"', s)
if n == 0:
    raise SystemExit("HATA: gun_ayar.py içinde 'HEDEF_TARIH = \"...\"' satırı bulunamadı.")
open(p, "w", encoding="utf-8").write(s2)
print(f"[GÜN] HEDEF_TARIH = {d}  ({n} satır güncellendi)")
PY
else
  echo "[GÜN] Tarih verilmedi; gun_ayar.py'deki HEDEF_TARIH kullanılacak."
fi

echo ""
echo "================= 1/4  DERECE (+ stil / üçgen) ================="
caffeinate -i python3 ERTESI_GUN_MODU_OKA_BASAN_PLUS_FAZ30_STIL_TEK_PY.py

echo ""
echo "================= 2/4  SON 800 ================================"
caffeinate -i python3 son800_BUGUN_ILK_AT_STIL_FINAL_TEST.py

echo ""
echo "========= 2.5/4  SON 800  %45 STİL (ayrı video fazı) ========="
echo "  (OPTION B: derece 183 gün + son800 375 gün 1.'lik koşuların %45'i)"
caffeinate -i python3 son800_stil_ekle.py

echo ""
echo "================= 3/4  GALOP / SON GALOP ======================"
caffeinate -i python3 tjk_tum_akis_main_SELF_HEAL_KOSU_LINK_FIX_YARIN_TAM.py

echo ""
echo "================= 4/5  BİRLEŞTİR (yapılacak yer) =============="
caffeinate -i python3 tjk_yeni_yer.py

echo ""
echo "================= 5/5  WEB SAYFASI ============================"
if [ -d web ]; then
  (cd web && python3 parse_mockup.py ../yeni_yer_SONUC.xlsx && python3 prototip_uret.py) \
    && echo "web sayfası güncellendi: web/tjk_analiz_prototip.html" \
    || echo "UYARI: web sayfası üretilemedi (pipeline çıktısı etkilenmez)"
else
  echo "NOT: web klasörü yok, web adımı atlandı."
fi

echo ""
echo "=================  BİTTİ ✔  ->  yeni_yer_SONUC.xlsx  =========="
echo "Sayfalar: Sayfa1, Sayfa2, yeni yer, derece(+Stil Üçgen), 800, yapılacak yer, yapılacak yer 800"
echo "Web: web/tjk_analiz_prototip.html (çift tıkla aç)"
