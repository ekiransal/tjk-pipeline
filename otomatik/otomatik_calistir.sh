#!/bin/zsh
# GANYANRADAR OTOMATIK GUNLUK KOSU - her aksam launchd tarafindan calistirilir.
# Ertesi gunun verisini ceker, siteyi yayinlar, ozellikleri kaydeder,
# gunun sonuclarini esler, arsivler. Log: ~/Desktop/tjk/loglar/
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
export LANG=tr_TR.UTF-8

TJK="$HOME/Desktop/tjk"
LOGDIR="$TJK/loglar"
mkdir -p "$LOGDIR"
# HEDEF GUN: normalde YARIN (aksam 18:00 calismasi). Ama Mac 18:00'de uyuyup
# gorev SABAH uyaninca telafi olarak calisirsa (saat < 16) hedef BUGUN olmali;
# yoksa canli sayfa 1 gun ileri kayar (12.07 vakasi, 11.07.2026).
if [ "$(date +%H)" -lt 16 ]; then
  TARIH=$(date +%d.%m.%Y)              # sabah telafi calismasi: BUGUN
else
  TARIH=$(date -v+1d +%d.%m.%Y)        # normal aksam calismasi: YARIN
fi
LOG="$LOGDIR/otomatik_$(date +%Y-%m-%d_%H%M).log"

bildir() { osascript -e "display notification \"$1\" with title \"GanyanRadar\"" 2>/dev/null; }

{
  echo "================================================="
  echo "OTOMATIK BASLADI : $(date)"
  echo "HEDEF GUN        : $TARIH"
  echo "================================================="

  cd "$TJK" || { echo "HATA: tjk klasoru yok"; exit 1; }

  caffeinate -i bash GUNLUK_CALISTIR.sh "$TARIH"
  DURUM=$?
  echo "--- GUNLUK_CALISTIR bitti (kod: $DURUM) ---"

  cd "$TJK/web" || exit 1
  python3 parse_mockup.py ../yeni_yer_SONUC.xlsx \
    && python3 prototip_uret.py \
    && python3 ozellik_kaydet.py
  python3 birlestir.py                 # sonuc yoksa da devam
  scp tjk_analiz_prototip.html root@185.111.235.75:/opt/ganyanradar/analiz/tjk_analiz_prototip.html \
    && ssh root@185.111.235.75 "python3 /opt/ganyanradar/arsivle.py" \
    && echo "HEPSI_TAMAM"

  echo "================================================="
  echo "OTOMATIK BITTI   : $(date)"
  echo "================================================="
} >> "$LOG" 2>&1

if grep -q "HEPSI_TAMAM" "$LOG"; then
  bildir "✅ $TARIH verisi hazir, site yayinda."
else
  bildir "⚠️ Gunluk kosu HATALI bitti — loglar klasorune bak."
fi
