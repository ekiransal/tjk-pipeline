#!/bin/zsh
# GANYANRADAR OTOMATIK - her aksam: YARINI TAZELE + OBUR GUNU IHTIYATEN CEK.
# Ikisi de yayinlanip arsive yazilir; site varsayilani zaten "o anki gun"u acar.
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
export LANG=tr_TR.UTF-8
TJK="$HOME/Desktop/tjk"
LOGDIR="$TJK/loglar"; mkdir -p "$LOGDIR"
# Normal aksam kosusu: yarin + obur gun. Sabah telafi (saat<16): bugun + yarin.
if [ "$(date +%H)" -lt 16 ]; then
  GUNLER=("$(date +%d.%m.%Y)" "$(date -v+1d +%d.%m.%Y)")
else
  GUNLER=("$(date -v+1d +%d.%m.%Y)" "$(date -v+2d +%d.%m.%Y)")
fi
LOG="$LOGDIR/otomatik_$(date +%Y-%m-%d_%H%M).log"
bildir() { osascript -e "display notification \"$1\" with title \"GanyanRadar\"" 2>/dev/null; }
{
  echo "================================================="
  echo "OTOMATIK BASLADI : $(date)"
  echo "GUNLER           : ${GUNLER[@]}"
  echo "================================================="
  cd "$TJK" || exit 1
  for TARIH in "${GUNLER[@]}"; do
    echo ""; echo "########## $TARIH ##########"
    caffeinate -i bash GUNLUK_CALISTIR.sh "$TARIH"
    echo "--- GUNLUK_CALISTIR bitti ($TARIH, kod: $?) ---"
    cd "$TJK/web" || exit 1
    python3 parse_mockup.py ../yeni_yer_SONUC.xlsx \
      && python3 prototip_uret.py \
      && python3 ozellik_kaydet.py
    python3 birlestir.py
    scp tjk_analiz_prototip.html root@185.111.235.75:/opt/ganyanradar/analiz/tjk_analiz_prototip.html \
      && ssh root@185.111.235.75 "python3 /opt/ganyanradar/arsivle.py" \
      && echo "GUN_TAMAM:$TARIH"
    cd "$TJK"
  done
  echo "================================================="
  echo "OTOMATIK BITTI   : $(date)"
  echo "================================================="
} >> "$LOG" 2>&1
T1=$(grep -c "GUN_TAMAM" "$LOG")
if [ "$T1" -ge 2 ]; then
  bildir "✅ ${GUNLER[1]} tazelendi + ${GUNLER[2]} ihtiyata alindi."
elif [ "$T1" -eq 1 ]; then
  bildir "⚠️ Ilk gun tamam, ikinci gun HATALI - loglara bak."
else
  bildir "⚠️ Gunluk kosu HATALI bitti - loglara bak."
fi
