#!/bin/zsh
LOGDIR="$HOME/Desktop/tjk/loglar"
BUGUN=$(date +%Y-%m-%d)
pgrep -f otomatik_calistir.sh >/dev/null && exit 0
cnt=$(find "$LOGDIR" -name "otomatik_${BUGUN}_1[89]*.log" -o -name "otomatik_${BUGUN}_2*.log" 2>/dev/null | wc -l)
[ "$cnt" -gt 0 ] && exit 0
osascript -e 'display notification "18:00 gorevi kacti - BEKCI baslatti" with title "GanyanRadar"' 2>/dev/null
nohup /bin/zsh "$HOME/Desktop/tjk/otomatik/otomatik_calistir.sh" >/dev/null 2>&1 &
