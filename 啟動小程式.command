#!/bin/bash
# 在這台 Mac 上啟動 IE 人力需求小程式。
# 用 http://localhost 開啟，才能安裝成 App、離線使用。用完把這個視窗關掉即可。
cd "$(dirname "$0")" || exit 1
PORT=8768
PY=""
for c in python3 "$HOME/anaconda3/bin/python3" "$HOME/opt/anaconda3/bin/python3" /opt/anaconda3/bin/python3; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "找不到 Python。請改成直接點兩下 index.html 開啟（功能一樣，只是不能安裝成 App）。"
  read -r -p "按 Enter 關閉"
  exit 1
fi
echo "IE 人力需求小程式：http://localhost:$PORT"
echo "用完之後關掉這個視窗即可。"
( sleep 1; open "http://localhost:$PORT/" ) &
exec "$PY" -m http.server "$PORT" --bind 127.0.0.1
