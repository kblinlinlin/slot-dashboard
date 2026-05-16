#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
HOST="${HOST:-0.0.0.0}"
PORT="${PORT:-8765}"
PID_FILE="${PID_FILE:-$APP_DIR/.dashboard-server.pid}"
LOG_FILE="${LOG_FILE:-$APP_DIR/dashboard-server.log}"

cd "$APP_DIR"

if [[ ! -f "$APP_DIR/index.html" ]]; then
  echo "ERROR: index.html not found in $APP_DIR" >&2
  exit 1
fi

if [[ ! -d "$APP_DIR/data" ]]; then
  echo "ERROR: data directory not found in $APP_DIR" >&2
  exit 1
fi

if command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
  PYTHON_BIN="python"
else
  echo "ERROR: python3 or python is required to start this static site." >&2
  exit 1
fi

if [[ -f "$PID_FILE" ]]; then
  OLD_PID="$(cat "$PID_FILE" 2>/dev/null || true)"
  if [[ -n "$OLD_PID" ]] && kill -0 "$OLD_PID" 2>/dev/null; then
    echo "Dashboard server is already running."
    echo "PID: $OLD_PID"
    echo "URL: http://$HOST:$PORT/"
    exit 0
  fi
  rm -f "$PID_FILE"
fi

nohup "$PYTHON_BIN" -m http.server "$PORT" --bind "$HOST" >"$LOG_FILE" 2>&1 &
SERVER_PID="$!"
disown "$SERVER_PID" 2>/dev/null || true
echo "$SERVER_PID" > "$PID_FILE"

sleep 1

if ! kill -0 "$SERVER_PID" 2>/dev/null; then
  echo "ERROR: failed to start dashboard server. See log: $LOG_FILE" >&2
  rm -f "$PID_FILE"
  exit 1
fi

echo "Dashboard server started."
echo "PID: $SERVER_PID"
echo "Directory: $APP_DIR"
echo "URL: http://$HOST:$PORT/"
echo "Log: $LOG_FILE"
echo "Stop: kill $SERVER_PID && rm -f '$PID_FILE'"
