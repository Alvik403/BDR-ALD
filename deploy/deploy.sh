#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

# Две раскладки:
#  - архив дистрибутива: deploy.sh рядом с docker-compose, данные в ./deploy/;
#  - git-клон: deploy.sh лежит в deploy/, данные рядом с ним.
if [ -d deploy ]; then
  DATA_DIR="deploy"
else
  DATA_DIR="."
fi
CERT_DIR="${DATA_DIR}/certs"

if [ ! -f .env ]; then
  cp .env.example .env
  if command -v openssl >/dev/null 2>&1; then
    password="$(openssl rand -hex 16)"
  elif [ -r /dev/urandom ]; then
    password="$(head -c 16 /dev/urandom | od -An -tx1 | tr -d ' \n')"
  else
    password="$(python3 -c 'import secrets; print(secrets.token_hex(16))')"
  fi
  sed -i "s|^POSTGRES_PASSWORD=.*|POSTGRES_PASSWORD=${password}|" .env
  echo "Создан .env, пароль PostgreSQL сгенерирован автоматически."
fi

# Снимка базы может не быть (git-клон) — тогда база начнётся пустой,
# схему приложение создаст само при первом старте.
if [ ! -f "${DATA_DIR}/db.sql" ]; then
  touch "${DATA_DIR}/db.sql"
  echo "Снимок базы не найден — база начнётся пустой."
fi
mkdir -p "${DATA_DIR}/uploads"

# Самоподписанный сертификат, если корпоративного нет.
# Чтобы использовать свой: положите fullchain.pem и privkey.pem в ${CERT_DIR}.
if [ ! -f "${CERT_DIR}/fullchain.pem" ] || [ ! -f "${CERT_DIR}/privkey.pem" ]; then
  mkdir -p "${CERT_DIR}"
  cn="$(grep -E '^TLS_CN=' .env | tail -1 | cut -d= -f2 || true)"
  cn="${cn:-bdr.alabuga.ru}"
  openssl req -x509 -nodes -newkey rsa:2048 -days 825 \
    -keyout "${CERT_DIR}/privkey.pem" \
    -out "${CERT_DIR}/fullchain.pem" \
    -subj "/CN=${cn}"
  echo "Создан самоподписанный сертификат для ${cn} (${CERT_DIR}/)."
  echo "Браузер покажет предупреждение; корпоративный сертификат можно подложить туда же и перезапустить."
fi

docker compose -f docker-compose.deploy.yml up -d --build
docker compose -f docker-compose.deploy.yml ps

http_port="$(grep -E '^HTTP_PORT=' .env | tail -1 | cut -d= -f2 || true)"
http_port="${http_port:-80}"
https_port="$(grep -E '^HTTPS_PORT=' .env | tail -1 | cut -d= -f2 || true)"
https_port="${https_port:-443}"
app_port="$(grep -E '^APP_PORT=' .env | tail -1 | cut -d= -f2 || true)"
app_port="${app_port:-8010}"
echo
echo "Редирект:  curl -I http://127.0.0.1:${http_port}/healthz   → 301 на https"
echo "Проверка:  curl -k https://127.0.0.1:${https_port}/healthz"
echo "Сайт:      https://<IP-сервера или имя>:${https_port}"
echo "Отладка (только на сервере): http://127.0.0.1:${app_port}"
