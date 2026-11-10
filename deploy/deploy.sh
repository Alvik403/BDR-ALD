#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

# Из git-клона снимка базы нет: база начнётся пустой (схему приложение
# создаст само при первом старте). В архиве дистрибутива рядом лежит db.sql
# с данными — тогда он подхватится автоматически.
if [ ! -f db.sql ]; then
  touch db.sql
  echo "Снимок базы db.sql не найден — база начнётся пустой."
fi
mkdir -p uploads

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

# Самоподписанный сертификат, если корпоративного нет.
# Чтобы использовать свой: положите fullchain.pem и privkey.pem в deploy/certs/.
if [ ! -f deploy/certs/fullchain.pem ] || [ ! -f deploy/certs/privkey.pem ]; then
  mkdir -p deploy/certs
  cn="$(grep -E '^TLS_CN=' .env | tail -1 | cut -d= -f2 || true)"
  cn="${cn:-bdr.alabuga.ru}"
  openssl req -x509 -nodes -newkey rsa:2048 -days 825 \
    -keyout deploy/certs/privkey.pem \
    -out deploy/certs/fullchain.pem \
    -subj "/CN=${cn}"
  echo "Создан самоподписанный сертификат для ${cn} (deploy/certs/)."
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
