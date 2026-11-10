# Развёртывание ОСВ / БДР на сервере

Два способа установки: из git-репозитория (база начнётся пустой) или из
архива дистрибутива (внутри снимок PostgreSQL и загруженные Excel-файлы).

## Что нужно на сервере

- Linux (Ubuntu 22.04/24.04 или аналог), доступ `sudo`.
- Docker Engine 24+ и Docker Compose v2 (`docker compose`, не `docker-compose`).
- Свободные порты **80** и **443** (меняются в `.env`), а также локальный
  порт приложения `8010` (наружу не открывается).
- Доступ к Docker Hub для сборки образа и загрузки `postgres:16-alpine`.

### Ресурсы

| Вариант | CPU | RAM | Диск |
|---|---:|---:|---:|
| Минимум для внутреннего контура, 1–20 пользователей | 1 vCPU | 2 ГБ + swap 2 ГБ | 20 ГБ SSD |
| Рекомендуется | 2 vCPU | 4 ГБ | 40 ГБ SSD |
| Если рядом будет Keycloak, Vault-агент, SIEM/Kaspersky-агент | 4 vCPU | 8 ГБ | 60–80 ГБ SSD |

Основная нагрузка — импорт Excel и расчёт мэппинга; 2 vCPU / 4 ГБ — спокойный запас.

## Вариант 1 — установка из git

```bash
sudo git clone https://github.com/Alvik403/BDR-ALD.git /opt/osv-bdr
cd /opt/osv-bdr/deploy
sudo bash deploy.sh
```

Если репозиторий приватный, используйте токен GitHub или SSH-ключ:

```bash
sudo git clone https://<TOKEN>@github.com/Alvik403/BDR-ALD.git /opt/osv-bdr
# либо
sudo git clone git@github.com:Alvik403/BDR-ALD.git /opt/osv-bdr
```

При установке из git **база создаётся пустой**: приложение само создаёт
схему и добавляет шесть счетов при первом старте. Чтобы сразу загрузить
рабочие данные, положите снимок базы в `deploy/db.sql` **до первого запуска**
(файл подхватывается только на пустом томе PostgreSQL):

```bash
sudo cp /path/to/db.sql /opt/osv-bdr/deploy/db.sql
```

Обновление после изменений в репозитории:

```bash
cd /opt/osv-bdr
sudo git pull
cd deploy
sudo bash deploy.sh
```

## Вариант 2 — установка из архива

```bash
scp osv-bdr-deploy-2026-11-05.tar.gz user@server:/home/ubuntu/
sudo mkdir -p /opt/osv-bdr
sudo tar -xzf /home/ubuntu/osv-bdr-deploy-2026-11-05.tar.gz -C /opt/osv-bdr --strip-components=1
cd /opt/osv-bdr
sudo bash deploy.sh
```

В архиве уже есть `deploy/db.sql` (снимок на 05.11.2026: 8 загрузок,
8 735 строк ОСВ, 5 221 строка мэппинга) и `deploy/uploads/` с исходными
Excel-файлами. Они применяются автоматически при первом старте.

## Что делает deploy.sh

- создаёт `.env` из `.env.example` и генерирует случайный пароль PostgreSQL;
- если снимка базы нет — создаёт пустой `db.sql` (схема создастся приложением);
- создаёт самоподписанный TLS-сертификат, если нет корпоративного;
- собирает и поднимает `db`, `web`, `nginx`;
- печатает адреса для проверки.

## HTTP → HTTPS (80 и 443)

- порт **80** — только редирект `301` на `https://…`;
- порт **443** — TLS и проксирование в приложение (`client_max_body_size 30m`,
  чтобы проходил Excel до 25 МБ);
- приложение (`APP_PORT`, по умолчанию 8010) слушает только `127.0.0.1`
  сервера — снаружи доступно исключительно через 443.

Сертификат по умолчанию самоподписанный: `deploy/certs/fullchain.pem` и
`deploy/certs/privkey.pem`, имя из `TLS_CN` в `.env`. Браузер будет
предупреждать. Корпоративный сертификат: положите эти два файла в
`deploy/certs/` и перезапустите `sudo bash deploy.sh`.

Если 80/443 заняты — измените `HTTP_PORT`/`HTTPS_PORT` в `.env`.

## Проверка

```bash
curl -I  http://127.0.0.1/healthz     # HTTP/1.1 301, Location: https://…
curl -k https://127.0.0.1/healthz     # {"status":"ok"}
```

Сайт: `https://<IP-сервера или DNS-имя>/`.

## Данные, бэкапы, секреты

- `.env` — пароль PostgreSQL и порты; в git не попадает.
- `deploy/uploads/` — загруженные Excel-файлы.
- `deploy/db.sql` — снимок базы только для первого старта пустого тома.
- `backup.sh` — дамп базы и архив файлов:

```bash
cd /opt/osv-bdr/deploy
sudo bash backup.sh
```

Файлы появятся в `deploy/backups/` (в git и архив не входят).
