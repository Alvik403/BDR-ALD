#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
set -a
# shellcheck disable=SC1091
. ./.env
set +a

stamp="$(date +%Y%m%d_%H%M%S)"
mkdir -p backups

docker compose -f docker-compose.deploy.yml exec -T db \
  pg_dump -U "${POSTGRES_USER:-osv}" -d "${POSTGRES_DB:-osv}" \
  --no-owner --no-privileges > "backups/db_${stamp}.sql"

tar -czf "backups/uploads_${stamp}.tar.gz" -C deploy/uploads .

echo "backups/db_${stamp}.sql"
echo "backups/uploads_${stamp}.tar.gz"
