#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

python -u agrix_mig13_import_novac.py \
  --stage "../MigrationStage_v2" \
  --input-xlsm "../MigrationGoods_MIG11_PATCHED_66f7897f/AgriX_MIG11_LINKS_PATCHED.xlsm" \
  --master "../AgriX_MASTER_COMPILED_66f7897f.xlsm" \
  --financial-jsonl "../MigrationMIG12_Archive_MIG13_Audit_66f7897f/MIG13_SECONDARY_tblNovac_SOURCE.jsonl" \
  --out-dir "../MigrationGoods_MIG13_66f7897f"
