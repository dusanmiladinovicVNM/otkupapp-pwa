#!/usr/bin/env bash
# This run does NOT repeat MIG-09 or MIG-11 staging conversion.
# Input XLSM, compiled master, VBA, legacy data and original stage are READ-ONLY.
set -euo pipefail
cd "$HOME/Desktop/AgriX DEV/agrix_legacy_audit"

python -u agrix_mig11_write_links_ooxml.py \
  --input-dir "../MigrationGoods_MIG11_v1_1_66f7897f" \
  --master "../AgriX_MASTER_COMPILED_66f7897f.xlsm" \
  --stage "../MigrationMIG11_Links_v1_1_66f7897f/stage_MIG11_ready" \
  --out-dir "../MigrationGoods_MIG11_PATCHED_66f7897f"
