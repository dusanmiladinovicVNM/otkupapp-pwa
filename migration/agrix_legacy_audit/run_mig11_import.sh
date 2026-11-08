#!/usr/bin/env bash
# AGRIx MIG-11 v1.1: full local migration step, no repo/VBA/original XLSM writes.
set -euo pipefail
cd "$HOME/Desktop/AgriX DEV/agrix_legacy_audit"

LINKS_DIR="../MigrationMIG11_Links_v1_1_66f7897f"
XLSM_DIR="../MigrationGoods_MIG11_v1_1_66f7897f"

python -u agrix_mig11_build_links.py \
  --stage "../MigrationStage_v2" \
  --out "$LINKS_DIR" \
  --stage-copy

# Run OOXML import ONLY if all source-specific membership expectations hold.
python - <<'PY'
import json
from pathlib import Path
p=Path('../MigrationMIG11_Links_v1_1_66f7897f/mig11_build_summary.json')
r=json.loads(p.read_text(encoding='utf-8'))
expected={'tblOtpremnicaIzvori':1692, 'tblZbirnaIzvori':907}
checks={
    'no unresolved FK or membership conflicts':not r['unresolved'],
    'snapshot was built':r['snapshot']=='stage_MIG11_ready',
    'links match audited source':r['links']==expected and r['new_link_rows']==expected,
    'both historical number discrepancies preserved':r['historical_number_mismatches_preserved_by_explicit_id']==2,
    '10 historic unlinked purchases preserved':r['historically_unlinked'].get('OPI_NO_HISTORICAL_OTP_REFERENCE')==10,
    'source fingerprints match prior audit':
        r['input_sha256'].get('SOURCE__tblOtkup')=='3035123964c6d77f436c8765b792fcd7781d4c04b172b1a97647a7a9bed72c7c' and
        r['input_sha256'].get('SOURCE__tblOtpremnica')=='d02436aa32a0e14ae6ca19172227f83b358b3ab0e329a7f040f97d29748e611a'
}
for k,v in checks.items():print(('PASS' if v else 'BLOCKED')+': '+k)
if not all(checks.values()):
    raise SystemExit('MIG-11 input/preflight not identical to audited dataset; OOXML import NOT run.')
print('MIG-11 LINKS PRECHECK PASS; starting existing MIG-09 OOXML writer')
PY

python -u agrix_mig09_preserve_vba.py \
  --stage "$LINKS_DIR/stage_MIG11_ready" \
  --master "../AgriX_MASTER_COMPILED_66f7897f.xlsm" \
  --preflight "../MigrationAmb10WorkbookPreflight_66f7897f" \
  --schema "$HOME/otkupapp-pwa/schema/schema.json" \
  --out "$XLSM_DIR" \
  --write-copy

# Verify actual XLSM: exact 2599 link rows + identical compiled VBA binary.
python -u agrix_mig11_verify_xlsm.py \
  --xlsm-dir "$XLSM_DIR" \
  --master "../AgriX_MASTER_COMPILED_66f7897f.xlsm" \
  --stage "$LINKS_DIR/stage_MIG11_ready" \
  --out "../MigrationMIG11_Verify_v1_1_66f7897f.json"
