from __future__ import annotations
import tempfile
from pathlib import Path
from decimal import Decimal
from agrix_mig16_daily_close import compute_daily_policy

D=Decimal

def e(day,from_,to,q,kind='IZDATA_PRAZNA',typ='Gajbica'):
    return {'case':day+str(q), 'day':__import__('datetime').date.fromisoformat(day),'from':from_,'to':to,'q':D(q),'kind':kind,'typ':typ}

def test_all():
    k=('Kooperant','K1');s=('Stanica','S1')
    rows=[e('2026-08-01',k,s,'8','POVRAT_PRAZNE'),
          e('2026-08-02',s,k,'5'),
          e('2026-08-03',s,k,'10'),
          # same-day issue+return: daily closing only matters
          e('2026-08-04',s,k,'3'),
          e('2026-08-04',k,s,'3','POVRAT_PRAZNE')]
    d,stats=compute_daily_policy(rows)
    assert len(d)==4
    assert [x['ending_signed_balance'] for x in d]==['-8','-3','7','7']
    assert [x['historical_obligation_at_close'] for x in d]==['8','3','0','0']
    assert [x['obligation_increase'] for x in d]==['8','0','0','0']
    assert [x['obligation_decrease'] for x in d]==['0','5','3','0']
    assert stats['days_with_increased_historical_obligation']==1
    assert stats['days_with_decreased_historical_obligation']==2
    rev=list(reversed(rows))
    d2,s2=compute_daily_policy(rev)
    assert d==d2 and stats==s2
    # A debt falling due to something OTHER than partner issuance is not
    # represented as a foreign return in the canonical physical ledger.
    problem=[e('2026-08-01',k,s,'8','POVRAT_PRAZNE'),e('2026-08-02',s,k,'5','AMBALAZA_UZ_ROBU')]
    dg,st=compute_daily_policy(problem)
    assert dg[1]['not_assignable_decrease']=='5'
    assert st['days_obligation_reduction_without_matching_issuance']==1
    # Distinct packaging types and different kooperants never net together.
    other=[e('2026-08-01',k,s,'8','POVRAT_PRAZNE','Gajbica'),
           e('2026-08-01',s,k,'8','IZDATA_PRAZNA','Paleta')]
    dg,st=compute_daily_policy(other)
    assert len(dg)==2 and st['distinct_kooperant_type_pairs']==2
    print('MIG16 SYNTHETIC DAILY CLOSE TESTS PASS')

if __name__=='__main__':test_all()
