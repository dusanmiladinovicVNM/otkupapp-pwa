from decimal import Decimal
from datetime import datetime
from agrix_mig15c_order_sensitivity import replay

def ev(name,day,ord,frm,to,q):
 return {'case':name,'date':datetime(2026,8,day).date(),'moment':datetime(2026,8,day),
 'ord':ord,'q':Decimal(q),'kind':'IZDATA_PRAZNA' if frm[0]=='Stanica' else 'AMBALAZA_UZ_ROBU',
 'typ':'Gajbica','from':frm,'to':to}
K=('Kooperant','K1');S=('Stanica','S1')
rows=[ev('one',1,1,S,K,'5'),ev('two',2,2,K,S,'8'),ev('three',3,3,S,K,'4')]
parts,balances,days=replay(rows)
assert parts['one']==(Decimal(0),Decimal(0))
assert parts['two']==(Decimal('3'),Decimal(0))
assert parts['three']==(Decimal(0),Decimal('3'))
assert balances[('kooperant','k1','gajbica')]==Decimal('1')
# Intra-day order can change which legacy event creates negative balance,
# but final signed balance remains identical.
rows=[ev('out',1,1,K,S,'8'),ev('in',1,2,S,K,'4')]
a,ba,_=replay(rows);b,bb,_=replay(rows,True)
assert ba==bb
assert a['out']!=(b['out'])
print('MIG15C synthetic tests PASS')
