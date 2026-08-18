# -*- coding: utf-8 -*-
"""Command-line runner for the ERP <-> Tally reconciliation.

Usage:
    python reco.py                       # uses Book1.xlsx
    python reco.py <workbook.xlsx> [out.xlsx] [tally_sheet] [erp_sheet]
"""
import sys
import reco_engine as E

src = sys.argv[1] if len(sys.argv) > 1 else 'Book1.xlsx'
out = sys.argv[2] if len(sys.argv) > 2 else 'ERP_vs_Tally_Reconciliation.xlsx'
tsh = sys.argv[3] if len(sys.argv) > 3 else None
esh = sys.argv[4] if len(sys.argv) > 4 else None

R = E.reconcile(src, tsh, esh)
m = E.money

print('=' * 78)
print('ERP vs TALLY  -  %s' % src)
print('Tally sheet: %s   |   ERP sheet: %s' % (R['tally_sheet'], R['erp_sheet']))
print('=' * 78)
print('LEDGER CONTROL')
def dr(x):
    return abs(x) if x < 0 else 0.0


def cr(x):
    return x if x > 0 else 0.0


print('%-24s %11s %11s %11s %11s' % ('', 'Tally Dr', 'Tally Cr', 'ERP Dr', 'ERP Cr'))
print('%-24s %11s %11s %11s %11s' % ('1. Opening balance', m(dr(R['t_open'])),
                                     m(cr(R['t_open'])), m(dr(R['e_open'])),
                                     m(cr(R['e_open']))))
print('%-24s %11s %11s %11s %11s' % ('2. Total DEBIT', m(R['t_debit']), '-',
                                     m(R['e_debit']), '-'))
print('%-24s %11s %11s %11s %11s' % ('3. Total CREDIT', '-', m(R['t_credit']),
                                     '-', m(R['e_credit'])))
print('%-24s %11s %11s %11s %11s' % ('4. Closing balance', m(dr(R['t_close'])),
                                     m(cr(R['t_close'])), m(dr(R['e_close'])),
                                     m(cr(R['e_close']))))
print('%-24s %11s' % ('   Closing difference', m(R['gap'])))
print('%-30s %16s %16s' % ('   Control check (=0)',
                           m(R['control'][0][1]), m(R['control'][1][1])))
print('-' * 78)
for side in ('Debit', 'Credit'):
    v = R['sides'][side]
    print('%s SIDE' % side.upper())
    for cat, lab in (('matched', 'Matched'), ('unmatched', 'Mismatched'),
                     ('total', 'Total')):
        print('   %-27s %16s %16s' % (lab, m(v['tally_' + cat]), m(v['erp_' + cat])))
print('-' * 78)
for k in E.STATUS_ORDER:
    if k in R['buckets']:
        c, ta, ea = R['buckets'][k]
        print('%-56s %3d  T:%15s E:%15s' % (E.LABEL[k], c, m(ta), m(ea)))
print('-' * 78)
print('BALANCE BRIDGE')
print('%-58s %18s' % ('Tally closing balance', m(R['t_close'])))
for k, v in sorted(R['bridge'].items(), key=lambda x: -abs(x[1])):
    print('%-58s %18s' % ('  ' + k[:56], m(v)))
print('%-58s %18s' % ('= ERP closing balance', m(R['e_close'])))
print('%-58s %18s' % ('  UNEXPLAINED RESIDUAL', m(R['residual'])))
print('-' * 78)
print('Data-quality flags: %d' % len(R['dq']))
for w, v, i in R['dq']:
    print('  %-18s %-26s %s' % (w, v, i))
print('=' * 78)
E.write_report(R, out)
print('Report written to %s' % out)
if R['residual'] != 0:
    sys.exit(1)
