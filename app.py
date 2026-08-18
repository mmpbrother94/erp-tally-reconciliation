# -*- coding: utf-8 -*-
"""
ERP vs Tally Reconciliation - local web app.

Run:      python app.py
Then open http://127.0.0.1:5000 in a browser.

Upload a workbook containing a Tally ledger sheet and an ERP ledger sheet.
The sheets are auto-detected; you can override them and re-run.
"""
import os
import uuid
import glob
import time
import traceback

from flask import (Flask, request, render_template_string, send_file,
                   redirect, url_for, abort)
from werkzeug.utils import secure_filename

import reco_engine as E

BASE = os.path.dirname(os.path.abspath(__file__))
UPLOADS = os.path.join(BASE, 'uploads')
REPORTS = os.path.join(BASE, 'reports')
for p in (UPLOADS, REPORTS):
    os.makedirs(p, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024   # 32 MB


def ensure_dirs():
    for p in (UPLOADS, REPORTS):
        os.makedirs(p, exist_ok=True)


def purge(older_than_hours=24):
    """Keep the working folders from growing without bound."""
    ensure_dirs()
    cutoff = time.time() - older_than_hours * 3600
    for folder in (UPLOADS, REPORTS):
        for f in glob.glob(os.path.join(folder, '*')):
            try:
                if os.path.getmtime(f) < cutoff:
                    os.remove(f)
            except OSError:
                pass


# --------------------------------------------------------------------------
CSS = """
:root{
  --bg:#f4f6f9; --card:#ffffff; --ink:#16202c; --muted:#65707d; --line:#dfe4ea;
  --brand:#1f4e79; --brand2:#2d6ca8;
  --ok:#1f7a43; --okbg:#e6f4ec; --warn:#8a6100; --warnbg:#fdf3d8;
  --bad:#a32b2b; --badbg:#fbe9e9; --info:#1d5c8f; --infobg:#e4f0fa;
  --nil:#555; --nilbg:#eeeeee;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font:14px/1.55 "Segoe UI",system-ui,-apple-system,sans-serif}
a{color:var(--brand2)}
.wrap{max-width:1240px;margin:0 auto;padding:26px 20px 70px}
header.top{background:var(--brand);color:#fff;padding:18px 0}
header.top .wrap{padding:0 20px}
header.top h1{margin:0;font-size:19px;letter-spacing:.2px}
header.top p{margin:3px 0 0;opacity:.82;font-size:12.5px}
.card{background:var(--card);border:1px solid var(--line);border-radius:9px;
  padding:20px;margin-bottom:18px}
.card h2{margin:0 0 14px;font-size:15px;color:var(--brand);
  text-transform:uppercase;letter-spacing:.6px}
.drop{border:2px dashed #b9c4d0;border-radius:9px;padding:38px 20px;text-align:center;
  background:#fbfcfd;transition:.15s}
.drop.hot{border-color:var(--brand2);background:#eef5fb}
.drop p{margin:8px 0;color:var(--muted)}
input[type=file]{margin:10px 0}
button,.btn{background:var(--brand);color:#fff;border:0;border-radius:6px;
  padding:10px 20px;font-size:14px;cursor:pointer;text-decoration:none;
  display:inline-block}
button:hover,.btn:hover{background:var(--brand2)}
.btn.alt{background:#fff;color:var(--brand);border:1px solid var(--brand)}
select{padding:7px 9px;border:1px solid var(--line);border-radius:5px;font-size:13.5px;
  background:#fff}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(215px,1fr));gap:14px}
.kpi{background:#fff;border:1px solid var(--line);border-left:4px solid var(--brand);
  border-radius:7px;padding:14px 16px}
.kpi .lab{font-size:11.5px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)}
.kpi .val{font-size:21px;font-weight:600;margin-top:5px;font-variant-numeric:tabular-nums}
.kpi.good{border-left-color:var(--ok)} .kpi.good .val{color:var(--ok)}
.kpi.bad{border-left-color:var(--bad)} .kpi.bad .val{color:var(--bad)}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;
  vertical-align:top}
th{background:#eef2f6;font-weight:600;font-size:11.5px;text-transform:uppercase;
  letter-spacing:.4px;color:#3d4a58;position:sticky;top:0}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
tr:hover td{background:#fafbfc}
.scroll{max-height:560px;overflow:auto;border:1px solid var(--line);border-radius:7px}
.tag{display:inline-block;padding:2px 8px;border-radius:11px;font-size:11px;
  font-weight:600;white-space:nowrap}
.MATCHED{background:var(--okbg);color:var(--ok)}
.MATCHED_NIL{background:var(--nilbg);color:var(--nil)}
.DIFF_TDS,.DIFF_GST,.DIFF_DIRECTION{background:var(--warnbg);color:var(--warn)}
.DIFF_OPEN{background:var(--badbg);color:var(--bad)}
.TALLY_ONLY{background:var(--badbg);color:var(--bad)}
.ERP_ONLY{background:var(--infobg);color:var(--info)}
.neg{color:var(--bad)}
.sm{font-size:12px;color:var(--muted)}
.bridge td:first-child{padding-left:22px}
.bridge tr.h td{font-weight:700;padding-left:10px;background:#f2f5f8}
.err{background:var(--badbg);border:1px solid #e6b8b8;color:var(--bad);
  padding:14px 16px;border-radius:7px;white-space:pre-wrap;font-size:13px}
.note{color:var(--muted);font-size:12.5px;margin:10px 0 0}
.pills{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:12px}
.pill{border:1px solid var(--line);background:#fff;border-radius:16px;padding:5px 13px;
  font-size:12.5px;cursor:pointer;user-select:none}
.pill.on{background:var(--brand);color:#fff;border-color:var(--brand)}
.row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}
.ok-banner{background:var(--okbg);border:1px solid #b6ddc6;color:var(--ok);
  padding:11px 15px;border-radius:7px;font-weight:600}
.bad-banner{background:var(--badbg);border:1px solid #e6b8b8;color:var(--bad);
  padding:11px 15px;border-radius:7px;font-weight:600}
"""

SHELL = """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ERP vs Tally Reconciliation</title><style>""" + CSS + """</style></head><body>
<header class="top"><div class="wrap">
  <h1>ERP &nbsp;vs&nbsp; Tally &mdash; Supplier Ledger Reconciliation</h1>
  <p>Upload a workbook, get a matched / unmatched breakdown and a balance bridge that must tie to zero.</p>
</div></header><div class="wrap">{{body|safe}}</div></body></html>"""

UPLOAD_BODY = """
<div class="card">
  <h2>Upload workbook</h2>
  <form method="post" action="/run" enctype="multipart/form-data" id="f">
    <div class="drop" id="drop">
      <p><strong>Drag an .xlsx / .xlsm file here, or choose one below</strong></p>
      <input type="file" name="file" id="file" accept=".xlsx,.xlsm,.xls" required>
      <p>The workbook needs one Tally ledger sheet and one ERP ledger sheet.<br>
         They are detected automatically &mdash; you can override them after the first run.</p>
    </div>
    <p style="margin-top:16px"><button type="submit">Reconcile</button></p>
  </form>
</div>
<div class="card">
  <h2>What it checks</h2>
  <table>
    <tr><th>Parameter</th><th>How it is used</th></tr>
    <tr><td><strong>Direction (Dr / Cr)</strong></td><td>Both ledgers are converted to one signed number from the supplier's point of view, so <em>+ = payable up</em>, <em>&minus; = payable down</em>. A value that matches but sits on the wrong side is reported as a direction error, not left unmatched.</td></tr>
    <tr><td><strong>Document number</strong></td><td>Reduced to <code>number/FY</code>, so <code>EEPLB/006/25-26</code>, <code>EEPL/006/25-26</code> and <code>EEPL2/006/25-26</code> all match. Survives series typos.</td></tr>
    <tr><td><strong>Amount</strong></td><td>Tolerance of &plusmn; Rs 1.00, so Tally 10.46 against ERP 11.00 counts as matched. Bank entries are the exception &mdash; they must agree exactly.</td></tr>
    <tr><td><strong>Purchase</strong></td><td>Must match on the invoice / document reference. No invoice number means no match.</td></tr>
    <tr><td><strong>Bank</strong></td><td>Date + amount in both systems, exact, with the receipt/payment wording cross-checked.</td></tr>
    <tr><td><strong>Date</strong></td><td>Used as a fallback window (45 days) for payments, because ERP books them under its own reference while Tally quotes the invoice number.</td></tr>
    <tr><td><strong>TDS</strong></td><td>If a document's gap equals the TDS legs on the Tally voucher, it is reported as <em>ERP booked GROSS, Tally booked NET of TDS</em> &mdash; not as a mismatch.</td></tr>
    <tr><td><strong>Entry nature</strong></td><td>Read from the contra ledger, not the voucher type &mdash; so journals are split into real transactions (service bills) vs settlements (bill discounting).</td></tr>
  </table>
</div>
<script>
var dz=document.getElementById('drop'),fi=document.getElementById('file');
['dragenter','dragover'].forEach(function(e){dz.addEventListener(e,function(ev){
  ev.preventDefault();dz.classList.add('hot');});});
['dragleave','drop'].forEach(function(e){dz.addEventListener(e,function(ev){
  ev.preventDefault();dz.classList.remove('hot');});});
dz.addEventListener('drop',function(ev){if(ev.dataTransfer.files.length){
  fi.files=ev.dataTransfer.files;}});
</script>
"""

RESULT_BODY = """
<div class="card">
  <div class="row" style="justify-content:space-between">
    <div>
      <h2 style="margin:0">Result &mdash; {{fname}}</h2>
      <p class="note">Tally sheet: <strong>{{R.tally_sheet}}</strong> &nbsp;|&nbsp;
         ERP sheet: <strong>{{R.erp_sheet}}</strong> &nbsp;|&nbsp;
         {{R.tally|length}} Tally vouchers vs {{R.erp|length}} ERP lines</p>
    </div>
    <div class="row">
      <a class="btn" href="/report/{{token}}">Download Excel report</a>
      <a class="btn alt" href="/">New file</a>
    </div>
  </div>
  <form method="post" action="/run" style="margin-top:14px" class="row">
    <input type="hidden" name="token" value="{{token}}">
    <span class="note">Wrong sheets picked?</span>
    <label class="note">Tally
      <select name="tally_sheet">
        {% for s in sheets %}<option {% if s==R.tally_sheet %}selected{% endif %}>{{s}}</option>{% endfor %}
      </select></label>
    <label class="note">ERP
      <select name="erp_sheet">
        {% for s in sheets %}<option {% if s==R.erp_sheet %}selected{% endif %}>{{s}}</option>{% endfor %}
      </select></label>
    <button type="submit">Re-run</button>
  </form>
</div>

<div class="kpis">
  <div class="kpi"><div class="lab">Tally closing balance</div>
    <div class="val">{{ money(R.t_close) }}</div></div>
  <div class="kpi"><div class="lab">ERP closing balance</div>
    <div class="val">{{ money(R.e_close) }}</div></div>
  <div class="kpi bad"><div class="lab">Difference to explain</div>
    <div class="val">{{ money(R.gap) }}</div></div>
  <div class="kpi {{ 'good' if R.residual==0 else 'bad' }}">
    <div class="lab">Unexplained residual</div>
    <div class="val">{{ money(R.residual) }}</div></div>
</div>

<div class="card" style="margin-top:18px">
  {% if R.residual == 0 %}
    <div class="ok-banner">Reconciliation is complete &mdash; every rupee of the
      {{ money(R.gap) }} gap is classified below and the bridge ties to zero.</div>
  {% else %}
    <div class="bad-banner">{{ money(R.residual) }} is still unexplained. Do not
      sign off &mdash; check the Exceptions list.</div>
  {% endif %}
</div>


<div class="card">
  <h2>Ledger control comparison</h2>
  <table>
    <tr><th rowspan="2">&nbsp;</th><th colspan="2" style="text-align:center">Tally</th>
        <th colspan="2" style="text-align:center">ERP</th><th rowspan="2">Note</th></tr>
    <tr><th class="n">Debit</th><th class="n">Credit</th>
        <th class="n">Debit</th><th class="n">Credit</th></tr>
    <tr><td><strong>1. Opening balance</strong></td>
        <td class="n">{{ money(dr(R.t_open)) }}</td><td class="n">{{ money(cr(R.t_open)) }}</td>
        <td class="n">{{ money(dr(R.e_open)) }}</td><td class="n">{{ money(cr(R.e_open)) }}</td>
        <td class="sm">From the export header; 0.00 when not stated</td></tr>
    <tr><td><strong>2. Total DEBIT</strong></td>
        <td class="n">{{ money(R.t_debit) }}</td><td class="n">&mdash;</td>
        <td class="n">{{ money(R.e_debit) }}</td><td class="n">&mdash;</td>
        <td class="sm">Payments, TDS, bill discounting, bank settlements</td></tr>
    <tr><td><strong>3. Total CREDIT</strong></td>
        <td class="n">&mdash;</td><td class="n">{{ money(R.t_credit) }}</td>
        <td class="n">&mdash;</td><td class="n">{{ money(R.e_credit) }}</td>
        <td class="sm">Purchases and other supplier charges</td></tr>
    <tr style="background:#f2f5f8"><td><strong>4. Closing balance</strong></td>
        <td class="n"><strong>{{ money(dr(R.t_close)) }}</strong></td>
        <td class="n"><strong>{{ money(cr(R.t_close)) }}</strong></td>
        <td class="n"><strong>{{ money(dr(R.e_close)) }}</strong></td>
        <td class="n"><strong>{{ money(cr(R.e_close)) }}</strong></td>
        <td class="sm">Opening + Credit &minus; Debit</td></tr>
    <tr><td class="sm">Control check (must be 0.00)</td>
        <td class="n sm" colspan="2">{{ money(R.control[0][1]) }}</td>
        <td class="n sm" colspan="2">{{ money(R.control[1][1]) }}</td>
        <td class="sm">Opening + Credit &minus; Debit &minus; Closing</td></tr>
    <tr><td><strong>Closing difference (ERP &minus; Tally)</strong></td>
        <td class="n" colspan="4"><strong>{{ money(R.gap) }}</strong></td>
        <td class="sm">The amount this reconciliation has to explain</td></tr>
  </table>
</div>

<div class="card">
  <h2>Matching rules applied</h2>
  <table>
    <tr><th>Entry type</th><th>Rule</th></tr>
    <tr><td><strong>Purchase</strong></td><td>Must match on the invoice / document
        reference. A purchase with no invoice number is reported as unmatchable.</td></tr>
    <tr><td><strong>Bank</strong></td><td>Date + amount, <strong>exact</strong> &mdash;
        no tolerance is applied to bank entries. Date window {{ bank_win }} days,
        and the receipt/payment wording is cross-checked.</td></tr>
    <tr><td><strong>TDS</strong></td><td>Every TDS booked in Tally is searched for on
        the <strong>ERP Debit side only</strong>. Found on the credit side instead,
        it is reported as a direction error.</td></tr>
    <tr><td><strong>All others</strong></td><td>Amount within <strong>&plusmn; Rs
        {{ tol }}</strong> (so Tally 10.46 vs ERP 11.00 counts as matched), date
        within {{ win }} days, then the reference number is confirmed.</td></tr>
    <tr><td><strong>Unmatched</strong></td><td>Anything without a counterpart goes
        straight to Mismatched.</td></tr>
  </table>
</div>

{% for side in ['Debit','Credit'] %}
{% set v = R.sides[side] %}
<div class="card">
  <h2>{{ side }} side &mdash; matched vs mismatched</h2>
  <table style="margin-bottom:16px">
    <tr><th>&nbsp;</th><th class="n">Tally</th><th class="n">ERP</th>
        <th class="n">Difference</th><th>Meaning</th></tr>
    <tr><td><span class="tag MATCHED">Matched</span></td>
        <td class="n">{{ money(v.tally_matched) }}</td>
        <td class="n">{{ money(v.erp_matched) }}</td>
        <td class="n">{{ money(v.erp_matched - v.tally_matched) }}</td>
        <td class="sm">Counterpart found and values agree</td></tr>
    <tr><td><span class="tag TALLY_ONLY">Mismatched</span></td>
        <td class="n">{{ money(v.tally_unmatched) }}</td>
        <td class="n">{{ money(v.erp_unmatched) }}</td>
        <td class="n">{{ money(v.erp_unmatched - v.tally_unmatched) }}</td>
        <td class="sm">No counterpart, or values disagree</td></tr>
    <tr style="background:#f2f5f8"><td><strong>Total {{ side }}</strong></td>
        <td class="n"><strong>{{ money(v.tally_total) }}</strong></td>
        <td class="n"><strong>{{ money(v.erp_total) }}</strong></td>
        <td class="n"><strong>{{ money(v.diff) }}</strong></td>
        <td class="sm">Ties to line {{ '2' if side=='Debit' else '3' }} above</td></tr>
  </table>

  <div class="pills" data-for="t{{side}}">
    <span class="pill on" data-f="ALL">All</span>
    <span class="pill" data-f="matched">Matched</span>
    <span class="pill" data-f="unmatched">Mismatched</span>
    <span class="pill" data-f="Tally">Tally rows</span>
    <span class="pill" data-f="ERP">ERP rows</span>
  </div>
  <div class="scroll"><table id="t{{side}}">
    <tr><th>Result</th><th>Status</th><th>System</th><th>Date</th><th>Reference</th>
        <th>Doc key</th><th>Description</th><th class="n">{{ side }} amount</th>
        <th>Counterpart in other system</th><th class="n">Doc diff</th>
        <th>Criteria checked</th><th>Reason</th></tr>
    {% for x in R.ledger if x.side == side %}
    <tr data-c="{{x.cat}}" data-y="{{x.system}}">
      <td><span class="tag {{ 'MATCHED' if x.cat=='matched' else 'TALLY_ONLY' }}">{{ x.cat|title }}</span>
          {% if x.ambiguous %}<br><span class="tag DIFF_TDS">Ambiguous</span>{% endif %}</td>
      <td><span class="tag {{x.status}}">{{x.status}}</span></td>
      <td>{{x.system}}</td><td>{{x.date}}</td><td>{{x.ref}}</td><td>{{x.key}}</td>
      <td>{{x.desc}}</td><td class="n">{{ money(x.amount) }}</td>
      <td class="sm">{{x.counterpart}}</td>
      <td class="n {{ 'neg' if x.diff<0 }}">{{ money(x.diff) }}</td>
      <td class="sm">{{x.criteria}}</td><td class="sm">{{x.reason}}</td>
    </tr>{% endfor %}
  </table></div>
</div>
{% endfor %}

<div class="card">
  <h2>Match result</h2>
  <table>
    <tr><th>Outcome</th><th class="n">Count</th><th class="n">Tally value</th>
        <th class="n">ERP value</th></tr>
    {% for k in order %}{% if k in R.buckets %}
    <tr><td><span class="tag {{k}}">{{ label[k] }}</span></td>
        <td class="n">{{ R.buckets[k][0] }}</td>
        <td class="n">{{ money(R.buckets[k][1]) }}</td>
        <td class="n">{{ money(R.buckets[k][2]) }}</td></tr>
    {% endif %}{% endfor %}
  </table>
</div>

<div class="card">
  <h2>Balance bridge &mdash; why the two systems differ</h2>
  <table class="bridge">
    <tr class="h"><td>Tally closing balance</td><td class="n">{{ money(R.t_close) }}</td></tr>
    {% for k,v in bridge %}
    <tr><td>{{k}}</td><td class="n {{ 'neg' if v<0 }}">{{ money(v) }}</td></tr>
    {% endfor %}
    <tr class="h"><td>ERP closing balance</td><td class="n">{{ money(R.e_close) }}</td></tr>
    <tr class="h"><td>Unexplained residual</td>
        <td class="n" style="color:{{ '#1f7a43' if R.residual==0 else '#a32b2b' }}">
        {{ money(R.residual) }}</td></tr>
  </table>
</div>

<div class="card">
  <h2>Exceptions &mdash; {{ exc|length }} items needing action</h2>
  <div class="pills" id="pills">
    <span class="pill on" data-f="ALL">All</span>
    {% for k in order %}{% if k!='MATCHED' and k in R.buckets %}
      <span class="pill" data-f="{{k}}">{{ label[k] }} ({{ R.buckets[k][0] }})</span>
    {% endif %}{% endfor %}
  </div>
  <div class="scroll"><table id="exc">
    <tr><th>Status</th><th>Document</th><th>Date</th><th>Description</th>
        <th class="n">Tally</th><th class="n">ERP</th><th class="n">Difference</th>
        <th>Reason</th><th>Action</th></tr>
    {% for r in exc %}
    <tr data-s="{{r.status}}">
      <td><span class="tag {{r.status}}">{{r.status}}</span></td>
      <td>{{r.key}}</td><td>{{ r.date }}</td><td>{{r.desc}}</td>
      <td class="n {{ 'neg' if r.t_amt<0 }}">{{ money(r.t_amt) }}</td>
      <td class="n {{ 'neg' if r.e_amt<0 }}">{{ money(r.e_amt) }}</td>
      <td class="n {{ 'neg' if r.diff<0 }}">{{ money(r.diff) }}</td>
      <td>{{r.reason}}</td><td>{{r.action}}</td>
    </tr>{% endfor %}
  </table></div>
</div>

{% if R.dq %}
<div class="card">
  <h2>Data quality &mdash; fix at source</h2>
  <table><tr><th>Where</th><th>Value</th><th>Issue</th></tr>
  {% for a,b,c in R.dq %}<tr><td>{{a}}</td><td>{{b}}</td><td>{{c}}</td></tr>{% endfor %}
  </table>
</div>{% endif %}

<script>
document.getElementById('pills').addEventListener('click',function(ev){
  var p=ev.target.closest('.pill'); if(!p) return;
  this.querySelectorAll('.pill').forEach(function(x){x.classList.remove('on')});
  p.classList.add('on'); var f=p.dataset.f;
  document.querySelectorAll('#exc tr[data-s]').forEach(function(tr){
    tr.style.display=(f=='ALL'||tr.dataset.s==f)?'':'none';});
});
// Debit / Credit detail filters: category pills and system pills
document.querySelectorAll('.pills[data-for]').forEach(function(box){
  box.addEventListener('click',function(ev){
    var p=ev.target.closest('.pill'); if(!p) return;
    box.querySelectorAll('.pill').forEach(function(x){x.classList.remove('on')});
    p.classList.add('on'); var f=p.dataset.f;
    document.querySelectorAll('#'+box.dataset.for+' tr[data-c]').forEach(function(tr){
      var show = (f=='ALL') || (tr.dataset.c==f) || (tr.dataset.y==f);
      tr.style.display = show ? '' : 'none';});
  });
});
</script>
"""


def money(x):
    try:
        x = float(x)
    except (TypeError, ValueError):
        return ''
    return ('-' if x < 0 else '') + format(abs(round(x, 2)), ',.2f')


def page(body):
    return render_template_string(SHELL, body=body)


@app.errorhandler(Exception)
def on_error(ex):
    tb = traceback.format_exc()
    app.logger.error(tb)
    code = getattr(ex, 'code', 500)
    if code == 404:
        return page('<div class="card"><h2>Not found</h2>'
                    '<p><a class="btn" href="/">Back</a></p></div>'), 404
    return page('<div class="card"><h2>Something went wrong</h2>'
                '<div class="err">%s</div>'
                '<p style="margin-top:14px"><a class="btn alt" href="/">Back</a></p>'
                '</div>' % tb), 500


@app.route('/')
def index():
    purge()
    return page(render_template_string(UPLOAD_BODY))


@app.route('/run', methods=['POST'])
def run():
    ensure_dirs()
    token = request.form.get('token')
    if token:
        matches = glob.glob(os.path.join(UPLOADS, secure_filename(token) + '__*'))
        if not matches:
            return redirect(url_for('index'))
        path = matches[0]
    else:
        f = request.files.get('file')
        if not f or not f.filename:
            return redirect(url_for('index'))
        name = secure_filename(f.filename) or 'upload.xlsx'
        if not name.lower().endswith(('.xlsx', '.xlsm', '.xls')):
            return page('<div class="card"><div class="err">Please upload an '
                        '.xlsx, .xlsm or .xls file.</div>'
                        '<p><a class="btn" href="/">Back</a></p></div>')
        token = uuid.uuid4().hex[:12]
        path = os.path.join(UPLOADS, token + '__' + name)
        f.save(path)

    fname = os.path.basename(path).split('__', 1)[-1]
    try:
        R = E.reconcile(path,
                        request.form.get('tally_sheet') or None,
                        request.form.get('erp_sheet') or None)
    except Exception as ex:
        try:
            sheets = E.detect_sheets(path)['all']
        except Exception:
            sheets = []
        hint = ''
        if sheets:
            hint = ('<form method="post" action="/run" class="row" style="margin-top:14px">'
                    '<input type="hidden" name="token" value="%s">'
                    '<span class="note">Pick the sheets manually:</span>'
                    '<label class="note">Tally <select name="tally_sheet">%s</select></label>'
                    '<label class="note">ERP <select name="erp_sheet">%s</select></label>'
                    '<button type="submit">Re-run</button></form>'
                    % (token,
                       ''.join('<option>%s</option>' % s for s in sheets),
                       ''.join('<option>%s</option>' % s for s in sheets)))
        return page('<div class="card"><h2>Could not reconcile</h2>'
                    '<div class="err">%s</div>%s'
                    '<p style="margin-top:14px"><a class="btn alt" href="/">Back</a></p>'
                    '</div>' % (str(ex) or traceback.format_exc(), hint))

    report = os.path.join(REPORTS, token + '__ERP_vs_Tally_Reconciliation.xlsx')
    E.write_report(R, report)

    exc = []
    for r in R['results']:
        if r['status'] == 'MATCHED':
            continue
        desc = (' + '.join(sorted({t['nature'] for t in r['t_rows']}))
                or ' + '.join(sorted({e['particulars'] for e in r['e_rows']})))
        exc.append(dict(status=r['status'], key=r['key'],
                        date=(r['t_date'] or r['e_date']),
                        desc=desc, t_amt=r['t_amt'], e_amt=r['e_amt'],
                        diff=r['diff'], reason=r['reason'],
                        action=E.ACTION.get(r['status'], '')))

    body = render_template_string(
        RESULT_BODY, R=R, token=token, fname=fname, exc=exc, money=money,
        dr=lambda x: abs(x) if x < 0 else 0.0,
        cr=lambda x: x if x > 0 else 0.0,
        tol='%.2f' % E.TOL, win=E.DATE_WINDOW, bank_win=E.BANK_DATE_WINDOW,
        order=E.STATUS_ORDER, label=E.LABEL,
        sheets=E.detect_sheets(path)['all'],
        bridge=sorted(R['bridge'].items(), key=lambda x: -abs(x[1])))
    return page(body)


@app.route('/report/<token>')
def report(token):
    token = secure_filename(token)
    matches = glob.glob(os.path.join(REPORTS, token + '__*'))
    if not matches:
        abort(404)
    return send_file(matches[0], as_attachment=True,
                     download_name='ERP_vs_Tally_Reconciliation.xlsx')


if __name__ == '__main__':
    print('\n  ERP vs Tally Reconciliation')
    print('  Open this in your browser:  http://127.0.0.1:5000\n')
    try:
        app.run(host='127.0.0.1', port=5000, debug=False)
    except OSError as ex:
        print('  Could not start on port 5000: %s' % ex)
        print('  Another copy is probably already running - just open the URL above.')
