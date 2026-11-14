"""Chama WhatsApp bot - command logic. Backend-agnostic (Google Sheets or local .xlsx)."""
import re, time, datetime as dt, subprocess, json, os

EPOCH = dt.date(1899, 12, 30)
CATS = {'s': 'Savings', 'sav': 'Savings', 'savings': 'Savings', 'w': 'Welfare', 'wel': 'Welfare', 'welfare': 'Welfare',
        'm': 'Merry-go-round', 'mgr': 'Merry-go-round', 'merry': 'Merry-go-round',
        'f': 'Fine Payment', 'fine': 'Fine Payment', 'l': 'Loan Repayment', 'loan': 'Loan Repayment'}

def today():
    return (dt.datetime.utcnow() + dt.timedelta(hours=3)).date()  # Nairobi, UTC+3

def to_date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, (int, float)) and v > 20000: return EPOCH + dt.timedelta(days=int(v))
    return None

def num(v): return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0
def money(v): return f"{num(v):,.0f}"
def cell(row, i): return row[i] if i < len(row) and row[i] not in (None, '') else None
def last9(p): return re.sub(r'\D', '', str(p or ''))[-9:]

def split_messages(text, limit=1500):
    messages = []
    current = ''
    for line in text.splitlines():
        pieces = [line[i:i + limit] for i in range(0, len(line), limit)] or ['']
        for piece in pieces:
            addition = ('\n' if current else '') + piece
            if current and len(current) + len(addition) > limit:
                messages.append(current)
                current = piece
            else:
                current += addition
    if current or not messages: messages.append(current)
    return messages

def parse_amount(t):
    t = t.lower().replace(',', '')
    m = re.fullmatch(r'(\d+(?:\.\d+)?)(k?)', t)
    if not m: return None
    return float(m.group(1)) * (1000 if m.group(2) else 1)

# ---------------------------------------------------------------- backends
class XlsxBackend:
    """Local test backend: edits an .xlsx with openpyxl and recalculates via LibreOffice."""
    def __init__(self, path, recalc_script='/mnt/skills/public/xlsx/scripts/recalc.py'):
        self.path, self.recalc_script, self.dirty = path, recalc_script, False
    def _recalc(self):
        if self.dirty:
            subprocess.run(['python3', self.recalc_script, self.path, '180'], capture_output=True)
            self.dirty = False
    def get(self, sheet, a1):
        from openpyxl import load_workbook
        self._recalc()
        wb = load_workbook(self.path, data_only=True)
        return [[c.value for c in r] for r in wb[sheet][a1]]
    def write(self, sheet, row, values):
        from openpyxl import load_workbook
        wb = load_workbook(self.path); ws = wb[sheet]
        for col, v in values.items(): ws[f'{col}{row}'].value = v
        wb.save(self.path); self.dirty = True
    def clear(self, sheet, row, cols):
        self.write(sheet, row, {c: None for c in cols})

class SheetsBackend:
    """Production backend: Google Sheets via a service account (pip install gspread)."""
    def __init__(self, sheet_id, creds_file):
        import gspread
        self.sh = gspread.service_account(filename=creds_file).open_by_key(sheet_id)
        self._ws = {}
    def ws(self, name):
        if name not in self._ws: self._ws[name] = self.sh.worksheet(name)
        return self._ws[name]
    def get(self, sheet, a1):
        return self.ws(sheet).get(a1, value_render_option='UNFORMATTED_VALUE')
    def write(self, sheet, row, values):
        data = []
        for col, v in values.items():
            if isinstance(v, dt.date): v = (v - EPOCH).days  # serial number: locale-proof
            data.append({'range': f'{col}{row}', 'values': [[v if v is not None else '']]})
        self.ws(sheet).batch_update(data, value_input_option='USER_ENTERED')
    def clear(self, sheet, row, cols):
        self.ws(sheet).batch_clear([f'{c}{row}' for c in cols])

# ---------------------------------------------------------------- bot
HELP = """*Chama bot commands*
*REG001* - type a member ID for their details + statement\n*members* - list all members and their status\n*0712345678* - look up by phone number\n*bal john* - balance for a member
*statement john* - full statement
*daily [date]* - contributions for a day (e.g. daily yesterday)
*monthly [month]* - contributions for a month (e.g. monthly last)
*group [month]* - formatted group statement to forward
*arrears* - who owes what
*overdue* - overdue loans
*summary* - group totals
*pay john 2000* - record savings
*pay john 200 welfare*
*pay john 500 mgr*
*pay john 100 fine*
*pay john 4000 loan L001 REFCODE*
*fine john 200 absent from meeting*
*loan john 20000 4 [5]* - amount, months, optional rate %
*welfare faith 10000 bereavement* - welfare payout
*undo* - reverse your last entry
Paste an M-Pesa receipt SMS and I'll suggest the entry.
Members can send *bal* to see their own balance."""

class Bot:
    def __init__(self, backend, admins):
        self.b = backend
        self.admins = {last9(a) for a in admins}
        self.pending = {}   # sender -> (expiry, summary, fn)
        self.last = {}      # sender -> (sheet, row, cols)

    # ---- reads
    def setup(self):
        r = self.b.get('Setup', 'B3:B13')
        g = lambda i: cell(r[i], 0) if i < len(r) else None
        return {'name': g(0) or 'Chama', 'cur': g(1) or 'KES', 'asof': to_date(g(2)) or today(), 'pay': g(8) or '', 'tname': g(9), 'tphone': g(10)}
    def members(self):
        out = []
        for r in self.b.get('Members', 'A6:H65'):
            if cell(r, 1): out.append({'id': str(cell(r, 0) or ''), 'name': str(r[1]), 'phone': cell(r, 2),
                                       'join': to_date(cell(r, 3)), 'sav': cell(r, 4), 'wel': cell(r, 5), 'exit': to_date(cell(r, 6))})
        return out
    def find(self, q, pool=None):
        q = q.strip().lower(); ms = self.members() if pool is None else pool
        ex = [m for m in ms if m['name'].lower() == q or m['id'].lower() == q]
        if ex: return ex
        qt = q.split()
        return [m for m in ms if all(any(nt.startswith(t) for nt in m['name'].lower().split()) for t in qt)]
    def summary(self, mid):
        for r in self.b.get('Summary', 'A8:V67'):
            if cell(r, 0) == mid: return r
    def loans(self, mid=None):
        out = []
        for r in self.b.get('Loans', 'A6:Q205'):
            if cell(r, 0) and (mid is None or cell(r, 2) == mid): out.append(r)
        return out

    # ---- text builders
    def statement(self, m):
        s, st = self.summary(m['id']), self.setup()
        if not s: return f"No summary row for {m['name']}."
        L = [f"*{st['name']}*", f"Statement for *{m['name']}* ({m['id']})", f"As at {st['asof']:%d %b %Y}", ""]
        def line(label, exp, paid, bal):
            bal = num(bal)
            tail = f" - *owe {money(bal)}*" if bal > 0 else (f" (ahead {money(-bal)})" if bal < 0 else "")
            L.append(f"{'⚠️' if bal > 0 else '✅'} {label}: paid {money(paid)} of {money(exp)}{tail}")
        line('Savings', s[5], s[6], s[7]); line('Welfare', s[8], s[9], s[10])
        line('Merry-go-round', s[11], s[12], s[13])
        if num(s[14]): line('Fines', s[14], s[15], s[16])
        for ln in self.loans(m['id']):
            if num(cell(ln, 12)) > 0:
                d = to_date(cell(ln, 10)); od = num(cell(ln, 13))
                L.append(f"💰 Loan {ln[0]}: balance {money(ln[12])}, due {d:%d %b %Y}" + (f" - *{int(od)} days overdue*, penalty {money(ln[14])}" if od else ""))
        L.append("")
        L.append(f"*TOTAL OWED: {st['cur']} {money(s[19])}*" if num(s[19]) > 0 else "*All clear ✅*")
        if num(s[19]) > 0 and st['pay']: L += ["", st['pay']]
        return "\n".join(L)
    def details(self, m):
        head = [f"👤 *{m['name']}* ({m['id']})"]
        if m.get('phone'): head.append(f"📱 {m['phone']}")
        info = []
        if m.get('join'): info.append(f"Joined {m['join']:%d %b %Y}")
        if m.get('sav') is not None: info.append(f"Savings {money(m['sav'])}/month")
        if m.get('wel') is not None: info.append(f"Welfare {money(m['wel'])}/month")
        if m.get('exit'): info.append(f"Exited {m['exit']:%d %b %Y}")
        if info: head.append(" | ".join(info))
        return "\n".join(head) + "\n\n" + self.statement(m)
    def balance(self, m):
        s = self.summary(m['id'])
        if not s: return f"No data for {m['name']}."
        parts = []
        for lab, i in [('Savings', 7), ('Welfare', 10), ('MGR', 13), ('Fines', 16)]:
            if num(s[i]) > 0: parts.append(f"{lab} arrears {money(s[i])}")
        if num(s[17]) > 0: parts.append(f"Loan balance {money(s[17])}")
        if num(s[18]) > 0: parts.append(f"Loan penalty {money(s[18])}")
        body = "\n".join(parts) if parts else "No arrears"
        return f"*{m['name']}*\n{body}\n*TOTAL OWED: {money(s[19])}*"
    def member_list(self):
        rows = [r for r in self.b.get('Summary', 'A8:V67') if cell(r, 0)]
        if not rows: return "No members found."
        out = [f"*Members ({len(rows)})*"]
        for r in rows:
            out.append(f"{r[0]} {r[1]}: " + (f"owes {money(r[19])}" if num(r[19]) > 0 else "up to date"))
        return "\n".join(out) + "\n\nSend a member ID (e.g. " + str(rows[0][0]) + ") for full details."
    def arrears(self):
        rows = [(r[1], num(r[19])) for r in self.b.get('Summary', 'A8:V67') if cell(r, 0) and num(r[19]) > 0]
        rows.sort(key=lambda x: -x[1])
        if not rows: return "Nobody owes anything 🎉"
        return "*Members owing*\n" + "\n".join(f"{n}: {money(v)}" for n, v in rows) + f"\n*Total: {money(sum(v for _, v in rows))}*"
    def overdue(self):
        rows = [r for r in self.loans() if cell(r, 16) == 'OVERDUE']
        if not rows: return "No overdue loans 🎉"
        return "*Overdue loans*\n" + "\n".join(f"{r[0]} {r[1]}: balance {money(r[12])}, {int(num(r[13]))} days late, penalty {money(r[14])}" for r in rows)
    def group_summary(self):
        d = {str(cell(r, 0)).strip(): cell(r, 1) for r in self.b.get('Dashboard', 'A3:B29') if cell(r, 0)}
        keys = ['Active members', 'Members currently owing something', 'Total savings collected', 'Contribution arrears (savings + welfare + MGR)',
                'Outstanding loan balance', 'of which OVERDUE', 'Fines outstanding', 'Welfare fund balance', 'EXPECTED CASH ON HAND']
        return "*Group summary*\n" + "\n".join(f"{k}: {money(d.get(k))}" for k in keys)

    def report_date(self, text):
        value = text.strip().lower()
        current = today()
        if not value or value == 'today': return current
        if value == 'yesterday': return current - dt.timedelta(days=1)
        for fmt in ('%Y-%m-%d', '%d %b %Y', '%d %B %Y'):
            try:
                return dt.datetime.strptime(value, fmt).date()
            except ValueError:
                pass
        for fmt in ('%d %b %Y', '%d %B %Y'):
            try:
                return dt.datetime.strptime(f'{value} {current.year}', fmt).date()
            except ValueError:
                pass
        return None

    def report_month(self, text):
        value = text.strip().lower()
        current = today().replace(day=1)
        if not value or value in ('this', 'current'): return current
        if value in ('last', 'previous'):
            return (current - dt.timedelta(days=1)).replace(day=1)
        for fmt in ('%Y-%m', '%b %Y', '%B %Y'):
            try:
                return dt.datetime.strptime(value, fmt).date().replace(day=1)
            except ValueError:
                pass
        for fmt in ('%b %Y', '%B %Y'):
            try:
                return dt.datetime.strptime(f'{value} {current.year}', fmt).date().replace(day=1)
            except ValueError:
                pass
        return None

    @staticmethod
    def month_end(month):
        return (month.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)

    def transaction_rows(self):
        return self.b.get('Transactions', 'A6:H3005')

    def collections(self, start, end):
        grouped = {}
        category_totals = {}
        for row in self.transaction_rows():
            date = to_date(cell(row, 0))
            name, category, amount = cell(row, 1), cell(row, 3), num(cell(row, 4))
            if date is None or not start <= date <= end or not name or not amount:
                continue
            category = str(category or 'Other')
            grouped.setdefault(str(name), {}).setdefault(category, 0)
            grouped[str(name)][category] += amount
            category_totals[category] = category_totals.get(category, 0) + amount
        return grouped, category_totals

    def daily_report(self, day):
        grouped, totals = self.collections(day, day)
        st = self.setup()
        lines = [f"*DAILY CONTRIBUTIONS - {day:%d %b %Y}*"]
        if not grouped:
            return "\n".join(lines + ["No payments recorded."])
        for name in sorted(grouped, key=str.casefold):
            cats = grouped[name]
            detail = ", ".join(f"{cat} {money(value)}" for cat, value in sorted(cats.items()))
            lines.append(f"{name}: {detail} - *{money(sum(cats.values()))}*")
        lines.extend(["", "*TOTALS*"])
        lines.extend(f"{cat}: {money(value)}" for cat, value in sorted(totals.items()))
        lines.append(f"*Total received: {st['cur']} {money(sum(totals.values()))}*")
        return "\n".join(lines)

    def monthly_report(self, month):
        end = self.month_end(month)
        grouped, totals = self.collections(month, end)
        st = self.setup()
        members = self.members()
        expected = sum(
            num(m['sav']) for m in members
            if (m['join'] is None or m['join'] <= end)
            and (m['exit'] is None or m['exit'] >= month)
        )
        savings_received = totals.get('Savings', 0)
        lines = [f"*MONTHLY CONTRIBUTIONS - {month:%B %Y}*"]
        if grouped:
            for name in sorted(grouped, key=str.casefold):
                cats = grouped[name]
                detail = ", ".join(f"{cat} {money(value)}" for cat, value in sorted(cats.items()))
                lines.append(f"{name}: {detail} - *{money(sum(cats.values()))}*")
        else:
            lines.append("No payments recorded.")
        lines.extend(["", "*TOTALS*"])
        if totals:
            lines.extend(f"{cat}: {money(value)}" for cat, value in sorted(totals.items()))
        else:
            lines.append("No collections.")
        pct = f"{savings_received / expected:.0%}" if expected else "n/a"
        lines.append(f"*Savings against expected: {st['cur']} {money(savings_received)} / {money(expected)} ({pct})*")
        no_savings = sorted(
            (m['name'] for m in members
             if (m['join'] is None or m['join'] <= end)
             and (m['exit'] is None or m['exit'] >= month)
             and not num(grouped.get(m['name'], {}).get('Savings'))),
            key=str.casefold,
        )
        if no_savings:
            lines.append("*No savings payment:* " + ", ".join(no_savings))
        lines.append(f"*Total received: {st['cur']} {money(sum(totals.values()))}*")
        return "\n".join(lines)

    def group_statement(self, month):
        st = self.setup()
        _, totals = self.collections(month, self.month_end(month))
        rows = [r for r in self.b.get('Summary', 'A8:V67') if cell(r, 0)]
        dashboard = {str(cell(r, 0)).strip(): cell(r, 1)
                     for r in self.b.get('Dashboard', 'A3:B29') if cell(r, 0)}
        lines = [
            f"*{st['name'].upper()}: GROUP STATEMENT*",
            f"_Collections for {month:%B %Y} | As at {st['asof']:%d %b %Y}_",
            "",
            "*MONTHLY COLLECTIONS*",
        ]
        if totals:
            lines.extend(f"{cat}: {money(value)}" for cat, value in sorted(totals.items()))
        else:
            lines.append("No collections recorded.")
        lines.append(f"*Total received: {st['cur']} {money(sum(totals.values()))}*")
        lines.extend(["", "*MEMBER STATUS*"])
        lines.extend(
            f"{'⚠️' if num(cell(row, 19)) > 0 else '✅'} {cell(row, 1)}: "
            + (f"owes {money(cell(row, 19))}" if num(cell(row, 19)) > 0 else "up to date")
            for row in rows
        )
        keys = [
            'Total savings collected',
            'Contribution arrears (savings + welfare + MGR)',
            'Outstanding loan balance',
            'of which OVERDUE',
            'Welfare fund balance',
            'EXPECTED CASH ON HAND',
        ]
        lines.extend(["", "*GROUP POSITION*"])
        lines.extend(f"{key}: {st['cur']} {money(dashboard.get(key))}" for key in keys)
        if st['pay']: lines.extend(["", st['pay']])
        if st['tname'] or st['tphone']:
            lines.append("Treasurer: " + " ".join(str(x) for x in (st['tname'], st['tphone']) if x))
        return "\n".join(lines)

    # ---- writes (all go through confirmation)
    def next_row(self, sheet, first=6, last=3005):
        data = self.b.get(sheet, f'A{first}:B{last}')
        for i in range(last - first + 1):
            row = data[i] if i < len(data) else []
            if not any(str(x).strip() for x in row if x is not None): return first + i
        raise RuntimeError(f'{sheet} sheet is full')
    def do_write(self, sender, sheet, values, last):
        row = self.next_row(sheet, 6, last)
        self.b.write(sheet, row, values)
        self.last[sender] = (sheet, row, list(values))
        return row
    def ask(self, sender, summary, fn):
        self.pending[sender] = (time.time() + 300, summary, fn)
        return f"{summary}\nReply *YES* to confirm or *NO* to cancel."

    def pick(self, q):
        ms = self.find(q)
        if not ms: return None, f"No member matches '{q}'."
        if len(ms) > 1: return None, "Which one? " + ", ".join(m['name'] for m in ms[:6])
        return ms[0], None
    def risk_note(self, m):
        s = self.summary(m['id'])
        return f"\n⚠️ Note: {m['name']} currently owes {money(s[19])}." if s and num(s[19]) > 0 else ""

    def cmd_pay(self, sender, args, note='via WhatsApp', ref_override=None):
        sp = self.split_name_amount(args)
        if not sp: return "Format: *pay john 2000* [savings|welfare|mgr|fine|loan] [LoanID] [ref]"
        q, amt, rest = sp
        m, err = self.pick(q)
        if err: return err
        cat, loan_id = 'Savings', None
        if rest and rest[0].lower() in CATS: cat = CATS[rest[0].lower()]; rest = rest[1:]
        if cat == 'Loan Repayment':
            open_ = [l for l in self.loans(m['id']) if num(cell(l, 12)) > 0]
            if rest and re.fullmatch(r'[Ll]\d+', rest[0]): loan_id = rest[0].upper(); rest = rest[1:]
            elif len(open_) == 1: loan_id = open_[0][0]
            elif not open_: return f"{m['name']} has no open loan."
            else: return f"{m['name']} has several loans ({', '.join(l[0] for l in open_)}). Add the Loan ID."
            bal = next((num(l[12]) for l in self.loans(m['id']) if l[0] == loan_id), None)
            if bal is None: return f"Loan {loan_id} not found for {m['name']}."
        ref = (ref_override or (' '.join(rest).upper() if rest else ''))
        d = today()
        vals = {'A': d, 'B': m['name'], 'D': cat, 'E': amt, 'F': ref, 'G': loan_id, 'H': note}
        summ = f"Record *{money(amt)}* {cat}" + (f" ({loan_id})" if loan_id else "") + f" from *{m['name']}*" + (f", ref {ref}" if ref else "") + "?"
        if cat == 'Loan Repayment' and amt > bal: summ += f"\n⚠️ That is more than the loan balance ({money(bal)})."
        def run():
            self.do_write(sender, 'Transactions', vals, 3005)
            return f"✅ Recorded. " + self.balance(m)
        return self.ask(sender, summ, run)

    def split_name_amount(self, tokens):
        for i, t in enumerate(tokens):
            if i > 0 and parse_amount(t) is not None: return ' '.join(tokens[:i]), parse_amount(t), tokens[i + 1:]

    def cmd_fine(self, sender, args):
        sp = self.split_name_amount(args)
        if not sp or not sp[2]: return "Format: *fine john 200 absent from meeting*"
        q, amt, rest = sp; m, err = self.pick(q)
        if err: return err
        reason = ' '.join(rest).capitalize()
        vals = {'A': today(), 'B': m['name'], 'D': reason, 'E': amt, 'F': 'via WhatsApp'}
        def run():
            self.do_write(sender, 'Fines', vals, 1005); return "✅ Fine recorded.\n" + self.balance(m)
        return self.ask(sender, f"Charge *{m['name']}* a fine of *{money(amt)}* for: {reason}?", run)

    def cmd_welfare(self, sender, args):
        sp = self.split_name_amount(args)
        if not sp or not sp[2]: return "Format: *welfare faith 10000 bereavement*"
        q, amt, rest = sp; m, err = self.pick(q)
        if err: return err
        reason = ' '.join(rest).capitalize()
        vals = {'A': today(), 'B': m['name'], 'D': reason, 'E': amt, 'F': 'via WhatsApp'}
        def run():
            self.do_write(sender, 'Welfare', vals, 205); return "✅ Welfare payout recorded.\n" + self.group_summary()
        return self.ask(sender, f"Pay out *{money(amt)}* welfare to *{m['name']}* ({reason})?", run)

    def cmd_loan(self, sender, args):
        nums = [i for i, t in enumerate(args) if i > 0 and parse_amount(t) is not None]
        if len(nums) < 2: return "Format: *loan john 20000 4* (amount, months) or add a rate: *loan john 20000 4 5*"
        q = ' '.join(args[:nums[0]]); amt = parse_amount(args[nums[0]]); term = int(parse_amount(args[nums[1]]))
        rate = parse_amount(args[nums[2]]) / 100 if len(nums) > 2 else None
        m, err = self.pick(q)
        if err: return err
        ids = [int(re.sub(r'\D', '', str(r[0])) or 0) for r in self.loans()]
        lid = f"L{(max(ids) if ids else 0) + 1:03d}"
        vals = {'A': lid, 'B': m['name'], 'D': today(), 'E': amt, 'F': term, 'G': rate}
        def run():
            self.do_write(sender, 'Loans', vals, 205)
            ln = next((r for r in self.loans() if r[0] == lid), None)
            extra = f"\nTotal due {money(ln[9])} (interest {money(ln[8])}), due {to_date(ln[10]):%d %b %Y}." if ln else ""
            return f"✅ Loan {lid} created for {m['name']}." + extra
        return self.ask(sender, f"Issue loan *{lid}* of *{money(amt)}* to *{m['name']}* for {term} months" + (f" at {rate*100:g}%/month" if rate else " at the default rate") + "?" + self.risk_note(m), run)

    def cmd_mpesa(self, sender, text):
        mm = re.search(r'([A-Z0-9]{10})\s+Confirmed\.? .*?Ksh\.?\s?([\d,]+(?:\.\d{1,2})?)\s+received from\s+(.+?)\s+(\+?\d{9,13})', text, re.S | re.I)
        if not mm: return None
        code, amt, name, phone = mm.group(1).upper(), float(mm.group(2).replace(',', '')), mm.group(3), mm.group(4)
        ms = [m for m in self.members() if last9(m['phone']) == last9(phone)] or self.find(name)
        if len(ms) != 1:
            return f"M-Pesa receipt {code}: KES {money(amt)} from {name} {phone}. I can't match one member. Use *pay <name> {amt:g} {code}*"
        m = ms[0]
        vals = {'A': today(), 'B': m['name'], 'D': 'Savings', 'E': amt, 'F': code, 'H': 'M-Pesa SMS via WhatsApp'}
        def run():
            self.do_write(sender, 'Transactions', vals, 3005); return "✅ Recorded. " + self.balance(m)
        return self.ask(sender, f"M-Pesa {code}: *{money(amt)}* from *{m['name']}*. Record as *Savings*?\n(For another category say NO, then send e.g. *pay {m['name'].split()[0].lower()} {amt:g} welfare {code}*)", run)

    def cmd_undo(self, sender):
        if sender not in self.last: return "Nothing to undo (I only remember your last entry since I started)."
        sheet, row, cols = self.last.pop(sender)
        self.b.clear(sheet, row, cols)
        return f"↩️ Removed your last entry ({sheet} row {row})."

    # ---- member self-service: a number can ONLY ever see records whose phone on the Members sheet matches it
    BAL_CMDS = ('bal', 'balance', 'b', 'statement', 'stmt', 'st')
    MEMBER_HELP = "*Your chama account*\nSend *bal* for your balance\nSend *statement* for your full statement\nOr send your member ID (e.g. REG001)."
    DENY = "I can only show your own balance and statement. Send *bal*, *statement* or your member ID."
    def member_flow(self, sender, parts):
        cmd = parts[0].lower()
        if cmd in ('help', 'menu', 'hi', 'hello'): return self.MEMBER_HELP
        key = last9(sender)
        own = [m for m in self.members() if len(key) == 9 and last9(m['phone']) == key]
        if not own: return "Your number is not registered on the member list. Please contact the treasurer."
        args = ' '.join(parts[1:])
        if cmd in self.BAL_CMDS:
            if args:
                hit = self.find(args, own)           # searched ONLY inside the person's own records
                if len(hit) != 1: return self.DENY    # same reply whether or not that member exists
                own = hit
            full = cmd in ('statement', 'stmt', 'st')
            return "\n\n".join(self.statement(m) if full else self.balance(m) for m in own)
        if len(parts) == 1:
            hit = [m for m in own if m['id'].lower() == cmd]
            return self.details(hit[0]) if hit else self.DENY
        return self.DENY

    # ---- entry point
    def handle(self, sender, text):
        text = (text or '').strip()
        if not text: return HELP
        is_admin = last9(sender) in self.admins
        low = text.lower()
        if low in ('yes', 'y', 'confirm', 'ok'):
            p = self.pending.pop(sender, None)
            if not p or p[0] < time.time(): return "Nothing waiting for confirmation."
            return p[2]()
        if low in ('no', 'n', 'cancel'):
            return "Cancelled." if self.pending.pop(sender, None) else "Nothing to cancel."
        if 'confirmed' in low and 'received from' in low:
            if not is_admin: return "Only the treasurer can record payments."
            return self.cmd_mpesa(sender, text) or "Couldn't read that M-Pesa message. Use *pay john 2000*."
        parts = text.split(); cmd, args = parts[0].lower(), parts[1:]
        if not is_admin:
            try: return self.member_flow(sender, parts)
            except Exception: return "Something went wrong. Please try again."
        try:
            if len(parts) == 1:  # a bare member ID, e.g. REG001 -> member details
                hit = [m for m in self.members() if m['id'].lower() == low]
                if hit:
                    if not is_admin and last9(hit[0]['phone']) != last9(sender): return "You can only view your own details."
                    return self.details(hit[0])
                digits = re.sub(r'\D', '', parts[0])   # admin: a bare phone number -> that member's details
                if len(digits) >= 9 and re.fullmatch(r'\+?\d+', parts[0]):
                    hit = [m for m in self.members() if last9(m['phone']) == digits[-9:]]
                    if hit: return "\n\n".join(self.details(m) for m in hit)
                    return "No member has that phone number."
            if cmd in ('members', 'list'): return self.member_list()
            if cmd in ('help', 'menu', 'hi', 'hello'): return HELP
            if cmd in ('bal', 'balance', 'b', 'statement', 'stmt', 'st'):
                if args: ms = self.find(' '.join(args))
                else: ms = [m for m in self.members() if last9(m['phone']) == last9(sender)]
                if not ms: return "I couldn't find that member." if args else "Your number isn't on the member list. Ask the treasurer."
                if len(ms) > 1: return "Which one? " + ", ".join(m['name'] for m in ms[:6])
                if not is_admin and last9(ms[0]['phone']) != last9(sender): return "You can only view your own balance."
                return self.statement(ms[0]) if cmd in ('statement', 'stmt', 'st') else self.balance(ms[0])
            if not is_admin: return "Sorry, only the chama officials can do that. Send *bal* for your own balance."
            if cmd in ('arrears', 'owing'): return self.arrears()
            if cmd == 'overdue': return self.overdue()
            if cmd in ('summary', 'report'): return self.group_summary()
            if cmd == 'daily':
                day = self.report_date(' '.join(args))
                return self.daily_report(day) if day else "Format: *daily* or *daily 2026-11-14* (also accepts *daily yesterday*)."
            if cmd == 'monthly':
                month = self.report_month(' '.join(args))
                return self.monthly_report(month) if month else "Format: *monthly* or *monthly 2026-11* (also accepts *monthly last*)."
            if cmd == 'group':
                month = self.report_month(' '.join(args))
                return self.group_statement(month) if month else "Format: *group* or *group 2026-11* (also accepts *group last*)."
            if cmd in ('pay', 'paid', 'p'): return self.cmd_pay(sender, args)
            if cmd == 'fine': return self.cmd_fine(sender, args)
            if cmd == 'welfare': return self.cmd_welfare(sender, args)
            if cmd == 'loan': return self.cmd_loan(sender, args)
            if cmd == 'undo': return self.cmd_undo(sender)
        except Exception as e:  # never crash the webhook
            return f"Something went wrong: {e}"
        return "I didn't understand that. Send *help* to see commands."
