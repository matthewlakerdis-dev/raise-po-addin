/* Raise PO add-in - readers for PO request emails and supplier quote / invoice PDFs.
   Same reading logic as the Nexvia "Quote to PO Lines" script (v1.25):
   Lysaght-style line blocks, Knotwood-style columns, general tables (Baker, Nut & Bolt, Bradstone),
   then a line-by-line fallback. */
(function (g) {
  'use strict';

  const num = (s) => {
    if (s == null) return NaN;
    const t = String(s).replace(/[$,\s]/g, '');
    return /^-?\d+(\.\d+)?$/.test(t) ? parseFloat(t) : NaN;
  };
  const isNum = (s) => !isNaN(num(s));
  const isUom = (s) => /^(ea|each|m|lm|l\/m|m2|m²|sqm|mm|pcs?|pc|no|nr|item|set|lot|kg|hr|hrs|len|lengths?|sht|sheet|box|pk|pack|roll)$/i.test(String(s || ''));
  const near = (a, b) => Math.abs(a - b) <= Math.max(0.05, Math.abs(b) * 0.01);
  const r2 = (v) => Math.round(v * 100) / 100;
  const mk = (description, quantity, unitCost, measure) =>
    ({ description: String(description).trim(), quantity, unitCost, measure: measure || '', check: 'ok' });
  const skipLine = /\b(sub\s*-?total|total|gst|tax|balance|amount due|deposit|page \d)/i;

  // ---------- line-by-line fallback ----------
  function parseLines(lines) {
    const items = [];
    let started = false;
    for (const raw of lines) {
      const line = String(raw).replace(/\u00a0/g, ' ').trim();
      if (!line) continue;
      const tok = line.split(/\s+/);
      const nums = [];
      let uom = '';
      while (tok.length && (isNum(tok[tok.length - 1]) || (nums.length && !uom && isUom(tok[tok.length - 1])))) {
        const t = tok.pop();
        if (isNum(t)) nums.unshift(num(t)); else uom = t;
        if (nums.length >= 4) break;
      }
      const desc = tok.join(' ');
      if (skipLine.test(line) && nums.length <= 2) continue;
      if (/^(ph|phone|tel|fax|mob(ile)?|abn|acn|bsb|acc(ount)?|a\/c)\b[\s.:#]*/i.test(desc) || /:$/.test(desc)) continue;
      if (nums.length >= 3) {
        let [q, u, t] = nums.slice(-3);
        if (!near(q * u, t)) [q, u, t] = nums.slice(0, 3);
        if (!near(q * u, t)) {
          const [a, b, c] = nums.slice(-3);
          if (near(a * b, c)) { q = b; u = a; t = c; }
        }
        const it = mk(desc, q, u, uom);
        it.amount = t;
        it.check = near(q * u, t) ? 'ok' : 'check';
        items.push(it); started = true;
      } else if (nums.length === 2 && desc && isNum(tok[0]) && near(num(tok[0]) * nums[0], nums[1])) {
        const it = mk(tok.slice(1).join(' '), num(tok[0]), nums[0], uom);
        it.amount = nums[1];
        items.push(it); started = true;
      } else if (nums.length === 2 && desc) {
        const [a, b] = nums;
        const it = a === b ? mk(desc, 1, a, uom) : mk(desc, a, b, uom);
        it.check = 'check';
        items.push(it); started = true;
      } else if (started && !nums.length && items.length && !skipLine.test(line)) {
        items[items.length - 1].description += ' ' + line;
      }
    }
    return items.filter((i) => i.description && !isNaN(i.quantity) && !isNaN(i.unitCost));
  }

  // ---------- Lysaght-style: "Line | ... | Sales Qty | ... | Unit Pr | Line Amount" ----------
  function parseLineBlocks(all) {
    const out = [];
    const pages = [...new Set(all.map((i) => i.page))];
    for (const pg of pages) {
      const its = all.filter((i) => i.page === pg);
      const sameY = (a, b) => Math.abs(a.y - b.y) < 2.5;
      const hLine = its.find((i) => /^line$/i.test(i.s) &&
        its.some((j) => sameY(i, j) && /^unit/i.test(j.s)) && its.some((j) => sameY(i, j) && /^amount$/i.test(j.s)));
      if (!hLine) continue;
      const hdr = its.filter((j) => sameY(j, hLine)).sort((a, b) => a.x - b.x);
      const hUnit = hdr.find((j) => /^unit/i.test(j.s));
      const hAmt = hdr.filter((j) => /^amount$/i.test(j.s)).pop();
      const hSales = hdr.find((j, k) => /^sales$/i.test(j.s) && hdr[k + 1] && /^qty$/i.test(hdr[k + 1].s));
      const rows = [];
      its.filter((i) => i.y < hLine.y - 2).sort((a, b) => b.y - a.y || a.x - b.x).forEach((i) => {
        const r = rows.find((r) => Math.abs(r.y - i.y) < 2.5);
        if (r) r.t.push(i); else rows.push({ y: i.y, t: [i] });
      });
      rows.forEach((r) => r.t.sort((a, b) => a.x - b.x));
      const txt = (r) => r.t.map((i) => i.s).join(' ');
      const isAnchor = (r) => r.t.length > 1 && /^\d{1,3}$/.test(r.t[0].s) && Math.abs(r.t[0].x - hLine.x) < 12 &&
        /^[A-Z0-9][A-Z0-9.\-\/]{2,}$/i.test(r.t[1].s);
      const isStop = (r) => /^(cutting|freight|gst|total|sub\s*total)/i.test(txt(r));
      let cur = null;
      const blocks = [];
      for (const r of rows) {
        if (isAnchor(r)) { cur = { anchor: r, rows: [r] }; blocks.push(cur); continue; }
        if (isStop(r)) { cur = null; continue; }
        if (cur) cur.rows.push(r);
      }
      for (const b of blocks) {
        let unit = NaN, uom = '', amount = NaN, qty = NaN;
        for (const r of b.rows) {
          const u = r.t.find((i) => Math.abs(i.x - hUnit.x) < 30 && /^-?[\d,]*\.?\d+[A-Za-z]{0,4}$/.test(i.s) && /\d/.test(i.s));
          const a = r.t.filter((i) => i.x >= hAmt.x - 70 && isNum(i.s)).pop();
          if (u && a) {
            const m = u.s.match(/^(-?[\d,]*\.?\d+)([A-Za-z]*)$/);
            unit = num(m[1]); uom = m[2] || '';
            amount = num(a.s);
            if (hSales) {
              const q = r.t.find((i) => Math.abs(i.x - hSales.x) < 30 && isNum(i.s));
              if (q) qty = num(q.s);
            }
            break;
          }
        }
        if (isNaN(amount)) continue;
        if (isNaN(qty) && unit) qty = +(amount / unit).toFixed(3);
        if (isNaN(unit) && qty) unit = +(amount / qty).toFixed(4);
        const code = b.anchor.t[1].s;
        const comment = [];
        let on = false;
        for (const r of b.rows) {
          for (const i of r.t) {
            if (/^comments:?$/i.test(i.s)) { on = true; continue; }
            if (on && /^dlvy$/i.test(i.s)) { on = false; break; }
            if (on) comment.push(i.s);
          }
          if (!on && comment.length) break;
        }
        const rest = b.anchor.t.slice(2).map((i) => i.s).join(' ');
        const d = (code + ' - ' + (comment.length ? comment.join(' ') : rest)).replace(/\s+/g, ' ').trim();
        const it = mk(d, qty, unit, uom);
        it.amount = amount;
        it.check = near(qty * unit, amount) ? 'ok' : 'check';
        out.push(it);
      }
      rows.filter((r) => /^(freight|cutting)/i.test(r.t[0].s)).forEach((r) => {
        const n = r.t.filter((i) => isNum(i.s)).pop();
        if (!n || !num(n.s)) return;
        const d = r.t.filter((i) => !isNum(i.s)).map((i) => i.s).join(' ');
        const it = mk(d.charAt(0) + d.slice(1).toLowerCase(), 1, num(n.s), 'EA');
        it.amount = num(n.s);
        out.push(it);
      });
    }
    return out.length ? out : null;
  }

  // ---------- Knotwood-style: Qty first, Description, Unit Rate, Amount ----------
  function parseColumnar(all) {
    const find = (re) => all.find((i) => re.test(i.s));
    const hQty = find(/^(qty|quantity)$/i);
    const hDesc = find(/^(description|item description|details)$/i);
    const hAmt = find(/^(amount|total|line total|ext\.? ?price|value)(\s*(aud|\(?ex\.?.*|\(?inc\.?.*))?$/i);
    const hRate = find(/^(unit\s*rate|rate|unit\s*price|price|unit\s*cost)$/i);
    if (!hQty || !hDesc || !hAmt) return null;
    const numMinX = hRate ? hRate.x - 40 : hDesc.x + 120;
    const out = [];
    const pages = [...new Set(all.map((i) => i.page))];
    for (const pg of pages) {
      if (pg < hQty.page) continue;
      const its = all.filter((i) => i.page === pg);
      const topY = pg === hQty.page ? hQty.y - 2 : Infinity;
      const stops = its.filter((i) => i.y < topY &&
        /^(sub\s*-?total|note:|notes|terms|freight|delivery address|\*+\s*discount)/i.test(i.s)).map((i) => i.y);
      const botY = stops.length ? Math.max(...stops) + 2 : -Infinity;
      const band = its.filter((i) => i.y < topY && i.y > botY);
      const anchors = band.filter((i) => isNum(i.s) && Math.abs(i.x - hQty.x) < 25).map((a) => ({ a, parts: [] }));
      if (!anchors.length) continue;
      band.forEach((i) => {
        if (anchors.some((r) => r.a === i)) return;
        let best = anchors[0];
        anchors.forEach((r) => { if (Math.abs(r.a.y - i.y) < Math.abs(best.a.y - i.y)) best = r; });
        if (Math.abs(best.a.y - i.y) < 24) best.parts.push(i);
      });
      anchors.forEach(({ a, parts }) => {
        const nums = parts.filter((i) => i.x >= numMinX && isNum(i.s)).sort((p, q) => p.x - q.x);
        const text = parts.filter((i) => !nums.includes(i) && !/^\d+(\.\d+)?%$/.test(i.s));
        const left = text.filter((i) => i.x < hDesc.x - 5).sort((p, q) => p.x - q.x);
        const desc = text.filter((i) => i.x >= hDesc.x - 5).sort((p, q) => (q.y - p.y) || (p.x - q.x));
        const qty = num(a.s);
        const amount = nums.length ? num(nums[nums.length - 1].s) : NaN;
        let rate = nums.length >= 2 ? num(nums[nums.length - 2].s) : NaN;
        if (isNaN(rate) && !isNaN(amount) && qty) rate = +(amount / qty).toFixed(4);
        const d = [...left, ...desc].map((i) => i.s).join(' ').replace(/\s+/g, ' ');
        const it = mk(d, qty, rate, '');
        it.amount = amount;
        it.check = near(qty * rate, amount) ? 'ok' : 'check';
        out.push(it);
      });
    }
    return out.length ? out.filter((i) => i.description && !isNaN(i.unitCost)) : null;
  }

  // ---------- general one-row-per-item tables ----------
  function parseTable(all) {
    const RE = {
      qty: /^(qty|quantity)(\s+(ordered|ord\.?|req(uired)?|supplied|shipped))?$/i,
      price: /^(unit\s*(price|rate|cost|pr\.?)|price|rate|each)(\s*\$)?(\s*\(?(ex|inc).*)?$/i,
      unit: /^(unit|uom|units|measure|u\/m)$/i,
      desc: /^(.*\/\s*)?(item\s*|product\s*)?description$|^(details|product)$/i,
      code: /^(stock\s*code|item\s*code|product\s*code|part\s*(no\.?|number|#)|code|sku|item)$/i,
      tax: /^(tax|gst|tax\s*code)$/i,
      amount: /^((amount|total|line\s*total|ext\.?\s*(price|total)?|extended(\s*price)?|value|net|nett)(\s.*)?|\$|aud|\$\s*aud)$/i,
    };
    const stopRe = /^(sub\s*-?total|ex\.?\s*tax|total|tax:?|notes?:?|terms|payment|accepted|no\.?\s*of\s*items|sales order totals|prop:|tech:|account|bank)/i;
    const out = [];
    const pages = [...new Set(all.map((i) => i.page))];
    for (const pg of pages) {
      const rows = [];
      all.filter((i) => i.page === pg).sort((a, b) => b.y - a.y || a.x - b.x).forEach((i) => {
        const r = rows.find((r) => Math.abs(r.y - i.y) < 3);
        if (r) r.t.push(i); else rows.push({ y: i.y, t: [i] });
      });
      rows.forEach((r) => r.t.sort((a, b) => a.x - b.x));
      const hdr = rows.find((r) => r.t.some((i) => RE.qty.test(i.s)) &&
        r.t.some((i) => RE.amount.test(i.s) && !RE.price.test(i.s)));
      if (!hdr) continue;
      const cols = {};
      hdr.t.forEach((i) => {
        for (const k of ['qty', 'price', 'unit', 'desc', 'code', 'tax']) {
          if (!cols[k] && RE[k].test(i.s)) { cols[k] = i; return; }
        }
        if (RE.amount.test(i.s)) cols.amount = i;
      });
      if (!cols.amount) continue;
      const wOf = (i) => i.w || i.s.length * 5;
      const heads = Object.entries(cols).map(([k, h]) => ({ k, a: h.x, b: h.x + wOf(h) }));
      const NUMCOL = { qty: 1, price: 1, amount: 1 };
      const colOf = (i) => {
        const c = i.x + wOf(i) / 2;
        const numeric = isNum(i.s);
        let best = null, bd = Infinity;
        heads.forEach((h) => {
          if (!numeric && NUMCOL[h.k]) return;
          const d = c < h.a ? h.a - c : c > h.b ? c - h.b : 0;
          if (d < bd) { bd = d; best = h.k; }
        });
        return best || 'desc';
      };
      let cur = null, lastY = hdr.y;
      for (const r of rows.filter((r) => r.y < hdr.y - 2)) {
        const f = {};
        r.t.forEach((i) => { const k = colOf(i); (f[k] = f[k] || []).push(i.s); });
        const get = (k) => (f[k] || []).join(' ').replace(/\s+/g, ' ').trim();
        const qty = num(get('qty')), amount = num(get('amount'));
        if (r.t.some((i) => stopRe.test(i.s)) && isNaN(qty)) break;
        if (!isNaN(qty) && !isNaN(amount)) {
          if (!qty && !amount) continue;
          let price = num(get('price'));
          let q = qty;
          let d = [get('code'), get('desc')].filter(Boolean).join(' ') || r.t.map((i) => i.s).join(' ');
          if (isNaN(price)) {
            const each = qty ? amount / qty : amount;
            if (qty && Math.abs(each * 100 - Math.round(each * 100)) < 1e-6) price = r2(each);
            else { price = amount; if (qty && qty !== 1) d = qty + ' x ' + d; q = 1; }
          }
          const it = mk(d, q, price, get('unit'));
          it.amount = amount;
          it.check = near(q * price, amount) ? 'ok' : 'check';
          out.push(it);
          cur = it; lastY = r.y;
        } else if (!isNaN(amount)) {
          break;
        } else if (cur && get('desc') && !get('qty') && lastY - r.y < 22) {
          cur.description += ' ' + get('desc');
          lastY = r.y;
        }
      }
    }
    return out.length ? out : null;
  }

  // ---------- PDF ----------
  async function pdfToItems(bytes) {
    const pdf = await g.pdfjsLib.getDocument({ data: bytes, verbosity: 0 }).promise;
    const out = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const tc = await (await pdf.getPage(p)).getTextContent();
      tc.items.forEach((it) => {
        const t = String(it.str || '').trim();
        if (t) out.push({ page: p, x: it.transform[4], y: it.transform[5], w: it.width || 0, s: t });
      });
    }
    return out;
  }

  function textLines(all) {
    const out = [];
    [...new Set(all.map((i) => i.page))].forEach((p) => {
      const rows = {};
      all.filter((i) => i.page === p).forEach((i) => {
        const y = Math.round(i.y / 3) * 3;
        (rows[y] = rows[y] || []).push(i);
      });
      Object.keys(rows).map(Number).sort((a, b) => b - a)
        .forEach((y) => out.push(rows[y].sort((a, b) => a.x - b.x).map((r) => r.s).join(' ')));
    });
    return out;
  }

  async function readPdf(bytes) {
    const all = await pdfToItems(bytes);
    let lines = parseLineBlocks(all), how = 'line-numbered quote';
    if (!lines) { lines = parseColumnar(all); how = 'Qty / Description / Amount columns'; }
    if (!lines) { lines = parseTable(all); how = 'table, one row per item'; }
    const fromLayout = !!lines;
    if (!lines) { lines = parseLines(textLines(all)); how = 'line by line - check carefully'; }
    // "incl. GST" shown on its own and no separate GST amount -> convert to ex GST
    const incGst = fromLayout &&
      all.some((i) => /^inc(l(uding|\.)?)?\.?\s*gst\.?$/i.test(i.s)) &&
      !all.some((i) => /^(total\s*)?(gst|tax)(\s*10%)?:?$/i.test(i.s));
    if (incGst) {
      lines.forEach((it) => {
        const amt = !isNaN(it.amount) ? it.amount : it.quantity * it.unitCost;
        it.amount = r2(amt / 1.1);
        it.unitCost = it.quantity ? +(amt / 1.1 / it.quantity).toFixed(4) : it.amount;
      });
    }
    return { lines, how, incGst };
  }

  // ---------- EMAIL ----------
  const REQUEST_RE = new RegExp([
    'purchase order request',
    'received without (a )?purchase order',
    '(raise|create|do|organi[sz]e|need|send( through)?|give (you|them|us))\\s+(me\\s+)?(a\\s+|the\\s+|an\\s+)?(po|purchase order)\\b',
    'po (needed|required|please)',
  ].join('|'), 'i');
  const COMPANY_RE = /\b([A-Z][\w&.'’-]*(?:[ \t]+[A-Z&][\w&.'’-]*){0,5}[ \t]+(?:Pty\.?[ \t]*Ltd\.?|Pty[ \t]+Limited|Limited|Ltd\.?|Group))(?![\w])/g;
  const OURS = /lennox|jet facades/i;
  const cleanCo = (c) => c.replace(/^((tax|proforma|pro-forma)\s+)*(invoice|quotation|quote)\s+/i, '').replace(/\.$/, '').trim();

  function parseRequest(subject, body) {
    const t = (subject || '') + '\n' + (body || '');
    const first = (...res) => {
      for (const re of res) { const m = t.match(re); if (m && m[1]) return m[1].replace(/\s+/g, ' ').trim(); }
      return '';
    };
    let supplier = first(/invoice\s*#?\s*:?\s*[\w-]+\s+from\s+(.+?)(?=\s+for\s+|\s+is\s+due|\s*\n|$)/i);
    if (!supplier || OURS.test(supplier)) {
      const cos = [...t.matchAll(COMPANY_RE)].map((m) => cleanCo(m[1])).filter((c) => c && !OURS.test(c));
      supplier = cos[0] || '';
    }
    const ref = first(
      /invoice\s*(?:#|no\.?|number)?\s*:?\s*([A-Z]{0,4}-?\d{4,}(?:-\d+)?)\b/i,
      /\binv(?:oice)?[-\s]?#?\s*:?\s*(\d{4,})/i,
      /\b(INV-?\d{3,})\b/i,
      /\b(SO[A-Z]?\d{5,})\b/,
      /quot(?:e|ation)\s*(?:#|no\.?|number)?\s*:?\s*([A-Z]{0,4}[\/-]?\d{3,})\b/i,
    );
    const job = first(/\bjob(?:\s*ref)?\s*[:.\-]\s*(.+?)(?=\s+COS\b|\s{2,}|\n|$)/i);
    const cos = first(/\bCOS\s*[:.\-]\s*(.+?)(?=\s{2,}|\n|$)/i);
    const amount = first(/(?:amount due|invoice total|total(?: amount)?(?: \(?inc\.? ?(?:gst|tax)\)?)?)\s*\(?(?:aud)?\)?\s*:?\s*\$?\s*([\d,]+\.\d{2})/i);
    const existingPo = first(/\b(PO\d{4,})\b/);
    return { supplier, ref, job, cos, amount, existingPo };
  }

  g.QPO = {
    readPdf, parseRequest, isUom, r2,
    isRequest: (subject, body) => REQUEST_RE.test((subject || '') + '\n' + (body || '')),
  };
})(typeof window !== 'undefined' ? window : globalThis);
