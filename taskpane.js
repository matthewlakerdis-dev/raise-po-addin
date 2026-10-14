/* global Office, QPO */
/* Raise PO - Outlook task pane */
(function () {
  'use strict';

  const DEFAULT_NEXVIA = 'https://lennox.nexvia.com.au/staff/accounts/purchase-orders';
  const PDFJS = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/';
  const $ = (id) => document.getElementById(id);
  const esc = (x) => String(x == null ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const money = (v) => isNaN(v) ? '-' : '$' + Number(v).toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const state = { subject: '', lines: [], pdfs: [], pdfName: '', pdfBytes: null };
  let loadSeq = 0;

  // ---------- settings (roam with the mailbox) ----------
  function getSetting(k) {
    try { const v = Office.context.roamingSettings.get(k); if (v) return v; } catch (e) { /* ignore */ }
    return localStorage.getItem('raisepo_' + k);
  }
  function setSetting(k, v) {
    localStorage.setItem('raisepo_' + k, v);
    try { Office.context.roamingSettings.set(k, v); Office.context.roamingSettings.saveAsync(); } catch (e) { /* ignore */ }
  }

  const status = (msg, err) => { $('status').textContent = msg || ''; $('status').className = 'small center' + (err ? ' err' : ''); };
  const pdfStatus = (msg, err) => { $('pdfStatus').textContent = msg || ''; $('pdfStatus').className = 'small' + (err ? ' err' : ''); };

  // ---------- Office helpers ----------
  const ok = (r) => r.status === Office.AsyncResultStatus.Succeeded;
  function getBody(item) {
    return new Promise((res) => item.body.getAsync(Office.CoercionType.Text, (r) => res(ok(r) ? r.value : '')));
  }
  function getAttachment(item, id) {
    return new Promise((res, rej) => item.getAttachmentContentAsync(id, (r) =>
      ok(r) ? res(r.value) : rej(new Error((r.error && r.error.message) || 'Could not read the attachment'))));
  }
  function b64ToBytes(b64) {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  // ---------- load the open email ----------
  function resetUi() {
    ['supplier', 'ref', 'amount', 'job', 'cos'].forEach((id) => { $(id).value = ''; });
    $('warn').classList.add('hidden');
    $('pdfSel').innerHTML = '';
    $('dl').classList.add('hidden');
    pdfStatus('');
    status('');
    state.lines = []; state.pdfs = []; state.pdfName = ''; state.pdfBytes = null;
    renderLines();
  }

  async function load() {
    const seq = ++loadSeq;
    const item = Office.context.mailbox.item;
    resetUi();
    if (!item) { $('badge').textContent = 'Open an email to raise a PO from it.'; return; }

    state.subject = item.subject || '';
    const body = await getBody(item);
    if (seq !== loadSeq) return;

    const req = QPO.parseRequest(state.subject, body);
    $('badge').textContent = QPO.isRequest(state.subject, body)
      ? '📋 PO request found'
      : 'Not flagged as a PO request - you can still raise one';
    $('supplier').value = req.supplier;
    $('ref').value = req.ref;
    $('amount').value = req.amount;
    $('job').value = req.job;
    $('cos').value = req.cos;
    if (req.existingPo) {
      $('warn').textContent = `⚠ ${req.existingPo} is mentioned in this thread - check a PO hasn't already been raised.`;
      $('warn').classList.remove('hidden');
    }

    // PDFs attached to this email (ignores inline images and terms & conditions where possible)
    const File = Office.MailboxEnums.AttachmentType.File;
    state.pdfs = (item.attachments || []).filter((a) => a.attachmentType === File && /\.pdf$/i.test(a.name) && !a.isInline);
    if (!state.pdfs.length) {
      pdfStatus('No PDF attached to this email. If the quote is in an earlier email in the thread, open that one.');
      return;
    }
    if (!item.getAttachmentContentAsync) {
      pdfStatus('This version of Outlook can\'t read attachments from an add-in. Please update Outlook.', true);
      return;
    }
    const score = (n) => (/terms|condition|t&c|warranty|brochure|data ?sheet|sds/i.test(n) ? -10 : 0) +
      (/quot|invoice|inv|proforma|pro-forma|order|\bso\d|sor\d/i.test(n) ? 5 : 0) +
      (/purchase.?order|_PO\d|PO\d{4}/i.test(n) ? -3 : 0);
    const order = state.pdfs.map((a, i) => i).sort((a, b) => score(state.pdfs[b].name) - score(state.pdfs[a].name));
    $('pdfSel').innerHTML = state.pdfs.map((a, i) => `<option value="${i}">${esc(a.name)}</option>`).join('');
    $('pdfSel').value = String(order[0]);
    await readSelected(seq);
  }

  async function readSelected(seq) {
    const item = Office.context.mailbox.item;
    const att = state.pdfs[+$('pdfSel').value];
    if (!att) return;
    pdfStatus('Reading ' + att.name + '…');
    $('dl').classList.add('hidden');
    try {
      const content = await getAttachment(item, att.id);
      if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        throw new Error('This attachment type can\'t be read (' + content.format + ')');
      }
      const bytes = b64ToBytes(content.content);
      const res = await QPO.readPdf(bytes.slice(0));       // pdf.js takes ownership of the copy
      if (seq !== loadSeq) return;
      state.pdfBytes = bytes;
      state.pdfName = att.name;
      state.lines = res.lines;
      renderLines();
      const checks = res.lines.filter((l) => l.check !== 'ok').length;
      pdfStatus(`Read ${res.lines.length} line(s) - ${res.how}.` +
        (res.incGst ? ' Prices included GST, converted to ex GST.' : '') +
        (checks ? ` ${checks} to check (yellow).` : ''));
      $('dl').classList.remove('hidden');
    } catch (e) {
      if (seq !== loadSeq) return;
      pdfStatus('Couldn\'t read the PDF: ' + e.message, true);
    }
  }

  // ---------- lines table ----------
  function renderLines() {
    const box = $('lines');
    if (!state.lines.length) { box.className = 'small'; box.textContent = 'No lines yet.'; $('tot').innerHTML = ''; return; }
    box.className = '';
    let sum = 0;
    const rows = state.lines.map((l, i) => {
      const amt = QPO.r2(l.quantity * l.unitCost);
      sum += amt;
      return `<tr class="${l.check !== 'ok' ? 'check' : ''}">
        <td><input class="desc" data-i="${i}" data-k="description" value="${esc(l.description)}" title="${esc(l.description)}"></td>
        <td style="width:52px"><input data-i="${i}" data-k="quantity" value="${l.quantity}"></td>
        <td style="width:64px"><input data-i="${i}" data-k="unitCost" value="${l.unitCost}"></td>
        <td class="amt">${money(amt)}</td>
        <td><span class="del" data-del="${i}" title="Remove line">✕</span></td></tr>`;
    }).join('');
    box.innerHTML = `<table><tr><th>Description</th><th>Qty</th><th>Unit $</th><th class="amt">Amount</th><th></th></tr>${rows}</table>`;
    $('tot').innerHTML = `<span>${state.lines.length} line(s)${state.lines[0] && state.lines[0].measure ? '' : ''}</span><span>Total ex GST <b>${money(QPO.r2(sum))}</b></span>`;
    box.querySelectorAll('input').forEach((inp) => {
      inp.onchange = () => {
        const l = state.lines[+inp.dataset.i];
        const k = inp.dataset.k;
        l[k] = k === 'description' ? inp.value : parseFloat(String(inp.value).replace(/[$,\s]/g, ''));
        l.check = 'ok';
        renderLines();
      };
    });
    box.querySelectorAll('[data-del]').forEach((x) => {
      x.onclick = () => { state.lines.splice(+x.dataset.del, 1); renderLines(); };
    });
  }

  // ---------- hand-off to Nexvia ----------
  function buildUrl() {
    const base = ($('nexviaUrl').value || DEFAULT_NEXVIA).trim().split('#')[0];
    const clean = (s) => String(s || '').replace(/[\t\r\n]+/g, ' ').trim();
    const payload = {
      supplier: clean($('supplier').value), ref: clean($('ref').value), job: clean($('job').value),
      cos: clean($('cos').value), amount: clean($('amount').value),
      subject: state.subject.slice(0, 150), pdfs: state.pdfName ? [state.pdfName] : [],
      lines: state.lines.filter((l) => l.description && !isNaN(l.quantity) && !isNaN(l.unitCost)).map((l) => ({
        d: clean(l.description).slice(0, 250), q: l.quantity, u: l.unitCost, m: QPO.isUom(l.measure) ? l.measure : '',
      })),
      src: 'outlook-addin', t: Date.now(),
    };
    return base + '#qporeq=' + encodeURIComponent(JSON.stringify(payload));
  }

  function openInNexvia() {
    const url = buildUrl();
    try {
      if (Office.context.ui && Office.context.ui.openBrowserWindow &&
          Office.context.requirements.isSetSupported('OpenBrowserWindowApi', '1.1')) {
        Office.context.ui.openBrowserWindow(url);
      } else {
        window.open(url, '_blank');
      }
      status('Nexvia opened in your browser. Create the PO and the Quote → PO panel will fill in.');
    } catch (e) {
      window.open(url, '_blank');
    }
  }

  async function copyLink() {
    const url = buildUrl();
    try { await navigator.clipboard.writeText(url); status('Link copied - paste it into your browser.'); return; } catch (e) { /* fall back */ }
    const ta = document.createElement('textarea');
    ta.value = url; document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); status('Link copied - paste it into your browser.'); } catch (e) { status('Couldn\'t copy the link.', true); }
    ta.remove();
  }

  function savePdf() {
    if (!state.pdfBytes) return;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([state.pdfBytes], { type: 'application/pdf' }));
    a.download = state.pdfName || 'quote.pdf';
    document.body.appendChild(a); a.click(); a.remove();
  }

  // ---------- start ----------
  Office.onReady((info) => {
    if (window.pdfjsLib) window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS + 'pdf.worker.min.js';
    $('nexviaUrl').value = getSetting('nexviaUrl') || DEFAULT_NEXVIA;
    $('saveSettings').onclick = () => { setSetting('nexviaUrl', $('nexviaUrl').value.trim()); status('Settings saved.'); };
    $('open').onclick = openInNexvia;
    $('copyLink').onclick = copyLink;
    $('dl').onclick = savePdf;
    $('pdfSel').onchange = () => readSelected(++loadSeq);
    if (info.host !== Office.HostType.Outlook) { $('badge').textContent = 'Open this from Outlook.'; return; }
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => load()); } catch (e) { /* not pinned */ }
    load();
  });
})();
