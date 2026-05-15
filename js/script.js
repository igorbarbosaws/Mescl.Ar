// ─────────────────────────────────────────────────────────────────────────────
// STATE
// ─────────────────────────────────────────────────────────────────────────────
const state = {
  baseWb: null, baseData: [], baseHeaders: [], baseMode: 'exact',
  lookups: [],
  result: [],
  filtered: [],
  activeFilter: 'all',
  searchTerm: '',
  lookupColKeys: [],   // ordered list of lookup column names, set after consolidation
};

const COLORS = ['#3b82f6','#10b981','#a855f7','#f59e0b','#06b6d4','#f43f5e','#84cc16'];
let lookupCounter = 0;

// ─────────────────────────────────────────────────────────────────────────────
// CLOCK
// ─────────────────────────────────────────────────────────────────────────────
setInterval(() => {
  document.getElementById('clock').textContent = new Date().toLocaleTimeString('pt-BR');
}, 1000);

// ─────────────────────────────────────────────────────────────────────────────
// BASE SHEET
// ─────────────────────────────────────────────────────────────────────────────
function loadBase(input) {
  const file = input.files[0];
  if (!file) return;
  readWorkbook(file, (wb) => {
    state.baseWb = wb;
    populateSheetSel('baseSheet', wb, () => reloadBase());
  });
}

function reloadBase() {
  if (!state.baseWb) return;
  const sheetSel = document.getElementById('baseSheet');
  const sheetIdx = sheetSel.style.display !== 'none' ? parseInt(sheetSel.value) : 0;
  const hRow     = parseInt(document.getElementById('baseHeader').value);
  const { headers, data } = parseSheet(state.baseWb, sheetIdx, hRow);
  state.baseHeaders = headers;
  state.baseData    = data;
  populateSelect('baseKey', headers);
  document.getElementById('baseStat').textContent = `${data.length} linhas · ${headers.length} colunas`;
  document.getElementById('baseFname').style.display = 'block';
  document.getElementById('baseFname').textContent   = '📄 ' + state.baseWb.SheetNames[sheetIdx];
  document.getElementById('baseZone').classList.add('loaded');
  toast(`Base carregada — ${data.length} linhas`, 'ok');
  setStep(2);
  checkReady();
}

function setBaseMode(mode, el) {
  state.baseMode = mode;
  document.querySelectorAll('.mode-pills .mode-pill').forEach(b => b.classList.toggle('active', b === el));
  document.getElementById('baseRangeOpts').classList.toggle('visible', mode === 'range');
  checkReady();
}

function onBaseKeyChange() { checkReady(); }

// ─────────────────────────────────────────────────────────────────────────────
// LOOKUP MANAGEMENT
// ─────────────────────────────────────────────────────────────────────────────
function addLookup() {
  const id    = ++lookupCounter;
  const color = COLORS[(id - 1) % COLORS.length];
  const obj   = { id, wb: null, data: [], headers: [], keyCol: 0, cols: [], mode: 'exact', sep: 'a', name: `Lookup ${id}`, color };
  state.lookups.push(obj);
  renderLookupCard(obj);
  setStep(2);
  checkReady();
}

function renderLookupCard(lk) {
  const list = document.getElementById('lookupList');
  const card = document.createElement('div');
  card.className = 'lookup-card';
  card.id = `lk-${lk.id}`;
  card.style.borderColor = hexAlpha(lk.color, 0.25);

  card.innerHTML = `
  <div class="lookup-header" onclick="toggleLookup(${lk.id})">
    <div class="lookup-badge" style="background:${lk.color}">${lk.id}</div>
    <div class="lookup-title-area">
      <div class="lookup-title">
        <input type="text" value="${escHtml(lk.name)}"
          onclick="event.stopPropagation()"
          oninput="renameLookup(${lk.id}, this.value)"
          placeholder="Nome desta fonte">
      </div>
      <div class="lookup-sub" id="lk-sub-${lk.id}">Aguardando upload…</div>
    </div>
    <div class="lookup-actions">
      <button class="btn btn-ghost" style="font-size:0.75rem;padding:4px 10px"
        onclick="event.stopPropagation(); removeLookup(${lk.id})">✕ Remover</button>
      <span class="lookup-chevron open" id="lk-chev-${lk.id}">▼</span>
    </div>
  </div>
  <div class="lookup-body open" id="lk-body-${lk.id}">
    <div class="lookup-grid">
      <!-- Coluna esquerda -->
      <div>
        <div class="upload-zone" style="border-color:${hexAlpha(lk.color,0.3)}" id="lk-zone-${lk.id}"
          onclick="document.getElementById('lk-file-${lk.id}').click()">
          <input type="file" id="lk-file-${lk.id}" accept=".xls,.xlsx,.ods,.csv"
            onchange="loadLookup(${lk.id}, this)">
          <div class="uz-icon">🔎</div>
          <div class="uz-title">Arquivo desta fonte</div>
          <div class="uz-sub">ex: .xlsx, .xls, .csv</div>
          <button class="btn" style="font-size:0.75rem;padding:6px 14px">Selecionar</button>
          <div class="uz-fname" id="lk-fname-${lk.id}" style="display:none"></div>
        </div>
        <div class="field" style="margin-top:12px">
          <label>Aba (sheet)</label>
          <select id="lk-sheet-${lk.id}" onchange="reloadLookup(${lk.id})" style="display:none"></select>
          <div id="lk-sheet-none-${lk.id}" style="font-size:0.72rem;color:var(--muted);font-family:var(--mono)">—</div>
        </div>
        <div class="field">
          <label>Linha do cabeçalho</label>
          <select id="lk-header-${lk.id}" onchange="reloadLookup(${lk.id})">
            <option value="0">Linha 1</option>
            <option value="1">Linha 2</option>
            <option value="2">Linha 3</option>
            <option value="3">Linha 4</option>
            <option value="4">Linha 5</option>
          </select>
        </div>
        <div class="field">
          <label>Coluna-chave (para casar com a Base)</label>
          <select id="lk-key-${lk.id}" onchange="onLookupKeyChange(${lk.id})">
            <option value="">— aguardando upload —</option>
          </select>
        </div>
        <div class="field">
          <label>Modo de match</label>
          <div class="mode-pills" id="lk-mode-pills-${lk.id}">
            <button class="mode-pill active" onclick="setLookupMode(${lk.id},'exact',this)">Valor exato</button>
            <button class="mode-pill" onclick="setLookupMode(${lk.id},'range',this)">Intervalo / range</button>
          </div>
          <div class="range-opts" id="lk-range-${lk.id}">
            <label>Palavra separadora de intervalo</label>
            <input type="text" id="lk-sep-${lk.id}" value="a" placeholder="a"
              oninput="state.lookups.find(l=>l.id==${lk.id}).sep=this.value">
          </div>
        </div>
        <div class="preview-chip" id="lk-prev-${lk.id}"></div>
      </div>

      <!-- Coluna direita -->
      <div>
        <div class="field">
          <label>Colunas a trazer desta fonte</label>
          <div class="cols-bring-list" id="lk-cols-${lk.id}"></div>
          <button class="btn-add-col" onclick="addColBring(${lk.id})">＋ Adicionar coluna</button>
        </div>
        <div class="field" style="margin-top:16px">
          <label>Conflito de valores (quando houver múltiplos matches)</label>
          <select id="lk-agg-${lk.id}">
            <option value="join">Unir com vírgula</option>
            <option value="abbrev">Abreviar sequências (ex: 100 a 102, 104, 107)</option>
            <option value="first">Pegar o primeiro</option>
            <option value="last">Pegar o último</option>
            <option value="count">Contar ocorrências</option>
          </select>
        </div>
        <div class="field">
          <label>Prefixo nas colunas trazidas (opcional)</label>
          <input type="text" id="lk-prefix-${lk.id}" placeholder="ex: Redespacho_" style="padding:9px 12px;background:var(--surface);border:1px solid var(--border2);color:var(--text);border-radius:var(--radius-sm);font-family:var(--mono);font-size:0.8rem;outline:none;width:100%">
        </div>
      </div>
    </div>
  </div>`;

  list.appendChild(card);
}

function toggleLookup(id) {
  const body = document.getElementById(`lk-body-${id}`);
  const chev = document.getElementById(`lk-chev-${id}`);
  const open = body.classList.toggle('open');
  chev.classList.toggle('open', open);
}

function removeLookup(id) {
  state.lookups = state.lookups.filter(l => l.id !== id);
  document.getElementById(`lk-${id}`)?.remove();
  checkReady();
  toast('Fonte removida', 'ok');
}

function renameLookup(id, name) {
  const lk = state.lookups.find(l => l.id === id);
  if (lk) lk.name = name;
}

function loadLookup(id, input) {
  const file = input.files[0];
  if (!file) return;
  readWorkbook(file, (wb) => {
    const lk = state.lookups.find(l => l.id === id);
    if (!lk) return;
    lk.wb = wb;
    populateSheetSel(`lk-sheet-${id}`, wb, () => reloadLookup(id));
    document.getElementById(`lk-fname-${id}`).style.display  = 'block';
    document.getElementById(`lk-fname-${id}`).textContent    = '📄 ' + file.name;
  });
}

function reloadLookup(id) {
  const lk = state.lookups.find(l => l.id === id);
  if (!lk || !lk.wb) return;
  const sheetEl = document.getElementById(`lk-sheet-${id}`);
  const sheetIdx = sheetEl.style.display !== 'none' ? parseInt(sheetEl.value) : 0;
  const hRow     = parseInt(document.getElementById(`lk-header-${id}`).value);
  const { headers, data } = parseSheet(lk.wb, sheetIdx, hRow);
  lk.headers = headers;
  lk.data    = data;
  populateSelect(`lk-key-${id}`, headers);
  document.getElementById(`lk-sub-${id}`).textContent = `${data.length} linhas · ${headers.length} colunas`;
  document.getElementById(`lk-zone-${id}`).classList.add('loaded');
  // Rebuild cols bring
  lk.cols = [];
  const colsEl = document.getElementById(`lk-cols-${id}`);
  colsEl.innerHTML = '';
  addColBring(id);  // add one default
  updateLookupPreview(id);
  toast(`Lookup ${lk.name} carregado — ${data.length} linhas`, 'ok');
  checkReady();
}

function onLookupKeyChange(id) {
  const lk = state.lookups.find(l => l.id === id);
  if (!lk) return;
  lk.keyCol = parseInt(document.getElementById(`lk-key-${id}`).value);
  updateLookupPreview(id);
  checkReady();
}

function setLookupMode(id, mode, el) {
  const lk = state.lookups.find(l => l.id === id);
  if (!lk) return;
  lk.mode = mode;
  document.querySelectorAll(`#lk-mode-pills-${id} .mode-pill`).forEach(b => b.classList.toggle('active', b === el));
  document.getElementById(`lk-range-${id}`).classList.toggle('visible', mode === 'range');
}

// ─────────────────────────────────────────────────────────────────────────────
// COLUMNS TO BRING
// ─────────────────────────────────────────────────────────────────────────────
function addColBring(id) {
  const lk = state.lookups.find(l => l.id === id);
  if (!lk) return;
  const colsEl = document.getElementById(`lk-cols-${id}`);
  const rowIdx = lk.cols.length;
  const colObj = { colIdx: 0, label: '' };
  lk.cols.push(colObj);

  const row = document.createElement('div');
  row.className = 'col-bring-row';
  row.id = `lk-col-row-${id}-${rowIdx}`;
  row.innerHTML = `
    <select onchange="setColBring(${id},${rowIdx},'col',parseInt(this.value))">
      ${lk.headers.map((h,i) => `<option value="${i}">${escHtml(h || 'col'+i)}</option>`).join('')}
    </select>
    <input class="col-label-input" type="text" placeholder="Nome no resultado (opcional)"
      oninput="setColBring(${id},${rowIdx},'label',this.value)">
    <button class="btn-rm-col" onclick="removeColBring(${id},${rowIdx})">✕</button>`;
  colsEl.appendChild(row);
}

function setColBring(id, rowIdx, field, val) {
  const lk = state.lookups.find(l => l.id === id);
  if (!lk || !lk.cols[rowIdx]) return;
  lk.cols[rowIdx][field === 'col' ? 'colIdx' : 'label'] = val;
}

function removeColBring(id, rowIdx) {
  const lk = state.lookups.find(l => l.id === id);
  if (!lk) return;
  lk.cols.splice(rowIdx, 1);
  document.getElementById(`lk-col-row-${id}-${rowIdx}`)?.remove();
}

// ─────────────────────────────────────────────────────────────────────────────
// LOOKUP PREVIEW
// ─────────────────────────────────────────────────────────────────────────────
function updateLookupPreview(id) {
  const lk  = state.lookups.find(l => l.id === id);
  const el  = document.getElementById(`lk-prev-${id}`);
  if (!lk || !el || !lk.data.length) { el && (el.style.display = 'none'); return; }

  const samples = [];
  for (const row of lk.data) {
    if (samples.length >= 3) break;
    const key = String(row[lk.keyCol] ?? '').trim();
    if (!key) continue;
    samples.push(key);
  }

  el.style.display = 'block';
  el.className = 'preview-chip ok';
  el.innerHTML = `<b>Amostra de chaves:</b> ${samples.map(k => escHtml(k)).join(' · ')}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// CONSOLIDATION ENGINE
// ─────────────────────────────────────────────────────────────────────────────
function runConsolidate() {
  if (!state.baseData.length) return;

  setStep(3);
  progress(0, true);
  updateProcessStatus('Construindo índices…', '');

  setTimeout(() => {
    try {
      _consolidate();
    } catch (e) {
      toast('Erro: ' + e.message, 'err');
      console.error(e);
      progress(0, false);
    }
  }, 30);
}

function _consolidate() {
  const baseKeyCol = parseInt(document.getElementById('baseKey').value);
  const baseMode   = state.baseMode;

  // Collect active separators from checkboxes + custom field
  const baseSeps = getBaseSeps();

  // ── Build lookup indices ──────────────────────────────────────────
  const indices = state.lookups.map(lk => {
    const idx = {};
    lk.data.forEach(row => {
      const rawKey = String(row[lk.keyCol] ?? '').trim();
      if (!rawKey) return;
      // For lookup source: index by exact normalized key
      const nk = normKey(rawKey);
      if (!idx[nk]) idx[nk] = [];
      idx[nk].push(row);
    });
    return idx;
  });

  progress(20);

  // ── Process each base row ─────────────────────────────────────────
  const result = [];
  const total  = state.baseData.length;

  state.baseData.forEach((baseRow, ri) => {
    const rawCell = baseRow[baseKeyCol];

    // Keys from base (possibly expanded if range mode)
    let baseKeys;
    if (baseMode === 'range') {
      baseKeys = expandRange(rawCell, baseSeps);
    } else {
      const nk = normKey(String(rawCell ?? ''));
      baseKeys = nk ? [nk] : [];
    }

    // For each lookup, find all matching rows
    const lookupResults = state.lookups.map((lk, li) => {
      const idx    = indices[li];
      const agg    = document.getElementById(`lk-agg-${lk.id}`)?.value || 'join';
      // Auto-prefix: use user-defined prefix if set, otherwise "LookupName_"
      // This guarantees no key collision between different lookups
      const userPrefix = (document.getElementById(`lk-prefix-${lk.id}`)?.value || '').trim();
      const autoPrefix = userPrefix || (lk.name.replace(/\s+/g, '_') + '_');

      const hitRows = [];
      baseKeys.forEach(bk => {
        const hits = idx[bk];
        if (hits) hits.forEach(r => hitRows.push(r));
      });

      // Aggregate columns to bring
      // Use a Map to preserve insertion order and avoid key collisions
      const colVals = new Map();
      lk.cols.forEach((colCfg) => {
        const colIdx   = colCfg.colIdx;
        // Column name: user label → original header → fallback
        const rawLabel = (colCfg.label || lk.headers[colIdx] || `col${colIdx}`).trim();
        const colName  = autoPrefix + rawLabel;
        const rawVals  = hitRows.map(r => String(r[colIdx] ?? '').trim()).filter(Boolean);
        const unique   = [...new Set(rawVals)];

        let val = '';
        if (agg === 'join')   val = unique.join(', ');
        if (agg === 'abbrev') val = abbreviateNums(unique);
        if (agg === 'first')  val = unique[0] || '';
        if (agg === 'last')   val = unique[unique.length - 1] || '';
        if (agg === 'count')  val = unique.length;
        colVals.set(colName, val);
      });

      return { hitCount: hitRows.length, colVals };
    });

    // Determine overall status
    const anyMatch  = lookupResults.some(r => r.hitCount > 0);
    const allMatch  = lookupResults.every(r => r.hitCount > 0);
    const status    = state.lookups.length === 0 ? 'none'
                    : allMatch  ? 'match'
                    : anyMatch  ? 'part'
                    : 'none';

    // Build result row — column order: Base cols → Lookup1 cols → Lookup2 cols → ...
    // Never appended at the end; each lookup's columns come right after the previous one.
    const outRow = { _status: status, _raw_key: String(rawCell ?? ''), _expanded_keys: baseKeys.join(', ') };
    state.baseHeaders.forEach((h, i) => { outRow[`Base_${h || 'col'+i}`] = baseRow[i] ?? ''; });
    lookupResults.forEach(lr => {
      lr.colVals.forEach((val, key) => { outRow[key] = val; });
    });

    result.push(outRow);

    if (ri % 100 === 0) progress(20 + Math.round((ri / total) * 70));
  });

  progress(100);

  // Store the exact ordered list of lookup column keys for render/export
  state.lookupColKeys = state.lookups.flatMap((lk) => {
    const userPrefix = (document.getElementById(`lk-prefix-${lk.id}`)?.value || '').trim();
    const autoPrefix = userPrefix || (lk.name.replace(/\s+/g, '_') + '_');
    return lk.cols.map(colCfg => {
      const rawLabel = (colCfg.label || lk.headers[colCfg.colIdx] || `col${colCfg.colIdx}`).trim();
      return autoPrefix + rawLabel;
    });
  });

  state.result   = result;
  state.filtered = result;

  // Stats
  const matched = result.filter(r => r._status === 'match').length;
  const partial = result.filter(r => r._status === 'part').length;
  const none    = result.filter(r => r._status === 'none').length;
  const pct     = total ? Math.round((matched / total) * 100) : 0;

  document.getElementById('svTotal').textContent = total;
  document.getElementById('svMatch').textContent = matched;
  document.getElementById('svNone').textContent  = none + partial;
  document.getElementById('svPct').textContent   = pct + '%';
  document.getElementById('statsGrid').style.display = 'grid';

  renderTable(result);
  showResultsUI();
  setStep(4);

  const lkSummary = state.lookups.map(l => `${l.name}: ${l.data.length} linhas`).join(' | ');
  updateProcessStatus(`Consolidado — ${total} linhas base`, lkSummary);
  toast(`Concluído — ${matched} matches completos de ${total} linhas`, 'ok');

  setTimeout(() => progress(0, false), 800);
  document.getElementById('btnExport').style.display = 'inline-flex';
}

// ─────────────────────────────────────────────────────────────────────────────
// RANGE EXPANSION (generic, multi-separator)
// ─────────────────────────────────────────────────────────────────────────────
function normKey(s) {
  s = s.trim();
  if (/^\d+\.0+$/.test(s)) s = String(Math.round(parseFloat(s)));
  return s;
}

/** Returns active separators from the checkboxes + custom field (sorted longest first) */
function getBaseSeps() {
  const seps = [];
  document.querySelectorAll('#baseSepChecks input[type=checkbox]:checked').forEach(cb => {
    seps.push(cb.value);
  });
  const custom = (document.getElementById('baseSepCustom')?.value || '')
    .split(',').map(s => s.trim()).filter(Boolean);
  custom.forEach(s => { if (!seps.includes(s)) seps.push(s); });
  // Sort longest first so "até" is tested before "a"
  seps.sort((a, b) => b.length - a.length);
  return seps.length ? seps : ['a'];
}

/**
 * expandRange: split cell by hard delimiters (only , and ; that are NOT part of a sep word),
 * then expand each chunk using the full seps array.
 * seps: string[] sorted longest-first
 */
function expandRange(cellVal, seps) {
  if (cellVal === null || cellVal === undefined || cellVal === '') return [];
  const raw = String(cellVal).trim();
  if (!raw) return [];

  // Hard split only on , or ; that are standalone (not used as word-seps)
  // We split on comma/semicolon not covered by the seps list
  const hardSepRe = /[,;]+/;
  const results = [];
  const parts = raw.split(hardSepRe).map(s => s.trim()).filter(Boolean);
  for (const p of parts) results.push(...expandChunk(p, seps));
  return [...new Set(results)];
}

/**
 * expandChunk: tokenize expr with all seps, then walk tokens to expand ranges.
 * A token is: NUM | SEP (any word-sep) | HYPHEN (-) | COMMA/SEMICOLON
 */
function expandChunk(expr, seps) {
  if (/^\d+\.0+$/.test(expr.trim())) return [String(Math.round(parseFloat(expr)))];

  // Build regex alternating all word seps + hyphen + digits
  // Seps already sorted longest first
  const sepAlts = seps.filter(s => s !== '-').map(s => escRe(s));
  // hyphen handled separately as HYPHEN token
  const reStr = `(\\d+)` +
    (sepAlts.length ? `|(?:${sepAlts.join('|')})` : '') +
    `|(-)`;
  const re = new RegExp(reStr, 'gi');

  const tokens = [];
  let m;
  while ((m = re.exec(expr)) !== null) {
    if (m[1] !== undefined)                          tokens.push({ t: 'N', v: m[1] });
    else if (m[m.length - 1] !== undefined && m[m.length-1] === '-') tokens.push({ t: '-' });
    else                                              tokens.push({ t: 'S' });
  }
  if (!tokens.length) return [];

  function resolveEnd(start, end) {
    return end.length < start.length ? start.slice(0, start.length - end.length) + end : end;
  }

  const out = []; let i = 0; let lastFull = null;

  while (i < tokens.length) {
    const tok = tokens[i];
    if (tok.t !== 'N') { i++; continue; }
    let fullA = tok.v;
    if (lastFull && fullA.length < lastFull.length) fullA = resolveEnd(lastFull, fullA);

    const n1 = tokens[i+1], n2 = tokens[i+2];
    if (n1 && (n1.t === 'S' || n1.t === '-') && n2 && n2.t === 'N') {
      const fullB = resolveEnd(fullA, n2.v);
      const s = parseInt(fullA), e = parseInt(fullB);
      if (!isNaN(s) && !isNaN(e) && e >= s && (e - s) < 100000) {
        for (let n = s; n <= e; n++) out.push(String(n));
        lastFull = fullB.length >= fullA.length ? fullB : resolveEnd(fullA, fullB);
      } else {
        out.push(fullA); lastFull = fullA;
      }
      i += 3;
    } else {
      out.push(fullA); lastFull = fullA; i++;
    }
  }
  return out;
}

/**
 * abbreviateNums: unique values sorted numerically, consecutive runs contracted
 * into "start-end" notation (e.g. 100,101,102,104,107 → "100-102, 104, 107").
 * Non-numeric values are appended as-is after the numeric ranges.
 */
function abbreviateNums(vals) {
  const nums = [], nonNums = [];
  vals.forEach(v => {
    const n = parseInt(v);
    if (!isNaN(n) && String(n) === String(v).trim()) nums.push(n);
    else nonNums.push(v);
  });

  if (!nums.length) return [...new Set(nonNums)].join(', ');

  const sorted = [...new Set(nums)].sort((a, b) => a - b);

  const ranges = [];
  let start = sorted[0], end = sorted[0];
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === end + 1) {
      end = sorted[i];
    } else {
      ranges.push(start === end ? String(start) : `${start} a ${end}`);
      start = end = sorted[i];
    }
  }
  ranges.push(start === end ? String(start) : `${start} a ${end}`);

  return [...ranges, ...new Set(nonNums)].join(', ');
}

// ─────────────────────────────────────────────────────────────────────────────
// TABLE RENDER
// ─────────────────────────────────────────────────────────────────────────────
function renderTable(rows) {
  const outer = document.getElementById('tableOuter');
  const empty = document.getElementById('emptyState');
  const thead = document.getElementById('tHead');
  const tbody = document.getElementById('tBody');

  if (!rows.length) {
    outer.style.display = 'none';
    empty.style.display = 'block';
    return;
  }
  outer.style.display = 'block';
  empty.style.display = 'none';

  const allKeys    = Object.keys(rows[0]).filter(k => !k.startsWith('_'));
  const baseKeys   = allKeys.filter(k => k.startsWith('Base_'));
  // Lookup keys: use stored ordered list, filtered to what's actually in this result
  // Falls back to anything not Base_ if lookupColKeys is empty (e.g. after filter)
  const lookupKeys = state.lookupColKeys?.length
    ? state.lookupColKeys.filter(k => allKeys.includes(k))
    : allKeys.filter(k => !k.startsWith('Base_'));

  thead.innerHTML = '<tr>'
    + '<th>STATUS</th>'
    + '<th class="th-key">CHAVE BASE</th>'
    + baseKeys.map(k  => `<th>${escHtml(k.replace('Base_',''))}</th>`).join('')
    + lookupKeys.map(k => `<th class="th-lookup">${escHtml(k)}</th>`).join('')
    + '</tr>';

  tbody.innerHTML = rows.slice(0,3000).map(row => {
    const badge = row._status === 'match'
      ? '<span class="badge b-match">MATCH</span>'
      : row._status === 'part'
      ? '<span class="badge b-partial">PARCIAL</span>'
      : '<span class="badge b-none">SEM MATCH</span>';

    const baseCells   = baseKeys.map(k  => `<td title="${escHtml(String(row[k]??''))}">${escHtml(String(row[k]??''))}</td>`).join('');
    const lookupCells = lookupKeys.map(k => `<td style="color:var(--cyan);opacity:.85" title="${escHtml(String(row[k]??''))}">${escHtml(String(row[k]??''))}</td>`).join('');

    return `<tr>${badge ? `<td>${badge}</td>` : ''}<td style="font-family:var(--mono);font-size:.72rem;color:var(--muted2)">${escHtml(String(row._raw_key??''))}</td>${baseCells}${lookupCells}</tr>`;
  }).join('');

  if (rows.length > 3000) {
    tbody.innerHTML += `<tr><td colspan="99" style="text-align:center;color:var(--muted);padding:16px;font-family:var(--mono);font-size:.72rem">… +${rows.length-3000} linhas. Exporte para ver tudo.</td></tr>`;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FILTERS & SEARCH
// ─────────────────────────────────────────────────────────────────────────────
function setFilter(f) {
  state.activeFilter = f;
  ['all','match','none','part'].forEach(k => {
    document.getElementById(`fp_${k}`)?.classList.toggle('fp-active', k === f);
  });
  applyFilters();
}

function doSearch(val) {
  state.searchTerm = val.toLowerCase();
  applyFilters();
}

function applyFilters() {
  let rows = state.result;
  if (state.activeFilter === 'match') rows = rows.filter(r => r._status === 'match');
  else if (state.activeFilter === 'none')  rows = rows.filter(r => r._status === 'none');
  else if (state.activeFilter === 'part')  rows = rows.filter(r => r._status === 'part');
  if (state.searchTerm) {
    rows = rows.filter(r => Object.values(r).some(v => String(v).toLowerCase().includes(state.searchTerm)));
  }
  state.filtered = rows;
  renderTable(rows);
}

// ─────────────────────────────────────────────────────────────────────────────
// EXPORT
// ─────────────────────────────────────────────────────────────────────────────
function exportResult() {
  const rows = state.filtered;
  if (!rows.length) { toast('Nenhum dado para exportar', 'err'); return; }

  const allKeys    = Object.keys(rows[0]).filter(k => !k.startsWith('_'));
  const baseKeys   = allKeys.filter(k => k.startsWith('Base_'));
  const lookupKeys = state.lookupColKeys?.length
    ? state.lookupColKeys.filter(k => allKeys.includes(k))
    : allKeys.filter(k => !k.startsWith('Base_'));

  // Build export in correct order: STATUS, CHAVE_BASE, base cols, lookup cols (by lookup order)
  const exportData = rows.map(row => {
    const out = {
      STATUS:     row._status === 'match' ? 'MATCH' : row._status === 'part' ? 'PARCIAL' : 'SEM MATCH',
      CHAVE_BASE: row._raw_key ?? '',
    };
    baseKeys.forEach(k   => { out[k.replace('Base_', '')] = row[k] ?? ''; });
    lookupKeys.forEach(k => { out[k] = row[k] ?? ''; });
    return out;
  });

  const ws = XLSX.utils.json_to_sheet(exportData);
  ws['!cols'] = Array(Object.keys(exportData[0]).length).fill({ wch: 24 });

  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, 'Consolidado');

  // Stats sheet
  const statsData = [
    { Métrica: 'Total base',     Valor: state.result.length },
    { Métrica: 'Match completo', Valor: state.result.filter(r=>r._status==='match').length },
    { Métrica: 'Parcial',        Valor: state.result.filter(r=>r._status==='part').length },
    { Métrica: 'Sem match',      Valor: state.result.filter(r=>r._status==='none').length },
    { Métrica: 'Fontes usadas',  Valor: state.lookups.length },
    { Métrica: 'Exportado em',   Valor: new Date().toLocaleString('pt-BR') },
  ];
  XLSX.utils.book_append_sheet(wb2, XLSX.utils.json_to_sheet(statsData), 'Resumo');

  XLSX.writeFile(wb2, `datamesh_${new Date().toISOString().slice(0,10)}.xlsx`);
  toast('Exportado com sucesso!', 'ok');
}

// ─────────────────────────────────────────────────────────────────────────────
// WORKBOOK HELPERS
// ─────────────────────────────────────────────────────────────────────────────
function readWorkbook(file, cb) {
  const reader = new FileReader();
  reader.onerror = () => toast('Erro ao ler arquivo', 'err');
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), {
        type: 'array', cellDates: true, raw: false,
      });
      cb(wb);
    } catch (err) {
      toast('Erro ao processar planilha: ' + err.message, 'err');
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseSheet(wb, sheetIdx, hRow) {
  const ws  = wb.Sheets[wb.SheetNames[sheetIdx]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false, raw: false });
  const headers  = (raw[hRow] || []).map(h => String(h).trim());
  const data     = raw.slice(hRow + 1).filter(r => r.some(c => c !== '' && c !== null));
  return { headers, data };
}

function populateSheetSel(selId, wb, onChange) {
  const sel = document.getElementById(selId);
  if (!sel) return;
  const names = wb.SheetNames;
  sel.innerHTML = names.map((n,i) => `<option value="${i}">${escHtml(n)}</option>`).join('');
  if (names.length > 1) {
    sel.style.display = 'block';
    // hide the "—" placeholder if present
    const noneId = selId.replace('Sheet', 'SheetNone').replace('-sheet-', '-sheet-none-');
    const noneEl = document.getElementById(noneId);
    if (noneEl) noneEl.style.display = 'none';
  } else {
    sel.style.display = 'none';
    const noneId = selId.replace('Sheet', 'SheetNone').replace('-sheet-', '-sheet-none-');
    const noneEl = document.getElementById(noneId);
    if (noneEl) { noneEl.style.display = 'block'; noneEl.textContent = names[0]; }
  }
  sel.onchange = onChange;
  onChange();
}

function populateSelect(selId, headers, defaultVal) {
  const sel = document.getElementById(selId);
  if (!sel) return;
  sel.innerHTML = headers.map((h, i) => `<option value="${i}">${escHtml(h || 'col'+i)}</option>`).join('');
}

// ─────────────────────────────────────────────────────────────────────────────
// UI HELPERS
// ─────────────────────────────────────────────────────────────────────────────
function checkReady() {
  const hasBase = state.baseData.length > 0;
  const hasLookup = state.lookups.length > 0 && state.lookups.every(l => l.data.length > 0);
  document.getElementById('btnRun').disabled = !(hasBase && hasLookup);
  updateProcessStatus(
    hasBase && hasLookup ? 'Pronto para consolidar' : 'Aguardando configuração',
    hasBase ? `Base: ${state.baseData.length} linhas · ${state.lookups.length} lookup(s)` : 'Carregue a planilha base para começar'
  );
}

function updateProcessStatus(title, sub) {
  document.getElementById('processTitle').textContent = title;
  document.getElementById('processSub').textContent   = sub;
}

function showResultsUI() {
  document.getElementById('resultsHdr').style.display = 'flex';
  document.getElementById('tableOuter').style.display = 'block';
}

function setStep(n) {
  const labels = ['sp1','sp2','sp3','sp4'];
  labels.forEach((id, i) => {
    const el = document.getElementById(id);
    el.classList.remove('active','done');
    if      (i + 1 < n) el.classList.add('done');
    else if (i + 1 === n) el.classList.add('active');
  });
}

function progress(pct, visible = true) {
  const track = document.getElementById('progressTrack');
  const fill  = document.getElementById('progressFill');
  track.classList.toggle('visible', visible && pct > 0);
  fill.style.width = pct + '%';
}

let _toastTimer;
function toast(msg, type='ok') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = `toast show t-${type}`;
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 4000);
}

function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }

function hexAlpha(hex, a) {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${a})`;
}

// ─────────────────────────────────────────────────────────────────────────────
// PAGE NAVIGATION
// ─────────────────────────────────────────────────────────────────────────────
function switchPage(n) {
  [1,2].forEach(i => {
    document.getElementById(`page${i}`).classList.toggle('active', i === n);
    document.getElementById(`nav${i}`).classList.toggle('active', i === n);
  });
  const subtitles = { 1: '— Consolidador', 2: '— Unificador' };
  document.getElementById('pageTitle').innerHTML =
    `Mescl<em>.</em>ar <span style="font-weight:400;opacity:.5;font-size:.9em">${subtitles[n]}</span>`;
}

// ─────────────────────────────────────────────────────────────────────────────
// PAGE 2 — EXTRATOR DE RELATÓRIOS DE FRETE (FORMATO VERTICAL)
// Lê múltiplas planilhas com múltiplas abas, extrai campos de formulários
// verticais (col A = rótulo, col B/C = valor) e consolida em uma tabela.
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
// STATE — página 2
// ─────────────────────────────────────────────────────────────────────────────
const unifState = {
  files: [],    // [{ id, name, wb }]
  result: [],
  filtered: [],
};
let unifFileCounter = 0;

// ─────────────────────────────────────────────────────────────────────────────
// MAPA DE CAMPOS — rótulos das planilhas → nome canônico da coluna de saída
// Cada entrada: { key, re, pct? }
//   pct: true  → captura o valor percentual da linha
//   pct: false → captura o valor monetário da linha
//   pct: undefined → captura qualquer valor numérico encontrado
// ─────────────────────────────────────────────────────────────────────────────
const FIELD_MAP = [
  { key: 'Recibo',              re: /recibo/i },
  { key: 'Rota',                re: /\brota\b/i },
  { key: 'Data Emissão',        re: /recife[,\s]/i },
  { key: 'Percurso',            re: /percurso|trajeto/i },
  { key: 'Valor Frete (R$)',    re: /valor\s+frete/i },
  { key: 'Dif. Diesel (R$)',    re: /diferen[cç]a\s+de\s+diesel|dif.*diesel/i },
  { key: 'Adiantamento (R$)',   re: /adiantamento/i },
  { key: 'Ad-valorem (%)',      re: /ad.?valorem/i,   pct: true  },
  { key: 'Ad-valorem (R$)',     re: /ad.?valorem/i,   pct: false },
  { key: 'Carga/Descarga (R$)', re: /carga\s*[\/e]\s*descarga|carga\s*desc/i },
  { key: 'Pedágio (R$)',        re: /ped[aá]gio/i },
  { key: 'Taxa Adm. (%)',       re: /taxa\s+adm/i,    pct: true  },
  { key: 'Taxa Adm. (R$)',      re: /taxa\s+adm/i,    pct: false },
  { key: 'ICMS (R$)',           re: /\bicms\b/i },
  { key: 'Saldo a Pagar (R$)',  re: /saldo\s+a\s+pagar/i },
  { key: 'Total Frete Pago (R$)', re: /valor\s+total\s+frete\s+pago|total.*frete.*pago/i },
  { key: 'Conhecimento/Fatura', re: /conhecimento\s*[\/e]\s*fatura|conhecimento|fatura/i },
  { key: 'Frete s/ ICMS (R$)',  re: /frete\s+s\/?\.?\s*icms/i },
  { key: 'Valor da Carga (R$)', re: /valor\s+da\s+carga/i },
  { key: 'Percentual (%)',      re: /percentual\s+que\s+representa|percentual/i },
  { key: 'Notas Fiscais',       re: /notas?\s+fiscais?/i },
  { key: 'Entregas',            re: /\bentrega\b/i },
  { key: 'KM',                  re: /\bkm\b/i },
  { key: 'Tipo Veículo',        re: /tipo\s+ve[íi]culo/i },
  { key: 'CD',                  re: /^cd[-\s]?\d+$/i },
];

// Ordem das colunas no resultado final
const CANONICAL_COLS = [
  'Arquivo', 'Aba',
  'Recibo', 'Rota', 'Data Emissão', 'Percurso',
  'Valor Frete (R$)', 'Dif. Diesel (R$)', 'Adiantamento (R$)',
  'Ad-valorem (%)', 'Ad-valorem (R$)',
  'Carga/Descarga (R$)', 'Pedágio (R$)',
  'Taxa Adm. (%)', 'Taxa Adm. (R$)',
  'ICMS (R$)', 'Saldo a Pagar (R$)', 'Total Frete Pago (R$)',
  'Conhecimento/Fatura', 'Frete s/ ICMS (R$)', 'Valor da Carga (R$)',
  'Percentual (%)', 'Notas Fiscais', 'Entregas', 'KM', 'Tipo Veículo', 'CD',
];

// ─────────────────────────────────────────────────────────────────────────────
// PARSER DE ABA VERTICAL
// Cada aba é um formulário: col A = rótulo, col B/C = valor (às vezes % na col C)
// Retorna um objeto com os campos canônicos preenchidos.
// ─────────────────────────────────────────────────────────────────────────────
function parseFreightSheet(raw, fileName, sheetName) {
  const rec = {};
  CANONICAL_COLS.forEach(c => { rec[c] = ''; });
  rec['Arquivo'] = fileName;
  rec['Aba']     = sheetName;

  // ── Cabeçalho: Recibo, Rota e Data nas primeiras linhas ──────────
  for (let i = 0; i < Math.min(raw.length, 8); i++) {
    const line = (raw[i] || []).map(c => String(c ?? '').trim()).join(' ').trim();

    if (!rec['Recibo']) {
      const m = line.match(/recibo\s+([\d]+\s*\/\s*[\d]+)/i);
      if (m) rec['Recibo'] = m[1].replace(/\s/g, '');
    }
    if (!rec['Rota']) {
      const m = line.match(/rota\s+([\d]+\s*\/\s*[\d]+)/i);
      if (m) rec['Rota'] = m[1].replace(/\s/g, '');
    }
    if (!rec['Data Emissão']) {
      const m = line.match(/recife[,\s]+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})/i);
      if (m) rec['Data Emissão'] = m[1];
    }
  }

  // ── Varredura linha a linha ──────────────────────────────────────
  for (let i = 0; i < raw.length; i++) {
    const row      = raw[i] || [];
    const labelRaw = String(row[0] ?? row[1] ?? '').trim();
    if (!labelRaw) continue;

    // Percurso: "Percurso: Recife PE / Garanhuns PE"
    if (/percurso/i.test(labelRaw) && !rec['Percurso']) {
      const full = row.map(c => String(c ?? '').trim()).filter(Boolean).join(' ');
      const m = full.match(/percurso\s*:\s*(.+)/i);
      if (m) rec['Percurso'] = m[1].trim();
      continue;
    }

    // Extrai valor monetário e percentual das colunas B em diante
    let valBRL = '', valPct = '';
    for (let c = 1; c < row.length; c++) {
      const cell = String(row[c] ?? '').trim();
      if (!cell) continue;
      if (/^\d[\d.,]*%$/.test(cell) && !valPct) { valPct = cell; continue; }
      if (/^[\d.,]+$/.test(cell.replace(/\s/g, '')) && !valBRL) { valBRL = cell; continue; }
    }

    // Casa o rótulo com o mapa de campos
    for (const field of FIELD_MAP) {
      if (!field.re.test(labelRaw)) continue;
      if (field.pct === true)  { if (valPct && !rec[field.key]) rec[field.key] = valPct; }
      else if (field.pct === false) { if (valBRL && !rec[field.key]) rec[field.key] = valBRL; }
      else {
        const val = valBRL || valPct;
        if (val && !rec[field.key]) rec[field.key] = val;
      }
    }

    // Campos de texto livre — Conhecimento/Fatura
    if (/conhecimento\s*[\/e]\s*fatura|conhecimento|fatura/i.test(labelRaw) && !rec['Conhecimento/Fatura']) {
      const v = row.slice(1).map(c => String(c ?? '').trim()).filter(Boolean).join(' / ');
      if (v) rec['Conhecimento/Fatura'] = v;
    }
    // Notas Fiscais
    if (/notas?\s+fiscais?/i.test(labelRaw) && !rec['Notas Fiscais']) {
      const v = row.slice(1).map(c => String(c ?? '').trim()).filter(Boolean).join(', ');
      if (v) rec['Notas Fiscais'] = v;
    }
    // Entrega
    if (/\bentrega\b/i.test(labelRaw) && !rec['Entregas']) {
      const v = row.slice(1).map(c => String(c ?? '').trim()).filter(Boolean).join('');
      if (v) rec['Entregas'] = v;
    }
    // KM
    if (/\bkm\b/i.test(labelRaw) && !rec['KM']) {
      const v = row.slice(1).map(c => String(c ?? '').trim()).filter(Boolean).join('');
      if (v) rec['KM'] = v;
    }
    // Tipo Veículo
    if (/tipo\s+ve[íi]culo/i.test(labelRaw) && !rec['Tipo Veículo']) {
      const v = row.slice(1).map(c => String(c ?? '').trim()).filter(Boolean).join('');
      if (v) rec['Tipo Veículo'] = v;
    }
    // CD (ex: "CD-03" aparece como rótulo sem valor)
    if (/^cd[-\s]?\d+$/i.test(labelRaw) && !rec['CD']) {
      rec['CD'] = labelRaw;
    }
  }

  // Percentual — fallback: busca "X.XX%" em qualquer linha
  if (!rec['Percentual (%)']) {
    for (const row of raw) {
      const line = (row || []).map(c => String(c ?? '').trim()).join(' ');
      const m = line.match(/percentual\s+que\s+representa[^\d]*([\d.,]+%)/i);
      if (m) { rec['Percentual (%)'] = m[1]; break; }
    }
  }

  return rec;
}

// ─────────────────────────────────────────────────────────────────────────────
// DRAG & DROP / UPLOAD
// ─────────────────────────────────────────────────────────────────────────────
function handleUnifDrop(e) {
  e.preventDefault();
  document.getElementById('unifDropZone').classList.remove('drag-over');
  loadUnifFiles(e.dataTransfer.files);
}

function loadUnifFiles(fileList) {
  Array.from(fileList).forEach(file => {
    const id    = ++unifFileCounter;
    const entry = { id, name: file.name, wb: null };
    unifState.files.push(entry);
    renderUnifFileCard(entry);

    readWorkbook(file, wb => {
      entry.wb = wb;
      const sheetCount = wb.SheetNames.length;
      const meta = document.getElementById(`uf-meta-${id}`);
      if (meta) meta.textContent = `${sheetCount} aba(s) — pronto`;
      document.getElementById(`uf-card-${id}`)?.classList.add('loaded');
      toast(`${file.name} — ${sheetCount} aba(s) carregada(s)`, 'ok');
      updateUnifReadyState();
    });
  });
  setUnifStep(1);
}

function renderUnifFileCard(entry) {
  const grid = document.getElementById('unifFilesGrid');
  const card = document.createElement('div');
  card.className = 'unif-file-card';
  card.id = `uf-card-${entry.id}`;
  const shortName = entry.name.length > 28 ? entry.name.slice(0, 26) + '…' : entry.name;
  card.innerHTML = `
    <div class="unif-file-top">
      <div class="unif-file-icon">📄</div>
      <div style="flex:1">
        <div class="unif-file-name">${escHtml(shortName)}</div>
        <div class="unif-file-meta" id="uf-meta-${entry.id}">Lendo todas as abas…</div>
      </div>
      <button class="unif-file-rm" onclick="removeUnifFile(${entry.id})">✕</button>
    </div>`;
  grid.appendChild(card);
}

function removeUnifFile(id) {
  unifState.files = unifState.files.filter(f => f.id !== id);
  document.getElementById(`uf-card-${id}`)?.remove();
  updateUnifReadyState();
  toast('Arquivo removido', 'ok');
}

function updateUnifReadyState() {
  const loadedFiles = unifState.files.filter(f => f.wb !== null);
  const ready = loadedFiles.length > 0;
  const btnUnif = document.getElementById('btnUnif');
  if (btnUnif) btnUnif.disabled = !ready;

  if (ready) {
    const totalSheets = loadedFiles.reduce((s, f) => s + f.wb.SheetNames.length, 0);
    document.getElementById('unifProcessTitle').textContent = 'Pronto para consolidar';
    document.getElementById('unifProcessSub').textContent =
      `${loadedFiles.length} planilha(s) · ${totalSheets} aba(s) detectada(s) — clique Consolidar.`;
    document.getElementById('unifDetectedSection').style.display = 'block';

    // Atualiza tags de campos detectados
    renderDetectedFieldTags(loadedFiles);
    setUnifStep(2);
  } else {
    document.getElementById('unifDetectedSection').style.display = 'none';
    setUnifStep(1);
  }
}

// Mostra as tags dos campos que serão extraídos
function renderDetectedFieldTags(loadedFiles) {
  const tagsEl = document.getElementById('unifDetectedTags');
  const countEl = document.getElementById('unifDetectedCount');
  if (!tagsEl) return;

  const fieldCount = CANONICAL_COLS.length - 2; // exclui Arquivo e Aba
  if (countEl) countEl.textContent = `${fieldCount} campos · ${loadedFiles.length} planilha(s)`;

  tagsEl.innerHTML = CANONICAL_COLS
    .filter(c => c !== 'Arquivo' && c !== 'Aba')
    .map(c => `<span style="
      display:inline-flex;align-items:center;gap:5px;
      padding:4px 10px;border-radius:20px;font-size:0.7rem;font-family:var(--mono);
      background:var(--cyan-dim);border:1px solid rgba(0,150,199,0.2);color:var(--cyan)
    ">✦ ${escHtml(c)}</span>`)
    .join('');
}

// ─────────────────────────────────────────────────────────────────────────────
// MOTOR DE CONSOLIDAÇÃO — página 2
// ─────────────────────────────────────────────────────────────────────────────
function runUnify() {
  const files = unifState.files.filter(f => f.wb !== null);
  if (!files.length) { toast('Nenhuma planilha carregada', 'err'); return; }

  setUnifStep(3);
  unifProgress(0, true);
  document.getElementById('unifProcessTitle').textContent = 'Consolidando…';
  document.getElementById('unifProcessSub').textContent   = 'Varrendo abas…';

  setTimeout(() => {
    try {
      _runUnify(files);
    } catch (e) {
      toast('Erro: ' + e.message, 'err');
      console.error(e);
      unifProgress(0, false);
    }
  }, 30);
}

function _runUnify(files) {
  const result = [];
  let totalSheets = 0;
  let processed   = 0;

  files.forEach(f => { totalSheets += f.wb.SheetNames.length; });

  files.forEach(file => {
    file.wb.SheetNames.forEach(sheetName => {
      const ws  = file.wb.Sheets[sheetName];
      const raw = XLSX.utils.sheet_to_json(ws, {
        header: 1, defval: '', blankrows: false, raw: false,
      });

      if (!raw || raw.length < 3) { processed++; return; }

      const rec = parseFreightSheet(raw, file.name, sheetName);

      // Inclui apenas abas com dados relevantes
      const hasData = rec['Recibo'] || rec['Valor Frete (R$)'] || rec['Percurso'] || rec['Conhecimento/Fatura'];
      if (hasData) result.push(rec);

      processed++;
      unifProgress(Math.round((processed / totalSheets) * 90));
    });
  });

  unifProgress(100);
  unifState.result   = result;
  unifState.filtered = result;

  // Conta campos vazios para o stat
  const emptyCount = result.reduce((s, row) =>
    s + CANONICAL_COLS.filter(c => c !== 'Arquivo' && c !== 'Aba' && !row[c]).length, 0);

  document.getElementById('usvTotal').textContent  = result.length;
  document.getElementById('usvFiles').textContent  = `${files.length} · ${totalSheets}`;
  document.getElementById('usvCols').textContent   = CANONICAL_COLS.length - 2;
  document.getElementById('usvEmpty').textContent  = emptyCount;
  document.getElementById('unifStatsGrid').style.display = 'grid';

  renderUnifTable(result);
  document.getElementById('unifResultsHdr').style.display = 'flex';
  document.getElementById('unifTableOuter').style.display = 'block';
  document.getElementById('btnUnifExport').style.display  = 'inline-flex';
  setUnifStep(4);

  document.getElementById('unifProcessTitle').textContent = `Consolidado — ${result.length} registros`;
  document.getElementById('unifProcessSub').textContent   =
    `${CANONICAL_COLS.length - 2} campos · ${totalSheets} aba(s) de ${files.length} planilha(s)`;
  toast(`Concluído — ${result.length} registros extraídos`, 'ok');
  setTimeout(() => unifProgress(0, false), 800);
}

// ─────────────────────────────────────────────────────────────────────────────
// TABELA — página 2
// ─────────────────────────────────────────────────────────────────────────────
function renderUnifTable(rows) {
  const outer = document.getElementById('unifTableOuter');
  const thead = document.getElementById('unifTHead');
  const tbody = document.getElementById('unifTBody');
  if (!rows.length) { outer.style.display = 'none'; return; }
  outer.style.display = 'block';

  // Mostra apenas colunas que têm ao menos um valor preenchido
  const activeCols = CANONICAL_COLS.filter(col =>
    rows.some(r => r[col] && String(r[col]).trim() !== '')
  );

  thead.innerHTML = '<tr>'
    + activeCols.map(k => {
        const isSource = k === 'Arquivo' || k === 'Aba';
        return `<th style="color:${isSource ? 'var(--muted2)' : 'var(--cyan)'};white-space:nowrap">${escHtml(k)}</th>`;
      }).join('')
    + '</tr>';

  tbody.innerHTML = rows.slice(0, 3000).map(row =>
    '<tr>'
    + activeCols.map(k => {
        const v = String(row[k] ?? '').trim();
        const isSource = k === 'Arquivo' || k === 'Aba';
        const style = isSource
          ? 'font-size:.68rem;color:var(--muted);white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis'
          : '';
        return `<td style="${style}" title="${escHtml(v)}">${escHtml(v)}</td>`;
      }).join('')
    + '</tr>'
  ).join('');

  if (rows.length > 3000) {
    tbody.innerHTML +=
      `<tr><td colspan="99" style="text-align:center;color:var(--muted);padding:16px;font-family:var(--mono);font-size:.72rem">…+${rows.length - 3000} linhas. Exporte para ver tudo.</td></tr>`;
  }
}

function unifSearch(val) {
  const term = val.toLowerCase();
  unifState.filtered = term
    ? unifState.result.filter(r => Object.values(r).some(v => String(v).toLowerCase().includes(term)))
    : unifState.result;
  renderUnifTable(unifState.filtered);
}

// ─────────────────────────────────────────────────────────────────────────────
// EXPORTAR — página 2
// ─────────────────────────────────────────────────────────────────────────────
function exportUnif() {
  const rows = unifState.filtered;
  if (!rows.length) { toast('Nada para exportar', 'err'); return; }

  const activeCols = CANONICAL_COLS.filter(col =>
    rows.some(r => r[col] && String(r[col]).trim() !== '')
  );

  const exportData = rows.map(r => {
    const obj = {};
    activeCols.forEach(k => { obj[k] = r[k] ?? ''; });
    return obj;
  });

  const ws = XLSX.utils.json_to_sheet(exportData);
  ws['!cols'] = activeCols.map(k => ({ wch: Math.max(k.length + 2, 18) }));

  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, 'Fretes Consolidados');

  const loadedFiles = unifState.files.filter(f => f.wb !== null);
  XLSX.utils.book_append_sheet(wb2, XLSX.utils.json_to_sheet([
    { Métrica: 'Total de registros', Valor: unifState.result.length },
    { Métrica: 'Planilhas',          Valor: loadedFiles.length },
    { Métrica: 'Abas varridas',      Valor: loadedFiles.reduce((s, f) => s + f.wb.SheetNames.length, 0) },
    { Métrica: 'Campos extraídos',   Valor: activeCols.length - 2 },
    { Métrica: 'Exportado em',       Valor: new Date().toLocaleString('pt-BR') },
  ]), 'Resumo');

  XLSX.writeFile(wb2, `fretes_consolidados_${new Date().toISOString().slice(0, 10)}.xlsx`);
  toast('Exportado com sucesso!', 'ok');
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP & PROGRESS — página 2
// ─────────────────────────────────────────────────────────────────────────────
function setUnifStep(n) {
  ['usp1', 'usp2', 'usp3', 'usp4'].forEach((id, i) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.remove('active', 'done');
    if      (i + 1 < n)  el.classList.add('done');
    else if (i + 1 === n) el.classList.add('active');
  });
}

function unifProgress(pct, visible = true) {
  const track = document.getElementById('unifProgressTrack');
  const fill  = document.getElementById('unifProgressFill');
  if (!track || !fill) return;
  track.classList.toggle('visible', visible && pct > 0);
  fill.style.width = pct + '%';
}

// Stubs para compatibilidade com o HTML (slider de threshold não é mais usado)
function resetAndRedetect() { updateUnifReadyState(); }
function redetectFields()   { updateUnifReadyState(); }
function onSimThresholdChange(val) {
  const el = document.getElementById('simThresholdVal');
  if (el) el.textContent = val + '%';
}

