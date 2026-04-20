/* =============================================================================
 * Druk rezerwacji · Gazetka FAMIX (v3)
 * -----------------------------------------------------------------------------
 * Układ shuttle:
 *   [Params] | [Picker — lista dostępnych indeksów] | [Editor — pozycje w druku]
 *
 * - Cena FAMIX = kolumna cena_s ; VAT = kolumna vat
 * - Zaznaczasz w Pickerze → "Dodaj zaznaczone →" → pozycje lecą do Editora
 * - W Editorze każda komórka edytowalna (Cena, VAT, Rabaty, Promocja, Uwagi)
 * - Override'y per towar_id → localStorage
 * - Podgląd 1:1 z wzorem FAMIX + eksport XLSX + druk PDF
 * ===========================================================================*/

(function () {
  "use strict";

  const STATE = {
    rows: [],
    producers: [],
    producerIndex: new Map(),
    current: {
      producer: "",
      module: "",
      fee: "",
      contact: "",
      newsletter: defaultNewsletterName(),
      dateFrom: "",
      dateTo: "",
      dateDoc: todayISO(),
      selectedIds: new Set(), // pozycje dodane do druku
      search: "",
      overrides: {},          // edytowane wartości per towar_id
      calcMode: "cascade",    // "cascade" (mnożąco) | "sum" (sumarycznie)
    },
    pickerChecked: new Set(), // transient: zaznaczone w pickerze do dodania
  };

  const STORAGE_KEY = "famix_druk_rezerwacji_v3";

  // -------------------------------------------------------------------------
  // Pomocnicze
  // -------------------------------------------------------------------------
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  function todayISO() { return new Date().toISOString().slice(0, 10); }
  function defaultNewsletterName() {
    const m = ["Styczeń","Luty","Marzec","Kwiecień","Maj","Czerwiec","Lipiec","Sierpień","Wrzesień","Październik","Listopad","Grudzień"];
    const d = new Date();
    return `Gazetka ${m[d.getMonth()]} ${d.getFullYear()}`;
  }
  function fmtNum(n, d = 2) {
    if (n === null || n === undefined || n === "" || Number.isNaN(+n)) return "";
    return (+n).toLocaleString("pl-PL", { minimumFractionDigits: d, maximumFractionDigits: d });
  }
  function fmtMoney(n) {
    if (n === null || n === undefined || n === "") return "0,00 zł";
    return fmtNum(n, 2) + " zł";
  }
  function fmtPct(n) {
    if (n === null || n === undefined || n === "" || Number.isNaN(+n)) return "";
    return fmtNum((+n) * 100, 2) + "%";
  }
  function fmtDatePL(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    return d.toLocaleDateString("pl-PL", { day: "2-digit", month: "2-digit", year: "numeric" });
  }
  function parseNum(v) {
    if (v === null || v === undefined || v === "") return null;
    const n = parseFloat(String(v).replace(",", ".").replace(/\s/g, ""));
    return Number.isFinite(n) ? n : null;
  }
  function parsePct(v) {
    const n = parseNum(v);
    if (n === null) return null;
    return n > 1 ? n / 100 : n;
  }
  function showToast(msg) {
    const t = $("#toast");
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(showToast._tid);
    showToast._tid = setTimeout(() => (t.hidden = true), 2400);
  }
  function setStatus(text, kind = "muted") {
    const c = $("#statusChip");
    c.textContent = text;
    c.className = "chip " + (kind === "ok" ? "chip-ok" : kind === "err" ? "chip-err" : "chip-muted");
  }
  function escapeHtml(s) {
    if (s == null) return "";
    return String(s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }
  function escapeAttr(s) { return escapeHtml(s); }

  // -------------------------------------------------------------------------
  // Baza
  // -------------------------------------------------------------------------
  async function loadDatabase() {
    setStatus("Ładowanie bazy…", "muted");
    $("#app").setAttribute("aria-busy", "true");
    try {
      const res = await fetch("./data/SC.xlsx", { cache: "no-cache" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
      prepareData(rows);
      setStatus(
        `Baza: ${STATE.rows.length.toLocaleString("pl-PL")} indeksów · ${STATE.producers.length} producentów`,
        "ok"
      );
      $("#app").setAttribute("aria-busy", "false");
      renderProducerList();
      restoreState();
      showToast("Baza załadowana");
    } catch (err) {
      console.error(err);
      setStatus("Błąd ładowania bazy (./data/SC.xlsx)", "err");
      showToast("Nie udało się załadować ./data/SC.xlsx");
    }
  }

  function prepareData(rows) {
    STATE.rows = rows
      .filter((r) => r.nazwa && r.znacznik)
      .map((r) => ({
        towar_id: r.towar_id ?? "",
        nazwa: String(r.nazwa).trim(),
        znacznik: String(r.znacznik).trim(),
        kod: r.kod ?? "",
        jm: r.jm ?? "",
        vat: r.vat ?? "",                    // 5, 8, 23
        cena_famix: r.cena_s ?? null,        // Cena FAMIX z kolumny cena_s
        cena_z: r.cena_z ?? null,
        cena_wyl: r.cena_wyliczona ?? null,
        stale: r.stale ?? null,              // brak w nowym pliku – pole ręczne
        refundacje: r.refundacje ?? null,    // brak w nowym pliku – pole ręczne
        grupa: r.grupa ?? "",
        podgrupa: r.podgrupa ?? "",
      }));

    const pset = new Map();
    STATE.producerIndex = new Map();
    STATE.rows.forEach((r, i) => {
      pset.set(r.znacznik, (pset.get(r.znacznik) || 0) + 1);
      if (!STATE.producerIndex.has(r.znacznik)) STATE.producerIndex.set(r.znacznik, []);
      STATE.producerIndex.get(r.znacznik).push(i);
    });
    STATE.producers = Array.from(pset.keys()).sort((a, b) =>
      a.localeCompare(b, "pl", { sensitivity: "base" })
    );
  }

  function renderProducerList() {
    const dl = $("#producerList");
    dl.innerHTML = STATE.producers
      .map((p) => {
        const count = STATE.producerIndex.get(p).length;
        return `<option value="${escapeAttr(p)}">${count} ind.</option>`;
      })
      .join("");
  }

  // -------------------------------------------------------------------------
  // Merge: baza + override
  // -------------------------------------------------------------------------
  function getMerged(row) {
    const id = String(row.towar_id);
    const ov = STATE.current.overrides[id] || {};
    return {
      ...row,
      cena_famix: ov.cena_famix !== undefined ? ov.cena_famix : row.cena_famix,
      vat: ov.vat !== undefined ? ov.vat : row.vat,
      stale: ov.stale !== undefined ? ov.stale : row.stale,
      rabat_prom: ov.rabat_prom !== undefined ? ov.rabat_prom : null,
      rabat_zo: ov.rabat_zo !== undefined ? ov.rabat_zo : "",
      promocja_netto: ov.promocja_netto !== undefined ? ov.promocja_netto : null,
      refundacja: ov.refundacja !== undefined ? ov.refundacja : row.refundacje,
      prom_rabat: ov.prom_rabat !== undefined ? ov.prom_rabat : "",
      prom_pakietowa: ov.prom_pakietowa !== undefined ? ov.prom_pakietowa : "",
      uwagi: ov.uwagi !== undefined ? ov.uwagi : "",
    };
  }

  function setOverride(id, field, rawValue) {
    if (!STATE.current.overrides[id]) STATE.current.overrides[id] = {};
    const slot = STATE.current.overrides[id];

    if (field === "uwagi" || field === "prom_rabat" || field === "prom_pakietowa") {
      const v = String(rawValue || "").trim();
      if (v === "") delete slot[field]; else slot[field] = v;
    } else if (field === "rabat_zo") {
      const v = String(rawValue || "").toUpperCase();
      if (v !== "Z" && v !== "O") delete slot[field]; else slot[field] = v;
    } else if (field === "vat") {
      if (rawValue === "" || rawValue == null) delete slot[field];
      else slot[field] = String(rawValue);
    } else if (field === "stale" || field === "rabat_prom") {
      const v = parsePct(rawValue);
      if (v === null) delete slot[field]; else slot[field] = v;
    } else {
      const v = parseNum(rawValue);
      if (v === null) delete slot[field]; else slot[field] = v;
    }
    if (Object.keys(slot).length === 0) delete STATE.current.overrides[id];
    persistState();
  }

  /**
   * Centralne wyliczenia dla jednego wiersza (UI, preview, XLSX).
   *  Tryb kaskadowy: Po rab. stał. * (1 - rab.prom) - refund
   *  Tryb sumaryczny: cena_famix * (1 - rab.stały - rab.prom) - refund
   *  Marża % = (C. po rab. prom. - cena_z) / C. po rab. prom.
   */
  function computeRow(m, mode) {
    const cena = m.cena_famix != null ? +m.cena_famix : null;
    const stale = m.stale != null ? +m.stale : null;
    const prom = m.rabat_prom != null ? +m.rabat_prom : null;
    const refund = m.refundacja != null ? +m.refundacja : 0;

    const afterStale = (cena != null && stale != null) ? cena * (1 - stale) : null;

    let afterProm = null;
    if (cena != null && prom != null) {
      if (mode === "sum") {
        const total = (stale || 0) + prom;
        afterProm = cena * (1 - total);
      } else {
        afterProm = afterStale != null ? afterStale * (1 - prom) : cena * (1 - prom);
      }
      afterProm -= refund;
    } else if (refund > 0 && afterStale != null) {
      // brak rabatu prom, ale jest refundacja → C. po rab. prom. = Po rab. stał. - refund
      afterProm = afterStale - refund;
    }

    const vatFrac = parsePct(m.vat);
    const brutto = m.promocja_netto != null && vatFrac != null
      ? (+m.promocja_netto) * (1 + vatFrac)
      : null;

    let marza = null;
    if (afterProm != null && afterProm !== 0 && m.cena_z != null && +m.cena_z > 0) {
      marza = (afterProm - (+m.cena_z)) / afterProm;
    }

    return { afterStale, afterProm, brutto, marza, vatFrac };
  }

  // -------------------------------------------------------------------------
  // PICKER — lewa lista
  // -------------------------------------------------------------------------
  function renderPicker() {
    const body = $("#pickerRows");
    const producer = STATE.current.producer;
    const search = STATE.current.search.trim().toLowerCase();

    if (!producer || !STATE.producerIndex.has(producer)) {
      $("#pickerTitle").textContent = "Dostępne indeksy";
      $("#pickerSub").textContent = "Wybierz producenta, aby zobaczyć indeksy.";
      body.innerHTML = `<div class="empty"><div class="empty-ic">📦</div><p>Wybierz producenta z listy po lewej.</p></div>`;
      updatePickerFooter(0, 0);
      return;
    }

    const ids = STATE.producerIndex.get(producer);
    const rows = ids
      .map((i) => STATE.rows[i])
      .filter((r) => {
        if (!search) return true;
        const hay = (r.nazwa + " " + (r.kod ?? "") + " " + (r.towar_id ?? "")).toLowerCase();
        return hay.includes(search);
      })
      .sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));

    $("#pickerTitle").textContent = `Producent: ${producer}`;
    $("#pickerSub").textContent = `${rows.length} pozycji · ${ids.length} w bazie`;

    if (rows.length === 0) {
      body.innerHTML = `<div class="empty"><div class="empty-ic">🔍</div><p>Brak pozycji pasujących do filtra.</p></div>`;
      updatePickerFooter(0, 0);
      return;
    }

    const frag = document.createDocumentFragment();
    let addable = 0;
    let checked = 0;
    rows.forEach((r) => {
      const id = String(r.towar_id);
      const isAdded = STATE.current.selectedIds.has(id);
      const isChecked = STATE.pickerChecked.has(id);
      if (!isAdded) addable++;
      if (isChecked) checked++;

      const row = document.createElement("div");
      row.className = "pick-row" + (isAdded ? " added" : (isChecked ? " picked" : ""));
      row.dataset.id = id;

      const vatDisplay = r.vat != null && r.vat !== "" ? String(r.vat) + "%" : "—";
      const checkboxHTML = isAdded
        ? `<span class="chip chip-added" title="Już w druku">✓</span>`
        : `<input type="checkbox" ${isChecked ? "checked" : ""} aria-label="Zaznacz do dodania" />`;

      row.innerHTML = `
        <div class="c-chk">${checkboxHTML}</div>
        <div class="c-id">${escapeHtml(id)}</div>
        <div class="c-name" title="${escapeAttr(r.nazwa)}"><span>${escapeHtml(r.nazwa)}</span></div>
        <div class="c-num">${r.cena_famix != null ? fmtNum(r.cena_famix, 2) : "—"}</div>
        <div class="c-num" title="Cena zakupu AC">${r.cena_z != null ? fmtNum(r.cena_z, 2) : "—"}</div>
        <div class="c-num">${escapeHtml(vatDisplay)}</div>
      `;
      frag.appendChild(row);
    });
    body.innerHTML = "";
    body.appendChild(frag);

    // Checkbox "zaznacz wszystkie widoczne"
    const addableVisible = rows.filter((r) => !STATE.current.selectedIds.has(String(r.towar_id)));
    const allChecked = addableVisible.length > 0 &&
      addableVisible.every((r) => STATE.pickerChecked.has(String(r.towar_id)));
    const someChecked = addableVisible.some((r) => STATE.pickerChecked.has(String(r.towar_id)));
    const cball = $("#pickCheckAll");
    cball.checked = allChecked;
    cball.indeterminate = !allChecked && someChecked;
    cball.disabled = addableVisible.length === 0;

    updatePickerFooter(checked, addable);
  }

  function updatePickerFooter(checked, addable) {
    $("#pickerFooterInfo").textContent = addable > 0
      ? `${checked} / ${addable} zaznaczonych`
      : STATE.current.producer ? "Wszystkie pozycje już dodane" : "—";
    $("#btnAddSelected").disabled = checked === 0;
  }

  // -------------------------------------------------------------------------
  // EDITOR — prawa lista z edytowalnymi polami
  // -------------------------------------------------------------------------
  function renderEditor() {
    const body = $("#editorRows");
    const producer = STATE.current.producer;
    const selectedIds = STATE.current.selectedIds;

    if (!producer || selectedIds.size === 0) {
      body.innerHTML = `<div class="empty"><div class="empty-ic">📋</div><p>Brak pozycji. Zaznacz indeksy w lewym panelu i kliknij <strong>"Dodaj zaznaczone"</strong>.</p></div>`;
      $("#editorSub").textContent = "Dodaj pozycje z lewej strony, a potem uzupełnij wartości promocyjne.";
      updateSummary();
      return;
    }

    const rows = STATE.rows
      .filter((r) => r.znacznik === producer && selectedIds.has(String(r.towar_id)))
      .sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));

    $("#editorSub").textContent = `${rows.length} pozycji — edytuj komórki, aby uzupełnić wartości promocyjne.`;

    const frag = document.createDocumentFragment();
    rows.forEach((r) => {
      const id = String(r.towar_id);
      const m = getMerged(r);
      const calc = computeRow(m, STATE.current.calcMode);

      const row = document.createElement("div");
      row.className = "edit-row";
      row.dataset.id = id;
      const marzaClass = calc.marza == null ? "" : (calc.marza >= 0 ? " pos" : " neg");
      const marzaText = calc.marza == null ? "—" : fmtPct(calc.marza);

      row.innerHTML = `
        <div class="c-id">${escapeHtml(id)}</div>
        <div class="c-name" title="${escapeAttr(r.nazwa)}"><span>${escapeHtml(r.nazwa)}</span></div>
        <div class="c-center" title="Jednostka miary">${escapeHtml(String(r.jm || ""))}</div>

        <div class="c-cell">
          <input class="cell-input num" data-id="${escapeAttr(id)}" data-field="cena_famix"
                 type="text" inputmode="decimal"
                 value="${m.cena_famix != null ? fmtNum(m.cena_famix, 2) : ""}"
                 placeholder="0,00" title="Cena FAMIX (zł)" />
        </div>

        <div class="c-cell">
          <input class="cell-input num vat-input" data-id="${escapeAttr(id)}" data-field="vat"
                 type="text" list="vatOpts"
                 value="${m.vat != null && m.vat !== "" ? escapeAttr(m.vat) : ""}"
                 placeholder="—" title="VAT %" />
        </div>

        <div class="c-cell">
          <input class="cell-input num" data-id="${escapeAttr(id)}" data-field="stale"
                 type="text" inputmode="decimal"
                 value="${m.stale != null ? fmtNum((+m.stale) * 100, 2) : ""}"
                 placeholder="0,00" title="Rabat stały %" />
        </div>

        <div class="c-computed" data-computed="po_rab" title="Cena po rabacie stałym">${calc.afterStale != null ? fmtNum(calc.afterStale, 2) : "—"}</div>

        <div class="c-cell">
          <input class="cell-input num" data-id="${escapeAttr(id)}" data-field="rabat_prom"
                 type="text" inputmode="decimal"
                 value="${m.rabat_prom != null ? fmtNum((+m.rabat_prom) * 100, 2) : ""}"
                 placeholder="0,00" title="Rabat promocyjny %" />
        </div>

        <div class="c-zo" data-zo-row="${escapeAttr(id)}">
          <div class="zo-group" role="radiogroup" aria-label="Rabat Z lub O">
            <button type="button" class="zo-btn${m.rabat_zo === "Z" ? " active" : ""}" data-zo="Z" title="Rabat zakupowy">Z</button>
            <button type="button" class="zo-btn${m.rabat_zo === "O" ? " active" : ""}" data-zo="O" title="Rabat odsprzedażowy">O</button>
          </div>
        </div>

        <div class="c-computed" data-computed="po_prom" title="Cena po rabacie promocyjnym (z uwzględnieniem refundacji)">${calc.afterProm != null ? fmtNum(calc.afterProm, 2) : "—"}</div>

        <div class="c-cell">
          <input class="cell-input num" data-id="${escapeAttr(id)}" data-field="refundacja"
                 type="text" inputmode="decimal"
                 value="${m.refundacja != null ? fmtNum(m.refundacja, 2) : ""}"
                 placeholder="0,00" title="Refundacja odsprzedażowa (zł/szt)" />
        </div>

        <div class="c-cell">
          <input class="cell-input num" data-id="${escapeAttr(id)}" data-field="promocja_netto"
                 type="text" inputmode="decimal"
                 value="${m.promocja_netto != null ? fmtNum(m.promocja_netto, 2) : ""}"
                 placeholder="0,00" title="Promocja cenowa netto (zł)" />
        </div>

        <div class="c-computed" data-computed="brutto" title="Promocja cenowa brutto (netto × VAT)">${calc.brutto != null ? fmtNum(calc.brutto, 2) : "—"}</div>

        <div class="c-cell">
          <input class="cell-input text" data-id="${escapeAttr(id)}" data-field="prom_rabat"
                 type="text" value="${escapeAttr(m.prom_rabat || "")}"
                 placeholder="—" title="Promocyjny rabat (np. 15%, 1+1)" />
        </div>

        <div class="c-cell">
          <input class="cell-input text" data-id="${escapeAttr(id)}" data-field="prom_pakietowa"
                 type="text" value="${escapeAttr(m.prom_pakietowa || "")}"
                 placeholder="—" title="Promocja pakietowa" />
        </div>

        <div class="c-cell">
          <input class="cell-input text" data-id="${escapeAttr(id)}" data-field="uwagi"
                 type="text" value="${escapeAttr(m.uwagi || "")}"
                 placeholder="—" title="Uwagi" />
        </div>

        <div class="c-marza${marzaClass}" data-computed="marza" title="Marża % = (C. po rab. prom. − Cena zakupu AC) / C. po rab. prom.">${marzaText}</div>

        <div class="c-chk">
          <button class="btn btn-icon" data-remove="${escapeAttr(id)}" title="Usuń z druku" aria-label="Usuń">
            <svg width="14" height="14" viewBox="0 0 16 16" fill="none"><path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
          </button>
        </div>
      `;
      frag.appendChild(row);
    });
    body.innerHTML = "";
    body.appendChild(frag);
    updateSummary();
  }

  function recomputeEditorRow(rowEl) {
    const id = rowEl.dataset.id;
    const base = STATE.rows.find((r) => String(r.towar_id) === id);
    if (!base) return;
    const m = getMerged(base);
    const calc = computeRow(m, STATE.current.calcMode);

    const setCalc = (key, val) => {
      const cell = rowEl.querySelector(`[data-computed="${key}"]`);
      if (cell) cell.textContent = val != null ? fmtNum(val, 2) : "—";
    };
    setCalc("po_rab", calc.afterStale);
    setCalc("po_prom", calc.afterProm);
    setCalc("brutto", calc.brutto);

    const marzaCell = rowEl.querySelector('[data-computed="marza"]');
    if (marzaCell) {
      marzaCell.textContent = calc.marza == null ? "—" : fmtPct(calc.marza);
      marzaCell.classList.remove("pos", "neg");
      if (calc.marza != null) marzaCell.classList.add(calc.marza >= 0 ? "pos" : "neg");
    }
  }

  function recomputeAllEditorRows() {
    document.querySelectorAll("#editorRows .edit-row").forEach(recomputeEditorRow);
  }

  function updateSummary() {
    $("#sumSelected").textContent = String(STATE.current.selectedIds.size);
    $("#sumModule").textContent = STATE.current.module || "—";
    const fee = parseFloat(STATE.current.fee);
    $("#sumFee").textContent = Number.isFinite(fee) ? fmtMoney(fee) : "0,00 zł";
  }

  // -------------------------------------------------------------------------
  // Zdarzenia
  // -------------------------------------------------------------------------
  function bindEvents() {
    // Producent
    const producerInput = $("#producerSearch");
    producerInput.addEventListener("change", () => {
      const v = producerInput.value.trim();
      if (v && !STATE.producers.includes(v)) {
        const match = STATE.producers.find((p) => p.toLowerCase() === v.toLowerCase());
        if (match) { producerInput.value = match; setProducer(match); return; }
        $("#producerHint").textContent = `Nie znaleziono: „${v}"`;
        return;
      }
      setProducer(v);
    });
    producerInput.addEventListener("input", () => {
      const v = producerInput.value.trim();
      if (STATE.producers.includes(v)) setProducer(v);
    });

    // Moduł
    $("#moduleGroup").addEventListener("click", (e) => {
      const btn = e.target.closest(".seg");
      if (!btn) return;
      const val = btn.dataset.val;
      STATE.current.module = STATE.current.module === val ? "" : val;
      updateModuleUI();
      updateSummary();
      persistState();
    });

    // Pola formularza w sidebarze
    const bindVal = (sel, field) => {
      $(sel).addEventListener("input", (e) => {
        STATE.current[field] = e.target.value;
        updateSummary();
        persistState();
      });
    };
    bindVal("#fee", "fee");
    bindVal("#contact", "contact");
    bindVal("#newsletter", "newsletter");
    bindVal("#dateFrom", "dateFrom");
    bindVal("#dateTo", "dateTo");
    bindVal("#dateDoc", "dateDoc");

    // Filtr w pickerze
    $("#productSearch").addEventListener("input", (e) => {
      STATE.current.search = e.target.value;
      renderPicker();
    });

    // === PICKER ===
    const picker = $("#pickerRows");

    picker.addEventListener("change", (e) => {
      if (e.target.matches('.c-chk input[type="checkbox"]')) {
        const row = e.target.closest(".pick-row");
        if (!row) return;
        const id = row.dataset.id;
        if (e.target.checked) STATE.pickerChecked.add(id);
        else STATE.pickerChecked.delete(id);
        row.classList.toggle("picked", e.target.checked);
        syncPickerFooter();
      }
    });

    picker.addEventListener("click", (e) => {
      if (e.target.closest("input")) return;
      const row = e.target.closest(".pick-row");
      if (!row || row.classList.contains("added")) return;
      const id = row.dataset.id;
      const cb = row.querySelector('input[type="checkbox"]');
      if (!cb) return;
      cb.checked = !cb.checked;
      cb.dispatchEvent(new Event("change", { bubbles: true }));
    });

    $("#pickCheckAll").addEventListener("change", (e) => {
      const rows = $$("#pickerRows .pick-row").filter((r) => !r.classList.contains("added"));
      rows.forEach((r) => {
        const id = r.dataset.id;
        const cb = r.querySelector('input[type="checkbox"]');
        if (!cb) return;
        cb.checked = e.target.checked;
        r.classList.toggle("picked", e.target.checked);
        if (e.target.checked) STATE.pickerChecked.add(id);
        else STATE.pickerChecked.delete(id);
      });
      syncPickerFooter();
    });

    // Dodaj zaznaczone → editor
    $("#btnAddSelected").addEventListener("click", () => {
      if (STATE.pickerChecked.size === 0) return;
      const added = STATE.pickerChecked.size;
      STATE.pickerChecked.forEach((id) => STATE.current.selectedIds.add(id));
      STATE.pickerChecked.clear();
      renderPicker();
      renderEditor();
      persistState();
      showToast(`Dodano ${added} ${added === 1 ? "pozycję" : (added < 5 ? "pozycje" : "pozycji")}`);
    });

    // === EDITOR ===
    const editor = $("#editorRows");

    editor.addEventListener("input", (e) => {
      const inp = e.target.closest(".cell-input");
      if (!inp) return;
      const id = inp.dataset.id;
      const field = inp.dataset.field;
      setOverride(id, field, inp.value);
      const rowEl = inp.closest(".edit-row");
      // Dowolna zmiana może wpłynąć na którąkolwiek wyliczaną komórkę
      recomputeEditorRow(rowEl);
    });

    editor.addEventListener("click", (e) => {
      // Remove row
      const btn = e.target.closest("[data-remove]");
      if (btn) {
        const id = btn.dataset.remove;
        STATE.current.selectedIds.delete(id);
        delete STATE.current.overrides[id];
        renderEditor();
        renderPicker();
        persistState();
        return;
      }
      // Toggle Z/O
      const zoBtn = e.target.closest(".zo-btn");
      if (zoBtn) {
        const group = zoBtn.closest("[data-zo-row]");
        if (!group) return;
        const id = group.dataset.zoRow;
        const val = zoBtn.dataset.zo;
        const current = STATE.current.overrides[id]?.rabat_zo || "";
        const next = current === val ? "" : val;
        setOverride(id, "rabat_zo", next);
        group.querySelectorAll(".zo-btn").forEach((b) => {
          b.classList.toggle("active", next !== "" && b.dataset.zo === next);
        });
        return;
      }
    });

    $("#btnClearAll").addEventListener("click", () => {
      if (STATE.current.selectedIds.size === 0) return;
      if (!confirm("Usunąć wszystkie pozycje z druku?")) return;
      STATE.current.selectedIds.clear();
      STATE.current.overrides = {};
      renderEditor();
      renderPicker();
      persistState();
    });

    // Przełącznik trybu kalkulacji (kaskadowo / sumarycznie)
    document.querySelectorAll(".calc-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const mode = btn.dataset.calc;
        if (STATE.current.calcMode === mode) return;
        STATE.current.calcMode = mode;
        document.querySelectorAll(".calc-btn").forEach((b) => {
          b.classList.toggle("active", b.dataset.calc === mode);
        });
        recomputeAllEditorRows();
        persistState();
      });
    });

    $("#btnReload").addEventListener("click", loadDatabase);

    $("#btnPreview").addEventListener("click", openPreview);
    $("#btnPdf").addEventListener("click", () => { openPreview(); setTimeout(() => window.print(), 250); });
    $("#btnExport").addEventListener("click", exportXLSX);
    $("#btnPreviewClose").addEventListener("click", closePreview);
    $("#btnPreviewPrint").addEventListener("click", () => window.print());
    $("#btnPreviewExport").addEventListener("click", exportXLSX);

    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && !$("#previewOverlay").hidden) closePreview();
    });
  }

  function syncPickerFooter() {
    const checked = STATE.pickerChecked.size;
    const rows = $$("#pickerRows .pick-row").filter((r) => !r.classList.contains("added"));
    updatePickerFooter(checked, rows.length);

    const all = rows.length > 0 &&
      rows.every((r) => STATE.pickerChecked.has(r.dataset.id));
    const some = rows.some((r) => STATE.pickerChecked.has(r.dataset.id));
    const cball = $("#pickCheckAll");
    cball.checked = all;
    cball.indeterminate = !all && some;
  }

  function setProducer(name) {
    if (STATE.current.producer === name) return;
    STATE.current.producer = name;
    STATE.current.selectedIds.clear();
    STATE.current.overrides = {};
    STATE.pickerChecked.clear();
    $("#producerSearch").value = name;
    $("#producerHint").textContent = name
      ? `${STATE.producerIndex.get(name)?.length ?? 0} indeksów w bazie`
      : "—";
    renderPicker();
    renderEditor();
    persistState();
  }

  function updateModuleUI() {
    $$("#moduleGroup .seg").forEach((b) => {
      b.setAttribute("aria-checked", String(b.dataset.val === STATE.current.module));
    });
  }

  // -------------------------------------------------------------------------
  // Persistencja
  // -------------------------------------------------------------------------
  function persistState() {
    try {
      const payload = {
        ...STATE.current,
        selectedIds: Array.from(STATE.current.selectedIds),
        overrides: STATE.current.overrides,
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (e) { /* ignore */ }
  }
  function restoreState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        $("#newsletter").value = STATE.current.newsletter;
        $("#dateDoc").value = STATE.current.dateDoc;
        updateSummary();
        renderPicker();
        renderEditor();
        return;
      }
      const data = JSON.parse(raw);
      Object.assign(STATE.current, data);
      STATE.current.selectedIds = new Set(data.selectedIds || []);
      STATE.current.overrides = data.overrides || {};
      $("#fee").value = data.fee ?? "";
      $("#contact").value = data.contact ?? "";
      $("#newsletter").value = data.newsletter ?? defaultNewsletterName();
      $("#dateFrom").value = data.dateFrom ?? "";
      $("#dateTo").value = data.dateTo ?? "";
      $("#dateDoc").value = data.dateDoc ?? todayISO();
      $("#producerSearch").value = data.producer ?? "";
      if (data.producer) {
        $("#producerHint").textContent = `${STATE.producerIndex.get(data.producer)?.length ?? 0} indeksów w bazie`;
      }
      updateModuleUI();
      // Synchronizacja przełącznika trybu kalkulacji
      const mode = STATE.current.calcMode || "cascade";
      STATE.current.calcMode = mode;
      document.querySelectorAll(".calc-btn").forEach((b) => {
        b.classList.toggle("active", b.dataset.calc === mode);
      });
      renderPicker();
      renderEditor();
      updateSummary();
    } catch (e) { /* ignore */ }
  }

  // -------------------------------------------------------------------------
  // Podgląd
  // -------------------------------------------------------------------------
  function selectedProducts() {
    if (!STATE.current.producer) return [];
    return STATE.rows
      .filter((r) => r.znacznik === STATE.current.producer && STATE.current.selectedIds.has(String(r.towar_id)))
      .map(getMerged)
      .sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));
  }

  function openPreview() {
    if (!STATE.current.producer) { showToast("Najpierw wybierz producenta."); return; }
    const prods = selectedProducts();
    if (prods.length === 0) { showToast("Brak pozycji w druku."); return; }
    $("#previewSheet").innerHTML = buildSheetHTML(prods);
    $("#previewChip").textContent = `${STATE.current.producer} · ${prods.length} poz. · ${STATE.current.module || "—"}`;
    $("#previewOverlay").hidden = false;
  }
  function closePreview() { $("#previewOverlay").hidden = true; }

  function buildSheetHTML(products) {
    const c = STATE.current;
    const dateRange =
      (c.dateFrom ? fmtDatePL(c.dateFrom) : "—") + " – " + (c.dateTo ? fmtDatePL(c.dateTo) : "—");

    const rowsHTML = products
      .map((p) => {
        const calc = computeRow(p, STATE.current.calcMode);
        const vatStr = p.vat != null && p.vat !== ""
          ? (calc.vatFrac != null ? fmtNum(calc.vatFrac * 100, 0) + "%" : String(p.vat))
          : "";
        const marzaStr = calc.marza == null ? "" : fmtPct(calc.marza);
        const marzaColor = calc.marza == null ? "" :
          (calc.marza >= 0 ? "color:#1e7e32;font-weight:700" : "color:#c1272d;font-weight:700");
        return `
          <tr>
            <td class="center">${escapeHtml(String(p.towar_id))}</td>
            <td class="name">${escapeHtml(p.nazwa)}</td>
            <td class="center">${escapeHtml(String(p.jm ?? ""))}</td>
            <td class="num">${fmtNum(p.cena_famix, 2)}</td>
            <td class="center">${escapeHtml(vatStr)}</td>
            <td class="num">${p.stale != null ? fmtPct(p.stale) : ""}</td>
            <td class="num">${calc.afterStale != null ? fmtNum(calc.afterStale, 2) : ""}</td>
            <td class="num">${p.rabat_prom != null ? fmtPct(p.rabat_prom) : ""}</td>
            <td class="center" style="font-weight:700">${escapeHtml(p.rabat_zo || "")}</td>
            <td class="num">${calc.afterProm != null ? fmtNum(calc.afterProm, 2) : ""}</td>
            <td class="num">${p.refundacja != null ? fmtNum(p.refundacja, 2) : ""}</td>
            <td class="num">${p.promocja_netto != null ? fmtNum(p.promocja_netto, 2) : ""}</td>
            <td class="num">${calc.brutto != null ? fmtNum(calc.brutto, 2) : ""}</td>
            <td class="center">${escapeHtml(p.prom_rabat || "")}</td>
            <td class="center">${escapeHtml(p.prom_pakietowa || "")}</td>
            <td class="name">${escapeHtml(p.uwagi || "")}</td>
            <td class="num" style="${marzaColor}">${marzaStr}</td>
          </tr>
        `;
      })
      .join("");

    return `
      <div class="sheet" id="printSheet">
        <div class="sheet-head">
          <div class="sheet-company">
            <div>Przedsiębiorstwo Handlowe</div>
            <div class="cname">"FAMIX" Sp. z o.o.</div>
            <div>35-234 Rzeszów, ul. Trembeckiego 11</div>
            <div style="margin-top:6px">Data promocji: <strong>${dateRange}</strong></div>
          </div>
          <div class="sheet-title">
            <div class="conf">POTWIERDZENIE UDZIAŁU W PROMOCJI:</div>
            <div class="nl">${escapeHtml(c.newsletter || "Gazetka")}</div>
            <div class="dates">Pozycje w druku: ${products.length}</div>
          </div>
          <div class="sheet-date">
            <div>Rzeszów, dnia:</div>
            <div style="font-weight:700">${fmtDatePL(c.dateDoc) || "—"}</div>
          </div>
        </div>

        <div class="sheet-meta">
          <div class="m-field"><span class="m-label">Producent:</span>
            <span class="m-value highlight">${escapeHtml(c.producer || "—")}</span></div>
          <div class="m-field"><span class="m-label">Moduł:</span>
            <span class="m-value highlight">${escapeHtml(c.module || "—")}</span></div>
          <div class="m-field"><span class="m-label">Opłata:</span>
            <span class="m-value">${c.fee ? fmtMoney(c.fee) : "—"}</span></div>
          <div class="m-field"><span class="m-label">Osoba kontaktowa produc.:</span>
            <span class="m-value">${escapeHtml(c.contact || "—")}</span></div>
        </div>

        <table class="sheet-table">
          <thead>
            <tr>
              <th rowspan="2">Indeks</th>
              <th rowspan="2">Nazwa</th>
              <th rowspan="2">Jm</th>
              <th rowspan="2">Cena Fam.</th>
              <th rowspan="2">VAT</th>
              <th rowspan="2">Rabat stały</th>
              <th rowspan="2">Cena po rab. stał.</th>
              <th rowspan="2" class="orange">Rabat prom</th>
              <th rowspan="2" class="orange">Rab. Z/O</th>
              <th rowspan="2" class="orange">C. po rab. prom.</th>
              <th rowspan="2">Refund. odsp. (zł)</th>
              <th colspan="2" class="orange">Promocja cenowa</th>
              <th rowspan="2" class="orange">Prom. rabat.</th>
              <th rowspan="2" class="orange">Promocja Pakietowa</th>
              <th rowspan="2">Uwagi dot. pozycji lub modułu</th>
              <th rowspan="2" style="background:#e6e6e6">Marża %</th>
            </tr>
            <tr>
              <th class="orange">Netto</th>
              <th class="orange">Brutto</th>
            </tr>
          </thead>
          <tbody>${rowsHTML}</tbody>
        </table>

        <div class="sheet-footer">
          <div>Wygenerowano: ${new Date().toLocaleString("pl-PL")}</div>
          <div>Osoba kontaktowa: ${escapeHtml(c.contact || "—")}</div>
        </div>
      </div>
    `;
  }

  // -------------------------------------------------------------------------
  // Eksport XLSX
  // -------------------------------------------------------------------------
  function exportXLSX() {
    if (!STATE.current.producer) { showToast("Najpierw wybierz producenta."); return; }
    const prods = selectedProducts();
    if (prods.length === 0) { showToast("Brak pozycji w druku."); return; }

    const c = STATE.current;
    const wb = XLSX.utils.book_new();

    const border = { top:{style:"thin",color:{rgb:"000000"}}, bottom:{style:"thin",color:{rgb:"000000"}}, left:{style:"thin",color:{rgb:"000000"}}, right:{style:"thin",color:{rgb:"000000"}} };
    const styleTitle = { font:{bold:true,sz:14}, alignment:{horizontal:"center",vertical:"center",wrapText:true} };
    const styleNl = { font:{bold:true,sz:16}, alignment:{horizontal:"center",vertical:"center"} };
    const styleLabel = { font:{bold:true,sz:10} };
    const styleBox = { border, alignment:{horizontal:"left",vertical:"center",wrapText:true}, font:{sz:11,bold:true} };
    const styleBoxHl = { ...styleBox, fill:{fgColor:{rgb:"FDE1CF"}} };
    const styleHeaderGray = { border, fill:{fgColor:{rgb:"D9D9D9"}}, font:{bold:true,sz:10}, alignment:{horizontal:"center",vertical:"center",wrapText:true} };
    const styleHeaderOrange = { border, fill:{fgColor:{rgb:"F2B07A"}}, font:{bold:true,sz:10}, alignment:{horizontal:"center",vertical:"center",wrapText:true} };
    const styleCell = { border, font:{sz:10}, alignment:{horizontal:"center",vertical:"center",wrapText:true} };
    const styleCellName = { border, font:{sz:10}, alignment:{horizontal:"left",vertical:"center",wrapText:true} };
    const styleCellNum = { border, font:{sz:10}, alignment:{horizontal:"right",vertical:"center"}, numFmt:"#,##0.00" };
    const styleCellPct = { border, font:{sz:10}, alignment:{horizontal:"right",vertical:"center"}, numFmt:"0.00%" };

    const ws = {};
    const setCell = (addr, v, s) => {
      ws[addr] = { v, s, t: typeof v === "number" ? "n" : "s" };
    };

    // --- Nagłówek (firma / tytuł / data) ---
    setCell("A1", "Przedsiębiorstwo Handlowe", { font:{sz:10} });
    setCell("A2", '"FAMIX" Sp. z o.o.', { font:{sz:12,bold:true} });
    setCell("A3", "35-234 Rzeszów, ul. Trembeckiego 11", { font:{sz:10} });
    setCell("E1", "POTWIERDZENIE UDZIAŁU W PROMOCJI:", styleTitle);
    setCell("E2", c.newsletter || "Gazetka", styleNl);
    setCell("E3", "Data promocji: " + (c.dateFrom ? fmtDatePL(c.dateFrom) : "—") + " – " + (c.dateTo ? fmtDatePL(c.dateTo) : "—"),
      { font:{sz:10,italic:true}, alignment:{horizontal:"center"} });
    setCell("N1", "Rzeszów, dnia:", { font:{sz:10}, alignment:{horizontal:"right"} });
    setCell("N2", fmtDatePL(c.dateDoc) || "", { font:{sz:11,bold:true}, alignment:{horizontal:"right"} });

    // --- Meta (producent / moduł / opłata / kontakt) ---
    setCell("A5", "Producent:", styleLabel);
    setCell("A6", c.producer || "", styleBoxHl);
    setCell("D5", "Moduł:", styleLabel);
    setCell("D6", c.module || "", styleBoxHl);
    setCell("G5", "Opłata:", styleLabel);
    setCell("G6", c.fee ? Number(c.fee) : "", { ...styleBox, numFmt:'#,##0.00" zł"' });
    setCell("J5", "Osoba kontaktowa producenta:", styleLabel);
    setCell("J6", c.contact || "", styleBox);

    // --- Nagłówki tabeli (wiersze 8-9) ---
    // Row 8: główne nagłówki (większość z rowspan=2, stąd row 9 puste dla tych)
    const mainHeaders = [
      ["A8","Indeks",false],
      ["B8","Nazwa",false],
      ["C8","Jm",false],
      ["D8","Cena Fam.",false],
      ["E8","VAT",false],
      ["F8","Rabat stały",false],
      ["G8","Cena po rab. stał.",false],
      ["H8","Rabat prom",true],
      ["I8","Rab. Z/O",true],
      ["J8","C. po rab. prom.",true],
      ["K8","Refund. odsp. (zł)",false],
      ["L8","Promocja cenowa",true],   // merged L8:M8 (colspan 2)
      ["N8","Prom. rabat.",true],
      ["O8","Promocja Pakietowa",true],
      ["P8","Uwagi dot. pozycji lub modułu",false],
      ["Q8","Marża %",false],
    ];
    mainHeaders.forEach(([a,v,orange]) => setCell(a, v, orange ? styleHeaderOrange : styleHeaderGray));
    // M8 to prawa część merge'u L8:M8 — puste, ale ostylowane
    setCell("M8", "", styleHeaderOrange);
    // Row 9: pod-nagłówki dla kolumn L/M (Netto/Brutto) + puste ostylowane dla reszty (żeby merge rowspan=2 dobrze się odwzorował)
    setCell("L9", "Netto", styleHeaderOrange);
    setCell("M9", "Brutto", styleHeaderOrange);
    ["A","B","C","D","E","F","G","H","I","J","K","N","O","P","Q"].forEach((col) => {
      if (!ws[col + "9"]) setCell(col + "9", "", styleHeaderGray);
    });

    // --- Wiersze danych (od 10) ---
    const startRow = 10;
    prods.forEach((p, i) => {
      const r = startRow + i;
      const calc = computeRow(p, STATE.current.calcMode);

      setCell("A" + r, p.towar_id, styleCell);                                                 // Indeks
      setCell("B" + r, p.nazwa, styleCellName);                                                // Nazwa
      setCell("C" + r, p.jm || "", styleCell);                                                 // Jm
      setCell("D" + r, p.cena_famix != null ? Number(p.cena_famix) : "", styleCellNum);        // Cena Fam.
      if (calc.vatFrac != null) {
        ws["E" + r] = { v: calc.vatFrac, s: styleCellPct, t: "n" };                             // VAT (frakcja)
      } else {
        setCell("E" + r, p.vat || "", styleCell);
      }
      setCell("F" + r, p.stale != null ? Number(p.stale) : "", styleCellPct);                   // Rabat stały
      setCell("G" + r, calc.afterStale != null ? Number(calc.afterStale) : "", styleCellNum);   // Cena po rab. stał.
      setCell("H" + r, p.rabat_prom != null ? Number(p.rabat_prom) : "", styleCellPct);         // Rabat prom
      if (p.rabat_zo === "Z" || p.rabat_zo === "O") {                                           // Rab. Z/O
        ws["I" + r] = {
          v: p.rabat_zo,
          s: { border, font:{sz:11,bold:true}, alignment:{horizontal:"center",vertical:"center"} },
          t: "s",
        };
      } else {
        setCell("I" + r, "", styleCell);
      }
      setCell("J" + r, calc.afterProm != null ? Number(calc.afterProm) : "", styleCellNum);     // C. po rab. prom. (z refund.)
      setCell("K" + r, p.refundacja != null ? Number(p.refundacja) : "", styleCellNum);         // Refund. odsp.
      setCell("L" + r, p.promocja_netto != null ? Number(p.promocja_netto) : "", styleCellNum); // Prom. Netto
      setCell("M" + r, calc.brutto != null ? Number(calc.brutto) : "", styleCellNum);           // Prom. Brutto
      setCell("N" + r, p.prom_rabat || "", styleCell);                                          // Prom. rabat.
      setCell("O" + r, p.prom_pakietowa || "", styleCellName);                                  // Promocja Pakietowa
      setCell("P" + r, p.uwagi || "", styleCellName);                                           // Uwagi
      // Marża — kolorowana na zielono/czerwono
      if (calc.marza != null) {
        const marzaColor = calc.marza >= 0 ? "1E7E32" : "C1272D";
        ws["Q" + r] = {
          v: calc.marza,
          s: {
            border,
            font: { sz:10, bold:true, color:{ rgb: marzaColor } },
            alignment: { horizontal:"right", vertical:"center" },
            numFmt: "0.00%",
          },
          t: "n",
        };
      } else {
        setCell("Q" + r, "", styleCellPct);
      }
    });

    // --- Merges ---
    ws["!merges"] = [
      // Nagłówek tytułu (E1:M1, E2:M2, E3:M3)
      { s:{r:0,c:4}, e:{r:0,c:12} },
      { s:{r:1,c:4}, e:{r:1,c:12} },
      { s:{r:2,c:4}, e:{r:2,c:12} },
      // Meta: producent (A6:C6), moduł (D6:F6), opłata (G6:I6), kontakt (J6:P6)
      { s:{r:5,c:0}, e:{r:5,c:2} },
      { s:{r:5,c:3}, e:{r:5,c:5} },
      { s:{r:5,c:6}, e:{r:5,c:8} },
      { s:{r:5,c:9}, e:{r:5,c:15} },
      // Nagłówki tabeli: rowspan=2 dla kolumn które nie są pod "Promocja cenowa"
      ...["A","B","C","D","E","F","G","H","I","J","K","N","O","P","Q"].map((col) => {
        const ci = XLSX.utils.decode_col(col);
        return { s:{r:7,c:ci}, e:{r:8,c:ci} };
      }),
      // "Promocja cenowa" — colspan 2 (L8:M8)
      { s:{r:7,c:11}, e:{r:7,c:12} },
    ];

    // --- Szerokości kolumn ---
    ws["!cols"] = [
      { wch:10 },  // A Indeks
      { wch:40 },  // B Nazwa
      { wch:6 },   // C Jm
      { wch:10 },  // D Cena Fam.
      { wch:7 },   // E VAT
      { wch:10 },  // F Rabat stały
      { wch:12 },  // G Cena po rab. stał.
      { wch:10 },  // H Rabat prom
      { wch:10 },  // I Rab. Z/O
      { wch:12 },  // J C. po rab. prom.
      { wch:10 },  // K Refund.
      { wch:10 },  // L Netto
      { wch:10 },  // M Brutto
      { wch:10 },  // N Prom. rabat.
      { wch:16 },  // O Promocja Pakietowa
      { wch:30 },  // P Uwagi
      { wch:10 },  // Q Marża %
    ];
    ws["!rows"] = [ { hpx:18 },{ hpx:22 },{ hpx:18 },{ hpx:10 },{ hpx:16 },{ hpx:28 },{ hpx:10 },{ hpx:28 },{ hpx:18 } ];

    const lastRow = startRow + prods.length - 1;
    ws["!ref"] = `A1:Q${Math.max(lastRow, 10)}`;
    ws["!pageSetup"] = { orientation:"landscape", paperSize:9 };
    ws["!margins"] = { left:0.3, right:0.3, top:0.4, bottom:0.4, header:0.2, footer:0.2 };

    XLSX.utils.book_append_sheet(wb, ws, "DEKLARACJA");

    const filename = `Druk_rezerwacji_${safeName(c.producer)}_${(c.newsletter || "Gazetka").replace(/\s+/g, "_")}.xlsx`;
    XLSX.writeFile(wb, filename, { bookType:"xlsx" });
    showToast("Wyeksportowano " + filename);
  }

  function safeName(s) { return String(s || "druk").replace(/[^\w\-]+/g, "_"); }

  // Service worker
  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("./sw.js").catch((e) => console.warn("SW:", e.message));
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    loadDatabase();
  });
})();
