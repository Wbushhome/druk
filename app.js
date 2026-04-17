/* =============================================================================
 * Druk rezerwacji · Gazetka FAMIX (v2)
 * -----------------------------------------------------------------------------
 * - Ładuje bazę z ./data/SC.xlsx
 * - Cena FAMIX = cena_wyliczona (fallback: cena_s)
 * - Edytowalne komórki: Cena FAMIX, VAT, Rabat stały, Rabat prom.,
 *   Promocja netto, Refundacja producenta, Uwagi
 * - Po rab. = Cena FAMIX × (1 − Rabat stały) (obliczane na żywo)
 * - Override'y per towar_id, zapisywane w localStorage
 * - Podgląd 1:1 z wzorem FAMIX + eksport XLSX + druk PDF
 * ===========================================================================*/

(function () {
  "use strict";

  // -------------------------------------------------------------------------
  // Stan aplikacji
  // -------------------------------------------------------------------------
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
      selectedIds: new Set(),
      search: "",
      overrides: {}, // { "<towar_id>": { cena_famix, vat, stale, rabat_prom, promocja_netto, refundacja, uwagi } }
    },
  };

  const STORAGE_KEY = "famix_druk_rezerwacji_v2";

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
  // Rabat/VAT wpisywane jako procent (np. 5 = 5%). Wartości < 1 traktujemy jako już-ułamek.
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
  // Ładowanie bazy z ./data/SC.xlsx
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
      showToast("Nie udało się załadować ./data/SC.xlsx — sprawdź, czy plik jest w repo.");
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
        vat: r.vat ?? "",                       // 5, 8 lub 23
        cena_famix: r.cena_s ?? null,           // <-- Cena FAMIX z kolumny cena_s
        cena_z: r.cena_z ?? null,               // cena zakupu (podgląd, niewyświetlana)
        cena_wyl: r.cena_wyliczona ?? null,     // cena wyliczona (podgląd)
        stale: r.stale ?? null,                 // w nowym pliku brak – wpisywane ręcznie
        refundacje: r.refundacje ?? null,       // j/w
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
      promocja_netto: ov.promocja_netto !== undefined ? ov.promocja_netto : null,
      refundacja: ov.refundacja !== undefined ? ov.refundacja : row.refundacje,
      uwagi: ov.uwagi !== undefined ? ov.uwagi : "",
    };
  }

  function setOverride(id, field, rawValue) {
    if (!STATE.current.overrides[id]) STATE.current.overrides[id] = {};
    const slot = STATE.current.overrides[id];

    if (field === "uwagi") {
      const v = String(rawValue || "");
      if (v === "") delete slot[field]; else slot[field] = v;
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

  // -------------------------------------------------------------------------
  // Renderowanie listy produktów
  // -------------------------------------------------------------------------
  function renderProducts() {
    const body = $("#productRows");
    const producer = STATE.current.producer;
    const search = STATE.current.search.trim().toLowerCase();

    if (!producer || !STATE.producerIndex.has(producer)) {
      $("#contentTitle").textContent = "Produkty";
      $("#contentSub").textContent = "Wybierz producenta, aby zobaczyć dostępne indeksy.";
      body.innerHTML = `<div class="empty"><div class="empty-ic">📦</div><p>Wybierz producenta z listy po lewej.</p></div>`;
      updateSummary();
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

    $("#contentTitle").textContent = `Produkty producenta: ${producer}`;
    $("#contentSub").textContent = `${rows.length} pozycji widocznych · ${ids.length} w bazie · kliknij komórkę, aby nadpisać wartość`;

    if (rows.length === 0) {
      body.innerHTML = `<div class="empty"><div class="empty-ic">🔍</div><p>Brak pozycji pasujących do filtra.</p></div>`;
      updateSummary();
      return;
    }

    const frag = document.createDocumentFragment();
    rows.forEach((r) => {
      const m = getMerged(r);
      const id = String(r.towar_id);
      const selected = STATE.current.selectedIds.has(id);
      const afterStale =
        m.cena_famix != null && m.stale != null
          ? (+m.cena_famix) * (1 - (+m.stale))
          : null;

      const row = document.createElement("div");
      row.className = "row" + (selected ? " selected" : "");
      row.dataset.id = id;
      row.innerHTML = `
        <div class="c-chk"><input type="checkbox" ${selected ? "checked" : ""} aria-label="Wybierz" /></div>
        <div class="c-id">${escapeHtml(id)}</div>
        <div class="c-name" title="${escapeAttr(r.nazwa)}">${escapeHtml(r.nazwa)}</div>

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

        <div class="c-num c-computed" data-computed="po_rab" title="Cena po rabacie stałym">${afterStale != null ? fmtNum(afterStale, 2) : "—"}</div>

        <div class="c-cell">
          <input class="cell-input num" data-id="${escapeAttr(id)}" data-field="rabat_prom"
                 type="text" inputmode="decimal"
                 value="${m.rabat_prom != null ? fmtNum((+m.rabat_prom) * 100, 2) : ""}"
                 placeholder="0,00" title="Rabat promocyjny %" />
        </div>

        <div class="c-cell">
          <input class="cell-input num" data-id="${escapeAttr(id)}" data-field="promocja_netto"
                 type="text" inputmode="decimal"
                 value="${m.promocja_netto != null ? fmtNum(m.promocja_netto, 2) : ""}"
                 placeholder="0,00" title="Promocja cenowa netto (zł)" />
        </div>

        <div class="c-cell">
          <input class="cell-input num" data-id="${escapeAttr(id)}" data-field="refundacja"
                 type="text" inputmode="decimal"
                 value="${m.refundacja != null ? fmtNum(m.refundacja, 2) : ""}"
                 placeholder="0,00" title="Refundacja odsprzedaży (zł)" />
        </div>

        <div class="c-cell">
          <input class="cell-input text" data-id="${escapeAttr(id)}" data-field="uwagi"
                 type="text" value="${escapeAttr(m.uwagi || "")}"
                 placeholder="—" title="Uwagi" />
        </div>
      `;
      frag.appendChild(row);
    });
    body.innerHTML = "";
    body.appendChild(frag);

    // Nagłówkowy checkbox
    const visibleIds = rows.map((r) => String(r.towar_id));
    const allChecked = visibleIds.length > 0 && visibleIds.every((id) => STATE.current.selectedIds.has(id));
    const someChecked = visibleIds.some((id) => STATE.current.selectedIds.has(id));
    const thCheck = $("#thCheck");
    thCheck.checked = allChecked;
    thCheck.indeterminate = !allChecked && someChecked;

    updateSummary();
  }

  function updateSummary() {
    const totalForProducer = STATE.current.producer
      ? (STATE.producerIndex.get(STATE.current.producer)?.length ?? 0)
      : 0;
    $("#sumSelected").textContent = `${STATE.current.selectedIds.size} / ${totalForProducer}`;
    $("#sumModule").textContent = STATE.current.module || "—";
    const fee = parseFloat(STATE.current.fee);
    $("#sumFee").textContent = Number.isFinite(fee) ? fmtMoney(fee) : "0,00 zł";
  }

  function recomputeRow(rowEl) {
    const id = rowEl.dataset.id;
    const base = STATE.rows.find((r) => String(r.towar_id) === id);
    if (!base) return;
    const m = getMerged(base);
    const afterStale =
      m.cena_famix != null && m.stale != null
        ? (+m.cena_famix) * (1 - (+m.stale))
        : null;
    const cell = rowEl.querySelector('[data-computed="po_rab"]');
    if (cell) cell.textContent = afterStale != null ? fmtNum(afterStale, 2) : "—";
  }

  // -------------------------------------------------------------------------
  // Zdarzenia
  // -------------------------------------------------------------------------
  function bindEvents() {
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

    $("#moduleGroup").addEventListener("click", (e) => {
      const btn = e.target.closest(".seg");
      if (!btn) return;
      const val = btn.dataset.val;
      STATE.current.module = STATE.current.module === val ? "" : val;
      updateModuleUI();
      updateSummary();
      persistState();
    });

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

    $("#productSearch").addEventListener("input", (e) => {
      STATE.current.search = e.target.value;
      renderProducts();
    });

    const body = $("#productRows");

    // Checkbox wyboru
    body.addEventListener("change", (e) => {
      if (e.target.matches('.c-chk input[type="checkbox"]')) {
        const row = e.target.closest(".row");
        const id = row.dataset.id;
        if (e.target.checked) STATE.current.selectedIds.add(id);
        else STATE.current.selectedIds.delete(id);
        row.classList.toggle("selected", e.target.checked);
        updateHeaderCheckbox();
        updateSummary();
        persistState();
      }
    });

    // Edycja komórki
    body.addEventListener("input", (e) => {
      const inp = e.target.closest(".cell-input");
      if (!inp) return;
      const id = inp.dataset.id;
      const field = inp.dataset.field;
      setOverride(id, field, inp.value);
      const rowEl = inp.closest(".row");
      if (field === "cena_famix" || field === "stale") recomputeRow(rowEl);
    });

    $("#thCheck").addEventListener("change", (e) => {
      const rows = $$("#productRows .row");
      rows.forEach((r) => {
        const id = r.dataset.id;
        const cb = r.querySelector('.c-chk input[type="checkbox"]');
        cb.checked = e.target.checked;
        r.classList.toggle("selected", e.target.checked);
        if (e.target.checked) STATE.current.selectedIds.add(id);
        else STATE.current.selectedIds.delete(id);
      });
      updateSummary();
      persistState();
    });

    $("#btnSelectAll").addEventListener("click", () => {
      const rows = $$("#productRows .row");
      rows.forEach((r) => {
        STATE.current.selectedIds.add(r.dataset.id);
        r.classList.add("selected");
        r.querySelector('.c-chk input[type="checkbox"]').checked = true;
      });
      updateHeaderCheckbox();
      updateSummary();
      persistState();
    });
    $("#btnClear").addEventListener("click", () => {
      STATE.current.selectedIds.clear();
      renderProducts();
      persistState();
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

  function setProducer(name) {
    STATE.current.producer = name;
    STATE.current.selectedIds.clear();
    $("#producerSearch").value = name;
    $("#producerHint").textContent = name
      ? `${STATE.producerIndex.get(name)?.length ?? 0} indeksów w bazie`
      : "—";
    renderProducts();
    persistState();
  }

  function updateModuleUI() {
    $$("#moduleGroup .seg").forEach((b) => {
      b.setAttribute("aria-checked", String(b.dataset.val === STATE.current.module));
    });
  }

  function updateHeaderCheckbox() {
    const ids = $$("#productRows .row").map((r) => r.dataset.id);
    const allChecked = ids.length > 0 && ids.every((id) => STATE.current.selectedIds.has(id));
    const some = ids.some((id) => STATE.current.selectedIds.has(id));
    const th = $("#thCheck");
    th.checked = allChecked;
    th.indeterminate = !allChecked && some;
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
      renderProducts();
      updateSummary();
    } catch (e) { /* ignore */ }
  }

  // -------------------------------------------------------------------------
  // Podgląd
  // -------------------------------------------------------------------------
  function selectedProducts() {
    if (!STATE.current.producer) return [];
    const ids = STATE.producerIndex.get(STATE.current.producer) ?? [];
    return ids
      .map((i) => STATE.rows[i])
      .filter((r) => STATE.current.selectedIds.has(String(r.towar_id)))
      .map(getMerged)
      .sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));
  }

  function openPreview() {
    if (!STATE.current.producer) { showToast("Najpierw wybierz producenta."); return; }
    const prods = selectedProducts();
    if (prods.length === 0) { showToast("Nie zaznaczono żadnego produktu."); return; }
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
      .map((p, idx) => {
        const afterStale =
          p.cena_famix != null && p.stale != null
            ? (+p.cena_famix) * (1 - (+p.stale))
            : null;
        const afterProm =
          afterStale != null && p.rabat_prom != null
            ? afterStale * (1 - (+p.rabat_prom))
            : null;
        const vatFrac = parsePct(p.vat);
        const vatStr = p.vat != null && p.vat !== ""
          ? (vatFrac != null ? fmtNum(vatFrac * 100, 0) + "%" : String(p.vat))
          : "";
        const brutto = p.promocja_netto != null && vatFrac != null
          ? (+p.promocja_netto) * (1 + vatFrac)
          : null;
        return `
          <tr>
            <td class="center">${idx + 1}</td>
            <td class="center">${escapeHtml(String(p.towar_id))}</td>
            <td class="name">${escapeHtml(p.nazwa)}</td>
            <td class="center">${escapeHtml(String(p.jm ?? ""))}</td>
            <td class="num">${fmtNum(p.cena_famix, 2)}</td>
            <td class="center">${escapeHtml(vatStr)}</td>
            <td class="num">${p.stale != null ? fmtPct(p.stale) : ""}</td>
            <td class="num">${afterStale != null ? fmtNum(afterStale, 2) : ""}</td>
            <td class="num">${p.rabat_prom != null ? fmtPct(p.rabat_prom) : ""}</td>
            <td class="num"></td>
            <td class="num">${afterProm != null ? fmtNum(afterProm, 2) : ""}</td>
            <td class="num">${p.refundacja != null ? fmtNum(p.refundacja, 2) : ""}</td>
            <td class="num">${p.promocja_netto != null ? fmtNum(p.promocja_netto, 2) : ""}</td>
            <td class="num">${brutto != null ? fmtNum(brutto, 2) : ""}</td>
            <td class="num"></td>
            <td class="num"></td>
            <td class="name">${escapeHtml(p.uwagi || "")}</td>
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
            <div class="dates">Zaznaczone pozycje: ${products.length}</div>
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
              <th rowspan="2">Lp.</th>
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
    if (prods.length === 0) { showToast("Nie zaznaczono produktów."); return; }

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

    setCell("B1", "Przedsiębiorstwo Handlowe", { font:{sz:10} });
    setCell("B2", '"FAMIX" Sp. z o.o.', { font:{sz:12,bold:true} });
    setCell("B3", "35-234 Rzeszów, ul. Trembeckiego 11", { font:{sz:10} });
    setCell("F1", "POTWIERDZENIE UDZIAŁU W PROMOCJI:", styleTitle);
    setCell("F2", c.newsletter || "Gazetka", styleNl);
    setCell("F3", "Data promocji: " + (c.dateFrom ? fmtDatePL(c.dateFrom) : "—") + " – " + (c.dateTo ? fmtDatePL(c.dateTo) : "—"),
      { font:{sz:10,italic:true}, alignment:{horizontal:"center"} });
    setCell("O1", "Rzeszów, dnia:", { font:{sz:10}, alignment:{horizontal:"right"} });
    setCell("O2", fmtDatePL(c.dateDoc) || "", { font:{sz:11,bold:true}, alignment:{horizontal:"right"} });

    setCell("B5", "Producent:", styleLabel);
    setCell("B6", c.producer || "", styleBoxHl);
    setCell("E5", "Moduł:", styleLabel);
    setCell("E6", c.module || "", styleBoxHl);
    setCell("H5", "Opłata:", styleLabel);
    setCell("H6", c.fee ? Number(c.fee) : "", { ...styleBox, numFmt:'#,##0.00" zł"' });
    setCell("K5", "Osoba kontaktowa producenta:", styleLabel);
    setCell("K6", c.contact || "", styleBox);

    const headersRow8 = [
      ["A8","Lp."],["B8","Indeks"],["C8","Nazwa"],["D8","Jm"],["E8","Cena Fam."],
      ["F8","VAT"],["G8","Rabat stały"],["H8","Cena po rab. stał."],
      ["I8","Rabat prom",true],["J8","Rab. Z/O",true],["K8","C. po rab. prom.",true],
      ["L8","Refund. odsp. (zł)"],["M8","Promocja cenowa",true],
      ["O8","Prom. rabat.",true],["P8","Promocja Pakietowa",true],
      ["Q8","Uwagi dot. pozycji lub modułu"],
    ];
    headersRow8.forEach(([a,v,orange]) => setCell(a, v, orange ? styleHeaderOrange : styleHeaderGray));
    setCell("N8", "", styleHeaderOrange);
    setCell("M9", "Netto", styleHeaderOrange);
    setCell("N9", "Brutto", styleHeaderOrange);
    ["A","B","C","D","E","F","G","H","I","J","K","L","O","P","Q"].forEach((col) => {
      if (!ws[col + "9"]) setCell(col + "9", "", styleHeaderGray);
    });

    const startRow = 10;
    prods.forEach((p, i) => {
      const r = startRow + i;
      const afterStale = p.cena_famix != null && p.stale != null ? (+p.cena_famix) * (1 - (+p.stale)) : null;
      const afterProm = afterStale != null && p.rabat_prom != null ? afterStale * (1 - (+p.rabat_prom)) : null;
      const vatFrac = parsePct(p.vat);
      const brutto = p.promocja_netto != null && vatFrac != null ? (+p.promocja_netto) * (1 + vatFrac) : null;

      setCell("A" + r, i + 1, styleCell);
      setCell("B" + r, p.towar_id, styleCell);
      setCell("C" + r, p.nazwa, styleCellName);
      setCell("D" + r, p.jm || "", styleCell);
      setCell("E" + r, p.cena_famix != null ? Number(p.cena_famix) : "", styleCellNum);
      if (vatFrac != null) {
        ws["F" + r] = { v: vatFrac, s: styleCellPct, t: "n" };
      } else {
        setCell("F" + r, p.vat || "", styleCell);
      }
      setCell("G" + r, p.stale != null ? Number(p.stale) : "", styleCellPct);
      setCell("H" + r, afterStale != null ? Number(afterStale) : "", styleCellNum);
      setCell("I" + r, p.rabat_prom != null ? Number(p.rabat_prom) : "", styleCellPct);
      setCell("J" + r, "", styleCellPct);
      setCell("K" + r, afterProm != null ? Number(afterProm) : "", styleCellNum);
      setCell("L" + r, p.refundacja != null ? Number(p.refundacja) : "", styleCellNum);
      setCell("M" + r, p.promocja_netto != null ? Number(p.promocja_netto) : "", styleCellNum);
      setCell("N" + r, brutto != null ? Number(brutto) : "", styleCellNum);
      setCell("O" + r, "", styleCellPct);
      setCell("P" + r, "", styleCellName);
      setCell("Q" + r, p.uwagi || "", styleCellName);
    });

    ws["!merges"] = [
      { s:{r:0,c:5}, e:{r:0,c:13} },
      { s:{r:1,c:5}, e:{r:1,c:13} },
      { s:{r:2,c:5}, e:{r:2,c:13} },
      { s:{r:5,c:1}, e:{r:5,c:3} },
      { s:{r:5,c:4}, e:{r:5,c:6} },
      { s:{r:5,c:7}, e:{r:5,c:9} },
      { s:{r:5,c:10}, e:{r:5,c:16} },
      ...["A","B","C","D","E","F","G","H","I","J","K","L","O","P","Q"].map((col) => {
        const c = XLSX.utils.decode_col(col);
        return { s:{r:7,c}, e:{r:8,c} };
      }),
      { s:{r:7,c:12}, e:{r:7,c:13} },
    ];

    ws["!cols"] = [
      { wch:5 },{ wch:10 },{ wch:40 },{ wch:6 },{ wch:10 },{ wch:7 },{ wch:10 },{ wch:12 },
      { wch:10 },{ wch:10 },{ wch:12 },{ wch:10 },{ wch:10 },{ wch:10 },{ wch:10 },{ wch:16 },{ wch:30 },
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
