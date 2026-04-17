/* =============================================================================
 * Druk rezerwacji · Gazetka FAMIX
 * -----------------------------------------------------------------------------
 * - Ładuje bazę produktów z ./data/SC.xlsx
 * - Filtrowanie po producencie (kolumna `znacznik`)
 * - Wybór produktów do druku (checkboxy)
 * - Moduł (1/4, 1/2, 3/4, 1) i opłata dla całego druku
 * - Podgląd 1:1 z wzorem FAMIX + wydruk (PDF z dialogu druku)
 * - Eksport XLSX w formie zgodnej z wzorem DEKLARACJA
 * - Zapisywanie stanu w localStorage (rozpoczęte rezerwacje)
 * ===========================================================================*/

(function () {
  "use strict";

  // -------------------------------------------------------------------------
  // Stan aplikacji
  // -------------------------------------------------------------------------
  const STATE = {
    rows: [],                 // wszystkie rekordy z SC.xlsx
    producers: [],            // posortowane unikalne nazwy producentów
    producerIndex: new Map(), // nazwa -> tablica indeksów w STATE.rows
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
    },
  };

  const STORAGE_KEY = "famix_druk_rezerwacji_v1";

  // -------------------------------------------------------------------------
  // Pomocnicze
  // -------------------------------------------------------------------------
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  function todayISO() {
    const d = new Date();
    return d.toISOString().slice(0, 10);
  }

  function defaultNewsletterName() {
    const months = [
      "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
      "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień",
    ];
    const d = new Date();
    return `Gazetka ${months[d.getMonth()]} ${d.getFullYear()}`;
  }

  function fmtNum(n, digits = 2) {
    if (n === null || n === undefined || n === "" || Number.isNaN(+n)) return "";
    return (+n).toLocaleString("pl-PL", {
      minimumFractionDigits: digits,
      maximumFractionDigits: digits,
    });
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

  function showToast(msg, type = "info") {
    const t = $("#toast");
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(showToast._tid);
    showToast._tid = setTimeout(() => (t.hidden = true), 2600);
  }

  function setStatus(text, kind = "muted") {
    const c = $("#statusChip");
    c.textContent = text;
    c.className = "chip " + (kind === "ok" ? "chip-ok" : kind === "err" ? "chip-err" : "chip-muted");
  }

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
      setStatus(`Baza: ${STATE.rows.length.toLocaleString("pl-PL")} indeksów · ${STATE.producers.length} producentów`, "ok");
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
    // Tylko wiersze z nazwą produktu i znacznikiem (producentem)
    STATE.rows = rows
      .filter((r) => r.nazwa && r.znacznik)
      .map((r) => ({
        towar_id: r.towar_id ?? "",
        nazwa: String(r.nazwa).trim(),
        znacznik: String(r.znacznik).trim(),
        kod: r.kod ?? "",
        jm: r.jm ?? "",
        vat: r.vat ?? "",
        cena_s: r.cena_s ?? null,
        stale: r.stale ?? null,
        refundacje: r.refundacje ?? null,
        grupa: r.grupa ?? "",
        podgrupa: r.podgrupa ?? "",
        cena_wyliczona: r.cena_wyliczona ?? null,
        cena_min: r.cena_min ?? null,
      }));

    const pset = new Map();
    STATE.producerIndex = new Map();
    STATE.rows.forEach((r, i) => {
      if (!pset.has(r.znacznik)) pset.set(r.znacznik, 0);
      pset.set(r.znacznik, pset.get(r.znacznik) + 1);
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
        return `<option value="${escapeHtml(p)}">${count} indeks${count === 1 ? "" : "ów"}</option>`;
      })
      .join("");
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
      body.innerHTML =
        `<div class="empty"><div class="empty-ic">📦</div><p>Wybierz producenta z listy po lewej.</p></div>`;
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
    $("#contentSub").textContent = `${rows.length} pozycji widocznych · ${ids.length} w bazie`;

    if (rows.length === 0) {
      body.innerHTML =
        `<div class="empty"><div class="empty-ic">🔍</div><p>Brak pozycji pasujących do filtra.</p></div>`;
      updateSummary();
      return;
    }

    const frag = document.createDocumentFragment();
    rows.forEach((r) => {
      const selected = STATE.current.selectedIds.has(String(r.towar_id));
      const row = document.createElement("div");
      row.className = "row" + (selected ? " selected" : "");
      row.dataset.id = String(r.towar_id);
      const afterStale =
        r.cena_s != null && r.stale != null
          ? (+r.cena_s) * (1 - (+r.stale))
          : null;

      row.innerHTML = `
        <div class="c-chk"><input type="checkbox" ${selected ? "checked" : ""} data-id="${String(r.towar_id)}" aria-label="Wybierz" /></div>
        <div class="c-id">${escapeHtml(String(r.towar_id))}</div>
        <div class="c-name" title="${escapeHtml(r.nazwa)}">${escapeHtml(r.nazwa)}</div>
        <div class="c-num">${fmtNum(r.cena_s, 2)}</div>
        <div class="c-num c-sub">${r.vat ?? ""}</div>
        <div class="c-num">${r.stale != null ? fmtPct(r.stale) : ""}</div>
        <div class="c-num">${afterStale != null ? fmtNum(afterStale, 2) : ""}</div>
        <div class="c-num">${r.refundacje != null ? fmtNum(r.refundacje, 2) : ""}</div>
        <div class="c-sub" title="${escapeHtml(String(r.kod ?? ""))}">${escapeHtml(String(r.kod ?? ""))}</div>
        <div class="c-sub" title="${escapeHtml(String(r.grupa ?? ""))}">${escapeHtml(String(r.grupa ?? ""))}</div>
      `;
      frag.appendChild(row);
    });
    body.innerHTML = "";
    body.appendChild(frag);

    // Header checkbox state
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

  // -------------------------------------------------------------------------
  // Obsługa zdarzeń
  // -------------------------------------------------------------------------
  function bindEvents() {
    // Producent — datalist + input
    const producerInput = $("#producerSearch");
    producerInput.addEventListener("change", () => {
      const v = producerInput.value.trim();
      if (v && !STATE.producers.includes(v)) {
        // spróbuj dopasować case-insensitive
        const match = STATE.producers.find((p) => p.toLowerCase() === v.toLowerCase());
        if (match) {
          producerInput.value = match;
          setProducer(match);
          return;
        }
        $("#producerHint").textContent = `Nie znaleziono: „${v}”`;
        return;
      }
      setProducer(v);
    });
    producerInput.addEventListener("input", () => {
      const v = producerInput.value.trim();
      if (STATE.producers.includes(v)) setProducer(v);
    });

    // Moduł — segmented
    $("#moduleGroup").addEventListener("click", (e) => {
      const btn = e.target.closest(".seg");
      if (!btn) return;
      const val = btn.dataset.val;
      STATE.current.module = STATE.current.module === val ? "" : val;
      updateModuleUI();
      updateSummary();
      persistState();
    });

    // Pola formularza
    const bindVal = (selector, field, cast = (v) => v) => {
      $(selector).addEventListener("input", (e) => {
        STATE.current[field] = cast(e.target.value);
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

    // Filtr produktów
    $("#productSearch").addEventListener("input", (e) => {
      STATE.current.search = e.target.value;
      renderProducts();
    });

    // Checkbox w nagłówku i w wierszach
    $("#productRows").addEventListener("click", (e) => {
      const row = e.target.closest(".row");
      if (!row) return;
      const id = row.dataset.id;
      const cb = row.querySelector('input[type="checkbox"]');
      // Kliknięcie w checkbox sam toggle'uje, dla pozostałych kliknięć w wiersz też toggle
      if (e.target !== cb) {
        cb.checked = !cb.checked;
      }
      if (cb.checked) STATE.current.selectedIds.add(id);
      else STATE.current.selectedIds.delete(id);
      row.classList.toggle("selected", cb.checked);
      updateHeaderCheckbox();
      updateSummary();
      persistState();
    });

    $("#thCheck").addEventListener("change", (e) => {
      const rows = $$("#productRows .row");
      const ids = rows.map((r) => r.dataset.id);
      if (e.target.checked) {
        ids.forEach((id) => STATE.current.selectedIds.add(id));
      } else {
        ids.forEach((id) => STATE.current.selectedIds.delete(id));
      }
      rows.forEach((r) => {
        r.classList.toggle("selected", e.target.checked);
        r.querySelector('input[type="checkbox"]').checked = e.target.checked;
      });
      updateSummary();
      persistState();
    });

    $("#btnSelectAll").addEventListener("click", () => {
      const rows = $$("#productRows .row");
      rows.forEach((r) => {
        STATE.current.selectedIds.add(r.dataset.id);
        r.classList.add("selected");
        r.querySelector('input[type="checkbox"]').checked = true;
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
    $("#btnPdf").addEventListener("click", () => {
      openPreview();
      setTimeout(() => window.print(), 250);
    });
    $("#btnExport").addEventListener("click", exportXLSX);
    $("#btnPreviewClose").addEventListener("click", closePreview);
    $("#btnPreviewPrint").addEventListener("click", () => window.print());
    $("#btnPreviewExport").addEventListener("click", exportXLSX);

    // Zamknij modal po ESC
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && !$("#previewOverlay").hidden) closePreview();
    });
  }

  function setProducer(name) {
    STATE.current.producer = name;
    STATE.current.selectedIds.clear(); // reset wyboru przy zmianie producenta
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
  // Persistencja stanu
  // -------------------------------------------------------------------------
  function persistState() {
    try {
      const payload = {
        ...STATE.current,
        selectedIds: Array.from(STATE.current.selectedIds),
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (e) {
      /* ignore */
    }
  }
  function restoreState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        // zasilenie domyślnymi wartościami UI
        $("#newsletter").value = STATE.current.newsletter;
        $("#dateDoc").value = STATE.current.dateDoc;
        updateSummary();
        return;
      }
      const data = JSON.parse(raw);
      Object.assign(STATE.current, data);
      STATE.current.selectedIds = new Set(data.selectedIds || []);
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
    } catch (e) {
      /* ignore */
    }
  }

  // -------------------------------------------------------------------------
  // Podgląd druku
  // -------------------------------------------------------------------------
  function selectedProducts() {
    if (!STATE.current.producer) return [];
    const ids = STATE.producerIndex.get(STATE.current.producer) ?? [];
    const rows = ids
      .map((i) => STATE.rows[i])
      .filter((r) => STATE.current.selectedIds.has(String(r.towar_id)))
      .sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));
    return rows;
  }

  function openPreview() {
    if (!STATE.current.producer) {
      showToast("Najpierw wybierz producenta.");
      return;
    }
    const prods = selectedProducts();
    if (prods.length === 0) {
      showToast("Nie zaznaczono żadnego produktu.");
      return;
    }
    $("#previewSheet").innerHTML = buildSheetHTML(prods);
    $("#previewChip").textContent =
      `${STATE.current.producer} · ${prods.length} poz. · ${STATE.current.module || "—"}`;
    $("#previewOverlay").hidden = false;
  }
  function closePreview() { $("#previewOverlay").hidden = true; }

  function buildSheetHTML(products) {
    const c = STATE.current;
    const dateRange =
      (c.dateFrom ? fmtDatePL(c.dateFrom) : "—") +
      " – " +
      (c.dateTo ? fmtDatePL(c.dateTo) : "—");

    const rowsHTML = products
      .map((p, idx) => {
        const afterStale =
          p.cena_s != null && p.stale != null
            ? (+p.cena_s) * (1 - (+p.stale))
            : null;
        return `
          <tr>
            <td class="center">${idx + 1}</td>
            <td class="center">${escapeHtml(String(p.towar_id))}</td>
            <td class="name">${escapeHtml(p.nazwa)}</td>
            <td class="center">${escapeHtml(String(p.jm ?? ""))}</td>
            <td class="num">${fmtNum(p.cena_s, 2)}</td>
            <td class="center">${escapeHtml(String(p.vat ?? ""))}</td>
            <td class="num">${p.stale != null ? fmtPct(p.stale) : ""}</td>
            <td class="num">${afterStale != null ? fmtNum(afterStale, 2) : ""}</td>
            <td class="num"></td>
            <td class="num"></td>
            <td class="num"></td>
            <td class="num">${p.refundacje != null ? fmtNum(p.refundacje, 2) : ""}</td>
            <td class="num"></td>
            <td class="num"></td>
            <td class="num"></td>
            <td class="num"></td>
            <td class="name"></td>
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
          <div class="m-field">
            <span class="m-label">Producent:</span>
            <span class="m-value highlight">${escapeHtml(c.producer || "—")}</span>
          </div>
          <div class="m-field">
            <span class="m-label">Moduł:</span>
            <span class="m-value highlight">${escapeHtml(c.module || "—")}</span>
          </div>
          <div class="m-field">
            <span class="m-label">Opłata:</span>
            <span class="m-value">${c.fee ? fmtMoney(c.fee) : "—"}</span>
          </div>
          <div class="m-field">
            <span class="m-label">Osoba kontaktowa produc.:</span>
            <span class="m-value">${escapeHtml(c.contact || "—")}</span>
          </div>
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
  // Eksport XLSX (format zbliżony 1:1 do wzoru DEKLARACJA)
  // -------------------------------------------------------------------------
  function exportXLSX() {
    if (!STATE.current.producer) { showToast("Najpierw wybierz producenta."); return; }
    const prods = selectedProducts();
    if (prods.length === 0) { showToast("Nie zaznaczono produktów."); return; }

    const c = STATE.current;
    const wb = XLSX.utils.book_new();

    // Style
    const border = { top: { style: "thin", color: { rgb: "000000" } }, bottom: { style: "thin", color: { rgb: "000000" } }, left: { style: "thin", color: { rgb: "000000" } }, right: { style: "thin", color: { rgb: "000000" } } };
    const styleTitle = { font: { bold: true, sz: 14 }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
    const styleNl = { font: { bold: true, sz: 16 }, alignment: { horizontal: "center", vertical: "center" } };
    const styleLabel = { font: { bold: true, sz: 10 } };
    const styleBox = { border, alignment: { horizontal: "left", vertical: "center", wrapText: true }, font: { sz: 11, bold: true } };
    const styleBoxHl = { ...styleBox, fill: { fgColor: { rgb: "FDE1CF" } } };
    const styleHeaderGray = { border, fill: { fgColor: { rgb: "D9D9D9" } }, font: { bold: true, sz: 10 }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
    const styleHeaderOrange = { border, fill: { fgColor: { rgb: "F2B07A" } }, font: { bold: true, sz: 10 }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
    const styleCell = { border, font: { sz: 10 }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
    const styleCellName = { border, font: { sz: 10 }, alignment: { horizontal: "left", vertical: "center", wrapText: true } };
    const styleCellNum = { border, font: { sz: 10 }, alignment: { horizontal: "right", vertical: "center" }, numFmt: "#,##0.00" };
    const styleCellPct = { border, font: { sz: 10 }, alignment: { horizontal: "right", vertical: "center" }, numFmt: "0.00%" };

    // Arkusz
    const ws = {};
    const setCell = (addr, v, s) => { ws[addr] = { v, s, t: typeof v === "number" ? "n" : "s" }; };

    // Nagłówek firmy
    setCell("B1", "Przedsiębiorstwo Handlowe", { font: { sz: 10 } });
    setCell("B2", '"FAMIX" Sp. z o.o.', { font: { sz: 12, bold: true } });
    setCell("B3", "35-234 Rzeszów, ul. Trembeckiego 11", { font: { sz: 10 } });

    setCell("F1", "POTWIERDZENIE UDZIAŁU W PROMOCJI:", styleTitle);
    setCell("F2", c.newsletter || "Gazetka", styleNl);
    setCell("F3",
      "Data promocji: " + (c.dateFrom ? fmtDatePL(c.dateFrom) : "—") +
      " – " + (c.dateTo ? fmtDatePL(c.dateTo) : "—"),
      { font: { sz: 10, italic: true }, alignment: { horizontal: "center" } });

    setCell("O1", "Rzeszów, dnia:", { font: { sz: 10 }, alignment: { horizontal: "right" } });
    setCell("O2", fmtDatePL(c.dateDoc) || "", { font: { sz: 11, bold: true }, alignment: { horizontal: "right" } });

    // Meta (producent / moduł / opłata / osoba)
    setCell("B5", "Producent:", styleLabel);
    setCell("B6", c.producer || "", styleBoxHl);
    setCell("E5", "Moduł:", styleLabel);
    setCell("E6", c.module || "", styleBoxHl);
    setCell("H5", "Opłata:", styleLabel);
    setCell("H6", c.fee ? Number(c.fee) : "", { ...styleBox, numFmt: '#,##0.00" zł"' });
    setCell("K5", "Osoba kontaktowa producenta:", styleLabel);
    setCell("K6", c.contact || "", styleBox);

    // Nagłówki tabeli (wiersz 8 – merged headers, wiersz 9 – subcolumn dla Promocji cenowej Netto/Brutto)
    const headersRow8 = [
      ["A8", "Lp."],
      ["B8", "Indeks"],
      ["C8", "Nazwa"],
      ["D8", "Jm"],
      ["E8", "Cena Fam."],
      ["F8", "VAT"],
      ["G8", "Rabat stały"],
      ["H8", "Cena po rab. stał."],
      ["I8", "Rabat prom", true],
      ["J8", "Rab. Z/O", true],
      ["K8", "C. po rab. prom.", true],
      ["L8", "Refund. odsp. (zł)"],
      ["M8", "Promocja cenowa", true], // merged M8:N8
      ["O8", "Prom. rabat.", true],
      ["P8", "Promocja Pakietowa", true],
      ["Q8", "Uwagi dot. pozycji lub modułu"],
    ];
    headersRow8.forEach(([addr, val, orange]) => {
      setCell(addr, val, orange ? styleHeaderOrange : styleHeaderGray);
    });
    // N8 puste (zostanie mergowane z M8)
    setCell("N8", "", styleHeaderOrange);
    setCell("M9", "Netto", styleHeaderOrange);
    setCell("N9", "Brutto", styleHeaderOrange);

    // Uzupełnij puste nagłówki w wierszu 9 (merged z w. 8)
    ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "O", "P", "Q"].forEach((col) => {
      if (!ws[col + "9"]) setCell(col + "9", "", styleHeaderGray);
    });

    // Wiersze produktów – startują od 10
    const startRow = 10;
    prods.forEach((p, i) => {
      const r = startRow + i;
      const afterStale =
        p.cena_s != null && p.stale != null
          ? (+p.cena_s) * (1 - (+p.stale))
          : null;

      setCell("A" + r, i + 1, styleCell);
      setCell("B" + r, p.towar_id, styleCell);
      setCell("C" + r, p.nazwa, styleCellName);
      setCell("D" + r, p.jm || "", styleCell);
      setCell("E" + r, p.cena_s != null ? Number(p.cena_s) : "", styleCellNum);
      setCell("F" + r, p.vat || "", styleCell);
      setCell("G" + r, p.stale != null ? Number(p.stale) : "", styleCellPct);
      setCell("H" + r, afterStale != null ? Number(afterStale) : "", styleCellNum);
      setCell("I" + r, "", styleCellPct);      // Rabat prom
      setCell("J" + r, "", styleCellPct);      // Rab. Z/O
      setCell("K" + r, "", styleCellNum);      // C. po rab. prom.
      setCell("L" + r, p.refundacje != null ? Number(p.refundacje) : "", styleCellNum);
      setCell("M" + r, "", styleCellNum);      // Netto
      setCell("N" + r, "", styleCellNum);      // Brutto
      setCell("O" + r, "", styleCellPct);      // Prom. rabat.
      setCell("P" + r, "", styleCellName);     // Promocja Pakietowa
      setCell("Q" + r, "", styleCellName);     // Uwagi
    });

    // Merge ranges
    ws["!merges"] = [
      // Tytuł POTWIERDZENIE
      { s: { r: 0, c: 5 }, e: { r: 0, c: 13 } },   // F1:N1
      { s: { r: 1, c: 5 }, e: { r: 1, c: 13 } },   // F2:N2
      { s: { r: 2, c: 5 }, e: { r: 2, c: 13 } },   // F3:N3
      // Meta – pola wartości rozciągnięte na kilka kolumn
      { s: { r: 5, c: 1 }, e: { r: 5, c: 3 } },    // B6:D6 producent
      { s: { r: 5, c: 4 }, e: { r: 5, c: 6 } },    // E6:G6 moduł
      { s: { r: 5, c: 7 }, e: { r: 5, c: 9 } },    // H6:J6 opłata
      { s: { r: 5, c: 10 }, e: { r: 5, c: 16 } },  // K6:Q6 kontakt
      // Nagłówki tabeli merged pionowo dla kolumn bez podkolumn
      ...["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "O", "P", "Q"].map(
        (col) => {
          const c = XLSX.utils.decode_col(col);
          return { s: { r: 7, c }, e: { r: 8, c } };
        }
      ),
      // Nagłówek "Promocja cenowa" scalony poziomo M8:N8
      { s: { r: 7, c: 12 }, e: { r: 7, c: 13 } },
    ];

    // Szerokości kolumn
    ws["!cols"] = [
      { wch: 5 },   // A Lp.
      { wch: 10 },  // B Indeks
      { wch: 40 },  // C Nazwa
      { wch: 6 },   // D Jm
      { wch: 10 },  // E Cena Fam.
      { wch: 6 },   // F VAT
      { wch: 10 },  // G Rabat stały
      { wch: 12 },  // H Cena po rab. stał.
      { wch: 10 },  // I Rabat prom
      { wch: 10 },  // J Rab. Z/O
      { wch: 12 },  // K C. po rab. prom.
      { wch: 10 },  // L Refund.
      { wch: 10 },  // M Netto
      { wch: 10 },  // N Brutto
      { wch: 10 },  // O Prom. rabat.
      { wch: 16 },  // P Pakietowa
      { wch: 30 },  // Q Uwagi
    ];

    ws["!rows"] = [
      { hpx: 18 },  // 1
      { hpx: 22 },  // 2
      { hpx: 18 },  // 3
      { hpx: 10 },
      { hpx: 16 },  // 5
      { hpx: 28 },  // 6 meta
      { hpx: 10 },
      { hpx: 28 },  // 8 header
      { hpx: 18 },  // 9 subheader
    ];

    // Zakres
    const lastRow = startRow + prods.length - 1;
    ws["!ref"] = `A1:Q${Math.max(lastRow, 10)}`;

    // Wydruk: A4 landscape
    ws["!pageSetup"] = { orientation: "landscape", paperSize: 9 };
    ws["!margins"] = { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 };

    XLSX.utils.book_append_sheet(wb, ws, "DEKLARACJA");

    const filename =
      `Druk_rezerwacji_${safeName(c.producer)}_${(c.newsletter || "Gazetka")
        .replace(/\s+/g, "_")}.xlsx`;
    XLSX.writeFile(wb, filename, { bookType: "xlsx" });
    showToast("Wyeksportowano " + filename);
  }

  function safeName(s) {
    return String(s || "druk").replace(/[^\w\-]+/g, "_");
  }

  function escapeHtml(s) {
    if (s == null) return "";
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  // -------------------------------------------------------------------------
  // Service worker (offline)
  // -------------------------------------------------------------------------
  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .register("./sw.js")
        .catch((e) => console.warn("SW:", e.message));
    });
  }

  // -------------------------------------------------------------------------
  // Start
  // -------------------------------------------------------------------------
  document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    loadDatabase();
  });
})();
