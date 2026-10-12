(function () {
  "use strict";

  var DATA_KEY = "compta_data_v1";
  var LOG_KEY = "compta_audit_log_v1";

  // ---------- State ----------

  var state = loadData();
  var auditLog = loadLog();
  var pendingImport = null; // holds {rows, duplicates} while duplicate modal is open
  var editingId = null; // id of the row currently being edited inline

  var ICON_PENCIL = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"></path><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"></path></svg>';
  var ICON_CHECK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>';

  // ---------- Persistence ----------

  function defaultData() {
    var now = new Date();
    var monthKey = toMonthKey(now.getFullYear(), now.getMonth() + 1);
    return {
      currentMonth: monthKey,
      currentType: "ventes",
      months: {},
      excelMapping: { date: "A", amount: "R", label: "B", constants: [] },
      multiplier: { enabled: true, value: 1000 },
      paymentMode: "ESP",
      excelPaymentFilter: "ALL"
    };
  }

  function loadData() {
    try {
      var raw = localStorage.getItem(DATA_KEY);
      if (!raw) return defaultData();
      var parsed = JSON.parse(raw);
      if (!parsed.months) parsed.months = {};
      if (!parsed.currentMonth) parsed.currentMonth = defaultData().currentMonth;
      if (!parsed.currentType) parsed.currentType = "ventes";
      if (!parsed.excelMapping || !parsed.excelMapping.date || !parsed.excelMapping.amount) {
        parsed.excelMapping = { date: "A", amount: "R", label: "B", constants: [] };
      }
      if (!Array.isArray(parsed.excelMapping.constants)) {
        parsed.excelMapping.constants = [];
      } else {
        parsed.excelMapping.constants = parsed.excelMapping.constants.filter(function (c) {
          return c && typeof c.column === "string" && /^[A-Za-z]{1,3}$/.test(c.column);
        });
      }
      // Libellé: "" means explicitly disabled; missing/invalid falls back to the default "B"
      // unless that column is already taken by a saved mapping.
      var lbl = parsed.excelMapping.label;
      if (lbl !== "" && !(typeof lbl === "string" && /^[A-Za-z]{1,3}$/.test(lbl))) {
        lbl = "B";
      }
      lbl = lbl.toUpperCase();
      var lblTaken = lbl !== "" && (
        lbl === String(parsed.excelMapping.date).toUpperCase() ||
        lbl === String(parsed.excelMapping.amount).toUpperCase() ||
        parsed.excelMapping.constants.some(function (c) { return c.column.toUpperCase() === lbl; })
      );
      parsed.excelMapping.label = lblTaken ? "" : lbl;
      if (!parsed.multiplier || typeof parsed.multiplier.value !== "number" || isNaN(parsed.multiplier.value) || parsed.multiplier.value <= 0) {
        parsed.multiplier = { enabled: true, value: 1000 };
      }
      if (!isPaymentMode(parsed.paymentMode)) parsed.paymentMode = "ESP";
      if (parsed.excelPaymentFilter !== "ALL" && !isPaymentMode(parsed.excelPaymentFilter)) {
        parsed.excelPaymentFilter = "ALL";
      }
      return parsed;
    } catch (e) {
      console.error("[COMPTA] Erreur de chargement des données, réinitialisation.", e);
      return defaultData();
    }
  }

  function saveData() {
    localStorage.setItem(DATA_KEY, JSON.stringify(state));
  }

  function loadLog() {
    try {
      var raw = localStorage.getItem(LOG_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      return [];
    }
  }

  function saveLog() {
    localStorage.setItem(LOG_KEY, JSON.stringify(auditLog));
  }

  function audit(action, details) {
    var entry = {
      ts: new Date().toISOString(),
      action: action,
      details: details || {}
    };
    auditLog.push(entry);
    saveLog();
    console.log("[AUDIT]", entry.ts, action, details);
  }

  // ---------- Helpers ----------

  var PAYMENT_MODES = ["ESP", "OV"]; // espèces, ordre de virement

  function isPaymentMode(value) {
    return PAYMENT_MODES.indexOf(value) !== -1;
  }

  // Entries saved before the payment mode existed have no `payment` field.
  function paymentBadge(payment) {
    if (!isPaymentMode(payment)) return '<span class="payment-badge">—</span>';
    return '<span class="payment-badge payment-' + payment.toLowerCase() + '">' + payment + '</span>';
  }

  function toMonthKey(year, month) {
    return year + "-" + String(month).padStart(2, "0");
  }

  function daysInMonth(monthKey) {
    var parts = monthKey.split("-").map(Number);
    return new Date(parts[0], parts[1], 0).getDate();
  }

  function getMonthBucket(monthKey) {
    if (!state.months[monthKey]) {
      state.months[monthKey] = { ventes: [], achats: [] };
    }
    return state.months[monthKey];
  }

  function getCurrentList() {
    return getMonthBucket(state.currentMonth)[state.currentType];
  }

  function genId() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  function parseAmount(raw) {
    if (typeof raw !== "string") return NaN;
    var cleaned = raw.trim().replace(/\s/g, "").replace(",", ".");
    if (cleaned === "") return NaN;
    var num = Number(cleaned);
    return num;
  }

  function formatAmount(num) {
    return num.toLocaleString("fr-FR", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }

  function formatDateFR(monthKey, day) {
    var parts = monthKey.split("-");
    return String(day).padStart(2, "0") + "/" + parts[1] + "/" + parts[0];
  }

  // Dash-separated on purpose: Excel/LibreOffice Calc remember the delimiters
  // from the last "Text to Columns" / unformatted-text-paste import and silently
  // re-split pasted values containing "/" into extra cells on subsequent pastes.
  function formatDateForSpreadsheetPaste(monthKey, day) {
    var parts = monthKey.split("-");
    return String(day).padStart(2, "0") + "-" + parts[1] + "-" + parts[0];
  }

  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function colLetterToIndex(letters) {
    var clean = String(letters || "").trim().toUpperCase();
    var index = 0;
    for (var i = 0; i < clean.length; i++) {
      var code = clean.charCodeAt(i) - 64; // A=1
      if (code < 1 || code > 26) return -1;
      index = index * 26 + code;
    }
    return clean.length ? index - 1 : -1; // 0-based
  }

  function showToast(message, isError) {
    var container = document.getElementById("toastContainer");
    var el = document.createElement("div");
    el.className = "toast" + (isError ? " error" : "");
    el.textContent = message;
    container.appendChild(el);
    setTimeout(function () {
      el.remove();
    }, 3200);
  }

  function csvEscape(value) {
    var str = String(value);
    if (/[;,"\n]/.test(str)) {
      str = '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  }

  function downloadFile(filename, content, mime) {
    var blob = new Blob([content], { type: mime || "text/csv;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  // ---------- Rendering ----------

  var els = {};

  function cacheEls() {
    els.monthInput = document.getElementById("monthInput");
    els.tabBtns = Array.prototype.slice.call(document.querySelectorAll(".tab-btn"));
    els.tabsTotal = document.getElementById("tabsTotal");
    els.tableWrapper = document.getElementById("tableWrapper");
    els.entriesBody = document.getElementById("entriesBody");
    els.emptyState = document.getElementById("emptyState");
    els.entryForm = document.getElementById("entryForm");
    els.dayInput = document.getElementById("dayInputField");
    els.amountInput = document.getElementById("amountInputField");
    els.paymentRadios = Array.prototype.slice.call(document.querySelectorAll('input[name="paymentMode"]'));
    els.excelPaymentFilter = document.getElementById("excelPaymentFilter");
    els.btnNote = document.getElementById("btnNote");
    els.noteField = document.getElementById("noteField");
    els.noteInput = document.getElementById("noteInputField");
    els.formError = document.getElementById("formError");
    els.multiplierEnabled = document.getElementById("multiplierEnabled");
    els.multiplierValue = document.getElementById("multiplierValue");
    els.btnCopyCsv = document.getElementById("btnCopyCsv");
    els.btnDownloadCsv = document.getElementById("btnDownloadCsv");
    els.fileImportCsv = document.getElementById("fileImportCsv");
    els.btnCopyExcel = document.getElementById("btnCopyExcel");
    els.btnExcelSettings = document.getElementById("btnExcelSettings");
    els.btnExportJournal = document.getElementById("btnExportJournal");
    els.modalBackdrop = document.getElementById("modalBackdrop");
    els.modalTitle = document.getElementById("modalTitle");
    els.modalBody = document.getElementById("modalBody");
    els.modalExportExisting = document.getElementById("modalExportExisting");
    els.modalCancel = document.getElementById("modalCancel");
    els.modalConfirm = document.getElementById("modalConfirm");
    els.settingsModalBackdrop = document.getElementById("settingsModalBackdrop");
    els.mapDateColumn = document.getElementById("mapDateColumn");
    els.btnTransferMonth = document.getElementById("btnTransferMonth");
    els.transferModalBackdrop = document.getElementById("transferModalBackdrop");
    els.transferSummary = document.getElementById("transferSummary");
    els.transferTargetMonth = document.getElementById("transferTargetMonth");
    els.transferError = document.getElementById("transferError");
    els.transferCancel = document.getElementById("transferCancel");
    els.transferConfirm = document.getElementById("transferConfirm");
    els.mapAmountColumn = document.getElementById("mapAmountColumn");
    els.mapLabelColumn = document.getElementById("mapLabelColumn");
    els.settingsError = document.getElementById("settingsError");
    els.settingsCancel = document.getElementById("settingsCancel");
    els.settingsSave = document.getElementById("settingsSave");
    els.constantColumnsList = document.getElementById("constantColumnsList");
    els.btnAddConstantColumn = document.getElementById("btnAddConstantColumn");
  }

  function render() {
    els.monthInput.value = state.currentMonth;
    els.tabBtns.forEach(function (btn) {
      btn.classList.toggle("active", btn.dataset.type === state.currentType);
    });

    var list = getCurrentList().slice().sort(function (a, b) {
      if (a.day !== b.day) return a.day - b.day;
      return a.createdAt.localeCompare(b.createdAt);
    });

    els.entriesBody.innerHTML = "";
    list.forEach(function (entry) {
      var tr = document.createElement("tr");
      tr.className = "row-" + state.currentType;
      if (entry.id === editingId) {
        tr.classList.add("row-editing");
        tr.dataset.id = entry.id;
        tr.innerHTML =
          '<td class="col-day"><input type="number" class="edit-day" min="1" max="' + daysInMonth(state.currentMonth) + '" value="' + entry.day + '"></td>' +
          '<td class="col-amount"><input type="text" inputmode="decimal" class="edit-amount" value="' + escapeHtml(String(entry.amount).replace(".", ",")) + '"></td>' +
          '<td class="col-payment"><select class="edit-payment">' +
            (isPaymentMode(entry.payment) ? "" : '<option value="" selected>—</option>') +
            PAYMENT_MODES.map(function (m) {
              return '<option value="' + m + '"' + (entry.payment === m ? " selected" : "") + ">" + m + "</option>";
            }).join("") +
          '</select></td>' +
          '<td class="col-label"><input type="text" class="edit-label" maxlength="120" placeholder="Libellé" value="' + escapeHtml(entry.label || "") + '"></td>' +
          '<td class="col-actions">' +
            '<button class="btn-edit-save" title="Enregistrer (Entrée)" aria-label="Enregistrer" data-id="' + entry.id + '">' + ICON_CHECK + '</button>' +
            '<button class="btn-edit-cancel" title="Annuler (Échap)" aria-label="Annuler">✕</button>' +
          '</td>';
      } else {
        tr.innerHTML =
          '<td class="col-day">' + entry.day + '</td>' +
          '<td class="col-amount">' + formatAmount(entry.amount) + '</td>' +
          '<td class="col-payment">' + paymentBadge(entry.payment) + '</td>' +
          '<td class="col-label">' + escapeHtml(entry.label || "") + '</td>' +
          '<td class="col-actions">' +
            '<button class="btn-edit" title="Modifier" aria-label="Modifier" data-id="' + entry.id + '">' + ICON_PENCIL + '</button>' +
            '<button class="btn-delete" title="Supprimer" aria-label="Supprimer" data-id="' + entry.id + '">✕</button>' +
          '</td>';
      }
      els.entriesBody.appendChild(tr);
    });

    if (editingId) {
      var firstInput = els.entriesBody.querySelector(".row-editing .edit-amount");
      if (firstInput) { firstInput.focus(); firstInput.select(); }
    }

    els.emptyState.style.display = list.length ? "none" : "block";

    var total = list.reduce(function (sum, e) { return sum + e.amount; }, 0);
    els.tabsTotal.textContent = list.length
      ? (list.length + " élément" + (list.length > 1 ? "s" : "") + " · Total : " + formatAmount(total))
      : "";

    els.tableWrapper.scrollTop = els.tableWrapper.scrollHeight;
  }

  // ---------- Entry CRUD ----------

  function addEntry(day, amount, meta) {
    var entry = {
      id: genId(),
      day: day,
      amount: amount,
      label: (meta && meta.label) || "",
      payment: meta && meta.payment,
      createdAt: new Date().toISOString()
    };
    getCurrentList().push(entry);
    saveData();
    var details = {
      month: state.currentMonth,
      type: state.currentType,
      day: day,
      amount: amount,
      payment: entry.payment
    };
    if (entry.label) details.label = entry.label;
    if (meta && meta.multiplierApplied) {
      details.rawInput = meta.rawInput;
      details.multiplier = meta.multiplier;
    }
    audit("entry_added", details);
    render();
    return entry;
  }

  function deleteEntry(id) {
    var list = getCurrentList();
    var idx = list.findIndex(function (e) { return e.id === id; });
    if (idx === -1) return;
    var removed = list.splice(idx, 1)[0];
    saveData();
    audit("entry_deleted", {
      month: state.currentMonth,
      type: state.currentType,
      day: removed.day,
      amount: removed.amount,
      payment: removed.payment || "",
      label: removed.label || ""
    });
    render();
    showToast("Ligne supprimée.");
  }

  function startEdit(id) {
    editingId = id;
    render();
  }

  function cancelEdit() {
    editingId = null;
    render();
    els.dayInput.focus();
  }

  function saveEdit() {
    var tr = els.entriesBody.querySelector(".row-editing");
    if (!tr) return;
    var entry = getCurrentList().find(function (e) { return e.id === editingId; });
    if (!entry) { cancelEdit(); return; }

    var day = parseInt(tr.querySelector(".edit-day").value, 10);
    var maxDay = daysInMonth(state.currentMonth);
    if (isNaN(day) || day < 1 || day > maxDay) {
      showToast("Jour invalide (1-" + maxDay + ").", true);
      tr.querySelector(".edit-day").focus();
      return;
    }
    var amount = parseAmount(tr.querySelector(".edit-amount").value);
    if (isNaN(amount) || amount <= 0) {
      showToast("Montant invalide.", true);
      tr.querySelector(".edit-amount").focus();
      return;
    }
    var label = tr.querySelector(".edit-label").value.replace(/\s+/g, " ").trim();
    var payment = tr.querySelector(".edit-payment").value;

    var from = { day: entry.day, amount: entry.amount, payment: entry.payment || "", label: entry.label || "" };
    entry.day = day;
    entry.amount = amount;
    if (isPaymentMode(payment)) entry.payment = payment;
    entry.label = label;
    editingId = null;
    saveData();
    audit("entry_edited", {
      month: state.currentMonth,
      type: state.currentType,
      from: from,
      to: { day: day, amount: amount, payment: entry.payment || "", label: label }
    });
    render();
    showToast("Ligne modifiée.");
    els.dayInput.focus();
  }

  // ---------- Form logic ----------

  function clampDay(day, monthKey) {
    var max = daysInMonth(monthKey);
    if (day < 1) return 1;
    if (day > max) return max;
    return day;
  }

  function setFormError(msg) {
    els.formError.textContent = msg || "";
  }

  function handleDayKeydown(e) {
    if (e.key === "Enter") {
      e.preventDefault();
      var day = parseInt(els.dayInput.value, 10);
      if (isNaN(day) || day < 1 || day > daysInMonth(state.currentMonth)) {
        setFormError("Jour invalide (1-" + daysInMonth(state.currentMonth) + ").");
        return;
      }
      setFormError("");
      els.amountInput.focus();
      els.amountInput.select();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      var d = parseInt(els.dayInput.value, 10);
      if (isNaN(d)) d = 1;
      els.dayInput.value = clampDay(d + 1, state.currentMonth);
    } else if (e.key === "ArrowDown") {
      e.preventDefault();
      var d2 = parseInt(els.dayInput.value, 10);
      if (isNaN(d2)) d2 = 1;
      els.dayInput.value = clampDay(d2 - 1, state.currentMonth);
    }
  }

  function handleAmountKeydown(e) {
    if (e.key === "Enter") {
      e.preventDefault();
      submitForm();
    }
  }

  function setNoteOpen(open) {
    els.noteField.hidden = !open;
    els.btnNote.setAttribute("aria-expanded", open ? "true" : "false");
    if (open) {
      els.noteInput.focus();
    } else {
      els.noteInput.value = "";
    }
  }

  function handleNoteKeydown(e) {
    if (e.key === "Enter") {
      e.preventDefault();
      submitForm();
    } else if (e.key === "Escape") {
      setNoteOpen(false);
      els.dayInput.focus();
    }
  }

  function submitForm() {
    var day = parseInt(els.dayInput.value, 10);
    var maxDay = daysInMonth(state.currentMonth);
    if (isNaN(day) || day < 1 || day > maxDay) {
      setFormError("Jour invalide (1-" + maxDay + ").");
      els.dayInput.focus();
      return;
    }
    var rawInput = els.amountInput.value;
    var rawAmount = parseAmount(rawInput);
    if (isNaN(rawAmount) || rawAmount <= 0) {
      setFormError("Montant invalide.");
      els.amountInput.focus();
      els.amountInput.select();
      return;
    }

    var multiplierApplied = state.multiplier.enabled;
    var amount = multiplierApplied ? rawAmount * state.multiplier.value : rawAmount;

    setFormError("");
    addEntry(day, amount, {
      multiplierApplied: multiplierApplied,
      rawInput: rawAmount,
      multiplier: state.multiplier.value,
      payment: state.paymentMode,
      label: els.noteInput.value.replace(/\s+/g, " ").trim()
    });
    els.amountInput.value = "";
    els.noteInput.value = "";
    els.dayInput.focus();
    els.dayInput.select();
  }

  // ---------- CSV export ----------

  function buildCsv() {
    var list = getCurrentList().slice().sort(function (a, b) {
      if (a.day !== b.day) return a.day - b.day;
      return a.createdAt.localeCompare(b.createdAt);
    });
    var lines = ["Date;Montant;Libellé;Mode"];
    list.forEach(function (e) {
      lines.push(formatDateFR(state.currentMonth, e.day) + ";" + String(e.amount).replace(".", ",") + ";" + csvEscape(e.label || "") + ";" + (e.payment || ""));
    });
    return lines.join("\r\n");
  }

  function exportCsvDownload() {
    var csv = buildCsv();
    var filename = "compta_" + state.currentType + "_" + state.currentMonth + ".csv";
    downloadFile(filename, csv);
    audit("csv_exported", { month: state.currentMonth, type: state.currentType, via: "download", count: getCurrentList().length });
    showToast("CSV téléchargé.");
  }

  function exportCsvClipboard() {
    var csv = buildCsv();
    navigator.clipboard.writeText(csv).then(function () {
      audit("csv_exported", { month: state.currentMonth, type: state.currentType, via: "clipboard", count: getCurrentList().length });
      showToast("CSV copié dans le presse-papier.");
    }, function () {
      showToast("Impossible de copier dans le presse-papier.", true);
    });
  }

  // ---------- Excel clipboard (column-mapped paste) ----------

  // Entries of the current list matching the payment filter chosen next to "Copier Excel".
  function getExcelList() {
    var filter = state.excelPaymentFilter;
    var list = getCurrentList();
    if (filter === "ALL") return list;
    return list.filter(function (e) { return e.payment === filter; });
  }

  // Plain text only, on purpose: an HTML clipboard flavor would let Excel apply
  // the source formatting and wipe the destination's borders/fonts. The date
  // display (JJ/MM/AAAA vs serial like 46223) therefore comes from the number
  // format of the destination column.
  function buildExcelClipboardText() {
    var dateIdx = colLetterToIndex(state.excelMapping.date);
    var amountIdx = colLetterToIndex(state.excelMapping.amount);
    var constants = (state.excelMapping.constants || []).map(function (c) {
      return { idx: colLetterToIndex(c.column), value: c.value };
    });
    var labelIdx = colLetterToIndex(state.excelMapping.label);
    var allIdx = [dateIdx, amountIdx].concat(constants.map(function (c) { return c.idx; }));
    if (labelIdx >= 0) allIdx.push(labelIdx);
    var width = Math.max.apply(null, allIdx) + 1;

    var list = getExcelList().slice().sort(function (a, b) {
      if (a.day !== b.day) return a.day - b.day;
      return a.createdAt.localeCompare(b.createdAt);
    });

    var lines = list.map(function (e) {
      var row = new Array(width).fill("");
      row[dateIdx] = formatDateForSpreadsheetPaste(state.currentMonth, e.day);
      row[amountIdx] = String(e.amount).replace(".", ",");
      constants.forEach(function (c) { row[c.idx] = c.value; });
      if (labelIdx >= 0) row[labelIdx] = (e.label || "").replace(/[\t\r\n]+/g, " ");
      return row.join("\t");
    });

    return lines.join("\r\n");
  }

  function exportExcelClipboard() {
    var list = getExcelList();
    if (list.length === 0) {
      showToast(state.excelPaymentFilter === "ALL"
        ? "Aucune ligne à copier."
        : "Aucune ligne en " + state.excelPaymentFilter + " à copier.", true);
      return;
    }
    navigator.clipboard.writeText(buildExcelClipboardText()).then(function () {
      audit("excel_copied", {
        month: state.currentMonth,
        type: state.currentType,
        count: list.length,
        paymentFilter: state.excelPaymentFilter,
        mapping: state.excelMapping
      });
      showToast(list.length + " ligne(s)" +
        (state.excelPaymentFilter === "ALL" ? "" : " " + state.excelPaymentFilter) +
        " copiée(s) pour Excel : Date en colonne " + state.excelMapping.date + ", Montant en colonne " + state.excelMapping.amount + ".");
    }, function () {
      showToast("Impossible de copier dans le presse-papier.", true);
    });
  }

  // ---------- CSV import ----------

  function parseCsv(text) {
    var lines = text.split(/\r?\n/).filter(function (l) { return l.trim() !== ""; });
    var delimiter = lines[0] && lines[0].indexOf(";") !== -1 ? ";" : ",";
    var rows = [];
    var errors = [];

    lines.forEach(function (line, idx) {
      var cells = line.split(delimiter).map(function (c) { return c.trim().replace(/^"|"$/g, ""); });
      if (idx === 0 && isNaN(Date.parse(normalizeDateForParse(cells[0])))) {
        // treat as header row, skip
        return;
      }
      var dateStr = cells[0];
      var amountStr = cells[1];
      var parsedDate = parseDateFlexible(dateStr);
      var amount = parseAmount(amountStr);
      if (!parsedDate || isNaN(amount)) {
        errors.push({ line: idx + 1, raw: line });
        return;
      }
      // Trailing "Mode" column (ESP/OV) is optional, for CSVs exported before it existed.
      var rest = cells.slice(2);
      var payment;
      if (rest.length > 1 && isPaymentMode(rest[rest.length - 1].toUpperCase())) {
        payment = rest.pop().toUpperCase();
      } else if (rest.length > 1 && rest[rest.length - 1] === "") {
        rest.pop();
      }
      rows.push({
        monthKey: toMonthKey(parsedDate.year, parsedDate.month),
        day: parsedDate.day,
        amount: amount,
        payment: payment,
        label: rest.join(delimiter).replace(/""/g, '"').trim()
      });
    });

    return { rows: rows, errors: errors };
  }

  function normalizeDateForParse(str) {
    if (!str) return "";
    return str;
  }

  function parseDateFlexible(str) {
    if (!str) return null;
    var m;
    if ((m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) {
      return { day: parseInt(m[1], 10), month: parseInt(m[2], 10), year: parseInt(m[3], 10) };
    }
    if ((m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) {
      return { day: parseInt(m[3], 10), month: parseInt(m[2], 10), year: parseInt(m[1], 10) };
    }
    return null;
  }

  function findDuplicates(rows) {
    var duplicates = [];
    rows.forEach(function (row) {
      var bucket = getMonthBucket(row.monthKey)[state.currentType];
      var match = bucket.some(function (e) { return e.day === row.day && e.amount === row.amount; });
      if (match) duplicates.push(row);
    });
    return duplicates;
  }

  function commitImport(rows) {
    var monthsTouched = {};
    rows.forEach(function (row) {
      var bucket = getMonthBucket(row.monthKey)[state.currentType];
      bucket.push({
        id: genId(),
        day: row.day,
        amount: row.amount,
        label: row.label || "",
        payment: row.payment,
        createdAt: new Date().toISOString()
      });
      monthsTouched[row.monthKey] = true;
    });
    saveData();
    audit("csv_imported", {
      type: state.currentType,
      count: rows.length,
      months: Object.keys(monthsTouched)
    });
    render();
    showToast(rows.length + " ligne(s) importée(s).");
  }

  function handleFileImport(file) {
    var reader = new FileReader();
    reader.onload = function () {
      var text = String(reader.result);
      var parsed = parseCsv(text);
      if (parsed.rows.length === 0) {
        showToast("Aucune ligne valide trouvée dans le fichier.", true);
        return;
      }
      var duplicates = findDuplicates(parsed.rows);
      if (duplicates.length > 0) {
        pendingImport = { rows: parsed.rows, duplicates: duplicates, errors: parsed.errors };
        openDuplicateModal(parsed.rows, duplicates, parsed.errors);
      } else {
        if (parsed.errors.length > 0) {
          showToast(parsed.errors.length + " ligne(s) ignorée(s) (format invalide).", true);
        }
        commitImport(parsed.rows);
      }
    };
    reader.readAsText(file, "UTF-8");
  }

  function openDuplicateModal(rows, duplicates, errors) {
    els.modalTitle.textContent = "Doublons détectés";
    var html = "<p>" + duplicates.length + " ligne(s) sur " + rows.length +
      " correspondent à des saisies déjà existantes (même jour et même montant) pour l'onglet <strong>" +
      state.currentType + "</strong>.</p>";
    html += "<ul>";
    duplicates.slice(0, 20).forEach(function (d) {
      html += "<li>" + formatDateFR(d.monthKey, d.day) + " — " + formatAmount(d.amount) + "</li>";
    });
    if (duplicates.length > 20) html += "<li>… et " + (duplicates.length - 20) + " autre(s)</li>";
    html += "</ul>";
    if (errors.length > 0) {
      html += "<p>" + errors.length + " ligne(s) au format invalide ont été ignorées.</p>";
    }
    html += "<p>Voulez-vous exporter l'existant avant de continuer, annuler l'import, ou confirmer malgré les doublons ?</p>";
    els.modalBody.innerHTML = html;
    els.modalBackdrop.classList.add("open");
  }

  function closeModal() {
    els.modalBackdrop.classList.remove("open");
    pendingImport = null;
  }

  // ---------- Excel column settings modal ----------

  function addConstantColumnRow(column, value) {
    var row = document.createElement("div");
    row.className = "constant-column-row";

    var letterInput = document.createElement("input");
    letterInput.type = "text";
    letterInput.className = "constant-col-letter";
    letterInput.maxLength = 3;
    letterInput.placeholder = "Q";
    letterInput.value = column || "";

    var valueInput = document.createElement("input");
    valueInput.type = "text";
    valueInput.className = "constant-col-value";
    valueInput.placeholder = "Valeur (ex : 0)";
    valueInput.value = value !== undefined && value !== null ? value : "";

    var removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "btn-remove-constant";
    removeBtn.title = "Supprimer cette colonne";
    removeBtn.textContent = "✕";
    removeBtn.addEventListener("click", function () {
      row.remove();
    });

    row.appendChild(letterInput);
    row.appendChild(valueInput);
    row.appendChild(removeBtn);
    els.constantColumnsList.appendChild(row);
    return row;
  }

  function renderConstantColumnRows(constants) {
    els.constantColumnsList.innerHTML = "";
    constants.forEach(function (c) {
      addConstantColumnRow(c.column, c.value);
    });
  }

  function openSettingsModal() {
    els.mapDateColumn.value = state.excelMapping.date;
    els.mapAmountColumn.value = state.excelMapping.amount;
    els.mapLabelColumn.value = state.excelMapping.label || "";
    renderConstantColumnRows(state.excelMapping.constants || []);
    els.settingsError.textContent = "";
    els.settingsModalBackdrop.classList.add("open");
    els.mapDateColumn.focus();
  }

  function closeSettingsModal() {
    els.settingsModalBackdrop.classList.remove("open");
  }

  function saveSettings() {
    var dateCol = els.mapDateColumn.value.trim().toUpperCase();
    var amountCol = els.mapAmountColumn.value.trim().toUpperCase();
    var validPattern = /^[A-Z]{1,3}$/;

    if (!validPattern.test(dateCol) || !validPattern.test(amountCol)) {
      els.settingsError.textContent = "Utilisez des lettres de colonnes Excel valides (ex : A, R, AA).";
      return;
    }
    if (dateCol === amountCol) {
      els.settingsError.textContent = "La colonne Date et la colonne Montant doivent être différentes.";
      return;
    }

    var labelCol = els.mapLabelColumn.value.trim().toUpperCase();
    if (labelCol !== "" && !validPattern.test(labelCol)) {
      els.settingsError.textContent = "Colonne Libellé invalide : \"" + labelCol + "\".";
      return;
    }
    if (labelCol !== "" && (labelCol === dateCol || labelCol === amountCol)) {
      els.settingsError.textContent = "La colonne Libellé doit être différente des colonnes Date et Montant.";
      return;
    }

    var usedCols = {};
    usedCols[dateCol] = true;
    usedCols[amountCol] = true;
    if (labelCol !== "") usedCols[labelCol] = true;

    var constants = [];
    var rows = Array.prototype.slice.call(els.constantColumnsList.querySelectorAll(".constant-column-row"));
    for (var i = 0; i < rows.length; i++) {
      var letter = rows[i].querySelector(".constant-col-letter").value.trim().toUpperCase();
      var value = rows[i].querySelector(".constant-col-value").value;

      if (letter === "" && value === "") continue; // ignore empty unused row

      if (!validPattern.test(letter)) {
        els.settingsError.textContent = "Colonne supplémentaire invalide : \"" + letter + "\".";
        return;
      }
      if (usedCols[letter]) {
        els.settingsError.textContent = "La colonne " + letter + " est utilisée plusieurs fois.";
        return;
      }
      usedCols[letter] = true;
      constants.push({ column: letter, value: value });
    }

    els.settingsError.textContent = "";
    var previous = state.excelMapping;
    state.excelMapping = { date: dateCol, amount: amountCol, label: labelCol, constants: constants };
    saveData();
    audit("excel_mapping_updated", { from: previous, to: state.excelMapping });
    closeSettingsModal();
    showToast("Mapping mis à jour : Date → " + dateCol + ", Montant → " + amountCol +
      (constants.length ? ", + " + constants.length + " colonne(s) constante(s)" : "") + ".");
  }

  // ---------- Month transfer ----------

  function nextMonthKey(monthKey) {
    var parts = monthKey.split("-").map(Number);
    return parts[1] === 12 ? toMonthKey(parts[0] + 1, 1) : toMonthKey(parts[0], parts[1] + 1);
  }

  function openTransferModal() {
    var bucket = getMonthBucket(state.currentMonth);
    var count = bucket.ventes.length + bucket.achats.length;
    if (count === 0) {
      showToast("Aucune saisie à transférer pour ce mois.", true);
      return;
    }
    els.transferSummary.innerHTML = "Transférer les <strong>" + count + "</strong> saisie(s) de <strong>" +
      state.currentMonth + "</strong> (Ventes : " + bucket.ventes.length + ", Achats : " + bucket.achats.length + ") vers :";
    els.transferTargetMonth.value = nextMonthKey(state.currentMonth);
    els.transferError.textContent = "";
    els.transferModalBackdrop.classList.add("open");
    els.transferTargetMonth.focus();
  }

  function closeTransferModal() {
    els.transferModalBackdrop.classList.remove("open");
  }

  function confirmTransfer() {
    var source = state.currentMonth;
    var target = els.transferTargetMonth.value;
    if (!/^\d{4}-\d{2}$/.test(target)) {
      els.transferError.textContent = "Choisissez un mois de destination.";
      return;
    }
    if (target === source) {
      els.transferError.textContent = "Le mois de destination doit être différent du mois source.";
      return;
    }
    var mode = document.querySelector('input[name="transferMode"]:checked').value;
    var maxDay = daysInMonth(target);
    var sourceBucket = getMonthBucket(source);
    var targetBucket = getMonthBucket(target);
    var count = 0;
    var clamped = 0;

    ["ventes", "achats"].forEach(function (type) {
      sourceBucket[type].forEach(function (entry) {
        var day = entry.day;
        if (day > maxDay) { day = maxDay; clamped++; }
        var moved = {
          id: mode === "move" ? entry.id : genId(),
          day: day,
          amount: entry.amount,
          label: entry.label || "",
          payment: entry.payment,
          createdAt: entry.createdAt
        };
        targetBucket[type].push(moved);
        count++;
      });
      if (mode === "move") sourceBucket[type] = [];
    });

    state.currentMonth = target;
    editingId = null;
    saveData();
    audit("month_transferred", { from: source, to: target, mode: mode, count: count, daysClamped: clamped });
    closeTransferModal();
    render();
    showToast(count + " saisie(s) " + (mode === "move" ? "déplacée(s)" : "copiée(s)") + " vers " + target +
      (clamped ? " (" + clamped + " jour(s) ramené(s) au dernier jour du mois)" : "") + ".");
  }

  // ---------- Journal export ----------

  function exportJournal() {
    var lines = ["Horodatage;Action;Details"];
    auditLog.forEach(function (e) {
      lines.push([e.ts, e.action, csvEscape(JSON.stringify(e.details))].join(";"));
    });
    var csv = lines.join("\r\n");
    downloadFile("compta_journal_audit_" + new Date().toISOString().slice(0, 10) + ".csv", csv);
    showToast("Journal d'audit exporté (à des fins d'audit technique).");
    console.log("[AUDIT] journal exporté,", auditLog.length, "entrées");
  }

  // ---------- Event wiring ----------

  function wireEvents() {
    els.monthInput.addEventListener("change", function () {
      var val = els.monthInput.value;
      if (!val) return;
      state.currentMonth = val;
      saveData();
      audit("month_changed", { month: val });
      render();
    });

    els.tabBtns.forEach(function (btn) {
      btn.addEventListener("click", function () {
        state.currentType = btn.dataset.type;
        saveData();
        audit("type_changed", { type: state.currentType });
        render();
        els.dayInput.focus();
      });
    });

    els.entriesBody.addEventListener("click", function (e) {
      var btn = e.target.closest("button");
      if (!btn) return;
      if (btn.classList.contains("btn-delete")) deleteEntry(btn.dataset.id);
      else if (btn.classList.contains("btn-edit")) startEdit(btn.dataset.id);
      else if (btn.classList.contains("btn-edit-save")) saveEdit();
      else if (btn.classList.contains("btn-edit-cancel")) cancelEdit();
    });

    els.entriesBody.addEventListener("keydown", function (e) {
      if (!e.target.closest(".row-editing")) return;
      if (e.key === "Enter") { e.preventDefault(); saveEdit(); }
      else if (e.key === "Escape") { e.preventDefault(); cancelEdit(); }
    });

    els.dayInput.addEventListener("keydown", handleDayKeydown);
    els.amountInput.addEventListener("keydown", handleAmountKeydown);
    els.noteInput.addEventListener("keydown", handleNoteKeydown);
    els.btnNote.addEventListener("click", function () {
      setNoteOpen(els.noteField.hidden);
    });

    els.entryForm.addEventListener("submit", function (e) {
      e.preventDefault();
      submitForm();
    });

    els.paymentRadios.forEach(function (radio) {
      radio.addEventListener("change", function () {
        if (!radio.checked) return;
        state.paymentMode = radio.value;
        saveData();
      });
    });

    els.excelPaymentFilter.addEventListener("change", function () {
      state.excelPaymentFilter = els.excelPaymentFilter.value;
      saveData();
    });

    els.multiplierEnabled.addEventListener("change", function () {
      state.multiplier.enabled = els.multiplierEnabled.checked;
      els.multiplierValue.disabled = !state.multiplier.enabled;
      saveData();
      audit("multiplier_updated", { multiplier: state.multiplier });
    });

    els.multiplierValue.addEventListener("change", function () {
      var val = parseAmount(els.multiplierValue.value);
      if (isNaN(val) || val <= 0) {
        els.multiplierValue.value = state.multiplier.value;
        return;
      }
      state.multiplier.value = val;
      els.multiplierValue.value = val;
      saveData();
      audit("multiplier_updated", { multiplier: state.multiplier });
    });

    els.btnCopyCsv.addEventListener("click", exportCsvClipboard);
    els.btnDownloadCsv.addEventListener("click", exportCsvDownload);
    els.btnCopyExcel.addEventListener("click", exportExcelClipboard);
    els.btnExcelSettings.addEventListener("click", openSettingsModal);
    els.settingsCancel.addEventListener("click", closeSettingsModal);
    els.settingsSave.addEventListener("click", saveSettings);
    els.mapAmountColumn.addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); saveSettings(); }
    });
    els.btnAddConstantColumn.addEventListener("click", function () {
      var row = addConstantColumnRow("", "");
      row.querySelector(".constant-col-letter").focus();
    });
    els.mapDateColumn.addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); els.mapAmountColumn.focus(); els.mapAmountColumn.select(); }
    });

    els.fileImportCsv.addEventListener("change", function () {
      var file = els.fileImportCsv.files[0];
      if (file) handleFileImport(file);
      els.fileImportCsv.value = "";
    });

    els.btnExportJournal.addEventListener("click", exportJournal);

    els.btnTransferMonth.addEventListener("click", openTransferModal);
    els.transferCancel.addEventListener("click", closeTransferModal);
    els.transferConfirm.addEventListener("click", confirmTransfer);
    els.transferTargetMonth.addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); confirmTransfer(); }
      else if (e.key === "Escape") { closeTransferModal(); }
    });

    els.modalExportExisting.addEventListener("click", function () {
      exportCsvDownload();
    });

    els.modalCancel.addEventListener("click", function () {
      audit("csv_import_cancelled", { reason: "duplicates" });
      closeModal();
    });

    els.modalConfirm.addEventListener("click", function () {
      var rows = pendingImport ? pendingImport.rows : [];
      closeModal();
      commitImport(rows);
    });
  }

  // ---------- Init ----------

  function initMultiplierControls() {
    els.multiplierEnabled.checked = state.multiplier.enabled;
    els.multiplierValue.value = state.multiplier.value;
    els.multiplierValue.disabled = !state.multiplier.enabled;
  }

  function initPaymentControls() {
    els.paymentRadios.forEach(function (radio) {
      radio.checked = radio.value === state.paymentMode;
    });
    els.excelPaymentFilter.value = state.excelPaymentFilter;
  }

  function init() {
    cacheEls();
    wireEvents();
    render();
    initMultiplierControls();
    initPaymentControls();
    els.dayInput.focus();
    audit("app_loaded", {});
  }

  document.addEventListener("DOMContentLoaded", init);
})();
