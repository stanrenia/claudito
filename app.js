(function () {
  "use strict";

  var DATA_KEY = "compta_data_v1";
  var LOG_KEY = "compta_audit_log_v1";

  // ---------- State ----------

  var state = loadData();
  var auditLog = loadLog();
  var pendingImport = null; // holds {rows, duplicates} while duplicate modal is open

  // ---------- Persistence ----------

  function defaultData() {
    var now = new Date();
    var monthKey = toMonthKey(now.getFullYear(), now.getMonth() + 1);
    return {
      currentMonth: monthKey,
      currentType: "ventes",
      months: {},
      excelMapping: { date: "A", amount: "R", constants: [] },
      multiplier: { enabled: true, value: 1000 }
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
        parsed.excelMapping = { date: "A", amount: "R", constants: [] };
      }
      if (!Array.isArray(parsed.excelMapping.constants)) {
        parsed.excelMapping.constants = [];
      } else {
        parsed.excelMapping.constants = parsed.excelMapping.constants.filter(function (c) {
          return c && typeof c.column === "string" && /^[A-Za-z]{1,3}$/.test(c.column);
        });
      }
      if (!parsed.multiplier || typeof parsed.multiplier.value !== "number" || isNaN(parsed.multiplier.value) || parsed.multiplier.value <= 0) {
        parsed.multiplier = { enabled: true, value: 1000 };
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
    els.mapAmountColumn = document.getElementById("mapAmountColumn");
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
      tr.innerHTML =
        '<td class="col-day">' + entry.day + '</td>' +
        '<td class="col-amount">' + formatAmount(entry.amount) + '</td>' +
        '<td class="col-actions"><button class="btn-delete" title="Supprimer" data-id="' + entry.id + '">✕</button></td>';
      els.entriesBody.appendChild(tr);
    });

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
      createdAt: new Date().toISOString()
    };
    getCurrentList().push(entry);
    saveData();
    var details = {
      month: state.currentMonth,
      type: state.currentType,
      day: day,
      amount: amount
    };
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
      amount: removed.amount
    });
    render();
    showToast("Ligne supprimée.");
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
      multiplier: state.multiplier.value
    });
    els.amountInput.value = "";
    els.dayInput.focus();
    els.dayInput.select();
  }

  // ---------- CSV export ----------

  function buildCsv() {
    var list = getCurrentList().slice().sort(function (a, b) {
      if (a.day !== b.day) return a.day - b.day;
      return a.createdAt.localeCompare(b.createdAt);
    });
    var lines = ["Date;Montant"];
    list.forEach(function (e) {
      lines.push(formatDateFR(state.currentMonth, e.day) + ";" + String(e.amount).replace(".", ","));
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

  function buildExcelClipboardText() {
    var dateIdx = colLetterToIndex(state.excelMapping.date);
    var amountIdx = colLetterToIndex(state.excelMapping.amount);
    var constants = (state.excelMapping.constants || []).map(function (c) {
      return { idx: colLetterToIndex(c.column), value: c.value };
    });
    var allIdx = [dateIdx, amountIdx].concat(constants.map(function (c) { return c.idx; }));
    var width = Math.max.apply(null, allIdx) + 1;

    var list = getCurrentList().slice().sort(function (a, b) {
      if (a.day !== b.day) return a.day - b.day;
      return a.createdAt.localeCompare(b.createdAt);
    });

    var lines = list.map(function (e) {
      var row = new Array(width).fill("");
      row[dateIdx] = formatDateFR(state.currentMonth, e.day);
      row[amountIdx] = String(e.amount).replace(".", ",");
      constants.forEach(function (c) { row[c.idx] = c.value; });
      return row.join("\t");
    });

    return lines.join("\r\n");
  }

  function exportExcelClipboard() {
    var list = getCurrentList();
    if (list.length === 0) {
      showToast("Aucune ligne à copier.", true);
      return;
    }
    var text = buildExcelClipboardText();
    navigator.clipboard.writeText(text).then(function () {
      audit("excel_copied", {
        month: state.currentMonth,
        type: state.currentType,
        count: list.length,
        mapping: state.excelMapping
      });
      showToast("Collé pour Excel : Date en colonne " + state.excelMapping.date + ", Montant en colonne " + state.excelMapping.amount + ".");
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
      rows.push({
        monthKey: toMonthKey(parsedDate.year, parsedDate.month),
        day: parsedDate.day,
        amount: amount
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

    var usedCols = {};
    usedCols[dateCol] = true;
    usedCols[amountCol] = true;

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
    state.excelMapping = { date: dateCol, amount: amountCol, constants: constants };
    saveData();
    audit("excel_mapping_updated", { from: previous, to: state.excelMapping });
    closeSettingsModal();
    showToast("Mapping mis à jour : Date → " + dateCol + ", Montant → " + amountCol +
      (constants.length ? ", + " + constants.length + " colonne(s) constante(s)" : "") + ".");
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
      var btn = e.target.closest(".btn-delete");
      if (!btn) return;
      deleteEntry(btn.dataset.id);
    });

    els.dayInput.addEventListener("keydown", handleDayKeydown);
    els.amountInput.addEventListener("keydown", handleAmountKeydown);

    els.entryForm.addEventListener("submit", function (e) {
      e.preventDefault();
      submitForm();
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

  function init() {
    cacheEls();
    wireEvents();
    render();
    initMultiplierControls();
    els.dayInput.focus();
    audit("app_loaded", {});
  }

  document.addEventListener("DOMContentLoaded", init);
})();
