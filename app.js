(function () {
  const STORAGE_KEY = "custom-module-drop-tracker-cache-v2";
  const LEGACY_STORAGE_KEY = "custom-module-drop-tracker-v1";
  const USER_KEY = "custom-module-drop-tracker-user";
  const USER_OPTIONS_KEY = "custom-module-drop-tracker-user-options";
  const DAY_SYNC_META_KEY = "custom-module-drop-tracker-day-sync-v1";
  const DROP_RATE = 0.58;
  const DROP_DISTRIBUTION = { 1: 0.700963, 2: 0.289037, 3: 0.01 };
  const PIECES_PER_CHALLENGE = 37;
  const INITIAL_DATA = window.__INITIAL_DATA__ || {};
  const BACKEND_CONFIG = window.__BACKEND_CONFIG__ || {
    appsScriptUrl: "",
    apiToken: "",
    defaultUserId: "",
    autoSync: true,
  };

  const state = {
    records: [],
    archives: [],
    selectedDate: todayIso(),
    selectedMonth: monthKey(todayIso()),
    historyBaseMonth: monthKey(todayIso()),
    userId: loadUserId(),
    pendingUserId: loadUserId(),
    userOptions: [],
    autoSyncTimer: null,
    dayWatcherTimer: null,
    needsSelectedDateSync: false,
  };

  const elements = {
    editorTitle: document.getElementById("editor-title"),
    dateInput: document.getElementById("date-input"),
    fullBurstToggle: document.getElementById("full-burst-toggle"),
    slotGrid: document.getElementById("slot-grid"),
    saveStatus: document.getElementById("save-status"),
    monthSelect: document.getElementById("month-select"),
    historyMonthSelect: document.getElementById("history-month-select"),
    monthlyStats: document.getElementById("monthly-stats"),
    lifetimeStats: document.getElementById("lifetime-stats"),
    historyList: document.getElementById("history-list"),
    todayButton: document.getElementById("today-button"),
    clearDayButton: document.getElementById("clear-day-button"),
    exportButton: document.getElementById("export-button"),
    importButton: document.getElementById("import-button"),
    importFileInput: document.getElementById("import-file-input"),
    syncUploadButton: document.getElementById("sync-upload-button"),
    userIdSelect: document.getElementById("user-id-select"),
    applyUserButton: document.getElementById("apply-user-button"),
    userSwitchStatus: document.getElementById("user-switch-status"),
    currentProfileChip: document.getElementById("current-profile-chip"),
  };

  initialize();

  async function initialize() {
    const localData = state.userId ? loadLocalData(state.userId) : { records: [], archives: [] };
    state.records = mergeRecords([], localData.records);
    state.archives = mergeArchives([], localData.archives);
    state.userOptions = loadUserOptions();
    state.selectedMonth = availableMonths().includes(state.selectedMonth) ? state.selectedMonth : latestMonth();
    state.historyBaseMonth = availableMonths().includes(state.historyBaseMonth) ? state.historyBaseMonth : latestMonth();
    elements.dateInput.value = state.selectedDate;
    attachEvents();
    renderUserSelect();
    syncEditorFromRecord();
    refreshSelectedDateSyncState();
    render();
    setUserSwitchStatus(state.userId ? "待機中" : "ユーザー未選択", state.userId ? "idle" : "error");
    startDayWatcher();

  }

  function attachEvents() {
    elements.dateInput.addEventListener("change", () => {
      state.selectedDate = normalizeIsoDate(elements.dateInput.value) || todayIso();
      state.selectedMonth = monthKey(state.selectedDate);
      syncEditorFromRecord();
      render();
    });

    elements.fullBurstToggle.addEventListener("change", () => {
      const record = getRecord(state.selectedDate) || createEmptyRecord(state.selectedDate);
      record.isFullBurst = elements.fullBurstToggle.checked;
      record.slots = normalizeSlots(record.slots, record.isFullBurst ? 6 : 3);
      upsertRecord(record);
      persist();
      render();
      scheduleAutoSync();
    });

    elements.monthSelect.addEventListener("change", () => {
      const nextMonth = normalizeMonthKey(elements.monthSelect.value);
      if (nextMonth) {
        state.selectedMonth = nextMonth;
        render();
      }
    });

    elements.historyMonthSelect.addEventListener("change", () => {
      const nextMonth = normalizeMonthKey(elements.historyMonthSelect.value);
      if (nextMonth) {
        state.historyBaseMonth = nextMonth;
        render();
      }
    });

    elements.todayButton.addEventListener("click", () => {
      state.selectedDate = todayIso();
      state.selectedMonth = monthKey(state.selectedDate);
      state.historyBaseMonth = state.selectedMonth;
      syncEditorFromRecord();
      render();
    });

    elements.clearDayButton.addEventListener("click", async () => {
      deleteRecord(state.selectedDate);
      persist();
      syncEditorFromRecord();
      refreshSelectedDateSyncState();
      render();
      if (hasBackendConfig() && state.userId) {
        await syncUpload(false);
        elements.saveStatus.textContent = `${formatDateLabel(state.selectedDate)} を削除してスプシへ反映しました`;
      }
    });

    elements.exportButton.addEventListener("click", exportRecords);
    elements.importButton.addEventListener("click", () => elements.importFileInput.click());
    elements.importFileInput.addEventListener("change", importRecords);
    elements.syncUploadButton.addEventListener("click", () => syncUpload(false));
    elements.userIdSelect.addEventListener("change", () => {
      state.pendingUserId = sanitizeUserId(elements.userIdSelect.value);
      updateUserSwitchUi();
    });
    elements.applyUserButton.addEventListener("click", applyUser);
  }

  function render() {
    renderUserSelect();
    updateSyncUploadButton();
    renderEditor();
    renderMonthSelect();
    renderHistoryMonthSelect();
    renderStats(elements.monthlyStats, buildMonthCards());
    renderStats(elements.lifetimeStats, buildLifetimeCards());
    renderHistory();
  }

  function updateSyncUploadButton() {
    elements.syncUploadButton.classList.toggle("attention", state.needsSelectedDateSync);
  }

  function renderUserSelect() {
    const options = buildUserOptions();
    const selectedValue = options.includes(state.pendingUserId)
      ? state.pendingUserId
      : (options.includes(state.userId) ? state.userId : "");

    elements.userIdSelect.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "ユーザーを選択";
    placeholder.selected = !selectedValue;
    elements.userIdSelect.appendChild(placeholder);
    options.forEach((userId) => {
      const option = document.createElement("option");
      option.value = userId;
      option.textContent = userId;
      option.selected = userId === selectedValue;
      elements.userIdSelect.appendChild(option);
    });

    updateCurrentProfileChip();
    updateUserSwitchUi();
  }

  function renderEditor() {
    const record = getRecord(state.selectedDate) || createEmptyRecord(state.selectedDate);
    const slotCount = record.isFullBurst ? 6 : 3;
    const slots = normalizeSlots(record.slots, slotCount);
    elements.editorTitle.textContent = `${formatDateLabel(state.selectedDate)} の入力`;
    elements.dateInput.value = state.selectedDate;
    elements.fullBurstToggle.checked = record.isFullBurst;
    elements.slotGrid.innerHTML = "";

    slots.forEach((value, index) => {
      const card = document.createElement("article");
      card.className = `slot-card${index >= 3 ? " full-burst" : ""}`;
      const slotLabel = index < 3 ? `通常枠 ${index + 1}` : `追加枠 ${index - 2}`;
      card.innerHTML = `
        <h3>${slotLabel}</h3>
        <p>${describeSlot(value)}</p>
      `;

      const buttons = document.createElement("div");
      buttons.className = "value-buttons";
      [0, 1, 2, 3].forEach((choice) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = `value-button${choice === value ? " is-active" : ""}`;
        button.textContent = String(choice);
        button.addEventListener("click", () => {
          updateSlot(index, choice);
        });
        buttons.appendChild(button);
      });

      card.appendChild(buttons);
      elements.slotGrid.appendChild(card);
    });
  }

  function renderMonthSelect() {
    const months = availableMonths();
    if (!months.includes(state.selectedMonth)) {
      state.selectedMonth = latestMonth();
    }

    elements.monthSelect.innerHTML = "";
    months.forEach((month) => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = formatMonthLabel(month);
      option.selected = month === state.selectedMonth;
      elements.monthSelect.appendChild(option);
    });
  }

  function renderHistoryMonthSelect() {
    const months = availableMonths();
    if (!months.includes(state.historyBaseMonth)) {
      state.historyBaseMonth = latestMonth();
    }

    elements.historyMonthSelect.innerHTML = "";
    months.forEach((month) => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = formatMonthLabel(month);
      option.selected = month === state.historyBaseMonth;
      elements.historyMonthSelect.appendChild(option);
    });
  }

  function renderStats(target, cards) {
    target.innerHTML = "";
    cards.forEach((card) => {
      const article = document.createElement("article");
      article.className = `stat-card${card.featured ? " featured" : ""}${card.spanAll ? " span-all" : ""}`;
      article.innerHTML = `
        <div class="label">${card.label}</div>
        <div class="value">${card.value}</div>
        ${card.note ? `<div class="note">${card.note}</div>` : ""}
      `;
      target.appendChild(article);
    });
  }

  function renderHistory() {
    const records = recordsForMonth(state.historyBaseMonth).slice().sort((left, right) => right.date.localeCompare(left.date));
    const archive = archiveForMonth(state.historyBaseMonth);
    elements.historyList.innerHTML = "";

    if (!records.length && !archive) {
      elements.historyList.innerHTML = [
        '<div class="history-empty">',
        "<strong>まだ記録がありません</strong>",
        "<div>左側の入力から記録を追加できます。</div>",
        "</div>",
      ].join("");
      return;
    }

    if (!records.length && archive) {
      const item = document.createElement("div");
      item.className = "history-item";
      item.innerHTML = `
        <div class="history-date">
          <strong>${formatMonthLabel(archive.month)}</strong>
          <span>月次アーカイブ / ${formatInteger(archive.elapsedDays)}日</span>
        </div>
        <div>
          <span class="pill">Archive</span>
        </div>
        <div class="history-breakdown">
          <strong>${formatInteger(archive.oneDrops)} / ${formatInteger(archive.twoDrops)} / ${formatInteger(archive.threeDrops)}</strong>
          <span>1 / 2 / 3 Drop</span>
        </div>
        <div class="history-total">
          <strong>${formatInteger(archive.totalModules)}</strong>
          <span>modules</span>
        </div>
      `;
      elements.historyList.appendChild(item);
      return;
    }

    records.forEach((record) => {
      const total = record.slots.reduce((sum, value) => sum + value, 0);
      const hitCount = record.slots.filter((value) => value > 0).length;
      const item = document.createElement("div");
      item.className = "history-item";
      item.innerHTML = `
        <div class="history-date">
          <strong>${formatDateLabel(record.date)}</strong>
          <span>${formatInteger(record.isFullBurst ? 6 : 3)}枠 / ヒット ${formatInteger(hitCount)}回</span>
        </div>
        <div>
          <span class="pill${record.isFullBurst ? " full" : ""}">${record.isFullBurst ? "Full Burst" : "Normal"}</span>
        </div>
        <div class="history-breakdown">
          <strong>${record.slots.map(formatInteger).join(" / ")}</strong>
          <span>各枠ドロップ数</span>
        </div>
        <div class="history-total">
          <strong>${formatInteger(total)}</strong>
          <span>modules</span>
        </div>
      `;
      elements.historyList.appendChild(item);
    });
  }

  function buildMonthCards() {
    return statCardsFromStats(computeMonthStats(state.selectedMonth));
  }

  function buildLifetimeCards() {
    return statCardsFromStats(computeLifetimeStats());
  }

  function statCardsFromStats(stats) {
    return [
      {
        label: "モジュール合計",
        value: `${formatInteger(stats.totalModules)} (${formatInteger(stats.totalWithPieces)})`,
        note: `期待値 ${stats.expectedModules.toFixed(2)}`,
        featured: true,
      },
      {
        label: "経過日数",
        value: formatInteger(stats.elapsedDays),
        note: stats.periodLabel,
      },
      {
        label: "チャレンジ回数",
        value: formatInteger(stats.challengeCount),
        note: "",
      },
      {
        label: "落ちた回数",
        value: formatInteger(stats.droppedChallenges),
        note: `${percent(stats.actualDropRate)} / ${percent(DROP_RATE)}`,
      },
      {
        label: "1 / 2 / 3 Drop",
        value: `${formatInteger(stats.oneDrops)} / ${formatInteger(stats.twoDrops)} / ${formatInteger(stats.threeDrops)}`,
        note: `${percent(stats.oneDropRate)} / ${percent(stats.twoDropRate)} / ${percent(stats.threeDropRate)} (${percent(DROP_DISTRIBUTION[1])} / ${percent(DROP_DISTRIBUTION[2])} / ${percent(DROP_DISTRIBUTION[3])})`,
        spanAll: true,
      },
      {
        label: "モジュールピース",
        value: formatInteger(stats.totalPieces),
        note: `37 x ${formatInteger(stats.challengeCount)}`,
      },
    ];
  }

  function computeStats(records, month) {
    const challengeCount = records.reduce((sum, record) => sum + record.slots.length, 0);
    const flattened = records.flatMap((record) => record.slots);
    const droppedChallenges = flattened.filter((value) => value > 0).length;
    const totalModules = flattened.reduce((sum, value) => sum + value, 0);
    const oneDrops = flattened.filter((value) => value === 1).length;
    const twoDrops = flattened.filter((value) => value === 2).length;
    const threeDrops = flattened.filter((value) => value === 3).length;

    return {
      elapsedDays: computeElapsedDays(month),
      periodLabel: month ? formatMonthLabel(month) : "全期間",
      challengeCount,
      droppedChallenges,
      actualDropRate: challengeCount ? droppedChallenges / challengeCount : 0,
      totalModules,
      totalWithPieces: totalModules + Math.floor((challengeCount * PIECES_PER_CHALLENGE) / 100),
      expectedModules: challengeCount * DROP_RATE * expectedPerDrop(),
      expectedTotal: challengeCount * DROP_RATE * expectedPerDrop() + (challengeCount * PIECES_PER_CHALLENGE) / 100,
      oneDrops,
      twoDrops,
      threeDrops,
      oneDropRate: droppedChallenges ? oneDrops / droppedChallenges : 0,
      twoDropRate: droppedChallenges ? twoDrops / droppedChallenges : 0,
      threeDropRate: droppedChallenges ? threeDrops / droppedChallenges : 0,
      totalPieces: challengeCount * PIECES_PER_CHALLENGE,
    };
  }

  function computeMonthStats(month) {
    const archive = archiveForMonth(month);
    if (archive) {
      return statsFromAggregate(archive, month);
    }
    return computeStats(recordsForMonth(month), month);
  }

  function computeLifetimeStats() {
    return statsFromAggregate(combineAggregates([
      aggregateRecords(state.records),
      aggregateArchives(state.archives),
    ]), null);
  }

  function statsFromAggregate(aggregate, month) {
    const challengeCount = Number(aggregate.challengeCount || 0);
    const droppedChallenges = Number(aggregate.droppedChallenges || 0);
    const totalModules = Number(aggregate.totalModules || 0);
    const oneDrops = Number(aggregate.oneDrops || 0);
    const twoDrops = Number(aggregate.twoDrops || 0);
    const threeDrops = Number(aggregate.threeDrops || 0);
    const totalPieces = Number(aggregate.totalPieces || 0);
    const elapsedDays = month ? Number(aggregate.elapsedDays || computeElapsedDays(month)) : Number(aggregate.elapsedDays || state.records.length);

    return {
      elapsedDays,
      periodLabel: month ? formatMonthLabel(month) : "全期間",
      challengeCount,
      droppedChallenges,
      actualDropRate: challengeCount ? droppedChallenges / challengeCount : 0,
      totalModules,
      totalWithPieces: totalModules + Math.floor(totalPieces / 100),
      expectedModules: challengeCount * DROP_RATE * expectedPerDrop(),
      expectedTotal: challengeCount * DROP_RATE * expectedPerDrop() + totalPieces / 100,
      oneDrops,
      twoDrops,
      threeDrops,
      oneDropRate: droppedChallenges ? oneDrops / droppedChallenges : 0,
      twoDropRate: droppedChallenges ? twoDrops / droppedChallenges : 0,
      threeDropRate: droppedChallenges ? threeDrops / droppedChallenges : 0,
      totalPieces,
    };
  }

  function aggregateRecords(records) {
    const flattened = records.flatMap((record) => record.slots);
    return {
      elapsedDays: records.length,
      challengeCount: flattened.length,
      droppedChallenges: flattened.filter((value) => value > 0).length,
      totalModules: flattened.reduce((sum, value) => sum + value, 0),
      oneDrops: flattened.filter((value) => value === 1).length,
      twoDrops: flattened.filter((value) => value === 2).length,
      threeDrops: flattened.filter((value) => value === 3).length,
      totalPieces: flattened.length * PIECES_PER_CHALLENGE,
    };
  }

  function aggregateArchives(archives) {
    return combineAggregates(archives);
  }

  function combineAggregates(aggregates) {
    return aggregates.reduce((sum, item) => ({
      elapsedDays: Number(sum.elapsedDays || 0) + Number(item.elapsedDays || 0),
      challengeCount: Number(sum.challengeCount || 0) + Number(item.challengeCount || 0),
      droppedChallenges: Number(sum.droppedChallenges || 0) + Number(item.droppedChallenges || 0),
      totalModules: Number(sum.totalModules || 0) + Number(item.totalModules || 0),
      oneDrops: Number(sum.oneDrops || 0) + Number(item.oneDrops || 0),
      twoDrops: Number(sum.twoDrops || 0) + Number(item.twoDrops || 0),
      threeDrops: Number(sum.threeDrops || 0) + Number(item.threeDrops || 0),
      totalPieces: Number(sum.totalPieces || 0) + Number(item.totalPieces || 0),
    }), {
      elapsedDays: 0,
      challengeCount: 0,
      droppedChallenges: 0,
      totalModules: 0,
      oneDrops: 0,
      twoDrops: 0,
      threeDrops: 0,
      totalPieces: 0,
    });
  }

  function computeElapsedDays(month) {
    if (!month) {
      return state.records.length;
    }

    const normalizedMonth = normalizeMonthKey(month);
    if (!normalizedMonth) {
      return 0;
    }

    const year = Number(normalizedMonth.slice(0, 4));
    const monthIndex = Number(normalizedMonth.slice(4, 6));
    const today = gameDateNow();
    const lastDay = new Date(year, monthIndex, 0).getDate();

    if (today.getFullYear() === year && today.getMonth() + 1 === monthIndex) {
      return Math.max(0, today.getDate() - 1);
    }

    return lastDay;
  }

  function expectedPerDrop() {
    return DROP_DISTRIBUTION[1] + DROP_DISTRIBUTION[2] * 2 + DROP_DISTRIBUTION[3] * 3;
  }

  function describeSlot(value) {
    return value === 0 ? "未ドロップ" : `${value}個ドロップ`;
  }

  function updateSlot(index, value) {
    const record = getRecord(state.selectedDate) || createEmptyRecord(state.selectedDate);
    record.slots = normalizeSlots(record.slots, record.isFullBurst ? 6 : 3);
    record.slots[index] = value;
    upsertRecord(record);
    persist();
    render();
    scheduleAutoSync();
  }

  function syncEditorFromRecord() {
    const record = getRecord(state.selectedDate) || createEmptyRecord(state.selectedDate);
    elements.dateInput.value = state.selectedDate;
    elements.fullBurstToggle.checked = record.isFullBurst;
  }

  async function applyUser() {
    const nextUserId = sanitizeUserId(state.pendingUserId || elements.userIdSelect.value);
    if (!nextUserId) {
      window.alert("ユーザー名を選択してください。");
      return;
    }

    setUserSwitchStatus("読み込み中", "loading");
    state.userId = nextUserId;
    state.pendingUserId = nextUserId;
    localStorage.setItem(USER_KEY, nextUserId);
    rememberUserOption(nextUserId);
    const localData = loadLocalData(nextUserId);
    state.records = mergeRecords(seedRecords(), localData.records);
    state.archives = mergeArchives([], localData.archives);
    state.selectedDate = todayIso();
    state.selectedMonth = latestMonth();
    state.historyBaseMonth = state.selectedMonth;
    refreshSelectedDateSyncState();
    elements.saveStatus.textContent = `${nextUserId} で使用中`;
    syncEditorFromRecord();
    render();

    if (hasBackendConfig()) {
      const success = await syncDownload(true);
      setUserSwitchStatus(success ? "読み込み完了" : "読込失敗", success ? "success" : "error");
      return;
    }

    persist(true);
    setUserSwitchStatus("読み込み完了", "success");
  }

  async function exportRecords() {
    if (state.userId && hasBackendConfig()) {
      await syncDownload(true);
    }

    const payload = {
      version: 1,
      exportedAt: new Date().toISOString(),
      userId: state.userId,
      records: state.records,
      archives: state.archives,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `custom-module-backup-${state.userId || "user"}-${todayIso()}.json`;
    anchor.click();
    URL.revokeObjectURL(url);
    elements.saveStatus.textContent = hasBackendConfig() && state.userId
      ? "スプシから最新を読んでJSONを書き出しました"
      : "現在のデータをJSONで書き出しました";
  }

  async function importRecords(event) {
    const file = (event.target.files || [])[0];
    if (!file) {
      return;
    }

    try {
      const text = await file.text();
      const parsed = parseImportPayload(text, file.name);
      if (!parsed || !Array.isArray(parsed.records)) {
        throw new Error("invalid");
      }

      if (parsed.userId) {
        state.userId = sanitizeUserId(parsed.userId);
        localStorage.setItem(USER_KEY, state.userId);
        rememberUserOption(state.userId);
      }

      state.records = mergeRecords([], parsed.records);
      state.archives = mergeArchives([], parsed.archives || []);
      state.selectedDate = todayIso();
      state.selectedMonth = latestMonth();
      state.historyBaseMonth = state.selectedMonth;
      persist(true);
      syncEditorFromRecord();
      refreshSelectedDateSyncState();
      render();

      if (state.userId && hasBackendConfig()) {
        await syncUpload(true);
        markSelectedDateSynced();
        elements.saveStatus.textContent = "JSONを読み込み、スプシへ同期しました";
      } else {
        elements.saveStatus.textContent = "JSONを読み込みました";
        scheduleAutoSync();
      }
      return true;
    } catch (error) {
      window.alert(`JSONの読み込みに失敗しました: ${String(error)}`);
      return false;
    } finally {
      elements.importFileInput.value = "";
    }
  }

  async function syncUpload(isSilent) {
    if (!assertBackendReady()) {
      return;
    }

    try {
      setBusy(true);
      const payload = {
        token: BACKEND_CONFIG.apiToken,
        userId: state.userId,
        records: state.records,
        archives: state.archives,
      };

      const response = await fetch(BACKEND_CONFIG.appsScriptUrl, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload),
      });
      const result = await response.json();

      if (!response.ok || !result.ok) {
        throw new Error(result.error || "sync_failed");
      }

      markSelectedDateSynced();
      render();

      if (!isSilent) {
        elements.saveStatus.textContent = `この日付をスプシに保存しました (${state.userId})`;
      }
    } catch (error) {
      if (!isSilent) {
        window.alert(`スプシ保存に失敗しました: ${String(error)}`);
      }
      return false;
    } finally {
      setBusy(false);
    }
  }

  async function syncDownload(isSilent) {
    if (!assertBackendReady()) {
      return false;
    }

    try {
      setBusy(true);
      const url = new URL(BACKEND_CONFIG.appsScriptUrl);
      url.searchParams.set("token", BACKEND_CONFIG.apiToken);
      url.searchParams.set("userId", state.userId);

      const response = await fetch(url.toString());
      const result = await response.json();

      if (!response.ok || !result.ok || !Array.isArray(result.records)) {
        throw new Error(result.error || "load_failed");
      }

      state.records = mergeRecords([], result.records);
      state.archives = mergeArchives([], result.archives || []);
      state.selectedDate = todayIso();
      state.selectedMonth = latestMonth();
      state.historyBaseMonth = state.selectedMonth;
      persist(true);
      markSelectedDateSynced();
      refreshSelectedDateSyncState();
      syncEditorFromRecord();
      render();
      return true;

      if (!isSilent) {
        elements.saveStatus.textContent = `スプシをアプリへ反映しました (${state.userId})`;
      }
    } catch (error) {
      if (!isSilent) {
        window.alert(`スプシ読込に失敗しました: ${String(error)}`);
      }
    } finally {
      setBusy(false);
    }
  }

  function scheduleAutoSync() {
    if (!BACKEND_CONFIG.autoSync || !hasBackendConfig() || !state.userId) {
      return;
    }

    clearTimeout(state.autoSyncTimer);
    state.autoSyncTimer = setTimeout(() => {
      syncUpload(true);
    }, 800);
  }

  function assertBackendReady() {
    if (!hasBackendConfig()) {
      window.alert("backend-config.js に Apps Script の URL と token を設定してください。");
      return false;
    }

    if (!state.userId) {
      window.alert("ユーザー名を設定してください。");
      return false;
    }

    return true;
  }

  function hasBackendConfig() {
    return Boolean(BACKEND_CONFIG.appsScriptUrl && BACKEND_CONFIG.apiToken);
  }

  function setBusy(isBusy) {
    [
      elements.syncUploadButton,
      elements.exportButton,
      elements.importButton,
      elements.applyUserButton,
    ].forEach((button) => {
      button.disabled = isBusy;
    });
  }

  function updateCurrentProfileChip() {
    if (!elements.currentProfileChip) {
      return;
    }
    elements.currentProfileChip.textContent = `使用中: ${state.userId || "-"}`;
  }

  function updateUserSwitchUi() {
    const selectedUserId = sanitizeUserId(state.pendingUserId || (elements.userIdSelect ? elements.userIdSelect.value : ""));
    const isPending = Boolean(selectedUserId && selectedUserId !== state.userId);
    elements.applyUserButton.classList.toggle("attention", isPending);

    if (!selectedUserId) {
      setUserSwitchStatus("未選択", "error");
      return;
    }

    if (isPending) {
      setUserSwitchStatus("切替待ち", "pending");
      return;
    }

    if (!elements.userSwitchStatus.classList.contains("is-loading")
      && !elements.userSwitchStatus.classList.contains("is-success")
      && !elements.userSwitchStatus.classList.contains("is-error")) {
      setUserSwitchStatus(state.userId ? "使用中" : "未選択", state.userId ? "idle" : "error");
    }
  }

  function setUserSwitchStatus(text, tone) {
    if (!elements.userSwitchStatus) {
      return;
    }

    const normalizedText = normalizeSwitchStatusText(text, tone);
    const normalizedTone = normalizeSwitchStatusTone(text, tone);

    elements.userSwitchStatus.textContent = normalizedText;
    elements.userSwitchStatus.classList.remove("is-idle", "is-pending", "is-loading", "is-success", "is-error");
    if (normalizedTone === "idle") {
      elements.userSwitchStatus.classList.add("is-idle");
    } else if (normalizedTone === "pending") {
      elements.userSwitchStatus.classList.add("is-pending");
    } else if (normalizedTone === "loading") {
      elements.userSwitchStatus.classList.add("is-loading");
    } else if (normalizedTone === "success") {
      elements.userSwitchStatus.classList.add("is-success");
    } else if (normalizedTone === "error") {
      elements.userSwitchStatus.classList.add("is-error");
    }
  }

  function normalizeSwitchStatusText(text, tone) {
    if (tone === "pending") {
      return "切替待ち";
    }
    if (tone === "loading") {
      return "読込中";
    }
    if (tone === "success") {
      return "反映済み";
    }
    if (tone === "error") {
      return text && text !== "idle" ? text : "未選択";
    }
    if (tone === "idle") {
      return "使用中";
    }

    const source = String(text || "");
    if (source.includes("待")) {
      return "切替待ち";
    }
    if (source.includes("読") || source.includes("込")) {
      return "読込中";
    }
    if (source.includes("完") || source.includes("反映")) {
      return "反映済み";
    }
    if (source.includes("失")) {
      return "読込失敗";
    }
    return state.userId ? "使用中" : "未選択";
  }

  function normalizeSwitchStatusTone(text, tone) {
    if (tone) {
      return tone;
    }

    const source = String(text || "");
    if (source.includes("失")) {
      return "error";
    }
    if (source.includes("完") || source.includes("反映")) {
      return "success";
    }
    if (source.includes("読") || source.includes("込")) {
      return "loading";
    }
    if (source.includes("待")) {
      return "pending";
    }
    return state.userId ? "idle" : "error";
  }

  function persist(skipStatusUpdate) {
    const cache = loadStorageCache();
    const cacheKey = currentCacheKey();
    cache[cacheKey] = {
      records: state.records,
      archives: state.archives,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cache));
    refreshSelectedDateSyncState();
    if (!skipStatusUpdate) {
      elements.saveStatus.textContent = `ローカル保存済み (${new Date().toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit" })})`;
    }
  }

  function loadLocalData(userId = state.userId) {
    try {
      const cache = loadStorageCache();
      const cacheKey = sanitizeUserId(userId) || "__default__";
      const entry = cache[cacheKey];
      if (entry && typeof entry === "object" && !Array.isArray(entry)) {
        return {
          records: Array.isArray(entry.records) ? entry.records : [],
          archives: Array.isArray(entry.archives) ? entry.archives : [],
        };
      }
      if (Array.isArray(entry)) {
        return { records: entry, archives: [] };
      }

      const legacyRaw = localStorage.getItem(LEGACY_STORAGE_KEY);
      return {
        records: legacyRaw ? JSON.parse(legacyRaw) : [],
        archives: [],
      };
    } catch {
      return { records: [], archives: [] };
    }
  }

  function loadUserId() {
    return sanitizeUserId(localStorage.getItem(USER_KEY) || BACKEND_CONFIG.defaultUserId || "");
  }

  function loadUserOptions() {
    const configured = Array.isArray(BACKEND_CONFIG.userOptions) ? BACKEND_CONFIG.userOptions : [];
    const remembered = loadRememberedUserOptions();
    const merged = new Set([...configured, ...remembered].map(sanitizeUserId).filter(Boolean));
    return Array.from(merged);
  }

  function loadRememberedUserOptions() {
    try {
      const raw = localStorage.getItem(USER_OPTIONS_KEY);
      const parsed = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  function rememberUserOption(userId) {
    const normalized = sanitizeUserId(userId);
    if (!normalized) {
      return;
    }
    if (!state.userOptions.includes(normalized)) {
      state.userOptions = [...state.userOptions, normalized].sort((left, right) => left.localeCompare(right, "ja"));
      localStorage.setItem(USER_OPTIONS_KEY, JSON.stringify(state.userOptions));
    }
  }

  function buildUserOptions() {
    const options = new Set([...state.userOptions].map(sanitizeUserId).filter(Boolean));
    return Array.from(options).sort((left, right) => left.localeCompare(right, "ja"));
  }

  function parseImportPayload(text, filename) {
    const trimmed = String(text || "").trim();
    if (!trimmed) {
      throw new Error("empty");
    }

    const lowerName = String(filename || "").toLowerCase();
    if (lowerName.endsWith(".csv")) {
      return { version: 1, records: parseCsvRecords(trimmed), archives: [] };
    }

    const parsed = JSON.parse(trimmed);
    if (parsed && Array.isArray(parsed.records)) {
      return parsed;
    }

    throw new Error("invalid");
  }

  function parseCsvRecords(csvText) {
    const rows = parseCsv(csvText);
    if (rows.length < 2) {
      return [];
    }

    const headers = rows[0].map(normalizeCsvHeader);
    return rows.slice(1)
      .map((row) => parseCsvRecordRow(headers, row))
      .filter(Boolean);
  }

  function parseCsv(csvText) {
    const rows = [];
    let row = [];
    let value = "";
    let inQuotes = false;

    for (let index = 0; index < csvText.length; index += 1) {
      const char = csvText[index];
      const next = csvText[index + 1];

      if (char === "\"") {
        if (inQuotes && next === "\"") {
          value += "\"";
          index += 1;
        } else {
          inQuotes = !inQuotes;
        }
        continue;
      }

      if (char === "," && !inQuotes) {
        row.push(value);
        value = "";
        continue;
      }

      if ((char === "\n" || char === "\r") && !inQuotes) {
        if (char === "\r" && next === "\n") {
          index += 1;
        }
        row.push(value);
        rows.push(row);
        row = [];
        value = "";
        continue;
      }

      value += char;
    }

    if (value !== "" || row.length) {
      row.push(value);
      rows.push(row);
    }

    return rows
      .map((recordRow) => recordRow.map((cell) => String(cell || "").trim()))
      .filter((recordRow) => recordRow.some(Boolean));
  }

  function parseCsvRecordRow(headers, row) {
    const get = (...keys) => {
      for (const key of keys) {
        const index = headers.indexOf(key);
        if (index >= 0) {
          return row[index];
        }
      }
      return "";
    };

    const date = normalizeIsoDate(
      get("date", "day", "日付", "日時", "challenge_date"),
    );
    if (!date) {
      return null;
    }

    const fullBurstRaw = get("isfullburst", "fullburst", "fullburstday", "フルバーストデイ", "フルバースト");
    const slotCandidates = [
      parseCsvNumber(get("slot1", "slot_1", "通常枠1", "枠1")),
      parseCsvNumber(get("slot2", "slot_2", "通常枠2", "枠2")),
      parseCsvNumber(get("slot3", "slot_3", "通常枠3", "枠3")),
      parseCsvNumber(get("slot4", "slot_4", "追加枠1", "枠4")),
      parseCsvNumber(get("slot5", "slot_5", "追加枠2", "枠5")),
      parseCsvNumber(get("slot6", "slot_6", "追加枠3", "枠6")),
    ];
    const hasExtraSlots = slotCandidates.slice(3).some((value) => value > 0)
      || isTruthyCsvValue(fullBurstRaw);

    return {
      date,
      isFullBurst: hasExtraSlots,
      slots: normalizeSlots(slotCandidates, hasExtraSlots ? 6 : 3),
    };
  }

  function normalizeCsvHeader(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "")
      .replace(/[()（）]/g, "")
      .replace(/[・._-]/g, "");
  }

  function parseCsvNumber(value) {
    const text = String(value || "").trim().replace(/,/g, "");
    const number = Number(text);
    return Number.isFinite(number) && number >= 0 ? number : 0;
  }

  function isTruthyCsvValue(value) {
    const text = String(value || "").trim().toLowerCase();
    return ["1", "true", "yes", "y", "on", "full", "fb", "はい"].includes(text);
  }

  function sanitizeUserId(value) {
    return String(value || "").trim();
  }

  function seedRecords() {
    if (Array.isArray(INITIAL_DATA.records)) {
      return INITIAL_DATA.records;
    }
    if (Array.isArray(INITIAL_DATA.months)) {
      return INITIAL_DATA.months.flatMap((month) => Array.isArray(month.records) ? month.records : []);
    }
    return [];
  }

  function mergeRecords(initialRecords, localRecords) {
    const merged = new Map();

    [...initialRecords, ...localRecords].forEach((record) => {
      const normalized = normalizeRecord(record);
      if (!normalized) {
        return;
      }
      merged.set(normalized.date, normalized);
    });

    return Array.from(merged.values()).sort((left, right) => left.date.localeCompare(right.date));
  }

  function mergeArchives(initialArchives, localArchives) {
    const merged = new Map();

    [...initialArchives, ...localArchives].forEach((archive) => {
      const normalized = normalizeArchive(archive);
      if (!normalized) {
        return;
      }
      merged.set(normalized.month, normalized);
    });

    return Array.from(merged.values()).sort((left, right) => left.month.localeCompare(right.month));
  }

  function normalizeRecord(record) {
    if (!record || typeof record !== "object") {
      return null;
    }

    const date = normalizeIsoDate(record.date);
    if (!date) {
      return null;
    }

    const inputSlots = Array.isArray(record.slots) ? record.slots : [];
    const slots = Array.from({ length: 6 }, (_, index) => Number(inputSlots[index] || 0));
    const explicitFlag = record.isFullBurst === true || record.isFullBurst === 1 || record.isFullBurst === "1";
    const inferredFlag = slots.slice(3).some((value) => Number(value || 0) > 0);
    const isFullBurst = explicitFlag || inferredFlag;

    return {
      date,
      isFullBurst,
      slots: normalizeSlots(slots, isFullBurst ? 6 : 3),
    };
  }

  function normalizeArchive(archive) {
    if (!archive || typeof archive !== "object") {
      return null;
    }
    const month = normalizeMonthKey(archive.month);
    if (!month) {
      return null;
    }
    return {
      month,
      elapsedDays: Number(archive.elapsedDays || 0),
      challengeCount: Number(archive.challengeCount || 0),
      droppedChallenges: Number(archive.droppedChallenges || 0),
      totalModules: Number(archive.totalModules || 0),
      oneDrops: Number(archive.oneDrops || 0),
      twoDrops: Number(archive.twoDrops || 0),
      threeDrops: Number(archive.threeDrops || 0),
      totalPieces: Number(archive.totalPieces || 0),
    };
  }

  function upsertRecord(record) {
    const normalized = normalizeRecord(record);
    if (!normalized) {
      return;
    }

    const index = state.records.findIndex((item) => item.date === normalized.date);
    if (index >= 0) {
      state.records.splice(index, 1, normalized);
    } else {
      state.records.push(normalized);
    }

    state.records.sort((left, right) => left.date.localeCompare(right.date));
    state.selectedMonth = monthKey(normalized.date);
  }

  function deleteRecord(isoDate) {
    const normalized = normalizeIsoDate(isoDate);
    state.records = state.records.filter((record) => record.date !== normalized);
    state.selectedMonth = latestMonth();
  }

  function getRecord(isoDate) {
    return state.records.find((record) => record.date === isoDate);
  }

  function createEmptyRecord(isoDate) {
    return { date: normalizeIsoDate(isoDate) || todayIso(), isFullBurst: false, slots: [0, 0, 0] };
  }

  function recordsForMonth(month) {
    const normalizedMonth = normalizeMonthKey(month);
    return state.records.filter((record) => monthKey(record.date) === normalizedMonth);
  }

  function archiveForMonth(month) {
    const normalizedMonth = normalizeMonthKey(month);
    return state.archives.find((archive) => archive.month === normalizedMonth) || null;
  }

  function availableMonths() {
    const months = new Set();
    state.records.forEach((record) => {
      const key = monthKey(record.date);
      if (key) {
        months.add(key);
      }
    });
    state.archives.forEach((archive) => {
      const key = normalizeMonthKey(archive.month);
      if (key) {
        months.add(key);
      }
    });
    months.add(monthKey(todayIso()));
    return Array.from(months).filter(Boolean).sort().reverse();
  }

  function latestMonth() {
    const months = availableMonths();
    return months[0] || monthKey(todayIso());
  }

  function loadStorageCache() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : {};
      return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
    } catch {
      return {};
    }
  }

  function currentCacheKey() {
    return state.userId || "__default__";
  }

  function refreshSelectedDateSyncState() {
    state.needsSelectedDateSync = hasUnsyncedSelectedDateRecord();
  }

  function hasUnsyncedSelectedDateRecord() {
    if (!state.userId) {
      return false;
    }
    const savedHash = loadDaySyncMeta()[selectedDateMetaKey()];
    const currentHash = selectedDateRecordHash(state.selectedDate);
    return currentHash !== savedHash;
  }

  function markSelectedDateSynced() {
    if (!state.userId) {
      return;
    }
    const meta = loadDaySyncMeta();
    meta[selectedDateMetaKey()] = selectedDateRecordHash(state.selectedDate);
    localStorage.setItem(DAY_SYNC_META_KEY, JSON.stringify(meta));
    refreshSelectedDateSyncState();
  }

  function loadDaySyncMeta() {
    try {
      const raw = localStorage.getItem(DAY_SYNC_META_KEY);
      const parsed = raw ? JSON.parse(raw) : {};
      return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
    } catch {
      return {};
    }
  }

  function selectedDateMetaKey() {
    return `${state.userId || "__default__"}:${normalizeIsoDate(state.selectedDate) || todayIso()}`;
  }

  function selectedDateRecordHash(isoDate) {
    const normalizedDate = normalizeIsoDate(isoDate) || todayIso();
    const record = getRecord(normalizedDate);
    if (!record) {
      return "none";
    }
    return JSON.stringify(normalizeRecord(record));
  }

  function normalizeSlots(slots, desiredLength) {
    return Array.from({ length: desiredLength }, (_, index) => {
      const value = Number(Array.isArray(slots) ? slots[index] : 0);
      return Number.isFinite(value) && value >= 0 ? value : 0;
    });
  }

  function normalizeIsoDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return formatIsoDate(value);
    }

    const text = String(value || "").trim();
    if (!text) {
      return "";
    }

    const directMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (directMatch) {
      const year = Number(directMatch[1]);
      const month = Number(directMatch[2]);
      const day = Number(directMatch[3]);
      if (isValidDateParts(year, month, day)) {
        return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      }
    }

    const parsed = new Date(text);
    if (Number.isNaN(parsed.getTime())) {
      return "";
    }
    return formatIsoDate(parsed);
  }

  function normalizeMonthKey(value) {
    const text = String(value || "").trim();
    if (/^\d{6}$/.test(text)) {
      return text;
    }

    const normalizedDate = normalizeIsoDate(text);
    if (!normalizedDate) {
      return "";
    }
    return monthKey(normalizedDate);
  }

  function isValidDateParts(year, month, day) {
    const date = new Date(year, month - 1, day);
    return (
      date.getFullYear() === year &&
      date.getMonth() === month - 1 &&
      date.getDate() === day
    );
  }

  function monthKey(isoDate) {
    const normalized = normalizeIsoDate(isoDate);
    return normalized ? normalized.slice(0, 7).replace("-", "") : "";
  }

  function todayIso() {
    return formatIsoDate(gameDateNow());
  }

  function startDayWatcher() {
    let lastGameDate = todayIso();
    clearInterval(state.dayWatcherTimer);
    state.dayWatcherTimer = setInterval(() => {
      const nextGameDate = todayIso();
      if (nextGameDate === lastGameDate) {
        return;
      }

      const previousGameDate = lastGameDate;
      const previousSelectedDate = state.selectedDate;
      lastGameDate = nextGameDate;

      if (previousSelectedDate === previousGameDate) {
        state.selectedDate = nextGameDate;
      }

      if (state.userId && hasBackendConfig()) {
        syncDownload(true);
      } else {
        refreshSelectedDateSyncState();
        syncEditorFromRecord();
        render();
      }
    }, 60 * 1000);
  }

  function gameDateNow() {
    return new Date(Date.now() - 5 * 60 * 60 * 1000);
  }

  function formatIsoDate(date) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
  }

  function formatDateLabel(isoDate) {
    const normalized = normalizeIsoDate(isoDate);
    if (!normalized) {
      return "日付不明";
    }

    return new Date(`${normalized}T00:00:00`).toLocaleDateString("ja-JP", {
      year: "numeric",
      month: "long",
      day: "numeric",
      weekday: "short",
    });
  }

  function formatMonthLabel(month) {
    const normalized = normalizeMonthKey(month);
    if (!normalized) {
      return "不明";
    }
    return `${normalized.slice(0, 4)}-${normalized.slice(4, 6)}`;
  }

  function percent(value) {
    return `${(value * 100).toFixed(2)}%`;
  }

  function formatInteger(value) {
    return Number(value || 0).toLocaleString("en-US");
  }
})();
