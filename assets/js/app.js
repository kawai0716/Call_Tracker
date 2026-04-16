const STORAGE_KEY = "personal-call-tracker-v2";
const RECORDING_DB_NAME = "personal-call-tracker-audio-v1";
const RECORDING_STORE_NAME = "recordings";
const MAX_RECORDING_HISTORY = 5;
const PAGE_KEYS = ["dashboard", "history", "recording"];

const elements = {
  todayLabel: document.getElementById("todayLabel"),
  updatedAtLabel: document.getElementById("updatedAtLabel"),
  storageStatusLabel: document.getElementById("storageStatusLabel"),
  heroTodayPanel: document.getElementById("heroTodayPanel"),
  exportButton: document.getElementById("exportButton"),
  advanceBusinessDayButton: document.getElementById("advanceBusinessDayButton"),
  resetTodayButton: document.getElementById("resetTodayButton"),
  todayCallsValue: document.getElementById("todayCallsValue"),
  todayConnectionsValue: document.getElementById("todayConnectionsValue"),
  todaySampleSentValue: document.getElementById("todaySampleSentValue"),
  todayIntroductionsValue: document.getElementById("todayIntroductionsValue"),
  todayConnectionRateValue: document.getElementById("todayConnectionRateValue"),
  todayPaceText: document.getElementById("todayPaceText"),
  todaySampleRateValue: document.getElementById("todaySampleRateValue"),
  todaySampleRateText: document.getElementById("todaySampleRateText"),
  callsValue: document.getElementById("callsValue"),
  connectionsValue: document.getElementById("connectionsValue"),
  sampleSentValue: document.getElementById("sampleSentValue"),
  introductionsValue: document.getElementById("introductionsValue"),
  connectionTotalRateValue: document.getElementById("connectionTotalRateValue"),
  paceText: document.getElementById("paceText"),
  sampleTotalRateValue: document.getElementById("sampleTotalRateValue"),
  sampleRateText: document.getElementById("sampleRateText"),
  historyMonthConnectionRateValue: document.getElementById("historyMonthConnectionRateValue"),
  historyMonthSampleRateValue: document.getElementById("historyMonthSampleRateValue"),
  monthLabel: document.getElementById("monthLabel"),
  goalCallsInput: document.getElementById("goalCallsInput"),
  goalSampleSentInput: document.getElementById("goalSampleSentInput"),
  goalIntroductionsInput: document.getElementById("goalIntroductionsInput"),
  goalConnectionsInput: document.getElementById("goalConnectionsInput"),
  saveGoalsButton: document.getElementById("saveGoalsButton"),
  goalsStatusText: document.getElementById("goalsStatusText"),
  locationSelect: document.getElementById("locationSelect"),
  newLocationInput: document.getElementById("newLocationInput"),
  addLocationButton: document.getElementById("addLocationButton"),
  workedHoursInput: document.getElementById("workedHoursInput"),
  reflectionInput: document.getElementById("reflectionInput"),
  saveMemoButton: document.getElementById("saveMemoButton"),
  copyTemplateButton: document.getElementById("copyTemplateButton"),
  templatePreview: document.getElementById("templatePreview"),
  recordingStatusText: document.getElementById("recordingStatusText"),
  startRecordingButton: document.getElementById("startRecordingButton"),
  stopRecordingButton: document.getElementById("stopRecordingButton"),
  recordingHistory: document.getElementById("recordingHistory"),
  pageTabs: Array.from(document.querySelectorAll("[data-page-tab]")),
  pagePanels: Array.from(document.querySelectorAll("[data-page-panel]")),
  historyMonthSelect: document.getElementById("historyMonthSelect"),
  historyBody: document.getElementById("historyBody"),
  toast: document.getElementById("toast"),
};

const state = loadState();
const recordingState = {
  isSupported: Boolean(window.indexedDB && navigator.mediaDevices?.getUserMedia && window.MediaRecorder),
  isRecording: false,
  dbPromise: null,
  mediaRecorder: null,
  mediaStream: null,
  chunks: [],
  recordings: [],
  objectUrls: new Map(),
};
let toastTimer = null;

function createDefaultState() {
  return {
    records: {},
    dailyGoals: {},
    locationOptions: ["自宅"],
    selectedHistoryMonth: "",
    activePage: "dashboard",
    businessDayOverride: null,
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);

    if (!raw) {
      return createDefaultState();
    }

    const parsed = JSON.parse(raw);
    return sanitizeState(parsed);
  } catch (error) {
    console.error("Failed to load local data:", error);
    return createDefaultState();
  }
}

function sanitizeState(candidate) {
  const base = createDefaultState();

  if (!candidate || typeof candidate !== "object") {
    return base;
  }

  if (candidate.records && typeof candidate.records === "object") {
    for (const [dateKey, record] of Object.entries(candidate.records)) {
      const sanitizedRecord = sanitizeRecord(record);

      if (isMeaningfulRecord(sanitizedRecord)) {
        base.records[dateKey] = sanitizedRecord;
      }
    }
  }

  if (candidate.dailyGoals && typeof candidate.dailyGoals === "object") {
    for (const [dateKey, goals] of Object.entries(candidate.dailyGoals)) {
      base.dailyGoals[dateKey] = sanitizeGoals(goals);
    }
  }

  if (candidate.monthlyGoals && typeof candidate.monthlyGoals === "object") {
    const todayKey = getTodayKey();
    const currentMonthKey = getMonthKey(todayKey);
    const legacyGoals = candidate.monthlyGoals[currentMonthKey];

    if (legacyGoals && !base.dailyGoals[todayKey]) {
      base.dailyGoals[todayKey] = sanitizeGoals(legacyGoals);
    }
  }

  if (Array.isArray(candidate.locationOptions)) {
    const options = candidate.locationOptions
      .map((option) => (typeof option === "string" ? option.trim().slice(0, 40) : ""))
      .filter(Boolean);
    base.locationOptions = Array.from(new Set(["自宅", ...options.filter((option) => option !== "TIB")]));
  }

  base.selectedHistoryMonth =
    typeof candidate.selectedHistoryMonth === "string" ? candidate.selectedHistoryMonth : "";
  base.activePage =
    typeof candidate.activePage === "string" && PAGE_KEYS.includes(candidate.activePage)
      ? candidate.activePage
      : "dashboard";

  if (
    candidate.businessDayOverride &&
    typeof candidate.businessDayOverride === "object" &&
    typeof candidate.businessDayOverride.dateKey === "string" &&
    typeof candidate.businessDayOverride.expiresAt === "string"
  ) {
    base.businessDayOverride = {
      dateKey: candidate.businessDayOverride.dateKey,
      expiresAt: candidate.businessDayOverride.expiresAt,
    };
  }

  return base;
}

function sanitizeRecord(record) {
  const introductions = Math.max(0, Number.parseInt(record?.introductions, 10) || 0);
  const sampleSent = Math.max(0, Number.parseInt(record?.sampleSent ?? record?.samples, 10) || 0);
  const connections = Math.max(0, Number.parseInt(record?.connections, 10) || 0);
  const calls = Math.max(0, Number.parseInt(record?.calls, 10) || 0);

  return {
    calls,
    connections,
    sampleSent,
    introductions,
    location: typeof record?.location === "string" ? record.location.trim().slice(0, 40) : "",
    workedHours:
      typeof record?.workedHours === "string" ? record.workedHours.trim().slice(0, 16) : "",
    reflection:
      typeof record?.reflection === "string" ? record.reflection.trim().slice(0, 4000) : "",
    updatedAt: typeof record?.updatedAt === "string" ? record.updatedAt : null,
  };
}

function isMeaningfulRecord(record) {
  return Boolean(
    record.updatedAt ||
      record.calls ||
      record.connections ||
      record.sampleSent ||
      record.introductions ||
      record.location ||
      record.workedHours ||
      record.reflection,
  );
}

function sanitizeGoals(goals) {
  return {
    calls: toIntOrNull(goals?.calls),
    sampleSent: toIntOrNull(goals?.sampleSent),
    introductions: toIntOrNull(goals?.introductions),
    connections: toIntOrNull(goals?.connections),
  };
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function renderActivePage() {
  const activePage = PAGE_KEYS.includes(state.activePage) ? state.activePage : "dashboard";
  elements.heroTodayPanel.hidden = activePage !== "dashboard";
  elements.heroTodayPanel.style.display = activePage === "dashboard" ? "" : "none";

  elements.pageTabs.forEach((button) => {
    const isActive = button.dataset.pageTab === activePage;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-current", isActive ? "page" : "false");
  });

  elements.pagePanels.forEach((panel) => {
    panel.hidden = panel.dataset.pagePanel !== activePage;
  });
}

function createRecordingId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }

  return `rec-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getPreferredRecordingMimeType() {
  if (!window.MediaRecorder?.isTypeSupported) {
    return "";
  }

  const candidates = [
    "audio/webm;codecs=opus",
    "audio/webm",
    "audio/mp4",
    "audio/ogg;codecs=opus",
  ];

  return candidates.find((mimeType) => window.MediaRecorder.isTypeSupported(mimeType)) || "";
}

function openRecordingDb() {
  if (!recordingState.isSupported) {
    return Promise.reject(new Error("Recording not supported"));
  }

  if (recordingState.dbPromise) {
    return recordingState.dbPromise;
  }

  recordingState.dbPromise = new Promise((resolve, reject) => {
    const request = window.indexedDB.open(RECORDING_DB_NAME, 1);

    request.onupgradeneeded = () => {
      const db = request.result;

      if (!db.objectStoreNames.contains(RECORDING_STORE_NAME)) {
        db.createObjectStore(RECORDING_STORE_NAME, { keyPath: "id" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Failed to open recording db"));
  });

  return recordingState.dbPromise;
}

function getAllRecordingsFromDb() {
  return openRecordingDb().then(
    (db) =>
      new Promise((resolve, reject) => {
        const transaction = db.transaction(RECORDING_STORE_NAME, "readonly");
        const store = transaction.objectStore(RECORDING_STORE_NAME);
        const request = store.getAll();

        request.onsuccess = () => resolve(Array.isArray(request.result) ? request.result : []);
        request.onerror = () => reject(request.error || new Error("Failed to load recordings"));
      }),
  );
}

function saveRecordingToDb(entry) {
  return openRecordingDb().then(
    (db) =>
      new Promise((resolve, reject) => {
        const transaction = db.transaction(RECORDING_STORE_NAME, "readwrite");
        transaction.objectStore(RECORDING_STORE_NAME).put(entry);
        transaction.oncomplete = () => resolve();
        transaction.onerror = () => reject(transaction.error || new Error("Failed to save recording"));
        transaction.onabort = () => reject(transaction.error || new Error("Failed to save recording"));
      }),
  );
}

function deleteRecordingsFromDb(ids) {
  if (!ids.length) {
    return Promise.resolve();
  }

  return openRecordingDb().then(
    (db) =>
      new Promise((resolve, reject) => {
        const transaction = db.transaction(RECORDING_STORE_NAME, "readwrite");
        const store = transaction.objectStore(RECORDING_STORE_NAME);

        ids.forEach((id) => store.delete(id));

        transaction.oncomplete = () => resolve();
        transaction.onerror = () => reject(transaction.error || new Error("Failed to delete recording"));
        transaction.onabort = () => reject(transaction.error || new Error("Failed to delete recording"));
      }),
  );
}

function revokeRecordingObjectUrls() {
  for (const url of recordingState.objectUrls.values()) {
    window.URL.revokeObjectURL(url);
  }

  recordingState.objectUrls.clear();
}

function getRecordingObjectUrl(entry) {
  if (!recordingState.objectUrls.has(entry.id)) {
    recordingState.objectUrls.set(entry.id, window.URL.createObjectURL(entry.blob));
  }

  return recordingState.objectUrls.get(entry.id);
}

function formatRecordingDateTime(isoString) {
  return new Date(isoString).toLocaleString("ja-JP", {
    month: "numeric",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatRecordingSize(bytes) {
  if (!bytes) {
    return "0 KB";
  }

  const kiloBytes = bytes / 1024;

  if (kiloBytes < 1024) {
    return `${Math.max(1, Math.round(kiloBytes))} KB`;
  }

  return `${(kiloBytes / 1024).toFixed(1)} MB`;
}

function getRecordingFileExtension(mimeType) {
  if (mimeType?.includes("mp4")) {
    return "mp4";
  }

  if (mimeType?.includes("ogg")) {
    return "ogg";
  }

  return "webm";
}

function getRecordingDownloadName(entry) {
  const stamp = new Date(entry.createdAt);
  const year = String(stamp.getFullYear());
  const month = String(stamp.getMonth() + 1).padStart(2, "0");
  const day = String(stamp.getDate()).padStart(2, "0");
  const hour = String(stamp.getHours()).padStart(2, "0");
  const minute = String(stamp.getMinutes()).padStart(2, "0");
  return `voice-${year}${month}${day}-${hour}${minute}.${getRecordingFileExtension(entry.mimeType)}`;
}

function releaseRecordingStream() {
  if (recordingState.mediaStream) {
    recordingState.mediaStream.getTracks().forEach((track) => track.stop());
  }

  recordingState.mediaStream = null;
}

function renderRecordingStatus(text, isLive = false) {
  elements.recordingStatusText.textContent = text;
  elements.recordingStatusText.classList.toggle("is-live", isLive);
}

function renderRecordingControls() {
  if (!recordingState.isSupported) {
    elements.startRecordingButton.disabled = true;
    elements.stopRecordingButton.disabled = true;
    renderRecordingStatus("このブラウザでは録音に対応していません");
    return;
  }

  elements.startRecordingButton.disabled = recordingState.isRecording;
  elements.stopRecordingButton.disabled = !recordingState.isRecording;

  if (recordingState.isRecording) {
    renderRecordingStatus("録音中です。1件分が終わったら停止してください。", true);
    return;
  }

  renderRecordingStatus("マイク録音は停止中です");
}

function createRecordingHistoryItem(entry) {
  const item = document.createElement("article");
  item.className = "recording-item";

  const head = document.createElement("div");
  head.className = "recording-item-head";

  const info = document.createElement("div");
  const title = document.createElement("p");
  title.className = "recording-title";
  title.textContent = `${formatShortDate(entry.dateKey)} ${formatTime(entry.createdAt)} の録音`;
  const meta = document.createElement("p");
  meta.className = "recording-meta";
  meta.textContent = `${formatRecordingDateTime(entry.createdAt)} / ${formatRecordingSize(entry.size)}`;

  info.append(title, meta);

  const actions = document.createElement("div");
  actions.className = "recording-item-actions";

  const downloadButton = document.createElement("button");
  downloadButton.className = "secondary-button";
  downloadButton.type = "button";
  downloadButton.dataset.recordingAction = "download";
  downloadButton.dataset.recordingId = entry.id;
  downloadButton.textContent = "ダウンロード";

  const deleteButton = document.createElement("button");
  deleteButton.className = "danger-button";
  deleteButton.type = "button";
  deleteButton.dataset.recordingAction = "delete";
  deleteButton.dataset.recordingId = entry.id;
  deleteButton.textContent = "削除";

  actions.append(downloadButton, deleteButton);
  head.append(info, actions);

  const audio = document.createElement("audio");
  audio.className = "recording-audio";
  audio.controls = true;
  audio.preload = "metadata";
  audio.src = getRecordingObjectUrl(entry);

  item.append(head, audio);
  return item;
}

function renderRecordingHistory() {
  renderRecordingControls();

  revokeRecordingObjectUrls();
  elements.recordingHistory.innerHTML = "";

  if (!recordingState.isSupported) {
    const empty = document.createElement("p");
    empty.className = "empty-recording";
    empty.textContent = "録音は PC の Chrome / Edge で使うのが安定です。";
    elements.recordingHistory.append(empty);
    return;
  }

  if (!recordingState.recordings.length) {
    const empty = document.createElement("p");
    empty.className = "empty-recording";
    empty.textContent = "録音履歴はまだありません";
    elements.recordingHistory.append(empty);
    return;
  }

  recordingState.recordings.forEach((entry) => {
    elements.recordingHistory.append(createRecordingHistoryItem(entry));
  });
}

async function refreshRecordingHistory() {
  if (!recordingState.isSupported) {
    renderRecordingHistory();
    return;
  }

  try {
    const entries = await getAllRecordingsFromDb();
    const sorted = entries.sort((left, right) => right.createdAt.localeCompare(left.createdAt));
    const overflow = sorted.slice(MAX_RECORDING_HISTORY);

    if (overflow.length) {
      await deleteRecordingsFromDb(overflow.map((entry) => entry.id));
    }

    recordingState.recordings = sorted.slice(0, MAX_RECORDING_HISTORY);
    renderRecordingHistory();
  } catch (error) {
    console.error("Failed to refresh recordings:", error);
    renderRecordingStatus("録音履歴を読み込めませんでした");
    elements.startRecordingButton.disabled = true;
    elements.stopRecordingButton.disabled = true;
  }
}

async function saveRecordingBlob(blob, mimeType) {
  const entry = {
    id: createRecordingId(),
    createdAt: new Date().toISOString(),
    dateKey: getTodayKey(),
    mimeType: mimeType || blob.type || "audio/webm",
    size: blob.size,
    blob,
  };

  await saveRecordingToDb(entry);
  await refreshRecordingHistory();
}

async function startRecording() {
  if (!recordingState.isSupported || recordingState.isRecording) {
    return;
  }

  try {
    const mediaStream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const mimeType = getPreferredRecordingMimeType();
    const mediaRecorder = mimeType
      ? new window.MediaRecorder(mediaStream, { mimeType })
      : new window.MediaRecorder(mediaStream);

    recordingState.mediaStream = mediaStream;
    recordingState.mediaRecorder = mediaRecorder;
    recordingState.chunks = [];
    recordingState.isRecording = true;

    mediaRecorder.ondataavailable = (event) => {
      if (event.data?.size) {
        recordingState.chunks.push(event.data);
      }
    };

    mediaRecorder.onerror = (event) => {
      console.error("Recording failed:", event.error);
      releaseRecordingStream();
      recordingState.mediaRecorder = null;
      recordingState.chunks = [];
      recordingState.isRecording = false;
      renderRecordingHistory();
      showToast("録音中にエラーが発生しました", true);
    };

    mediaRecorder.onstop = async () => {
      const chunks = recordingState.chunks.slice();
      const currentMimeType = mediaRecorder.mimeType || mimeType || "audio/webm";

      releaseRecordingStream();
      recordingState.mediaRecorder = null;
      recordingState.chunks = [];
      recordingState.isRecording = false;

      if (!chunks.length) {
        renderRecordingHistory();
        showToast("録音データがありませんでした", true);
        return;
      }

      try {
        const blob = new Blob(chunks, { type: currentMimeType });
        await saveRecordingBlob(blob, currentMimeType);
        showToast("録音を保存しました");
      } catch (error) {
        console.error("Failed to save recording:", error);
        renderRecordingHistory();
        showToast("録音を保存できませんでした", true);
      }
    };

    mediaRecorder.start();
    renderRecordingHistory();
    showToast("録音を開始しました");
  } catch (error) {
    console.error("Failed to start recording:", error);
    releaseRecordingStream();
    recordingState.mediaRecorder = null;
    recordingState.chunks = [];
    recordingState.isRecording = false;
    renderRecordingHistory();
    showToast("マイク権限を許可してから再度お試しください", true);
  }
}

function stopRecording() {
  if (!recordingState.mediaRecorder || !recordingState.isRecording) {
    return;
  }

  recordingState.mediaRecorder.stop();
  renderRecordingStatus("録音を保存しています…");
  elements.stopRecordingButton.disabled = true;
}

async function deleteRecording(recordingId) {
  try {
    await deleteRecordingsFromDb([recordingId]);
    await refreshRecordingHistory();
    showToast("録音を削除しました");
  } catch (error) {
    console.error("Failed to delete recording:", error);
    showToast("録音を削除できませんでした", true);
  }
}

function downloadRecording(recordingId) {
  const entry = recordingState.recordings.find((recording) => recording.id === recordingId);

  if (!entry) {
    showToast("録音データが見つかりませんでした", true);
    return;
  }

  const url = window.URL.createObjectURL(entry.blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = getRecordingDownloadName(entry);
  document.body.append(link);
  link.click();
  link.remove();
  window.URL.revokeObjectURL(url);
  showToast("録音をダウンロードしました");
}

function getCalendarDateKey(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getTodayFourteen(date = new Date()) {
  const cutoff = new Date(date);
  cutoff.setHours(14, 0, 0, 0);
  return cutoff;
}

function pruneBusinessDayOverride() {
  if (!state.businessDayOverride?.expiresAt) {
    return;
  }

  if (new Date(state.businessDayOverride.expiresAt) <= new Date()) {
    state.businessDayOverride = null;
    saveState();
  }
}

function getBusinessDate() {
  const now = new Date();

  pruneBusinessDayOverride();

  if (state.businessDayOverride?.dateKey && state.businessDayOverride?.expiresAt) {
    return new Date(`${state.businessDayOverride.dateKey}T00:00:00`);
  }

  if (now.getHours() < 14) {
    const previous = new Date(now);
    previous.setDate(previous.getDate() - 1);
    return previous;
  }

  return now;
}

function getTodayKey() {
  const businessDate = getBusinessDate();
  const year = businessDate.getFullYear();
  const month = String(businessDate.getMonth() + 1).padStart(2, "0");
  const day = String(businessDate.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getMonthKey(dateKey = getTodayKey()) {
  return dateKey.slice(0, 7);
}

function formatDate(dateKey) {
  if (!dateKey) {
    return "-";
  }

  const [year, month, day] = dateKey.split("-");
  return `${year}/${month}/${day}`;
}

function formatMonth(dateKey) {
  if (!dateKey) {
    return "-";
  }

  const [year, month] = dateKey.split("-");
  return `${year}/${month}`;
}

function formatMonthLabel(monthKey) {
  if (!monthKey) {
    return "-";
  }

  const [year, month] = monthKey.split("-");
  return `${year}/${Number(month)}`;
}

function formatShortDate(dateKey) {
  if (!dateKey) {
    return "/";
  }

  const [, month, day] = dateKey.split("-");
  return `${Number(month)}/${Number(day)}`;
}

function formatTime(isoString) {
  if (!isoString) {
    return "未更新";
  }

  return new Date(isoString).toLocaleTimeString("ja-JP", {
    hour: "2-digit",
    minute: "2-digit",
  });
}

function getRate(numerator, denominator) {
  if (!denominator) {
    return 0;
  }

  return Math.round((numerator / denominator) * 1000) / 10;
}

function incrementCalls(record, amount = 1) {
  record.calls += amount;
}

function incrementConnections(record, amount = 1) {
  record.calls += amount;
  record.connections += amount;
}

function incrementSampleSent(record, amount = 1) {
  record.calls += amount;
  record.connections += amount;
  record.sampleSent += amount;
}

function incrementIntroductions(record, amount = 1) {
  record.calls += amount;
  record.introductions += amount;
}

function decrementCalls(record, amount = 1) {
  record.calls = Math.max(0, record.calls - amount);
}

function decrementConnections(record, amount = 1) {
  record.calls = Math.max(0, record.calls - amount);
  record.connections = Math.max(0, record.connections - amount);
}

function decrementSampleSent(record, amount = 1) {
  record.calls = Math.max(0, record.calls - amount);
  record.connections = Math.max(0, record.connections - amount);
  record.sampleSent = Math.max(0, record.sampleSent - amount);
}

function decrementIntroductions(record, amount = 1) {
  record.calls = Math.max(0, record.calls - amount);
  record.introductions = Math.max(0, record.introductions - amount);
}

function toIntOrNull(value) {
  const normalized = String(value ?? "").trim();

  if (!normalized) {
    return null;
  }

  const parsed = Number.parseInt(normalized, 10);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : null;
}

function goalLabel(value) {
  return value === null ? "/" : String(value);
}

function getTodayRecord() {
  const todayKey = getTodayKey();

  if (!state.records[todayKey]) {
    return sanitizeRecord({});
  }

  return state.records[todayKey];
}

function getLatestGoalsBeforeDate(targetDateKey) {
  const previousDateKeys = Object.keys(state.dailyGoals)
    .filter((dateKey) => dateKey < targetDateKey)
    .sort((left, right) => right.localeCompare(left));

  if (!previousDateKeys.length) {
    return null;
  }

  return sanitizeGoals(state.dailyGoals[previousDateKeys[0]]);
}

function getCurrentGoals() {
  const todayKey = getTodayKey();

  if (!state.dailyGoals[todayKey]) {
    state.dailyGoals[todayKey] = getLatestGoalsBeforeDate(todayKey) || sanitizeGoals({});
    saveState();
  }

  return state.dailyGoals[todayKey];
}

function goalsMissing(goals) {
  return Object.values(goals).some((value) => value === null);
}

function showToast(message, isError = false) {
  elements.toast.textContent = message;
  elements.toast.style.background = isError ? "rgba(201, 86, 79, 0.94)" : "rgba(31, 35, 48, 0.92)";
  elements.toast.classList.add("is-visible");

  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => {
    elements.toast.classList.remove("is-visible");
  }, 2600);
}

function getDraftGoals() {
  return {
    calls: toIntOrNull(elements.goalCallsInput.value),
    sampleSent: toIntOrNull(elements.goalSampleSentInput.value),
    introductions: toIntOrNull(elements.goalIntroductionsInput.value),
    connections: toIntOrNull(elements.goalConnectionsInput.value),
  };
}

function getDraftReportFields() {
  const todayRecord = getTodayRecord();

  return {
    location: elements.locationSelect.value || todayRecord.location || "",
    workedHours: elements.workedHoursInput.value.trim() || todayRecord.workedHours || "",
    reflection: elements.reflectionInput.value.trim() || todayRecord.reflection || "",
  };
}

function ensureLocationOption(location) {
  const safeLocation = String(location || "").trim().slice(0, 40);

  if (!safeLocation) {
    return;
  }

  if (!state.locationOptions.includes(safeLocation)) {
    state.locationOptions.push(safeLocation);
  }
}

function renderLocationOptions(selectedLocation) {
  const value = String(selectedLocation || "").trim();

  if (value) {
    ensureLocationOption(value);
  }

  elements.locationSelect.innerHTML = state.locationOptions
    .map((option) => `<option value="${option}">${option}</option>`)
    .join("");

  elements.locationSelect.value = value && state.locationOptions.includes(value) ? value : "自宅";
}

function buildTemplate() {
  const todayKey = getTodayKey();
  const todayRecord = getTodayRecord();
  const draftGoals = getDraftGoals();
  const draftFields = getDraftReportFields();

  return [
    `日付：${formatShortDate(todayKey)}`,
    `稼働場所：${draftFields.location || "/"}`,
    `実稼働時間：${draftFields.workedHours || "/"}`,
    "",
    "【実績/目標】",
    `架電件数 ${todayRecord.calls}/${goalLabel(draftGoals.calls)}`,
    `サンプル送付数：${todayRecord.sampleSent}/${goalLabel(draftGoals.sampleSent)}`,
    `導入件数：${todayRecord.introductions}/${goalLabel(draftGoals.introductions)}`,
    `担当者接続 ${todayRecord.connections}/${goalLabel(draftGoals.connections)}`,
    "",
    "【振り返り】",
    draftFields.reflection || "/",
  ].join("\n");
}

function getAggregateTotals() {
  const totals = {
    calls: 0,
    connections: 0,
    sampleSent: 0,
    introductions: 0,
  };

  for (const record of Object.values(state.records)) {
    const safeRecord = sanitizeRecord(record);
    totals.calls += safeRecord.calls;
    totals.connections += safeRecord.connections;
    totals.sampleSent += safeRecord.sampleSent;
    totals.introductions += safeRecord.introductions;
  }

  return totals;
}

function getAvailableHistoryMonths() {
  const months = Object.keys(state.records).map((dateKey) => getMonthKey(dateKey));
  months.push(getMonthKey());
  return Array.from(new Set(months)).sort((left, right) => right.localeCompare(left));
}

function getSelectedHistoryMonth() {
  const availableMonths = getAvailableHistoryMonths();

  if (state.selectedHistoryMonth && availableMonths.includes(state.selectedHistoryMonth)) {
    return state.selectedHistoryMonth;
  }

  return availableMonths[0] || getMonthKey();
}

function renderHistoryMonthOptions(selectedMonth) {
  const months = getAvailableHistoryMonths();

  elements.historyMonthSelect.innerHTML = months
    .map((monthKey) => `<option value="${monthKey}">${formatMonthLabel(monthKey)}</option>`)
    .join("");

  elements.historyMonthSelect.value = months.includes(selectedMonth) ? selectedMonth : months[0];
}

function getHistoryMonthTotals(selectedMonth, rows) {
  const totals = {
    calls: 0,
    connections: 0,
    sampleSent: 0,
  };

  for (const [dateKey, record] of rows) {
    if (getMonthKey(dateKey) !== selectedMonth) {
      continue;
    }

    totals.calls += record.calls;
    totals.connections += record.connections;
    totals.sampleSent += record.sampleSent;
  }

  return totals;
}

function renderHistory() {
  const allRows = Object.entries(state.records)
    .map(([dateKey, record]) => [dateKey, sanitizeRecord(record)])
    .sort(([left], [right]) => right.localeCompare(left));
  const selectedMonth = getSelectedHistoryMonth();
  const rows = allRows.filter(([dateKey]) => getMonthKey(dateKey) === selectedMonth);
  const totals = getHistoryMonthTotals(selectedMonth, allRows);
  const aggregateTotals = getAggregateTotals();

  renderHistoryMonthOptions(selectedMonth);
  elements.callsValue.textContent = aggregateTotals.calls;
  elements.connectionsValue.textContent = aggregateTotals.connections;
  elements.sampleSentValue.textContent = aggregateTotals.sampleSent;
  elements.introductionsValue.textContent = aggregateTotals.introductions;
  elements.connectionTotalRateValue.textContent = `${getRate(
    aggregateTotals.connections,
    aggregateTotals.calls,
  )}%`;
  elements.sampleTotalRateValue.textContent = `${getRate(
    aggregateTotals.sampleSent,
    aggregateTotals.calls,
  )}%`;
  elements.paceText.textContent = "全期間の架電に対する担当者接続率です";
  elements.sampleRateText.textContent = "全期間の架電に対するサンプル送付率です";
  elements.historyMonthConnectionRateValue.textContent = `${getRate(totals.connections, totals.calls)}%`;
  elements.historyMonthSampleRateValue.textContent = `${getRate(totals.sampleSent, totals.calls)}%`;

  if (!rows.length) {
    elements.historyBody.innerHTML = `
      <tr>
        <td colspan="8" class="empty-cell">この月の記録はまだありません</td>
      </tr>
    `;
    return;
  }

  elements.historyBody.innerHTML = rows
    .map(([dateKey, record]) => {
      const connectionRate = getRate(record.connections, record.calls);
      const sampleRate = getRate(record.sampleSent, record.calls);

      return `
        <tr>
          <td>${formatDate(dateKey)}</td>
          <td>${record.calls}</td>
          <td>${record.connections}</td>
          <td>${connectionRate}%</td>
          <td>${record.sampleSent}</td>
          <td>${sampleRate}%</td>
          <td>${record.introductions}</td>
          <td>${formatTime(record.updatedAt)}</td>
        </tr>
      `;
    })
    .join("");
}

function render() {
  const todayKey = getTodayKey();
  const todayRecord = getTodayRecord();
  const currentGoals = getCurrentGoals();
  const connectionDayRate = getRate(todayRecord.connections, todayRecord.calls);
  const sampleDayRate = getRate(todayRecord.sampleSent, todayRecord.calls);

  elements.todayLabel.textContent = formatDate(todayKey);
  elements.updatedAtLabel.textContent = todayRecord.updatedAt
    ? `${formatDate(todayKey)} ${formatTime(todayRecord.updatedAt)}`
    : "まだ記録されていません";
  elements.storageStatusLabel.textContent = "ブラウザ内";
  elements.todayCallsValue.textContent = todayRecord.calls;
  elements.todayConnectionsValue.textContent = todayRecord.connections;
  elements.todaySampleSentValue.textContent = todayRecord.sampleSent;
  elements.todayIntroductionsValue.textContent = todayRecord.introductions;
  elements.todayConnectionRateValue.textContent = `${connectionDayRate}%`;
  elements.todayPaceText.textContent = "架電に対する担当者接続率です";
  elements.todaySampleRateValue.textContent = `${sampleDayRate}%`;
  elements.todaySampleRateText.textContent = "架電に対するサンプル送付率です";

  elements.monthLabel.textContent = `${formatDate(todayKey)} の目標`;
  elements.goalCallsInput.value = currentGoals.calls ?? "";
  elements.goalSampleSentInput.value = currentGoals.sampleSent ?? "";
  elements.goalIntroductionsInput.value = currentGoals.introductions ?? "";
  elements.goalConnectionsInput.value = currentGoals.connections ?? "";
  elements.goalsStatusText.textContent = goalsMissing(currentGoals)
    ? "今日の目標が未設定です。必要ならここで保存してください。"
    : "今日の目標は保存済みです。対象日ごとに保持されます。";

  renderLocationOptions(todayRecord.location || "自宅");
  elements.workedHoursInput.value = todayRecord.workedHours || "";
  elements.reflectionInput.value = todayRecord.reflection || "";
  elements.templatePreview.value = buildTemplate();

  renderActivePage();
  renderHistory();
}

function updateTodayRecord(mutator) {
  const todayKey = getTodayKey();
  const current = getTodayRecord();
  const next = {
    calls: current.calls,
    connections: current.connections,
    sampleSent: current.sampleSent,
    introductions: current.introductions,
    location: current.location,
    workedHours: current.workedHours,
    reflection: current.reflection,
  };

  mutator(next);

  const sanitized = sanitizeRecord(next);
  sanitized.updatedAt = new Date().toISOString();
  state.records[todayKey] = sanitized;
  saveState();
  render();
}

function saveGoals() {
  const todayKey = getTodayKey();
  state.dailyGoals[todayKey] = sanitizeGoals(getDraftGoals());
  saveState();
  render();
  showToast("今日の目標を保存しました");
}

function saveDailyMemo() {
  updateTodayRecord((record) => {
    record.location = elements.locationSelect.value;
    record.workedHours = elements.workedHoursInput.value;
    record.reflection = elements.reflectionInput.value;
  });
  showToast("日報メモを保存しました");
}

function advanceBusinessDay() {
  const now = new Date();
  const currentBusinessDateKey = getTodayKey();
  const currentCalendarDateKey = getCalendarDateKey(now);

  if (currentBusinessDateKey === currentCalendarDateKey) {
    showToast("すでに現在日の営業日に切り替わっています");
    return;
  }

  state.businessDayOverride = {
    dateKey: currentCalendarDateKey,
    expiresAt: getTodayFourteen(now).toISOString(),
  };
  saveState();
  render();
  showToast(`${formatDate(currentCalendarDateKey)} の営業日に切り替えました`);
}

function addLocationOption() {
  const location = elements.newLocationInput.value.trim().slice(0, 40);

  if (!location) {
    showToast("追加する稼働場所を入力してください", true);
    return;
  }

  ensureLocationOption(location);
  saveState();
  renderLocationOptions(location);
  elements.newLocationInput.value = "";
  elements.templatePreview.value = buildTemplate();
  showToast(`稼働場所「${location}」を追加しました`);
}

function alignBusinessDayToCurrentDate() {
  const now = new Date();
  const currentCalendarDateKey = getCalendarDateKey(now);
  const currentBusinessDateKey = getTodayKey();

  if (currentBusinessDateKey !== currentCalendarDateKey) {
    state.businessDayOverride = {
      dateKey: currentCalendarDateKey,
      expiresAt: getTodayFourteen(now).toISOString(),
    };
    return currentCalendarDateKey;
  }

  state.businessDayOverride = null;
  return currentCalendarDateKey;
}

function exportCsv() {
  const rows = [
    [
      "date",
      "calls",
      "connections",
      "sample_sent",
      "introductions",
      "connection_rate_percent",
      "location",
      "worked_hours",
      "reflection",
      "updated_at",
    ],
  ];

  const records = Object.entries(state.records).sort(([left], [right]) => right.localeCompare(left));

  for (const [dateKey, record] of records) {
    const safeRecord = sanitizeRecord(record);
    rows.push([
      dateKey,
      String(safeRecord.calls),
      String(safeRecord.connections),
      String(safeRecord.sampleSent),
      String(safeRecord.introductions),
      String(getRate(safeRecord.connections, safeRecord.calls)),
      JSON.stringify(safeRecord.location || ""),
      JSON.stringify(safeRecord.workedHours || ""),
      JSON.stringify(safeRecord.reflection || ""),
      safeRecord.updatedAt || "",
    ]);
  }

  const csv = rows.map((row) => row.join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = `call-tracker-${getTodayKey()}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function copyTemplate() {
  const text = buildTemplate();

  try {
    await navigator.clipboard.writeText(text);
  } catch (error) {
    elements.templatePreview.focus();
    elements.templatePreview.select();
    document.execCommand("copy");
  }

  elements.templatePreview.value = text;
  showToast("テンプレートをコピーしました");
}

document.addEventListener("click", (event) => {
  const action = event.target.closest("[data-action]")?.dataset.action;
  const recordingAction = event.target.closest("[data-recording-action]")?.dataset.recordingAction;
  const recordingId = event.target.closest("[data-recording-id]")?.dataset.recordingId;

  if (recordingAction && recordingId) {
    if (recordingAction === "download") {
      downloadRecording(recordingId);
      return;
    }

    if (recordingAction === "delete") {
      deleteRecording(recordingId);
      return;
    }
  }

  if (!action) {
    return;
  }

  if (action === "log-call-plus") {
    updateTodayRecord((record) => {
      incrementCalls(record);
    });
    return;
  }

  if (action === "log-call-minus") {
    updateTodayRecord((record) => {
      decrementCalls(record);
    });
    return;
  }

  if (action === "log-connect-plus") {
    updateTodayRecord((record) => {
      incrementConnections(record);
    });
    return;
  }

  if (action === "log-connect-minus") {
    updateTodayRecord((record) => {
      decrementConnections(record);
    });
    return;
  }

  if (action === "log-sample-plus") {
    updateTodayRecord((record) => {
      incrementSampleSent(record);
    });
    return;
  }

  if (action === "log-sample-minus") {
    updateTodayRecord((record) => {
      decrementSampleSent(record);
    });
    return;
  }

  if (action === "log-introduction-plus") {
    updateTodayRecord((record) => {
      incrementIntroductions(record);
    });
    return;
  }

  if (action === "log-introduction-minus") {
    updateTodayRecord((record) => {
      decrementIntroductions(record);
    });
  }
});

elements.exportButton.addEventListener("click", exportCsv);
elements.advanceBusinessDayButton.addEventListener("click", advanceBusinessDay);
elements.startRecordingButton.addEventListener("click", startRecording);
elements.stopRecordingButton.addEventListener("click", stopRecording);
elements.addLocationButton.addEventListener("click", addLocationOption);
elements.locationSelect.addEventListener("change", () => {
  elements.templatePreview.value = buildTemplate();
});
elements.historyMonthSelect.addEventListener("change", () => {
  state.selectedHistoryMonth = elements.historyMonthSelect.value;
  saveState();
  renderHistory();
});
elements.pageTabs.forEach((button) => {
  button.addEventListener("click", () => {
    const nextPage = button.dataset.pageTab;

    if (!PAGE_KEYS.includes(nextPage)) {
      return;
    }

    state.activePage = nextPage;
    saveState();
    renderActivePage();
  });
});

elements.resetTodayButton.addEventListener("click", () => {
  const confirmed = window.confirm("この営業日の件数と日報メモを 0 に戻し、対象日も更新します。");

  if (!confirmed) {
    return;
  }

  const todayKey = alignBusinessDayToCurrentDate();
  state.records[todayKey] = sanitizeRecord({});
  state.records[todayKey].updatedAt = new Date().toISOString();
  saveState();
  render();
});

elements.saveGoalsButton.addEventListener("click", saveGoals);
elements.saveMemoButton.addEventListener("click", saveDailyMemo);
elements.copyTemplateButton.addEventListener("click", copyTemplate);

[
  elements.goalCallsInput,
  elements.goalSampleSentInput,
  elements.goalIntroductionsInput,
  elements.goalConnectionsInput,
  elements.workedHoursInput,
  elements.reflectionInput,
].forEach((element) => {
  element.addEventListener("input", () => {
    elements.templatePreview.value = buildTemplate();
  });
});

render();
renderRecordingHistory();
refreshRecordingHistory();

if (goalsMissing(getCurrentGoals())) {
  showToast(`${formatDate(getTodayKey())} の目標を入力してください`, true);
}
