const SETTINGS_STORAGE_KEY = "settings";
const DEFAULT_LIST_NAME = "Reminder Gen Inbox";
const DEFAULT_REMINDER_TIME = "18:00";
const RELATIVE_DUE_MODES = ["today", "tomorrow", "week", "month", "year"];
const CUSTOM_DUE_MODE = "custom";

const elements = {
  authStatus: document.querySelector("#auth-status"),
  openSetup: document.querySelector("#open-setup"),
  connectMicrosoft: document.querySelector("#connect-microsoft"),
  disconnectMicrosoft: document.querySelector("#disconnect-microsoft"),
  sourceText: document.querySelector("#source-text"),
  sourceGuideLines: document.querySelectorAll(".source-guide-line"),
  defaultDueButtons: document.querySelectorAll("[data-default-due]"),
  customDueDate: document.querySelector("#custom-due-date"),
  reminderTime: document.querySelector("#reminder-time"),
  priorityOn: document.querySelector("#priority-on"),
  summary: document.querySelector("#summary"),
  pushTasks: document.querySelector("#push-tasks")
};

const state = {
  auth: { signedIn: false },
  pushed: false,
  settings: {
    clientId: "",
    tenant: "consumers",
    listName: DEFAULT_LIST_NAME,
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC",
    defaultDueMode: "today",
    customDueDate: "",
    reminderTime: DEFAULT_REMINDER_TIME,
    priorityOn: false
  }
};

void init();

async function init() {
  bindOptional(elements.openSetup, "click", openSetupPage);
  bindOptional(elements.connectMicrosoft, "click", openSetupPage);
  bindOptional(elements.disconnectMicrosoft, "click", disconnectMicrosoft);
  bindOptional(elements.pushTasks, "click", pushTasksToMicrosoft);
  bindOptional(elements.sourceText, "input", () => {
    state.pushed = false;
    updateSourceGuide();
    updateActionAvailability();
  });
  bindOptional(elements.customDueDate, "change", () => {
    state.settings.defaultDueMode = CUSTOM_DUE_MODE;
    state.settings.customDueDate = normalizeDateInput(elements.customDueDate.value) || dueDateFromMode("today");
    state.pushed = false;
    renderSettings();
    void saveSettings();
    updateActionAvailability();
  });
  bindOptional(elements.reminderTime, "change", () => {
    state.settings.reminderTime = normalizeTimeInput(elements.reminderTime.value);
    state.pushed = false;
    renderSettings();
    void saveSettings();
    updateActionAvailability();
  });
  bindOptional(elements.priorityOn, "change", () => {
    state.settings.priorityOn = elements.priorityOn.checked;
    state.pushed = false;
    void saveSettings();
    updateActionAvailability();
  });

  elements.defaultDueButtons.forEach((button) => {
    bindOptional(button, "click", () => {
      state.settings.defaultDueMode = normalizeDefaultDueMode(button.dataset.defaultDue);
      state.settings.customDueDate = "";
      state.pushed = false;
      renderSettings();
      void saveSettings();
      updateActionAvailability();
    });
  });

  await loadSettings();
  await refreshAuthStatus();
  updateSourceGuide();
  updateActionAvailability();

  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== "local" || !changes[SETTINGS_STORAGE_KEY]) {
      return;
    }

    state.settings = normalizeSettings({
      ...state.settings,
      ...(changes[SETTINGS_STORAGE_KEY].newValue || {})
    });
    renderSettings();
    renderAuthStatus();
    updateActionAvailability();
  });
}

function bindOptional(node, eventName, handler) {
  if (node) {
    node.addEventListener(eventName, handler);
  }
}

async function loadSettings() {
  const result = await chrome.storage.local.get(SETTINGS_STORAGE_KEY);
  state.settings = normalizeSettings({
    ...state.settings,
    ...(result[SETTINGS_STORAGE_KEY] || {})
  });
  // Default is always non-priority when the popup is opened.
  state.settings.priorityOn = false;
  await saveSettings();
  renderSettings();
}

function normalizeSettings(settings) {
  const mode = normalizeDefaultDueMode(settings.defaultDueMode);
  const customDueDate = normalizeDateInput(settings.customDueDate);
  return {
    ...settings,
    tenant: normalizeTenant(settings.tenant),
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC",
    defaultDueMode: mode,
    customDueDate: mode === CUSTOM_DUE_MODE ? customDueDate || dueDateFromMode("today") : "",
    reminderTime: normalizeTimeInput(settings.reminderTime),
    priorityOn: Boolean(settings.priorityOn)
  };
}

async function saveSettings() {
  await chrome.storage.local.set({ [SETTINGS_STORAGE_KEY]: state.settings });
}

function renderSettings() {
  elements.customDueDate.value = selectedDueDate();
  elements.reminderTime.value = state.settings.reminderTime;
  elements.priorityOn.checked = Boolean(state.settings.priorityOn);
  updateDefaultDueButtons();
}

function normalizeTenant(value) {
  const tenant = String(value || "").trim();
  return tenant === "common" || !tenant ? "consumers" : tenant;
}

function normalizeDefaultDueMode(value) {
  return [...RELATIVE_DUE_MODES, CUSTOM_DUE_MODE].includes(value) ? value : "today";
}

function normalizeDateInput(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(value || "") ? value : "";
}

function normalizeTimeInput(value) {
  return /^([01]\d|2[0-3]):[0-5]\d$/.test(value || "") ? value : DEFAULT_REMINDER_TIME;
}

function updateDefaultDueButtons() {
  elements.defaultDueButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.defaultDue === state.settings.defaultDueMode);
  });
}

async function refreshAuthStatus() {
  state.auth = await sendMessage({ type: "auth:status" });
  renderAuthStatus();
}

function renderAuthStatus(message = null, isError = false) {
  const account = state.auth?.account;
  const accountName = account?.userPrincipalName || account?.displayName || "Microsoft account";
  elements.authStatus.textContent =
    message ||
    (state.auth?.signedIn
      ? accountName
      : state.settings.clientId
        ? `Sign in required. Target list: ${state.settings.listName || DEFAULT_LIST_NAME}.`
        : "Setup required: add a Microsoft Application client ID.");
  elements.authStatus.classList.toggle("is-error", isError);
  elements.connectMicrosoft.disabled = Boolean(state.auth?.signedIn);
  elements.disconnectMicrosoft.disabled = !state.auth?.signedIn;
}

function openSetupPage() {
  chrome.runtime.openOptionsPage();
}

async function disconnectMicrosoft() {
  await sendMessage({ type: "auth:disconnect" });
  state.auth = { signedIn: false };
  renderAuthStatus();
  updateActionAvailability();
}

function updateSourceGuide() {
  const lines = elements.sourceText.value.split("\n");
  elements.sourceGuideLines.forEach((lineNode, index) => {
    lineNode.classList.toggle("is-hidden", Boolean(lines[index]?.trim()));
  });
}

function buildTasksFromSource() {
  const dueDate = selectedDueDate();
  const reminderTime = normalizeTimeInput(state.settings.reminderTime);
  const importance = state.settings.priorityOn ? "high" : "normal";
  return splitTaskLines(elements.sourceText.value).map((line) => ({
    id: crypto.randomUUID(),
    title: stripBullet(line),
    dueAt: `${dueDate}T00:00`,
    reminderAt: `${dueDate}T${reminderTime}`,
    dateOnly: true,
    importance,
    tags: [],
    notes: "",
    sourceText: line
  }));
}

function splitTaskLines(text) {
  return text
    .replace(/\r/g, "")
    .replace(/[；;]+/g, "\n")
    .split("\n")
    .map((line) => stripBullet(line.trim()))
    .filter(Boolean);
}

function stripBullet(line) {
  return line
    .replace(/^[-*•]\s+/, "")
    .replace(/^\[[ xX]\]\s+/, "")
    .replace(/^\d+[.)]\s+/, "")
    .trim();
}

function updateActionAvailability() {
  const hasClientId = Boolean(state.settings.clientId);
  const hasTasks = splitTaskLines(elements.sourceText.value).length > 0;
  elements.pushTasks.disabled = state.pushed || !hasClientId || !state.auth?.signedIn || !hasTasks;
}

async function pushTasksToMicrosoft() {
  state.settings = normalizeSettings(state.settings);
  await saveSettings();

  const tasks = buildTasksFromSource();
  if (!tasks.length) {
    setSummary("Add at least one task line.");
    updateActionAvailability();
    return;
  }

  setSummary(`Creating ${tasks.length} tasks in Microsoft To Do...`);
  elements.pushTasks.disabled = true;

  try {
    const result = await sendMessage({
      type: "todo:push",
      settings: state.settings,
      tasks
    });
    state.pushed = true;
    setSummary(`Created ${result.createdCount} tasks in "${result.listName}". Edit the text or due date to push more.`);
  } catch (error) {
    setSummary(error.message);
  } finally {
    updateActionAvailability();
  }
}

function setSummary(message) {
  elements.summary.textContent = message;
}

function dueDateFromMode(mode) {
  const today = startOfToday();
  switch (normalizeDefaultDueMode(mode)) {
    case "tomorrow":
      return formatDateOnly(addDays(today, 1));
    case "week":
      return formatDateOnly(addDays(today, 7));
    case "month":
      return formatDateOnly(addMonths(today, 1));
    case "year":
      return formatDateOnly(addYears(today, 1));
    default:
      return formatDateOnly(today);
  }
}

function selectedDueDate(settings = state.settings) {
  if (settings.defaultDueMode === CUSTOM_DUE_MODE) {
    return normalizeDateInput(settings.customDueDate) || dueDateFromMode("today");
  }
  return dueDateFromMode(settings.defaultDueMode);
}

function startOfToday() {
  const date = new Date();
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addDays(date, amount) {
  const result = new Date(date);
  result.setDate(result.getDate() + amount);
  return result;
}

function addMonths(date, amount) {
  const result = new Date(date);
  result.setMonth(result.getMonth() + amount);
  return result;
}

function addYears(date, amount) {
  const result = new Date(date);
  result.setFullYear(result.getFullYear() + amount);
  return result;
}

function formatDateOnly(date) {
  return [
    String(date.getFullYear()).padStart(4, "0"),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0")
  ].join("-");
}

function sendMessage(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      if (!response?.ok) {
        reject(new Error(response?.error || "Extension request failed."));
        return;
      }
      resolve(response.data);
    });
  });
}
