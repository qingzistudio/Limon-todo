const SETTINGS_STORAGE_KEY = "settings";
const DEFAULT_LIST_NAME = "Reminder Gen Inbox";
const SIGN_IN_WAIT_HINT_MS = 15000;
const SIGN_IN_STUCK_HINT_MS = 45000;

const elements = {
  redirectUri: document.querySelector("#redirect-uri"),
  copyRedirect: document.querySelector("#copy-redirect"),
  clientId: document.querySelector("#client-id"),
  tenant: document.querySelector("#tenant"),
  listName: document.querySelector("#list-name"),
  saveSettings: document.querySelector("#save-settings"),
  saveStatus: document.querySelector("#save-status"),
  connectMicrosoft: document.querySelector("#connect-microsoft"),
  disconnectMicrosoft: document.querySelector("#disconnect-microsoft"),
  authStatus: document.querySelector("#auth-status")
};

const state = {
  auth: { signedIn: false },
  settings: {
    clientId: "",
    tenant: "consumers",
    listName: DEFAULT_LIST_NAME
  }
};

void init();

async function init() {
  elements.copyRedirect.addEventListener("click", copyRedirectUri);
  elements.saveSettings.addEventListener("click", saveSettingsFromForm);
  elements.connectMicrosoft.addEventListener("click", connectMicrosoft);
  elements.disconnectMicrosoft.addEventListener("click", disconnectMicrosoft);
  document.querySelectorAll("[data-tenant-preset]").forEach((button) => {
    button.addEventListener("click", () => {
      elements.tenant.value = button.dataset.tenantPreset;
      renderAuthStatus();
    });
  });
  for (const input of [elements.clientId, elements.tenant, elements.listName]) {
    input.addEventListener("input", () => renderAuthStatus());
  }

  await loadSettings();
  await loadRedirectUri();
  await refreshAuthStatus();
}

async function loadSettings() {
  const result = await chrome.storage.local.get(SETTINGS_STORAGE_KEY);
  state.settings = {
    ...state.settings,
    ...(result[SETTINGS_STORAGE_KEY] || {})
  };
  state.settings.tenant = normalizeTenant(state.settings.tenant);
  renderSettings();
}

function renderSettings() {
  elements.clientId.value = state.settings.clientId || "";
  elements.tenant.value = state.settings.tenant || "consumers";
  elements.listName.value = state.settings.listName || DEFAULT_LIST_NAME;
}

function readSettingsFromForm() {
  return {
    ...state.settings,
    clientId: elements.clientId.value.trim(),
    tenant: normalizeTenant(elements.tenant.value),
    listName: elements.listName.value.trim() || DEFAULT_LIST_NAME
  };
}

function normalizeTenant(value) {
  const tenant = String(value || "").trim();
  return tenant === "common" || !tenant ? "consumers" : tenant;
}

async function saveSettingsFromForm() {
  state.settings = readSettingsFromForm();
  await chrome.storage.local.set({ [SETTINGS_STORAGE_KEY]: state.settings });
  renderSaveStatus("Saved.");
  renderAuthStatus();
}

async function loadRedirectUri() {
  const result = await sendMessage({ type: "auth:getRedirectUri" });
  elements.redirectUri.value = result.redirectUri;
}

async function copyRedirectUri() {
  await navigator.clipboard.writeText(elements.redirectUri.value);
  renderSaveStatus("Redirect URI copied.");
}

async function refreshAuthStatus() {
  state.auth = await sendMessage({ type: "auth:status" });
  renderAuthStatus();
}

function renderAuthStatus(message = null, isError = false) {
  const account = state.auth?.account;
  const configured = Boolean(readSettingsFromForm().clientId);
  elements.authStatus.textContent =
    message ||
    (state.auth?.signedIn
      ? `Connected as ${account?.userPrincipalName || account?.displayName || "Microsoft account"}.`
      : configured
        ? "Ready to sign in."
        : "Not connected. Paste the Application client ID first.");
  elements.authStatus.classList.toggle("is-error", isError);
  elements.connectMicrosoft.disabled = state.auth?.signedIn || !configured;
  elements.disconnectMicrosoft.disabled = !state.auth?.signedIn;
}

function renderSaveStatus(message) {
  elements.saveStatus.textContent = message;
  elements.saveStatus.classList.remove("is-error");
}

async function connectMicrosoft() {
  await saveSettingsFromForm();
  renderAuthStatus(`Opening Microsoft sign-in with tenant "${state.settings.tenant}"...`);
  const waitHint = setTimeout(() => {
    renderAuthStatus(
      "Still waiting for Microsoft to redirect back. If it froze after the email step, check Account type, SPA redirect URI, and supported account types in the Entra app."
    );
  }, SIGN_IN_WAIT_HINT_MS);
  const stuckHint = setTimeout(() => {
    renderAuthStatus(
      "Login is still not returning. Close the Microsoft window, try the Personal Account preset for Outlook/Hotmail, and make sure the Redirect URI is registered as Single-page application, not Web.",
      true
    );
  }, SIGN_IN_STUCK_HINT_MS);
  try {
    state.auth = await sendMessage({ type: "auth:connect", settings: state.settings });
    renderAuthStatus();
  } catch (error) {
    state.auth = { signedIn: false };
    renderAuthStatus(error.message, true);
  } finally {
    clearTimeout(waitHint);
    clearTimeout(stuckHint);
  }
}

async function disconnectMicrosoft() {
  await sendMessage({ type: "auth:disconnect" });
  state.auth = { signedIn: false };
  renderAuthStatus();
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
