const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const AUTH_STORAGE_KEY = "microsoftAuth";
const DEFAULT_SCOPES = ["openid", "profile", "offline_access", "User.Read", "Tasks.ReadWrite"];

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  void handleMessage(message)
    .then((data) => sendResponse({ ok: true, data }))
    .catch((error) => sendResponse({ ok: false, error: error.message || String(error) }));
  return true;
});

async function handleMessage(message) {
  switch (message?.type) {
    case "auth:getRedirectUri":
      return { redirectUri: chrome.identity.getRedirectURL("auth") };
    case "auth:status":
      return getAuthStatus();
    case "auth:connect":
      return connectMicrosoft(normalizeSettings(message.settings));
    case "auth:disconnect":
      await chrome.storage.local.remove(AUTH_STORAGE_KEY);
      return { signedIn: false };
    case "todo:push":
      return pushTasksToMicrosoft(normalizeSettings(message.settings), message.tasks || []);
    default:
      throw new Error("Unknown background message.");
  }
}

async function connectMicrosoft(settings) {
  validateClientId(settings.clientId);
  const token = await acquireInteractiveToken(settings);
  const profile = await graphFetch("/me?$select=displayName,userPrincipalName,mail", token);
  const auth = await getStoredAuth();
  const account = {
    displayName: profile.displayName || "",
    userPrincipalName: profile.userPrincipalName || profile.mail || ""
  };
  await chrome.storage.local.set({
    [AUTH_STORAGE_KEY]: {
      ...auth,
      account
    }
  });
  return { signedIn: true, account };
}

async function getAuthStatus() {
  const auth = await getStoredAuth();
  if (!auth?.accessToken && !auth?.refreshToken) {
    return { signedIn: false };
  }
  return {
    signedIn: true,
    account: auth.account || null,
    expiresAt: auth.expiresAt || 0
  };
}

async function pushTasksToMicrosoft(settings, tasks) {
  validateClientId(settings.clientId);
  if (!Array.isArray(tasks) || tasks.length === 0) {
    throw new Error("No tasks to create.");
  }

  const accessToken = await acquireToken(settings);
  const list = await ensureTodoList(accessToken, settings.listName);
  const results = [];

  for (const task of tasks) {
    if (!task?.title?.trim()) {
      continue;
    }

    const created = await createTodoTask(accessToken, list.id, task, settings);
    results.push({
      id: created.id,
      title: created.title
    });
  }

  return {
    createdCount: results.length,
    listName: list.displayName,
    tasks: results
  };
}

async function ensureTodoList(accessToken, listName) {
  const wantedName = (listName || "Reminder Gen Inbox").trim();
  const lists = await graphFetch("/me/todo/lists", accessToken);
  const existing = (lists.value || []).find(
    (list) => list.displayName?.toLowerCase() === wantedName.toLowerCase()
  );
  if (existing) {
    return existing;
  }

  return graphFetch("/me/todo/lists", accessToken, {
    method: "POST",
    body: JSON.stringify({ displayName: wantedName })
  });
}

async function createTodoTask(accessToken, listId, task, settings) {
  const payload = {
    title: task.title.trim()
  };

  const body = buildTaskBody(task);
  if (body) {
    payload.body = {
      contentType: "text",
      content: body
    };
  }

  if (task.importance && task.importance !== "normal") {
    payload.importance = task.importance;
  }

  if (task.dueAt) {
    payload.dueDateTime = {
      dateTime: toGraphDateTime(task.dueAt),
      timeZone: settings.timezone || "UTC"
    };
  }

  if (task.reminderAt) {
    payload.isReminderOn = true;
    payload.reminderDateTime = {
      dateTime: toGraphDateTime(task.reminderAt),
      timeZone: settings.timezone || "UTC"
    };
  }

  return graphFetch(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`, accessToken, {
    method: "POST",
    body: JSON.stringify(payload)
  });
}

function buildTaskBody(task) {
  const lines = [];
  if (task.notes?.trim()) {
    lines.push(task.notes.trim());
  }
  if (task.sourceText?.trim() && task.sourceText.trim() !== task.notes?.trim()) {
    lines.push(`Source: ${task.sourceText.trim()}`);
  }
  if (Array.isArray(task.tags) && task.tags.length) {
    lines.push(`Tags: ${task.tags.map((tag) => `#${tag}`).join(" ")}`);
  }
  return lines.join("\n\n");
}

function toGraphDateTime(value) {
  return value.length === 16 ? `${value}:00` : value;
}

async function acquireToken(settings) {
  const auth = await getStoredAuth();
  if (auth?.accessToken && auth.expiresAt && auth.expiresAt > Date.now() + 120000) {
    return auth.accessToken;
  }

  if (auth?.refreshToken) {
    try {
      return await refreshAccessToken(settings, auth.refreshToken);
    } catch (_error) {
      await chrome.storage.local.remove(AUTH_STORAGE_KEY);
    }
  }

  return acquireInteractiveToken(settings);
}

async function acquireInteractiveToken(settings) {
  const redirectUri = chrome.identity.getRedirectURL("auth");
  const state = createRandomToken(24);
  const codeVerifier = createRandomToken(64);
  const codeChallenge = await createCodeChallenge(codeVerifier);
  const authUrl = new URL(`https://login.microsoftonline.com/${settings.tenant}/oauth2/v2.0/authorize`);
  authUrl.searchParams.set("client_id", settings.clientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", redirectUri);
  authUrl.searchParams.set("response_mode", "query");
  authUrl.searchParams.set("scope", DEFAULT_SCOPES.join(" "));
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("code_challenge", codeChallenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("prompt", "select_account");

  const responseUrl = await chrome.identity.launchWebAuthFlow({
    url: authUrl.toString(),
    interactive: true
  });

  if (!responseUrl) {
    throw new Error("Microsoft sign-in was cancelled.");
  }

  const parsedUrl = new URL(responseUrl);
  const returnedState = parsedUrl.searchParams.get("state");
  if (returnedState !== state) {
    throw new Error("Microsoft sign-in state mismatch.");
  }

  const error = parsedUrl.searchParams.get("error");
  if (error) {
    throw new Error(parsedUrl.searchParams.get("error_description") || error);
  }

  const code = parsedUrl.searchParams.get("code");
  if (!code) {
    throw new Error("Microsoft did not return an authorization code.");
  }

  return exchangeAuthorizationCode(settings, code, codeVerifier, redirectUri);
}

async function exchangeAuthorizationCode(settings, code, codeVerifier, redirectUri) {
  const params = new URLSearchParams({
    client_id: settings.clientId,
    scope: DEFAULT_SCOPES.join(" "),
    code,
    redirect_uri: redirectUri,
    grant_type: "authorization_code",
    code_verifier: codeVerifier
  });

  const token = await tokenRequest(settings.tenant, params);
  await storeToken(token);
  return token.access_token;
}

async function refreshAccessToken(settings, refreshToken) {
  const params = new URLSearchParams({
    client_id: settings.clientId,
    scope: DEFAULT_SCOPES.join(" "),
    refresh_token: refreshToken,
    grant_type: "refresh_token"
  });

  const token = await tokenRequest(settings.tenant, params);
  await storeToken(token);
  return token.access_token;
}

async function tokenRequest(tenant, params) {
  const response = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: params.toString()
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.error_description || data.error || "Microsoft token request failed.");
  }
  return data;
}

async function storeToken(token) {
  const existing = await getStoredAuth();
  await chrome.storage.local.set({
    [AUTH_STORAGE_KEY]: {
      ...existing,
      accessToken: token.access_token,
      refreshToken: token.refresh_token || existing?.refreshToken || "",
      expiresAt: Date.now() + Number(token.expires_in || 3600) * 1000
    }
  });
}

async function graphFetch(path, accessToken, options = {}) {
  const response = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(options.headers || {})
    }
  });
  const text = await response.text();
  const data = text ? parseJsonResponse(text) : {};
  if (!response.ok) {
    throw new Error(data.error?.message || "Microsoft Graph request failed.");
  }
  return data;
}

function parseJsonResponse(text) {
  try {
    return JSON.parse(text);
  } catch (_error) {
    return {};
  }
}

async function getStoredAuth() {
  const result = await chrome.storage.local.get(AUTH_STORAGE_KEY);
  return result[AUTH_STORAGE_KEY] || null;
}

function normalizeSettings(settings = {}) {
  return {
    clientId: String(settings.clientId || "").trim(),
    tenant: normalizeTenant(settings.tenant),
    listName: String(settings.listName || "Reminder Gen Inbox").trim() || "Reminder Gen Inbox",
    timezone: String(settings.timezone || "UTC").trim() || "UTC"
  };
}

function normalizeTenant(value) {
  const tenant = String(value || "").trim();
  return tenant === "common" || !tenant ? "consumers" : tenant;
}

function validateClientId(clientId) {
  if (!/^[0-9a-f-]{30,}$/i.test(clientId || "")) {
    throw new Error("Paste a valid Microsoft Application (client) ID first.");
  }
}

function createRandomToken(byteLength) {
  const bytes = new Uint8Array(byteLength);
  crypto.getRandomValues(bytes);
  return base64Url(bytes);
}

async function createCodeChallenge(codeVerifier) {
  const data = new TextEncoder().encode(codeVerifier);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return base64Url(new Uint8Array(digest));
}

function base64Url(bytes) {
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}
