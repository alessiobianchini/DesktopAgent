const chatLog = document.getElementById('chatLog');
const chatForm = document.getElementById('chatForm');
const chatInput = document.getElementById('chatInput');
const statusText = document.getElementById('statusText');
const statusBadges = document.getElementById('statusBadges');
const versionText = document.getElementById('versionText');
const confirmArea = document.getElementById('confirmArea');
const auditLog = document.getElementById('auditLog');
const clearHistoryBtn = document.getElementById('clearHistory');
const filterCheckboxes = Array.from(document.querySelectorAll('[data-filter]'));
const tabButtons = Array.from(document.querySelectorAll('[data-tab-target]'));
const tabPanels = Array.from(document.querySelectorAll('[data-tab-panel]'));
const auditBadge = document.getElementById('auditBadge');
const llmEnabledToggle = document.getElementById('llmEnabled');
const llmProviderInput = document.getElementById('llmProvider');
const llmEndpointInput = document.getElementById('llmEndpoint');
const llmAllowRemoteToggle = document.getElementById('llmAllowRemote');
const llmModelInput = document.getElementById('llmModel');
const llmTimeoutInput = document.getElementById('llmTimeout');
const llmMaxTokensInput = document.getElementById('llmMaxTokens');
const auditLlmInteractionsToggle = document.getElementById('auditLlmInteractions');
const auditLlmIncludeRawTextToggle = document.getElementById('auditLlmIncludeRawText');
const profileModeEnabledToggle = document.getElementById('profileModeEnabled');
const activeProfileInput = document.getElementById('activeProfile');
const requireConfirmationToggle = document.getElementById('requireConfirmation');
const quizSafeModeToggle = document.getElementById('quizSafeMode');
const maxActionsInput = document.getElementById('maxActionsPerSecond');
const findRetryCountInput = document.getElementById('findRetryCount');
const findRetryDelayMsInput = document.getElementById('findRetryDelayMs');
const postCheckTimeoutMsInput = document.getElementById('postCheckTimeoutMs');
const postCheckPollMsInput = document.getElementById('postCheckPollMs');
const clipboardHistoryMaxItemsInput = document.getElementById('clipboardHistoryMaxItems');
const filesystemAllowedRootsInput = document.getElementById('filesystemAllowedRoots');
const allowedAppsInput = document.getElementById('allowedApps');
const appAliasesInput = document.getElementById('appAliases');
const ocrEnabledToggle = document.getElementById('ocrEnabled');
const ocrEngineInput = document.getElementById('ocrEngine');
const ocrTesseractInput = document.getElementById('ocrTesseractPath');
const adapterRestartCommandInput = document.getElementById('adapterRestartCommand');
const adapterRestartDirInput = document.getElementById('adapterRestartWorkingDir');
const saveConfigBtn = document.getElementById('saveConfig');
const reloadConfigBtn = document.getElementById('reloadConfig');
const restartServiceBtn = document.getElementById('restartService');
const restartAdapterBtn = document.getElementById('restartAdapter');
const configStatus = document.getElementById('configStatus');
const ocrRestartBadge = document.getElementById('ocrRestartBadge');
const ocrRestartAck = document.getElementById('ocrRestartAck');
const utilityRefreshBtn = document.getElementById('utilityRefresh');
const installFfmpegBtn = document.getElementById('installFfmpeg');
const installOcrBtn = document.getElementById('installOcr');
const installOcrEnableBtn = document.getElementById('installOcrEnable');
const enableOcrOnlyBtn = document.getElementById('enableOcrOnly');
const utilityStatus = document.getElementById('utilityStatus');
const ffmpegToolStatus = document.getElementById('ffmpegToolStatus');
const tesseractToolStatus = document.getElementById('tesseractToolStatus');
const taskNameInput = document.getElementById('taskName');
const taskIntentInput = document.getElementById('taskIntent');
const taskDescriptionInput = document.getElementById('taskDescription');
const taskSaveBtn = document.getElementById('taskSave');
const taskRunBtn = document.getElementById('taskRun');
const taskRefreshBtn = document.getElementById('taskRefresh');
const taskStatus = document.getElementById('taskStatus');
const taskList = document.getElementById('taskList');
const scheduleIdInput = document.getElementById('scheduleId');
const scheduleTaskNameInput = document.getElementById('scheduleTaskName');
const scheduleStartAtInput = document.getElementById('scheduleStartAt');
const scheduleIntervalSecondsInput = document.getElementById('scheduleIntervalSeconds');
const scheduleEnabledToggle = document.getElementById('scheduleEnabled');
const scheduleSaveBtn = document.getElementById('scheduleSave');
const scheduleRunNowBtn = document.getElementById('scheduleRunNow');
const scheduleRefreshBtn = document.getElementById('scheduleRefresh');
const scheduleStatus = document.getElementById('scheduleStatus');
const scheduleList = document.getElementById('scheduleList');
const recordNameInput = document.getElementById('recordName');
const recordStartBtn = document.getElementById('recordStart');
const recordStopBtn = document.getElementById('recordStop');
const recordSaveBtn = document.getElementById('recordSave');
const recordingStatus = document.getElementById('recordingStatus');
const inspectorQueryInput = document.getElementById('inspectorQuery');
const inspectorRefreshBtn = document.getElementById('inspectorRefresh');
const inspectorStatus = document.getElementById('inspectorStatus');
const inspectorOutput = document.getElementById('inspectorOutput');

let lastConfig = null;
const ocrAckKey = 'da_ocr_restart_ack';
let ocrAcked = sessionStorage.getItem(ocrAckKey) === '1';

const historyKey = 'da_chat_history';
const filterKey = 'da_audit_filters';
const tabKey = 'da_active_tab';

let chatHistory = [];
let auditItems = [];
let taskItems = [];
let scheduleItems = [];
let unseenAudit = 0;
let auditInitialized = false;
let lastAuditSize = 0;

const filterMap = {
  action: new Set(['action', 'action_failed', 'dry_run']),
  policy: new Set(['policy_block', 'user_declined']),
  system: new Set([
    'rate_limit', 'kill', 'kill_reset', 'arm', 'disarm', 'simulate_presence', 'presence_required',
    'llm_request', 'llm_response', 'llm_error', 'llm_rewrite_applied', 'llm_fallback_rule_based',
    'schedule_trigger', 'schedule_result'
  ])
};

async function fetchStatus() {
  const res = await fetch('/api/status');
  if (!res.ok) {
    statusText.textContent = 'Unable to contact adapter.';
    if (versionText) versionText.textContent = 'Version: unavailable';
    if (statusBadges) statusBadges.innerHTML = '';
    return;
  }
  const data = await res.json();
  if (versionText) {
    versionText.textContent = `Version: ${data.version || 'unknown'}`;
  }
  let lockText = 'ContextLock: off';
  if (data.contextLock && data.contextLock.enabled) {
    if (data.contextLock.scope === 'window') {
      lockText = `ContextLock: window ${data.contextLock.windowId || 'n/a'}`;
    } else {
      lockText = `ContextLock: app ${data.contextLock.appId || 'n/a'}`;
    }
  }
  statusText.textContent = lockText;
  renderStatusBadges(data.llm, data.adapter, data.killSwitch);
  if (data.recording) {
    setRecordingStatusFromData(data.recording);
  }
  if (data.restart) {
    setOcrRestartBadge(!!data.restart.ocrRequired);
  }
}

async function loadConfig() {
  if (!llmEnabledToggle) return;
  const res = await fetch('/api/config');
  if (!res.ok) {
    setConfigStatus('Config unavailable.', true);
    return;
  }
  const data = await res.json();
  lastConfig = data;
  const llm = data.llm || {};
  llmEnabledToggle.checked = !!llm.enabled;
  if (llmProviderInput) llmProviderInput.value = llm.provider || 'ollama';
  if (llmEndpointInput) llmEndpointInput.value = llm.endpoint || '';
  if (llmAllowRemoteToggle) llmAllowRemoteToggle.checked = !!llm.allowNonLoopbackEndpoint;
  if (llmModelInput) llmModelInput.value = llm.model || '';
  if (llmTimeoutInput) llmTimeoutInput.value = llm.timeoutSeconds || 10;
  if (llmMaxTokensInput) llmMaxTokensInput.value = llm.maxTokens || 128;
  if (auditLlmInteractionsToggle) auditLlmInteractionsToggle.checked = data.auditLlmInteractions !== false;
  if (auditLlmIncludeRawTextToggle) auditLlmIncludeRawTextToggle.checked = !!data.auditLlmIncludeRawText;
  if (profileModeEnabledToggle) profileModeEnabledToggle.checked = !!data.profileModeEnabled;
  if (activeProfileInput) activeProfileInput.value = data.activeProfile || 'balanced';
  if (requireConfirmationToggle) requireConfirmationToggle.checked = !!data.requireConfirmation;
  if (quizSafeModeToggle) quizSafeModeToggle.checked = !!data.quizSafeModeEnabled;
  if (maxActionsInput) maxActionsInput.value = data.maxActionsPerSecond || 3;
  if (findRetryCountInput) findRetryCountInput.value = data.findRetryCount ?? 2;
  if (findRetryDelayMsInput) findRetryDelayMsInput.value = data.findRetryDelayMs ?? 250;
  if (postCheckTimeoutMsInput) postCheckTimeoutMsInput.value = data.postCheckTimeoutMs ?? 900;
  if (postCheckPollMsInput) postCheckPollMsInput.value = data.postCheckPollMs ?? 120;
  if (clipboardHistoryMaxItemsInput) clipboardHistoryMaxItemsInput.value = data.clipboardHistoryMaxItems ?? 50;
  if (filesystemAllowedRootsInput) filesystemAllowedRootsInput.value = (data.filesystemAllowedRoots || []).join('\n');
  if (allowedAppsInput) allowedAppsInput.value = (data.allowedApps || []).join('\n');
  if (appAliasesInput) appAliasesInput.value = formatAliases(data.appAliases || {});
  if (ocrEnabledToggle) ocrEnabledToggle.checked = !!data.ocrEnabled;
  if (ocrEngineInput) ocrEngineInput.value = (data.ocr && data.ocr.engine) ? data.ocr.engine : 'tesseract';
  if (ocrTesseractInput) ocrTesseractInput.value = (data.ocr && data.ocr.tesseractPath) ? data.ocr.tesseractPath : 'tesseract';
  if (adapterRestartCommandInput) adapterRestartCommandInput.value = data.adapterRestartCommand || '';
  if (adapterRestartDirInput) adapterRestartDirInput.value = data.adapterRestartWorkingDir || '';
  if (data.ocrRestartRequired !== undefined) {
    setOcrRestartBadge(!!data.ocrRestartRequired);
  }
  setConfigStatus('Config loaded.');
}

async function saveConfig() {
  if (!llmEnabledToggle) return;
  const payload = {
    allowedApps: allowedAppsInput ? parseLines(allowedAppsInput.value) : null,
    appAliases: appAliasesInput ? parseAliases(appAliasesInput.value) : null,
    profileModeEnabled: profileModeEnabledToggle ? profileModeEnabledToggle.checked : null,
    activeProfile: activeProfileInput ? activeProfileInput.value.trim() : null,
    requireConfirmation: requireConfirmationToggle ? requireConfirmationToggle.checked : null,
    maxActionsPerSecond: maxActionsInput ? parseIntOrNull(maxActionsInput.value) : null,
    findRetryCount: findRetryCountInput ? parseIntOrNull(findRetryCountInput.value) : null,
    findRetryDelayMs: findRetryDelayMsInput ? parseIntOrNull(findRetryDelayMsInput.value) : null,
    postCheckTimeoutMs: postCheckTimeoutMsInput ? parseIntOrNull(postCheckTimeoutMsInput.value) : null,
    postCheckPollMs: postCheckPollMsInput ? parseIntOrNull(postCheckPollMsInput.value) : null,
    clipboardHistoryMaxItems: clipboardHistoryMaxItemsInput ? parseIntOrNull(clipboardHistoryMaxItemsInput.value) : null,
    filesystemAllowedRoots: filesystemAllowedRootsInput ? parseLines(filesystemAllowedRootsInput.value) : null,
    quizSafeModeEnabled: quizSafeModeToggle ? quizSafeModeToggle.checked : null,
    ocrEnabled: ocrEnabledToggle ? ocrEnabledToggle.checked : null,
    ocr: {
      engine: ocrEngineInput ? ocrEngineInput.value.trim() : null,
      tesseractPath: ocrTesseractInput ? ocrTesseractInput.value.trim() : null
    },
    adapterRestartCommand: adapterRestartCommandInput ? adapterRestartCommandInput.value.trim() : null,
    adapterRestartWorkingDir: adapterRestartDirInput ? adapterRestartDirInput.value.trim() : null,
    llm: {
      enabled: llmEnabledToggle.checked,
      allowNonLoopbackEndpoint: llmAllowRemoteToggle ? llmAllowRemoteToggle.checked : null,
      provider: llmProviderInput ? llmProviderInput.value.trim() : null,
      endpoint: llmEndpointInput ? llmEndpointInput.value.trim() : null,
      model: llmModelInput ? llmModelInput.value.trim() : null,
      timeoutSeconds: llmTimeoutInput ? parseIntOrNull(llmTimeoutInput.value) : null,
      maxTokens: llmMaxTokensInput ? parseIntOrNull(llmMaxTokensInput.value) : null
    },
    auditLlmInteractions: auditLlmInteractionsToggle ? auditLlmInteractionsToggle.checked : null,
    auditLlmIncludeRawText: auditLlmIncludeRawTextToggle ? auditLlmIncludeRawTextToggle.checked : null
  };

  if (!payload.ocr.engine) {
    payload.ocr.engine = null;
  }
  if (!payload.ocr.tesseractPath) {
    payload.ocr.tesseractPath = null;
  }

  const res = await fetch('/api/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!res.ok) {
    let msg = 'Failed to save config.';
    try {
      const err = await res.json();
      if (err.message) msg = err.message;
    } catch {
      // ignore
    }
    setConfigStatus(msg, true);
    return;
  }

  const updated = await res.json();
  const ocrChanged = hasOcrChanged(lastConfig, updated);
  lastConfig = updated;
  const restartRequired = updated.ocrRestartRequired !== undefined
    ? !!updated.ocrRestartRequired
    : ocrChanged;
  setConfigStatus(restartRequired ? 'Config saved. OCR changes require restart.' : 'Config saved.');
  setOcrRestartBadge(restartRequired);
  loadUtilitiesStatus();
  fetchStatus();
}

function setUtilityStatus(message, isError) {
  if (!utilityStatus) return;
  utilityStatus.textContent = message;
  utilityStatus.classList.toggle('error', !!isError);
}

function formatToolStatus(tool) {
  if (!tool) return 'Unknown';
  if (!tool.installed) return `Not installed (${tool.status || 'missing'})`;
  const version = tool.version ? ` | ${tool.version}` : '';
  return `Installed${version}`;
}

async function loadUtilitiesStatus() {
  if (!ffmpegToolStatus || !tesseractToolStatus) return;
  try {
    const res = await fetch('/api/utilities/status');
    if (!res.ok) {
      throw new Error('Utility status unavailable.');
    }
    const data = await res.json();
    ffmpegToolStatus.textContent = formatToolStatus(data.ffmpeg);
    tesseractToolStatus.textContent = formatToolStatus(data.tesseract);
  } catch {
    ffmpegToolStatus.textContent = 'Unavailable';
    tesseractToolStatus.textContent = 'Unavailable';
    setUtilityStatus('Unable to load utility status.', true);
  }
}

async function installUtility(tool, enableOcr) {
  setUtilityStatus(`Installing ${tool}...`);
  try {
    const res = await fetch('/api/utilities/install', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tool,
        enableOcr: !!enableOcr
      })
    });

    let payload = null;
    try {
      payload = await res.json();
    } catch {
      payload = null;
    }

    if (!res.ok) {
      setUtilityStatus(payload?.message || `Install ${tool} failed.`, true);
      await loadUtilitiesStatus();
      return;
    }

    setUtilityStatus(payload?.message || `${tool} installed.`);
    if (payload?.ocrRestartRequired) {
      setOcrRestartBadge(true);
    }
    await loadUtilitiesStatus();
    await loadConfig();
  } catch {
    setUtilityStatus(`Install ${tool} failed.`, true);
  }
}

async function enableOcrOnly() {
  const tesseractPath = ocrTesseractInput ? ocrTesseractInput.value.trim() : '';
  setUtilityStatus('Enabling OCR...');
  try {
    const res = await fetch('/api/utilities/enable-ocr', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tesseractPath: tesseractPath || null
      })
    });
    let payload = null;
    try {
      payload = await res.json();
    } catch {
      payload = null;
    }
    if (!res.ok) {
      setUtilityStatus(payload?.message || 'Enable OCR failed.', true);
      return;
    }

    setUtilityStatus(payload?.message || 'OCR enabled.');
    if (payload?.ocrRestartRequired) {
      setOcrRestartBadge(true);
    }
    await loadConfig();
    await loadUtilitiesStatus();
  } catch {
    setUtilityStatus('Enable OCR failed.', true);
  }
}

async function loadTasks() {
  if (!taskList) return;
  const res = await fetch('/api/tasks');
  if (!res.ok) {
    setTaskStatus('Task library unavailable.', true);
    return;
  }
  const data = await res.json();
  taskItems = Array.isArray(data.tasks) ? data.tasks : [];
  renderTaskList();
}

async function saveTask() {
  const name = taskNameInput ? taskNameInput.value.trim() : '';
  const intent = taskIntentInput ? taskIntentInput.value.trim() : '';
  const description = taskDescriptionInput ? taskDescriptionInput.value.trim() : '';
  if (!name || !intent) {
    setTaskStatus('Task name and intent are required.', true);
    return;
  }

  const res = await fetch('/api/tasks', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, intent, description: description || null })
  });

  if (!res.ok) {
    setTaskStatus('Failed to save task.', true);
    return;
  }

  setTaskStatus('Task saved.');
  loadTasks();
}

async function runTask(name, dryRun) {
  if (!name) {
    setTaskStatus('Select a task first.', true);
    return;
  }

  setTaskStatus(dryRun ? 'Running task (dry-run)...' : 'Running task...');
  const res = await fetch(`/api/tasks/${encodeURIComponent(name)}/run`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ dryRun: !!dryRun })
  });
  if (!res.ok) {
    setTaskStatus('Task execution failed.', true);
    return;
  }

  const data = await res.json();
  if (data.reply) {
    appendMessage(`[Task:${name}] ${data.reply}`, 'bot');
  }
  if (data.modeLabel) {
    appendMode(data.modeLabel);
  }
  if (data.steps && data.steps.length) {
    appendSteps(data.steps);
  }
  if (data.planJson) {
    appendPlan(data.planJson);
  }
  if (data.needsConfirmation) {
    showConfirm(data.token, data.reply, data.planJson || null);
  }
  setTaskStatus('Task executed.');
  fetchStatus();
}

async function deleteTask(name) {
  if (!name) return;
  const ok = window.confirm(`Delete task "${name}"?`);
  if (!ok) return;
  const res = await fetch(`/api/tasks/${encodeURIComponent(name)}`, { method: 'DELETE' });
  if (!res.ok) {
    setTaskStatus('Failed to delete task.', true);
    return;
  }
  setTaskStatus('Task deleted.');
  if (taskNameInput && taskNameInput.value.trim() === name) {
    taskNameInput.value = '';
  }
  loadTasks();
}

function setTaskStatus(message, isError) {
  if (!taskStatus) return;
  taskStatus.textContent = message;
  taskStatus.classList.toggle('error', !!isError);
}

function setRecordingStatus(message, isError) {
  if (!recordingStatus) return;
  recordingStatus.textContent = message;
  recordingStatus.classList.toggle('error', !!isError);
}

function setRecordingStatusFromData(recording) {
  if (!recordingStatus || !recording) return;
  const state = recording.isRecording ? 'ON' : 'OFF';
  const label = recording.name ? ` (${recording.name})` : '';
  const steps = Number.isFinite(recording.capturedSteps) ? recording.capturedSteps : 0;
  setRecordingStatus(`Recording ${state}${label} - steps: ${steps}`, false);
}

async function refreshRecordingStatus() {
  if (!recordingStatus) return;
  try {
    const res = await fetch('/api/recording/status');
    if (!res.ok) {
      setRecordingStatus('Recording status unavailable.', true);
      return;
    }
    const data = await res.json();
    setRecordingStatusFromData(data);
  } catch {
    setRecordingStatus('Recording status unavailable.', true);
  }
}

async function startRecording() {
  const name = recordNameInput ? recordNameInput.value.trim() : '';
  setRecordingStatus('Starting recording...');
  const res = await fetch('/api/recording/start', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: name || null })
  });
  if (!res.ok) {
    setRecordingStatus('Failed to start recording.', true);
    return;
  }
  const data = await res.json();
  setRecordingStatusFromData(data);
}

async function stopRecording() {
  setRecordingStatus('Stopping recording...');
  const res = await fetch('/api/recording/stop', { method: 'POST' });
  if (!res.ok) {
    setRecordingStatus('Failed to stop recording.', true);
    return;
  }
  const data = await res.json();
  if (data.status) {
    setRecordingStatusFromData(data.status);
  } else {
    refreshRecordingStatus();
  }
  if (data.planJson) {
    appendPlan(data.planJson);
  }
}

async function saveRecording() {
  const name = (recordNameInput && recordNameInput.value.trim())
    || (taskNameInput && taskNameInput.value.trim())
    || '';
  if (!name) {
    setRecordingStatus('Set a recording name first.', true);
    return;
  }
  setRecordingStatus('Saving recording...');
  const description = taskDescriptionInput ? taskDescriptionInput.value.trim() : '';
  const res = await fetch('/api/recording/save', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name,
      description: description || null
    })
  });
  if (!res.ok) {
    let msg = 'Failed to save recording.';
    try {
      const err = await res.json();
      if (err.message) msg = err.message;
    } catch {
      // ignore
    }
    setRecordingStatus(msg, true);
    return;
  }
  const data = await res.json();
  if (taskNameInput && !taskNameInput.value.trim()) {
    taskNameInput.value = name;
  }
  setRecordingStatus(data.message || `Recording saved as ${name}.`, false);
  loadTasks();
}

function renderTaskList() {
  if (!taskList) return;
  taskList.innerHTML = '';
  if (taskItems.length === 0) {
    taskList.textContent = 'No tasks yet.';
    return;
  }

  const fragment = document.createDocumentFragment();
  taskItems.forEach(item => {
    const row = document.createElement('div');
    row.className = 'task-row';

    const info = document.createElement('div');
    info.className = 'task-info';
    const title = document.createElement('strong');
    title.textContent = item.name;
    const intent = document.createElement('div');
    intent.className = 'task-intent';
    intent.textContent = item.intent;
    info.appendChild(title);
    info.appendChild(intent);

    const actions = document.createElement('div');
    actions.className = 'task-actions-inline';
    const fillBtn = document.createElement('button');
    fillBtn.textContent = 'Use';
    fillBtn.type = 'button';
    fillBtn.onclick = () => {
      if (taskNameInput) taskNameInput.value = item.name;
      if (taskIntentInput) taskIntentInput.value = item.intent;
      if (taskDescriptionInput) taskDescriptionInput.value = item.description || '';
    };

    const runBtn = document.createElement('button');
    runBtn.textContent = 'Run';
    runBtn.type = 'button';
    runBtn.onclick = () => runTask(item.name, false);

    const dryRunBtn = document.createElement('button');
    dryRunBtn.textContent = 'Dry-run';
    dryRunBtn.type = 'button';
    dryRunBtn.onclick = () => runTask(item.name, true);

    const deleteBtn = document.createElement('button');
    deleteBtn.textContent = 'Delete';
    deleteBtn.type = 'button';
    deleteBtn.className = 'danger-outline';
    deleteBtn.onclick = () => deleteTask(item.name);

    actions.appendChild(fillBtn);
    actions.appendChild(runBtn);
    actions.appendChild(dryRunBtn);
    actions.appendChild(deleteBtn);

    row.appendChild(info);
    row.appendChild(actions);
    fragment.appendChild(row);
  });

  taskList.appendChild(fragment);
}

async function loadSchedules() {
  if (!scheduleList) return;
  const res = await fetch('/api/schedules');
  if (!res.ok) {
    setScheduleStatus('Scheduler unavailable.', true);
    return;
  }

  const data = await res.json();
  scheduleItems = Array.isArray(data.schedules) ? data.schedules : [];
  renderScheduleList();
}

async function saveSchedule() {
  const taskName = scheduleTaskNameInput ? scheduleTaskNameInput.value.trim() : '';
  if (!taskName) {
    setScheduleStatus('Task name is required.', true);
    return;
  }

  const id = scheduleIdInput ? scheduleIdInput.value.trim() : '';
  const startAtRaw = scheduleStartAtInput ? scheduleStartAtInput.value.trim() : '';
  const intervalSeconds = scheduleIntervalSecondsInput ? parseIntOrNull(scheduleIntervalSecondsInput.value) : null;
  const payload = {
    id: id || null,
    taskName,
    startAtUtc: startAtRaw ? new Date(startAtRaw).toISOString() : null,
    intervalSeconds,
    enabled: scheduleEnabledToggle ? !!scheduleEnabledToggle.checked : true
  };

  const res = await fetch('/api/schedules', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!res.ok) {
    let msg = 'Failed to save schedule.';
    try {
      const err = await res.json();
      if (err.message) msg = err.message;
    } catch {
      // ignore
    }
    setScheduleStatus(msg, true);
    return;
  }

  const data = await res.json();
  setScheduleStatus(data.message || 'Schedule saved.');
  await loadSchedules();
}

async function runScheduleNow() {
  const id = scheduleIdInput ? scheduleIdInput.value.trim() : '';
  if (!id) {
    setScheduleStatus('Select a schedule first.', true);
    return;
  }

  const res = await fetch(`/api/schedules/${encodeURIComponent(id)}/run`, { method: 'POST' });
  if (!res.ok) {
    let msg = 'Failed to run schedule.';
    try {
      const err = await res.json();
      if (err.message) msg = err.message;
    } catch {
      // ignore
    }
    setScheduleStatus(msg, true);
    return;
  }

  const data = await res.json();
  setScheduleStatus(data.message || 'Schedule triggered.');
  await loadSchedules();
  fetchAudit();
}

async function deleteSchedule(id) {
  if (!id) return;
  const ok = window.confirm(`Delete schedule "${id}"?`);
  if (!ok) return;

  const res = await fetch(`/api/schedules/${encodeURIComponent(id)}`, { method: 'DELETE' });
  if (!res.ok) {
    setScheduleStatus('Failed to delete schedule.', true);
    return;
  }

  if (scheduleIdInput && scheduleIdInput.value.trim() === id) {
    scheduleIdInput.value = '';
  }
  setScheduleStatus('Schedule deleted.');
  await loadSchedules();
}

function renderScheduleList() {
  if (!scheduleList) return;
  scheduleList.innerHTML = '';
  if (scheduleItems.length === 0) {
    scheduleList.textContent = 'No schedules yet.';
    return;
  }

  const fragment = document.createDocumentFragment();
  scheduleItems.forEach(item => {
    const row = document.createElement('div');
    row.className = 'task-row';

    const info = document.createElement('div');
    info.className = 'task-info';
    const title = document.createElement('strong');
    const intervalLabel = item.intervalSeconds ? `every ${item.intervalSeconds}s` : 'one-shot';
    title.textContent = `${item.taskName} (${intervalLabel})`;
    const meta = document.createElement('div');
    meta.className = 'task-intent';
    const next = item.nextRunAtUtc ? new Date(item.nextRunAtUtc).toLocaleString() : 'none';
    const last = item.lastRunAtUtc ? new Date(item.lastRunAtUtc).toLocaleString() : 'never';
    const status = item.lastSuccess === true ? 'ok' : item.lastSuccess === false ? 'failed' : 'n/a';
    meta.textContent = `id=${item.id} | enabled=${item.enabled} | next=${next} | last=${last} (${status})`;
    info.appendChild(title);
    info.appendChild(meta);

    const actions = document.createElement('div');
    actions.className = 'task-actions-inline';

    const useBtn = document.createElement('button');
    useBtn.type = 'button';
    useBtn.textContent = 'Use';
    useBtn.onclick = () => fillScheduleForm(item);

    const runBtn = document.createElement('button');
    runBtn.type = 'button';
    runBtn.textContent = 'Run now';
    runBtn.onclick = () => {
      if (scheduleIdInput) scheduleIdInput.value = item.id;
      runScheduleNow();
    };

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.textContent = 'Delete';
    deleteBtn.className = 'danger-outline';
    deleteBtn.onclick = () => deleteSchedule(item.id);

    actions.appendChild(useBtn);
    actions.appendChild(runBtn);
    actions.appendChild(deleteBtn);
    row.appendChild(info);
    row.appendChild(actions);
    fragment.appendChild(row);
  });

  scheduleList.appendChild(fragment);
}

function fillScheduleForm(item) {
  if (scheduleIdInput) scheduleIdInput.value = item.id || '';
  if (scheduleTaskNameInput) scheduleTaskNameInput.value = item.taskName || '';
  if (scheduleStartAtInput) scheduleStartAtInput.value = toLocalDateTimeInput(item.startAtUtc);
  if (scheduleIntervalSecondsInput) scheduleIntervalSecondsInput.value = item.intervalSeconds || '';
  if (scheduleEnabledToggle) scheduleEnabledToggle.checked = item.enabled !== false;
}

function toLocalDateTimeInput(value) {
  if (!value) return '';
  const dt = new Date(value);
  if (Number.isNaN(dt.getTime())) return '';
  const local = new Date(dt.getTime() - dt.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 16);
}

function setScheduleStatus(message, isError) {
  if (!scheduleStatus) return;
  scheduleStatus.textContent = message;
  scheduleStatus.classList.toggle('error', !!isError);
}

async function refreshInspector() {
  if (!inspectorOutput) return;
  const query = inspectorQueryInput ? inspectorQueryInput.value.trim() : '';
  setInspectorStatus('Loading inspector...');
  const queryString = query ? `?query=${encodeURIComponent(query)}` : '';
  const res = await fetch(`/api/inspector${queryString}`);
  if (!res.ok) {
    setInspectorStatus('Inspector unavailable.', true);
    return;
  }

  const data = await res.json();
  inspectorOutput.textContent = JSON.stringify(data, null, 2);
  setInspectorStatus('Inspector updated.');
}

function setInspectorStatus(message, isError) {
  if (!inspectorStatus) return;
  inspectorStatus.textContent = message;
  inspectorStatus.classList.toggle('error', !!isError);
}

function setConfigStatus(message, isError) {
  if (!configStatus) return;
  configStatus.textContent = message;
  configStatus.classList.toggle('error', !!isError);
}

function setOcrRestartBadge(show) {
  if (!ocrRestartBadge) return;
  if (!show) {
    ocrAcked = false;
    sessionStorage.removeItem(ocrAckKey);
  }
  const visible = show && !ocrAcked;
  ocrRestartBadge.classList.toggle('hidden', !visible);
  if (ocrRestartAck) {
    ocrRestartAck.classList.toggle('hidden', !visible);
  }
}

function hasOcrChanged(prev, next) {
  if (!prev || !next) return false;
  const prevOcr = prev.ocr || {};
  const nextOcr = next.ocr || {};
  if ((prev.ocrEnabled ?? false) !== (next.ocrEnabled ?? false)) return true;
  if ((prevOcr.engine || '') !== (nextOcr.engine || '')) return true;
  if ((prevOcr.tesseractPath || '') !== (nextOcr.tesseractPath || '')) return true;
  return false;
}

function parseLines(value) {
  return value
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.length > 0);
}

function parseIntOrNull(value) {
  const parsed = parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseAliases(value) {
  const result = {};
  value.split('\n').forEach(raw => {
    const line = raw.trim();
    if (!line) return;
    const parts = line.split('=');
    if (parts.length < 2) return;
    const key = parts.shift().trim();
    const target = parts.join('=').trim();
    if (!key || !target) return;
    result[key] = target;
  });
  return result;
}

function formatAliases(map) {
  const keys = Object.keys(map).sort((a, b) => a.localeCompare(b));
  return keys.map(key => `${key}=${map[key]}`).join('\n');
}

async function fetchAudit() {
  const res = await fetch('/api/audit?take=120');
  if (!res.ok) {
    auditLog.textContent = 'Audit log unavailable.';
    return;
  }
  const data = await res.json();
  auditItems = parseAuditLines(data.lines || []);
  renderAudit();
  if (auditInitialized) {
    const delta = Math.max(0, auditItems.length - lastAuditSize);
    notifyAudit(delta);
  } else {
    auditInitialized = true;
  }
  lastAuditSize = auditItems.length;
}

function appendMessage(text, cls) {
  renderEntry({ kind: cls, text }, true);
}

function appendSteps(lines) {
  renderEntry({ kind: 'steps', text: lines.join('\n') }, true);
}

function appendMode(text) {
  renderEntry({ kind: 'mode', text }, true);
}

function appendPlan(planJson) {
  renderEntry({ kind: 'plan', text: planJson }, true);
}

function renderStatusBadges(llm, adapter, killSwitch) {
  if (!statusBadges) return;
  statusBadges.innerHTML = '';
  if (adapter) {
    const armedBadge = document.createElement('span');
    armedBadge.className = `badge ${adapter.armed ? 'ok' : 'muted'}`;
    armedBadge.textContent = `${adapter.armed ? '●' : '○'} Armed`;
    statusBadges.appendChild(armedBadge);

    const presenceBadge = document.createElement('span');
    presenceBadge.className = `badge ${adapter.requireUserPresence ? 'ok' : 'muted'}`;
    presenceBadge.textContent = `${adapter.requireUserPresence ? '●' : '○'} Presence`;
    statusBadges.appendChild(presenceBadge);
  }

  if (killSwitch) {
    const killBadge = document.createElement('span');
    const tripped = !!killSwitch.tripped;
    killBadge.className = `badge ${tripped ? 'warn' : 'muted'}`;
    killBadge.textContent = `${tripped ? '●' : '○'} Kill`;
    if (tripped && killSwitch.reason) {
      killBadge.title = killSwitch.reason;
    }
    statusBadges.appendChild(killBadge);
  }

  if (!llm) return;
  const enabled = !!llm.enabled;
  const available = !!llm.available;
  const provider = llm.provider || 'local';
  let label = 'LLM: Disabled (config)';
  let cls = 'muted';
  if (enabled) {
    if (available) {
      label = `LLM: ${provider}`;
      cls = 'ok';
    } else {
      label = `LLM: Offline`;
      cls = 'warn';
    }
  }
  const badge = document.createElement('span');
  badge.className = `badge ${cls}`;
  badge.textContent = label;
  const tooltipParts = [];
  if (llm.message) tooltipParts.push(llm.message);
  if (llm.endpoint) tooltipParts.push(llm.endpoint);
  if (tooltipParts.length) badge.title = tooltipParts.join(' | ');
  statusBadges.appendChild(badge);
}

async function sendChat(message) {
  appendMessage(message, 'user');
  const res = await fetch('/api/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ message })
  });
  const data = await res.json();
  appendMessage(data.reply, 'bot');
  if (data.modeLabel) {
    appendMode(data.modeLabel);
  }
  if (data.steps && data.steps.length) {
    appendSteps(data.steps);
  }
  if (data.planJson) {
    appendPlan(data.planJson);
  }
  if (data.needsConfirmation) {
    showConfirm(data.token, data.reply, data.planJson || null);
  } else {
    clearConfirm();
  }
  fetchStatus();
}

function resizeChatInput() {
  if (!chatInput) return;
  chatInput.style.height = 'auto';
  const maxHeight = 180;
  const nextHeight = Math.min(chatInput.scrollHeight, maxHeight);
  chatInput.style.height = `${Math.max(40, nextHeight)}px`;
  chatInput.style.overflowY = chatInput.scrollHeight > maxHeight ? 'auto' : 'hidden';
}

function showConfirm(token, message, planJson) {
  confirmArea.innerHTML = '';
  const text = document.createElement('div');
  text.textContent = message;
  confirmArea.appendChild(text);

  let editor = null;
  if (planJson) {
    const label = document.createElement('div');
    label.textContent = 'Edit plan before execution (optional):';
    label.className = 'confirm-plan-label';
    editor = document.createElement('textarea');
    editor.className = 'confirm-plan-editor';
    editor.value = planJson;
    editor.rows = 12;
    confirmArea.appendChild(label);
    confirmArea.appendChild(editor);
  }

  const approve = document.createElement('button');
  approve.textContent = editor ? 'Confirm edited plan' : 'Confirm';
  const reject = document.createElement('button');
  reject.textContent = 'Cancel';
  approve.onclick = () => confirmAction(token, true, editor ? editor.value : null);
  reject.onclick = () => confirmAction(token, false, null);
  confirmArea.appendChild(approve);
  confirmArea.appendChild(reject);
}

function clearConfirm() {
  confirmArea.innerHTML = '';
}

async function confirmAction(token, approve, planJson) {
  const res = await fetch('/api/confirm', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ token, approve, planJson: approve ? planJson : null })
  });
  if (!res.ok) {
    let msg = 'Confirmation failed.';
    try {
      const err = await res.json();
      if (err.message) msg = err.message;
    } catch {
      // ignore
    }
    appendMessage(msg, 'bot');
    return;
  }
  const data = await res.json();
  appendMessage(data.reply, 'bot');
  if (data.modeLabel) {
    appendMode(data.modeLabel);
  }
  if (data.steps && data.steps.length) {
    appendSteps(data.steps);
  }
  if (data.planJson) {
    appendPlan(data.planJson);
  }
  clearConfirm();
  fetchStatus();
}

chatForm.addEventListener('submit', (e) => {
  e.preventDefault();
  const value = chatInput.value.trim();
  if (!value) return;
  chatInput.value = '';
  resizeChatInput();
  sendChat(value);
});

if (chatInput) {
  chatInput.addEventListener('input', () => resizeChatInput());
  chatInput.addEventListener('keydown', (e) => {
    if (e.isComposing) return;
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      chatForm.requestSubmit();
    }
  });
}

tabButtons.forEach(btn => {
  btn.addEventListener('click', () => setActiveTab(btn.dataset.tabTarget || 'chat'));
});

for (const btn of document.querySelectorAll('[data-cmd]')) {
  btn.addEventListener('click', () => sendChat(btn.dataset.cmd));
}

filterCheckboxes.forEach(cb => cb.addEventListener('change', () => {
  saveFilters();
  renderAudit();
}));

if (clearHistoryBtn) {
  clearHistoryBtn.addEventListener('click', () => {
    chatHistory = [];
    saveHistory();
    renderHistory();
  });
}

if (saveConfigBtn) {
  saveConfigBtn.addEventListener('click', () => saveConfig());
}

if (reloadConfigBtn) {
  reloadConfigBtn.addEventListener('click', () => loadConfig());
}

if (restartServiceBtn) {
  restartServiceBtn.addEventListener('click', async () => {
    const ok = window.confirm('Restart the server now? This will close the UI until you start it again.');
    if (!ok) return;
    setConfigStatus('Restarting server...');
    try {
      await fetch('/api/restart', { method: 'POST' });
    } catch {
      // ignore
    }
  });
}

if (restartAdapterBtn) {
  restartAdapterBtn.addEventListener('click', async () => {
    const ok = window.confirm('Restart the adapter now? Make sure the restart command is configured.');
    if (!ok) return;
    setConfigStatus('Restarting adapter...');
    try {
      const res = await fetch('/api/adapter/restart', { method: 'POST' });
      if (!res.ok) {
        let msg = 'Adapter restart failed.';
        try {
          const err = await res.json();
          if (err.message) msg = err.message;
        } catch {
          // ignore
        }
        setConfigStatus(msg, true);
        return;
      }
      setConfigStatus('Adapter restart command launched.');
    } catch {
      setConfigStatus('Adapter restart failed.', true);
    }
  });
}

if (utilityRefreshBtn) {
  utilityRefreshBtn.addEventListener('click', () => loadUtilitiesStatus());
}

if (installFfmpegBtn) {
  installFfmpegBtn.addEventListener('click', () => installUtility('ffmpeg', false));
}

if (installOcrBtn) {
  installOcrBtn.addEventListener('click', () => installUtility('ocr', false));
}

if (installOcrEnableBtn) {
  installOcrEnableBtn.addEventListener('click', () => installUtility('ocr', true));
}

if (enableOcrOnlyBtn) {
  enableOcrOnlyBtn.addEventListener('click', () => enableOcrOnly());
}

if (taskSaveBtn) {
  taskSaveBtn.addEventListener('click', () => saveTask());
}

if (taskRunBtn) {
  taskRunBtn.addEventListener('click', () => {
    const name = taskNameInput ? taskNameInput.value.trim() : '';
    runTask(name, false);
  });
}

if (taskRefreshBtn) {
  taskRefreshBtn.addEventListener('click', () => loadTasks());
}

if (scheduleSaveBtn) {
  scheduleSaveBtn.addEventListener('click', () => saveSchedule());
}

if (scheduleRunNowBtn) {
  scheduleRunNowBtn.addEventListener('click', () => runScheduleNow());
}

if (scheduleRefreshBtn) {
  scheduleRefreshBtn.addEventListener('click', () => loadSchedules());
}

if (recordStartBtn) {
  recordStartBtn.addEventListener('click', () => startRecording());
}

if (recordStopBtn) {
  recordStopBtn.addEventListener('click', () => stopRecording());
}

if (recordSaveBtn) {
  recordSaveBtn.addEventListener('click', () => saveRecording());
}

if (inspectorRefreshBtn) {
  inspectorRefreshBtn.addEventListener('click', () => refreshInspector());
}

if (ocrRestartAck) {
  ocrRestartAck.addEventListener('click', () => {
    ocrAcked = true;
    sessionStorage.setItem(ocrAckKey, '1');
    setOcrRestartBadge(true);
  });
}

fetchStatus();
setInterval(fetchStatus, 5000);
loadConfig();
loadUtilitiesStatus();
loadTasks();
loadSchedules();
refreshRecordingStatus();
loadFilters();
fetchAudit();
startAuditStream();
refreshInspector();
chatHistory = loadHistory();
renderHistory();
setActiveTab(localStorage.getItem(tabKey) || 'chat');
resizeChatInput();

function parseAuditLines(lines) {
  return lines.map(line => {
    try {
      const obj = JSON.parse(line);
      const type = (obj.EventType || obj.eventType || '').toString();
      return { raw: line, type };
    } catch {
      return { raw: line, type: 'unknown' };
    }
  });
}

function renderAudit() {
  const filters = getFilters();
  const output = auditItems.filter(item => auditVisible(item, filters)).map(item => item.raw);
  auditLog.textContent = output.join('\n');
}

function auditVisible(item, filters) {
  if (item.type === 'unknown') {
    return filters.system;
  }
  if (filterMap.action.has(item.type)) return filters.action;
  if (filterMap.policy.has(item.type)) return filters.policy;
  return filters.system;
}

function getFilters() {
  const state = { action: true, policy: true, system: true };
  filterCheckboxes.forEach(cb => {
    state[cb.dataset.filter] = cb.checked;
  });
  return state;
}

function saveFilters() {
  const state = getFilters();
  localStorage.setItem(filterKey, JSON.stringify(state));
}

function loadFilters() {
  try {
    const raw = localStorage.getItem(filterKey);
    if (!raw) return;
    const state = JSON.parse(raw);
    filterCheckboxes.forEach(cb => {
      if (state[cb.dataset.filter] !== undefined) {
        cb.checked = state[cb.dataset.filter];
      }
    });
  } catch {
    // ignore
  }
}

function renderEntry(entry, persist) {
  if (persist) {
    chatHistory.push(entry);
    if (chatHistory.length > 200) {
      chatHistory = chatHistory.slice(-200);
    }
    saveHistory();
  }

  let node;
  if (entry.kind === 'mode') {
    const div = document.createElement('div');
    div.className = 'msg mode';
    div.textContent = entry.text;
    node = div;
  } else if (entry.kind === 'steps') {
    const div = document.createElement('div');
    div.className = 'msg steps';
    const list = document.createElement('ul');
    list.className = 'step-list';
    entry.text.split('\n').forEach(line => {
      const trimmed = line.trim();
      if (!trimmed) return;
      const li = document.createElement('li');
      li.className = 'step-line';
      const { label, className, icon, rest } = parseStepLine(trimmed);
      if (label) {
        const badge = document.createElement('span');
        badge.className = `step-badge${className ? ` step-${className}` : ''}`;
        badge.title = label;
        badge.innerHTML = iconSvg(icon || className || 'system') + `<span class=\"sr-only\">${label}</span>`;
        li.appendChild(badge);
      }
      const content = document.createElement('span');
      content.className = 'step-text';
      content.textContent = rest || trimmed;
      li.appendChild(content);
      list.appendChild(li);
    });
    div.appendChild(list);
    node = div;
  } else if (entry.kind === 'plan') {
    const details = document.createElement('details');
    details.className = 'msg plan';
    const summary = document.createElement('summary');
    summary.textContent = 'Plan JSON';
    const pre = document.createElement('pre');
    pre.textContent = entry.text;
    details.appendChild(summary);
    details.appendChild(pre);
    node = details;
  } else {
    const div = document.createElement('div');
    div.className = `msg ${entry.kind}`;
    div.textContent = entry.text;
    node = div;
  }

  chatLog.appendChild(node);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function setActiveTab(target) {
  const next = target || 'chat';
  tabButtons.forEach(btn => {
    const active = btn.dataset.tabTarget === next;
    btn.classList.toggle('active', active);
    btn.setAttribute('aria-selected', active ? 'true' : 'false');
  });
  tabPanels.forEach(panel => {
    panel.classList.toggle('active', panel.dataset.tabPanel === next);
  });
  if (next === 'tasks' && taskItems.length === 0) {
    loadTasks();
  }
  if (next === 'inspector') {
    refreshInspector();
  }
  if (next === 'audit') {
    clearAuditBadge();
  }
  localStorage.setItem(tabKey, next);
}

function parseStepLine(line) {
  const match = line.match(/^(?:\\d+\\.|\\[\\d+\\])\\s*([A-Za-z]+)\\s*(.*)$/);
  if (!match) {
    return { label: '', className: '', icon: 'system', rest: line };
  }
  const type = match[1];
  const remainder = (match[2] || '').trim();
  const info = mapStepLabel(type);
  return { label: info.label, className: info.className, icon: info.icon, rest: remainder };
}

function mapStepLabel(type) {
  const t = type.toLowerCase();
  const map = {
    openapp: { label: 'OPEN', className: 'app', icon: 'app' },
    click: { label: 'CLICK', className: 'action', icon: 'click' },
    typetext: { label: 'TYPE', className: 'input', icon: 'type' },
    find: { label: 'FIND', className: 'read', icon: 'find' },
    keycombo: { label: 'KEYS', className: 'key', icon: 'key' },
    readtext: { label: 'READ', className: 'read', icon: 'read' },
    invoke: { label: 'INVOKE', className: 'action', icon: 'action' },
    setvalue: { label: 'SET', className: 'action', icon: 'action' },
    waitfor: { label: 'WAIT', className: 'system', icon: 'wait' },
    capturescreen: { label: 'CAPTURE', className: 'read', icon: 'read' },
    getclipboard: { label: 'CLIP', className: 'read', icon: 'read' },
    setclipboard: { label: 'CLIP', className: 'action', icon: 'action' },
    clipboardhistory: { label: 'CLIP', className: 'read', icon: 'read' },
    openurl: { label: 'URL', className: 'app', icon: 'app' },
    filewrite: { label: 'FILE', className: 'action', icon: 'action' },
    fileappend: { label: 'FILE', className: 'action', icon: 'action' },
    fileread: { label: 'FILE', className: 'read', icon: 'read' },
    filelist: { label: 'FILE', className: 'read', icon: 'read' },
    notify: { label: 'NOTE', className: 'system', icon: 'system' },
    mousejiggle: { label: 'MOUSE', className: 'action', icon: 'action' }
  };
  return map[t] || { label: type.toUpperCase(), className: 'system', icon: 'system' };
}

function iconSvg(name) {
  switch (name) {
    case 'app':
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><rect x=\"3\" y=\"4\" width=\"18\" height=\"16\" rx=\"2\" ry=\"2\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/><line x1=\"3\" y1=\"9\" x2=\"21\" y2=\"9\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
    case 'click':
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><circle cx=\"9\" cy=\"9\" r=\"6\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/><line x1=\"14\" y1=\"14\" x2=\"21\" y2=\"21\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
    case 'type':
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><rect x=\"3\" y=\"5\" width=\"18\" height=\"14\" rx=\"2\" ry=\"2\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/><line x1=\"7\" y1=\"9\" x2=\"17\" y2=\"9\" stroke=\"currentColor\" stroke-width=\"2\"/><line x1=\"7\" y1=\"13\" x2=\"13\" y2=\"13\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
    case 'find':
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><circle cx=\"11\" cy=\"11\" r=\"6\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/><line x1=\"16\" y1=\"16\" x2=\"21\" y2=\"21\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
    case 'key':
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><circle cx=\"8\" cy=\"12\" r=\"3\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/><path d=\"M11 12h9v3h-3v3h-3v-3h-3z\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
    case 'read':
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><path d=\"M4 5h12a4 4 0 0 1 4 4v10H8a4 4 0 0 0-4 4V5z\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/><path d=\"M8 5v14\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
    case 'action':
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><path d=\"M12 3l9 5-9 5-9-5 9-5z\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/><path d=\"M3 13l9 5 9-5\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
    case 'wait':
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><circle cx=\"12\" cy=\"12\" r=\"9\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/><path d=\"M12 7v6l4 2\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
    default:
      return '<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><circle cx=\"12\" cy=\"12\" r=\"8\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"/></svg>';
  }
}

function saveHistory() {
  localStorage.setItem(historyKey, JSON.stringify(chatHistory));
}

function loadHistory() {
  try {
    const raw = localStorage.getItem(historyKey);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function renderHistory() {
  chatLog.innerHTML = '';
  chatHistory.forEach(entry => renderEntry(entry, false));
}

function startAuditStream() {
  if (!window.EventSource) {
    setInterval(fetchAudit, 5000);
    return;
  }
  const source = new EventSource('/api/audit/stream');
  source.onmessage = (event) => {
    if (event.data) {
      const items = parseAuditLines([event.data]);
      auditItems = auditItems.concat(items).slice(-200);
      renderAudit();
      notifyAudit(items.length);
    }
  };
  source.onerror = () => {
    source.close();
    setInterval(fetchAudit, 5000);
  };
}

function notifyAudit(count) {
  if (!auditInitialized || count <= 0) return;
  const active = localStorage.getItem(tabKey) || 'chat';
  if (active === 'audit') return;
  unseenAudit += count;
  updateAuditBadge();
}

function updateAuditBadge() {
  if (!auditBadge) return;
  if (unseenAudit <= 0) {
    auditBadge.classList.add('hidden');
    auditBadge.textContent = '0';
    return;
  }
  auditBadge.textContent = `${Math.min(unseenAudit, 99)}`;
  auditBadge.classList.remove('hidden');
}

function clearAuditBadge() {
  unseenAudit = 0;
  updateAuditBadge();
}
