
const elements = {
  appId: document.getElementById("appId"),
  appSecret: document.getElementById("appSecret"),
  redirectUri: document.getElementById("redirectUri"),
  scope: document.getElementById("scope"),
  authCode: document.getElementById("authCode"),
  tenantToken: document.getElementById("tenantToken"),
  userToken: document.getElementById("userToken"),
  sharedFolderToken: document.getElementById("sharedFolderToken"),
  log: document.getElementById("log"),
  result: document.getElementById("result"),
  treeRoot: document.getElementById("treeRoot"),
  selectionSummary: document.getElementById("selectionSummary"),
  migrationSteps: document.getElementById("migrationSteps"),
  nodeProgressBar: document.getElementById("nodeProgressBar"),
  nodeProgressText: document.getElementById("nodeProgressText"),
  moveProgressBar: document.getElementById("moveProgressBar"),
  moveProgressText: document.getElementById("moveProgressText"),
  btnQuickAuth: document.getElementById("btnQuickAuth"),
  btnAddFolder: document.getElementById("btnAddFolder"),
  btnMigrate: document.getElementById("btnMigrate"),
  btnStop: document.getElementById("btnStop"),
  btnClearSelection: document.getElementById("btnClearSelection"),
  btnClear: document.getElementById("btnClear"),
  btnRefreshTree: document.getElementById("btnRefreshTree"),
  btnCollapseAll: document.getElementById("btnCollapseAll"),
};

const OAUTH_STATE_KEY = "feishu.oauth.state";
const WIKI_SUPPORTED_TYPES = new Set([
  "doc",
  "docx",
  "sheet",
  "bitable",
  "mindnote",
]);
const MOVE_DOCS_BATCH_LIMIT = 90;
const MOVE_DOCS_PAUSE_MS = 60000;
const DEFAULT_SCOPE =
  "drive:drive:readonly drive:drive.metadata:readonly wiki:wiki space:folder:create docs:document:copy";
const WIKI_TASK_TYPE = "move_docs_to_wiki";

function nowStamp() {
  return new Date().toISOString();
}

function formatPayload(payload) {
  if (payload === undefined) {
    return "";
  }
  if (typeof payload === "string") {
    return payload;
  }
  return JSON.stringify(payload, null, 2);
}

function appendLog(title, payload) {
  const line = `[${nowStamp()}] ${title}`;
  const detail = payload !== undefined ? `\n${formatPayload(payload)}` : "";
  elements.log.textContent += `${line}${detail}\n\n`;
  elements.log.scrollTop = elements.log.scrollHeight;
}

function setResult(payload) {
  elements.result.textContent = formatPayload(payload);
}

function getValue(input) {
  return input.value.trim();
}

function requireValue(input, label) {
  const value = getValue(input);
  if (!value) {
    appendLog("缺少必填字段", { field: label });
    throw new Error(`${label} 为必填项`);
  }
  return value;
}

const treeState = {
  roots: [],
  rootTokens: new Set(),
  nodes: [],
  selectedRootToken: null,
  sharedTokens: new Set(),
};

const migrationState = {
  active: false,
  cancelled: false,
};

const authState = {
  active: false,
};

function setMigrationActive(active) {
  migrationState.active = active;
  if (active) {
    migrationState.cancelled = false;
  }
  if (elements.btnMigrate) {
    elements.btnMigrate.disabled = active;
  }
  if (elements.btnStop) {
    elements.btnStop.disabled = !active;
  }
}

function setAuthActive(active) {
  authState.active = active;
  if (elements.btnQuickAuth) {
    elements.btnQuickAuth.disabled = active;
  }
}

function ensureNotCancelled() {
  if (migrationState.cancelled) {
    throw new Error("用户已取消迁移。");
  }
}

async function sleepWithCancel(ms) {
  const start = Date.now();
  while (Date.now() - start < ms) {
    ensureNotCancelled();
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
}

function registerNode(node) {
  treeState.nodes.push(node);
  return node;
}

function createNode({ token, name, type, parent, raw }) {
  const normalizedType = String(type || "file").toLowerCase();
  return registerNode({
    token,
    name: name || token,
    type: normalizedType,
    isFolder:
      normalizedType === "folder" || normalizedType.endsWith(":folder"),
    expanded: false,
    loaded: false,
    element: null,
    childrenElement: null,
    toggleElement: null,
    checkboxElement: null,
    selected: false,
    indeterminate: false,
    parent: parent || null,
    children: [],
    raw: raw || null,
  });
}

function applySelectionState(node) {
  if (!node.checkboxElement) {
    return;
  }
  node.checkboxElement.checked = Boolean(node.selected);
  node.checkboxElement.indeterminate = Boolean(node.indeterminate);
}

function setNodeSelection(node, selected, options = {}) {
  node.selected = Boolean(selected);
  node.indeterminate = false;
  applySelectionState(node);
  if (options.propagate && node.children && node.children.length > 0) {
    node.children.forEach((child) =>
      setNodeSelection(child, selected, { propagate: true })
    );
  }
}

function updateParentSelection(node) {
  if (!node || !node.children || node.children.length === 0) {
    return;
  }
  const total = node.children.length;
  let selectedCount = 0;
  let indeterminateCount = 0;
  node.children.forEach((child) => {
    if (child.selected) {
      selectedCount += 1;
    }
    if (child.indeterminate) {
      indeterminateCount += 1;
    }
  });

  if (selectedCount === total) {
    node.selected = true;
    node.indeterminate = false;
  } else if (selectedCount === 0 && indeterminateCount === 0) {
    node.selected = false;
    node.indeterminate = false;
  } else {
    node.selected = false;
    node.indeterminate = true;
  }

  applySelectionState(node);
  updateParentSelection(node.parent);
}

function getRootNode(node) {
  let current = node;
  while (current && current.parent) {
    current = current.parent;
  }
  return current || null;
}

function refreshSelectedRootToken() {
  const selectedNodes = treeState.nodes.filter((node) => node.selected);
  if (selectedNodes.length === 0) {
    treeState.selectedRootToken = null;
    return;
  }
  const rootNode = getRootNode(selectedNodes[0]);
  treeState.selectedRootToken = rootNode ? rootNode.token : null;
}

function hasSelectedAncestor(node) {
  let current = node.parent;
  while (current) {
    if (current.selected) {
      return true;
    }
    current = current.parent;
  }
  return false;
}

function getSelectionRoots() {
  return treeState.nodes.filter((node) => node.selected && !hasSelectedAncestor(node));
}

function updateSelectionSummary() {
  const selectedNodes = treeState.nodes.filter((node) => node.selected);
  if (selectedNodes.length === 0) {
    elements.selectionSummary.textContent = "暂无选择。";
    return;
  }
  const folderCount = selectedNodes.filter((node) => node.isFolder).length;
  const fileCount = selectedNodes.length - folderCount;
  const rootNode = treeState.roots.find(
    (node) => node.token === treeState.selectedRootToken
  );
  const rootLabel = rootNode ? rootNode.name : treeState.selectedRootToken;
  elements.selectionSummary.textContent = `${rootLabel} | 已选择 ${folderCount} 个文件夹，${fileCount} 个文件。`;
}

function clearSelection() {
  treeState.nodes.forEach((node) => {
    node.selected = false;
    node.indeterminate = false;
    applySelectionState(node);
  });
  treeState.selectedRootToken = null;
  updateSelectionSummary();
}

function handleSelectionToggle(node, checked) {
  const rootNode = getRootNode(node);
  const rootToken = rootNode ? rootNode.token : null;
  if (checked && treeState.selectedRootToken && rootToken) {
    if (treeState.selectedRootToken !== rootToken) {
      appendLog("仅支持选择同一根目录", {
        current_root: treeState.selectedRootToken,
        attempted_root: rootToken,
      });
      applySelectionState(node);
      return;
    }
  }

  if (checked && rootToken && !treeState.selectedRootToken) {
    treeState.selectedRootToken = rootToken;
  }

  setNodeSelection(node, checked, { propagate: true });
  updateParentSelection(node.parent);
  refreshSelectedRootToken();
  updateSelectionSummary();
}

function createStatusItem(message) {
  const item = document.createElement("li");
  item.className = "tree-status";
  item.textContent = message;
  return item;
}

function renderNode(node) {
  const li = document.createElement("li");
  li.className = "tree-node";
  li.dataset.token = node.token;

  const row = document.createElement("div");
  row.className = "tree-row";

  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.className = "tree-check";
  checkbox.addEventListener("change", () => handleSelectionToggle(node, checkbox.checked));

  const toggle = document.createElement("button");
  toggle.className = "tree-toggle";
  toggle.type = "button";
  toggle.textContent = node.isFolder ? "▶" : "";
  toggle.disabled = !node.isFolder;
  toggle.addEventListener("click", () => toggleNode(node));

  const label = document.createElement("span");
  label.className = "tree-label";
  label.textContent = node.name || node.token;
  label.title = node.token;

  const meta = document.createElement("span");
  meta.className = "tree-meta";
  meta.textContent = node.type;

  row.append(checkbox, toggle, label, meta);

  const children = document.createElement("ul");
  children.className = "tree-children";

  li.append(row, children);

  node.element = li;
  node.childrenElement = children;
  node.toggleElement = toggle;
  node.checkboxElement = checkbox;
  applySelectionState(node);

  return li;
}
function renderTree() {
  elements.treeRoot.innerHTML = "";
  if (treeState.roots.length === 0) {
    elements.treeRoot.appendChild(createStatusItem("暂无文件夹。"));
    return;
  }
  treeState.roots.forEach((node) => {
    if (!node.element) {
      renderNode(node);
    }
    elements.treeRoot.appendChild(node.element);
  });
}

function resetTreeNodes() {
  treeState.roots = [];
  treeState.rootTokens.clear();
  treeState.nodes = [];
  treeState.selectedRootToken = null;
  elements.treeRoot.innerHTML = "";
  updateSelectionSummary();
}

function addRootNode({ token, name, type }) {
  if (!token) {
    appendLog("文件夹 Token 无效", { token });
    return;
  }
  if (treeState.rootTokens.has(token)) {
    appendLog("文件夹已存在于文件树", { token });
    return;
  }
  const node = createNode({ token, name, type: type || "folder" });
  treeState.rootTokens.add(token);
  treeState.roots.push(node);
  renderTree();
  updateSelectionSummary();
}

async function refreshTreeData() {
  const accessToken = requireValue(elements.userToken, "user_access_token");
  resetTreeNodes();
  appendLog("刷新文件树", { message: "正在重新获取根目录与共享文件夹信息。" });

  const rootResult = await postJson("/api/root_folder_meta", {
    user_access_token: accessToken,
  });
  logApiResult("根目录元数据", rootResult);
  if (!rootResult.ok) {
    throw new Error("刷新根目录失败。");
  }
  const rootToken = rootResult.response?.body?.data?.token;
  const rootName = rootResult.response?.body?.data?.name || "我的空间根目录";
  if (rootToken) {
    addRootNode({ token: rootToken, name: rootName, type: "folder" });
  }

  for (const token of treeState.sharedTokens) {
    const result = await postJson("/api/folder_meta", {
      user_access_token: accessToken,
      folder_token: token,
    });
    logApiResult("文件夹元数据", result);
    if (!result.ok) {
      appendLog("共享文件夹刷新失败", { folder_token: token });
      continue;
    }
    const name = result.response?.body?.data?.name || "共享文件夹";
    addRootNode({ token, name, type: "folder" });
  }
}

function setToggleState(node) {
  if (!node.toggleElement) {
    return;
  }
  node.toggleElement.textContent = node.expanded ? "▼" : "▶";
}

async function toggleNode(node) {
  if (!node.isFolder) {
    return;
  }
  node.expanded = !node.expanded;
  setToggleState(node);
  if (node.childrenElement) {
    node.childrenElement.classList.toggle("is-open", node.expanded);
  }
  if (node.expanded && !node.loaded) {
    await loadFolderChildren(node);
  }
}

async function fetchFolderPage({ token, folderToken, pageToken }) {
  const payload = {
    user_access_token: token,
    folder_token: folderToken,
    page_size: 200,
  };
  if (pageToken) {
    payload.page_token = pageToken;
  }
  return postJson("/api/folder_list", payload);
}

function normalizeFileType(item) {
  if (!item) {
    return "file";
  }
  const baseType = String(item.type || item.file_type || "file").toLowerCase();
  if (baseType === "shortcut" && item.shortcut_info?.target_type) {
    return `shortcut:${item.shortcut_info.target_type}`;
  }
  return baseType;
}

function resolveItemToken(item, type) {
  const rawToken = item.token || item.file_token || item.id;
  if (type?.startsWith("shortcut:") && item.shortcut_info?.target_token) {
    return item.shortcut_info.target_token;
  }
  return rawToken;
}

async function loadFolderChildren(node) {
  if (!node.childrenElement) {
    return;
  }
  node.childrenElement.innerHTML = "";
  node.childrenElement.appendChild(createStatusItem("加载中..."));

  try {
    const accessToken = requireValue(elements.userToken, "user_access_token");
    let pageToken = "";
    let pageIndex = 0;
    const allItems = [];

    while (true) {
      pageIndex += 1;
      const result = await fetchFolderPage({
        token: accessToken,
        folderToken: node.token,
        pageToken: pageToken || undefined,
      });

      logApiResult(`文件夹列表 第${pageIndex}页`, result);

      if (!result.ok) {
        throw new Error("文件夹列表请求失败");
      }

      const data = result.response?.body?.data || {};
      const files = Array.isArray(data.files) ? data.files : [];
      allItems.push(...files);
      if (!data.has_more || !data.page_token) {
        break;
      }
      pageToken = data.page_token;
      if (pageIndex > 20) {
        appendLog("分页已停止", { reason: "页数过多" });
        break;
      }
    }

    node.loaded = true;
    node.childrenElement.innerHTML = "";
    node.children = [];

    if (allItems.length === 0) {
      node.childrenElement.appendChild(createStatusItem("空文件夹。"));
      return;
    }

    const children = allItems
      .map((item) => {
        const type = normalizeFileType(item);
        const token = resolveItemToken(item, type);
        if (!token) {
          return null;
        }
        const nameBase = item.name || item.title || token;
        const name =
          type.startsWith("shortcut:") && item.name
            ? `${nameBase} (快捷方式)`
            : nameBase;
        return createNode({
          token,
          name,
          type,
          parent: node,
          raw: item,
        });
      })
      .filter(Boolean);

    children.forEach((child) => {
      node.children.push(child);
      renderNode(child);
      if (node.selected) {
        setNodeSelection(child, true, { propagate: true });
      }
      node.childrenElement.appendChild(child.element);
    });

    updateParentSelection(node);
  } catch (err) {
    node.childrenElement.innerHTML = "";
    node.childrenElement.appendChild(
      createStatusItem(`错误：${err.message}`)
    );
    appendLog("获取文件夹列表失败", err.message);
  }
}
function normalizeScope(scope) {
  return scope.trim().replace(/\s+/g, " ");
}

async function postJson(url, body) {
  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json; charset=utf-8",
    },
    body: JSON.stringify(body),
  });
  const data = await response.json();
  if (!response.ok) {
    return { ok: false, httpStatus: response.status, error: data };
  }
  return data;
}

function logApiResult(step, result) {
  if (!result || !result.request || !result.response) {
    appendLog(`${step} 错误`, result?.error || result);
    return;
  }
  appendLog(`${step} 请求`, result.request);
  appendLog(`${step} 响应`, result.response);
  if (!result.ok) {
    appendLog(`${step} 错误`, result.response?.body || result.error);
  }
}

function buildAuthorizeUrl({ clientId, redirectUri, scope, state }) {
  const url = new URL("https://accounts.feishu.cn/open-apis/authen/v1/authorize");
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("redirect_uri", redirectUri);
  url.searchParams.set("response_type", "code");
  if (scope) {
    url.searchParams.set("scope", scope);
  }
  if (state) {
    url.searchParams.set("state", state);
  }
  return url.toString();
}

function generateState() {
  const bytes = new Uint8Array(16);
  window.crypto.getRandomValues(bytes);
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function handleOauthMessage(event) {
  if (event.origin !== window.location.origin) {
    return;
  }
  const { type, payload } = event.data || {};
  if (type !== "feishu-oauth") {
    return;
  }
  const expectedState = localStorage.getItem(OAUTH_STATE_KEY);
  if (payload?.state && expectedState && payload.state !== expectedState) {
    appendLog("OAuth state 不一致", {
      expected: expectedState,
      received: payload.state,
    });
    return;
  }
  if (payload?.error) {
    appendLog("OAuth 授权失败", payload);
    return;
  }
  if (payload?.code) {
    elements.authCode.value = payload.code;
    appendLog("已收到 OAuth 授权码", payload);
  }
}

function resetStepper() {
  if (!elements.migrationSteps) {
    return;
  }
  elements.migrationSteps
    .querySelectorAll(".step")
    .forEach((step) => step.classList.remove("step--active", "step--done"));
}

function setStepState(stepId, state) {
  if (!elements.migrationSteps) {
    return;
  }
  const step = elements.migrationSteps.querySelector(
    `[data-step-id="${stepId}"]`
  );
  if (!step) {
    return;
  }
  step.classList.remove("step--active", "step--done");
  if (state === "active") {
    step.classList.add("step--active");
  }
  if (state === "done") {
    step.classList.add("step--done");
  }
}

function resetProgress() {
  updateProgress(elements.nodeProgressBar, elements.nodeProgressText, 0, 0);
  updateProgress(elements.moveProgressBar, elements.moveProgressText, 0, 0);
}

function updateProgress(bar, text, current, total) {
  if (!bar || !text) {
    return;
  }
  const safeTotal = total > 0 ? total : 0;
  const clampedCurrent = Math.min(current, safeTotal);
  const percent = safeTotal === 0 ? 0 : (clampedCurrent / safeTotal) * 100;
  bar.style.width = `${percent}%`;
  text.textContent = `${clampedCurrent} / ${safeTotal}`;
}

function normalizeWikiObjType(type) {
  if (!type) {
    return "";
  }
  const raw = String(type).toLowerCase();
  if (raw.startsWith("shortcut:")) {
    return raw.split(":")[1] || "";
  }
  return raw;
}

async function fetchFolderItems(accessToken, folderToken, label) {
  ensureNotCancelled();
  let pageToken = "";
  let pageIndex = 0;
  const allItems = [];

  while (true) {
    ensureNotCancelled();
    pageIndex += 1;
    const result = await fetchFolderPage({
      token: accessToken,
      folderToken,
      pageToken: pageToken || undefined,
    });

    logApiResult(`${label} 第${pageIndex}页`, result);

    if (!result.ok) {
      throw new Error(`${label} 请求失败`);
    }

    const data = result.response?.body?.data || {};
    const files = Array.isArray(data.files) ? data.files : [];
    allItems.push(...files);
    if (!data.has_more || !data.page_token) {
      break;
    }
    pageToken = data.page_token;
    if (pageIndex > 200) {
      appendLog("分页已停止", { label, reason: "页数过多" });
      break;
    }
  }

  return allItems;
}

function addFolderChainFromNode(node, folderMap, rootToken) {
  let current = node;
  while (current && current.token !== rootToken) {
    if (!folderMap.has(current.token)) {
      folderMap.set(current.token, {
        token: current.token,
        name: current.name || current.token,
        parentToken: current.parent ? current.parent.token : rootToken,
      });
    }
    current = current.parent;
  }
}

async function buildSelectionPlan(accessToken) {
  ensureNotCancelled();
  const selectionRoots = getSelectionRoots();
  if (selectionRoots.length === 0) {
    throw new Error("请至少选择一个文件或文件夹。");
  }

  const rootNode = getRootNode(selectionRoots[0]);
  const rootToken = rootNode ? rootNode.token : null;
  if (!rootToken) {
    throw new Error("无法识别所选根目录。");
  }

  for (const node of selectionRoots) {
    const nodeRoot = getRootNode(node);
    if (nodeRoot && nodeRoot.token !== rootToken) {
      throw new Error("选择范围跨越多个根目录。");
    }
  }

  const folderMap = new Map();
  const fileMap = new Map();
  const visited = new Set();
  let fileIndex = 0;

  async function traverseFolder(token, name, parentToken) {
    ensureNotCancelled();
    if (visited.has(token)) {
      return;
    }
    visited.add(token);

    if (token !== rootToken) {
      folderMap.set(token, {
        token,
        name: name || token,
        parentToken: parentToken || rootToken,
      });
    }

    const items = await fetchFolderItems(
      accessToken,
      token,
      `文件夹列表(${token})`
    );

    for (const item of items) {
      ensureNotCancelled();
      const type = normalizeFileType(item);
      const itemToken = resolveItemToken(item, type);
      if (!itemToken) {
        continue;
      }
      const itemName = item.name || item.title || itemToken;
      if (type === "folder" || type.endsWith(":folder")) {
        await traverseFolder(itemToken, itemName, token);
      } else {
        fileMap.set(itemToken, {
          token: itemToken,
          name: itemName,
          type,
          parentToken: token,
          index: fileIndex++,
        });
      }
    }
  }

  for (const node of selectionRoots) {
    if (node.isFolder) {
      addFolderChainFromNode(node.parent, folderMap, rootToken);
      await traverseFolder(node.token, node.name, node.parent?.token);
    } else {
      fileMap.set(node.token, {
        token: node.token,
        name: node.name || node.token,
        type: node.type,
        parentToken: node.parent ? node.parent.token : rootToken,
        index: fileIndex++,
      });
      addFolderChainFromNode(node.parent, folderMap, rootToken);
    }
  }

  return {
    rootToken,
    rootNode,
    folderMap,
    fileMap,
  };
}

async function fetchSelectedRootName(accessToken, rootToken, fallbackName) {
  ensureNotCancelled();
  const metaResult = await postJson("/api/folder_meta", {
    user_access_token: accessToken,
    folder_token: rootToken,
  });
  logApiResult("文件夹元数据", metaResult);
  const name = metaResult.response?.body?.data?.name;
  if (metaResult.ok && name) {
    return { name, meta: metaResult.response?.body?.data };
  }
  if (fallbackName) {
    return { name: fallbackName, meta: null };
  }
  throw new Error("获取选中根目录名称失败。");
}

async function fetchMyDriveRoot(accessToken) {
  ensureNotCancelled();
  const rootResult = await postJson("/api/root_folder_meta", {
    user_access_token: accessToken,
  });
  logApiResult("根目录元数据", rootResult);
  if (!rootResult.ok) {
    throw new Error("获取我的空间根目录失败。");
  }
  const data = rootResult.response?.body?.data || {};
  if (!data.token) {
    throw new Error("根目录信息缺少 token。");
  }
  return {
    token: data.token,
    name: data.name || "我的空间根目录",
    meta: data,
  };
}

async function createMigrationFolder(accessToken, parentToken, rootName) {
  ensureNotCancelled();
  const name = `${rootName}_to_migrate`;
  const result = await postJson("/api/create_folder", {
    user_access_token: accessToken,
    name,
    folder_token: parentToken,
  });
  logApiResult("创建文件夹", result);
  if (!result.ok) {
    throw new Error("创建 _to_migrate 失败。");
  }
  const data = result.response?.body?.data || {};
  const folderToken =
    data.folder_token || data.token || data.file_token || data.folder?.token;
  if (!folderToken) {
    throw new Error("创建文件夹响应缺少 token。");
  }
  return { name, token: folderToken, raw: data };
}
async function copySelectedFiles(accessToken, targetFolderToken, fileMap) {
  ensureNotCancelled();
  const files = Array.from(fileMap.values()).map((file) => {
    const wikiType = normalizeWikiObjType(file.type);
    return { ...file, wikiType };
  });
  const supported = files.filter((file) => WIKI_SUPPORTED_TYPES.has(file.wikiType));
  const skipped = files.filter((file) => !WIKI_SUPPORTED_TYPES.has(file.wikiType));
  const copyMap = new Map();
  const tasks = [];
  const errors = [];

  if (skipped.length > 0) {
    appendLog("已跳过不支持的文件类型", {
      count: skipped.length,
      items: skipped.map((item) => ({ token: item.token, type: item.type })),
    });
  }

  for (const file of supported) {
    ensureNotCancelled();
    const result = await postJson("/api/copy_file", {
      user_access_token: accessToken,
      file_token: file.token,
      name: file.name,
      type: file.wikiType,
      folder_token: targetFolderToken,
    });

    logApiResult("复制文件", result);

    if (!result.ok) {
      throw new Error(`复制失败：${file.name || file.token}`);
    }

    const data = result.response?.body?.data || {};
    const copiedToken = data.file?.token || data.token || data.file_token;
    if (copiedToken) {
      copyMap.set(file.token, copiedToken);
    } else if (data.task_id) {
      tasks.push({ taskId: data.task_id, file });
    }
  }

  return {
    supported,
    skipped,
    copyMap,
    tasks,
    errors,
  };
}

async function checkCopyTasks(accessToken, tasks) {
  const results = [];
  for (const task of tasks) {
    ensureNotCancelled();
    const result = await postJson("/api/drive_task_check", {
      user_access_token: accessToken,
      task_id: task.taskId,
    });
    logApiResult("复制任务检查", result);
    if (!result.ok) {
      throw new Error(`复制任务检查失败：${task.taskId}`);
    }
    results.push({ task, result });
  }
  return results;
}

function matchOriginalFileByNameType(item, remainingOriginals) {
  const type = normalizeWikiObjType(item.type);
  const index = remainingOriginals.findIndex(
    (original) =>
      original.name === item.name &&
      normalizeWikiObjType(original.type) === type
  );
  if (index === -1) {
    return null;
  }
  return remainingOriginals.splice(index, 1)[0];
}

async function verifyCopiedFiles(
  accessToken,
  migrateFolderToken,
  copyPlan,
  fileMap,
  rootToken
) {
  ensureNotCancelled();
  const copiedToOriginal = new Map();
  copyPlan.copyMap.forEach((copiedToken, originalToken) => {
    copiedToOriginal.set(copiedToken, originalToken);
  });

  const remainingOriginals = copyPlan.supported.map((file) => ({ ...file }));

  const items = await fetchFolderItems(
    accessToken,
    migrateFolderToken,
    "复制校验"
  );
  const files = items.filter((item) => {
    const type = normalizeFileType(item);
    return type !== "folder" && !type.endsWith(":folder");
  });

  const copiedItems = [];
  for (const item of files) {
    const type = normalizeFileType(item);
    const token = resolveItemToken(item, type);
    if (!token) {
      continue;
    }
    const name = item.name || item.title || token;
    const originalToken = copiedToOriginal.get(token);
    const original =
      (originalToken && fileMap.get(originalToken)) ||
      matchOriginalFileByNameType({ name, type }, remainingOriginals) ||
      null;

    copiedItems.push({
      token,
      name,
      type,
      parentToken: original?.parentToken || rootToken,
      originalToken: original?.token || originalToken || null,
      index: original?.index ?? Number.MAX_SAFE_INTEGER,
    });
  }

  copiedItems.sort((a, b) => a.index - b.index);

  return {
    listed: files.length,
    copiedItems,
  };
}
async function createWikiSpace(accessToken, name) {
  ensureNotCancelled();
  const result = await postJson("/api/wiki_space_create", {
    user_access_token: accessToken,
    name,
  });
  logApiResult("创建知识空间", result);
  if (!result.ok) {
    throw new Error("创建知识空间失败。");
  }
  const data = result.response?.body?.data || {};
  const spaceId = data.space_id || data.space?.space_id || data.space?.id || data.id;
  if (!spaceId) {
    throw new Error("知识空间返回缺少 space_id。");
  }
  return spaceId;
}

async function createWikiNodes(
  accessToken,
  spaceId,
  folderMap,
  rootToken,
  onProgress
) {
  ensureNotCancelled();
  const pending = new Map(folderMap);
  const wikiNodeMap = new Map();
  const total = pending.size;

  while (pending.size > 0) {
    ensureNotCancelled();
    let progressed = false;
    for (const [token, folder] of Array.from(pending.entries())) {
      ensureNotCancelled();
      const parentToken = folder.parentToken;
      const needsParent =
        parentToken && parentToken !== rootToken && !wikiNodeMap.has(parentToken);
      if (needsParent) {
        continue;
      }

      const payload = {
        user_access_token: accessToken,
        space_id: spaceId,
        obj_type: "docx",
        node_type: "origin",
        title: folder.name,
      };
      if (parentToken && parentToken !== rootToken) {
        payload.parent_node_token = wikiNodeMap.get(parentToken);
      }

      const result = await postJson("/api/wiki_node_create", payload);
      logApiResult("创建知识节点", result);
      if (!result.ok) {
        throw new Error(`创建知识节点失败：${folder.name}`);
      }

      const data = result.response?.body?.data || {};
      const nodeToken =
        data.node?.node_token ||
        data.node_token ||
        data.wiki_node?.node_token ||
        data.wiki_node?.token;
      if (!nodeToken) {
        throw new Error(`知识节点返回缺少 token：${folder.name}`);
      }

      wikiNodeMap.set(token, nodeToken);
      pending.delete(token);
      progressed = true;
      if (onProgress) {
        onProgress(wikiNodeMap.size, total);
      }
    }

    if (!progressed) {
      throw new Error("创建知识节点未完成，已停止。");
    }
  }

  return wikiNodeMap;
}

async function moveDocsToWiki(
  accessToken,
  spaceId,
  copiedItems,
  rootToken,
  wikiNodeMap,
  onProgress
) {
  ensureNotCancelled();
  const moved = [];
  const failed = [];
  const skipped = [];
  const taskIds = [];
  let requestCount = 0;
  const total = copiedItems.length;

  for (const item of copiedItems) {
    ensureNotCancelled();
    const objType = normalizeWikiObjType(item.type);
    if (!WIKI_SUPPORTED_TYPES.has(objType)) {
      skipped.push(item);
      continue;
    }

    const parentToken = item.parentToken;
    const payload = {
      user_access_token: accessToken,
      space_id: spaceId,
      obj_type: objType,
      obj_token: item.token,
      apply: true,
    };
    if (parentToken && parentToken !== rootToken) {
      payload.parent_wiki_token = wikiNodeMap.get(parentToken);
    }

    const result = await postJson("/api/wiki_move_docs", payload);
    logApiResult("移动文档到知识库", result);
    if (!result.ok) {
      throw new Error(`移动到知识库失败：${item.name || item.token}`);
    }

    const data = result.response?.body?.data || {};
    if (data.task_id) {
      taskIds.push(data.task_id);
    }
    moved.push(item);
    if (onProgress) {
      onProgress(moved.length, total);
    }

    requestCount += 1;
    if (requestCount % MOVE_DOCS_BATCH_LIMIT === 0) {
      appendLog("触发限流暂停", {
        message: "已发送 90 次移动请求，等待一段时间后继续，请耐心等待。",
        wait_ms: MOVE_DOCS_PAUSE_MS,
      });
      await sleepWithCancel(MOVE_DOCS_PAUSE_MS);
    }
  }

  return {
    moved,
    failed,
    skipped,
    taskIds,
  };
}

function extractWikiTaskSummary(responseBody) {
  const data = responseBody?.data || responseBody || {};
  const result = data.result || data;
  return {
    status: result.status || result.state || null,
    success: result.success_num ?? result.success_count ?? null,
    failed: result.fail_num ?? result.fail_count ?? null,
    fail_reason: result.fail_reason || result.fail_reasons || result.fail_msg || null,
    raw: result,
  };
}

async function checkWikiTasks(accessToken, taskIds) {
  const summaries = [];
  for (const taskId of taskIds) {
    ensureNotCancelled();
    const result = await postJson("/api/wiki_task_get", {
      user_access_token: accessToken,
      task_id: taskId,
      task_type: WIKI_TASK_TYPE,
    });
    logApiResult("查询迁移任务", result);
    if (!result.ok) {
      throw new Error(`迁移任务检查失败：${taskId}`);
    }
    summaries.push({
      task_id: taskId,
      ok: result.ok,
      summary: extractWikiTaskSummary(result.response?.body),
    });
  }
  return summaries;
}
elements.redirectUri.value = `${window.location.origin}/oauth/callback`;
elements.scope.value = DEFAULT_SCOPE;
elements.scope.readOnly = true;

function openOauthWindow(appId, redirectUri, scopeValue) {
  const state = generateState();
  localStorage.setItem(OAUTH_STATE_KEY, state);

  const authorizeUrl = buildAuthorizeUrl({
    clientId: appId,
    redirectUri,
    scope: scopeValue || undefined,
    state,
  });

  appendLog("打开 OAuth 授权页面", {
    url: authorizeUrl,
    state,
    scope: scopeValue,
  });

  const popup = window.open(
    authorizeUrl,
    "feishu-oauth",
    "width=520,height=720"
  );
  if (!popup) {
    appendLog("弹窗被拦截", {
      message: "请允许弹窗后重试授权。",
    });
  }
}

function waitForAuthCode(timeoutMs = 300000) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const timer = setInterval(() => {
      const code = getValue(elements.authCode);
      if (code) {
        clearInterval(timer);
        resolve(code);
        return;
      }
      if (Date.now() - start > timeoutMs) {
        clearInterval(timer);
        reject(new Error("等待授权码超时。"));
      }
    }, 800);
  });
}

async function runQuickAuthFlow() {
  const appId = requireValue(elements.appId, "APP_ID");
  const appSecret = requireValue(elements.appSecret, "APP_SECRET");
  const redirectUri = requireValue(elements.redirectUri, "Redirect URI");
  const scopeValue = normalizeScope(getValue(elements.scope));

  appendLog("开始获取 tenant_access_token", {
    url: "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
    app_id: appId,
  });

  const tenantResult = await postJson("/api/tenant_access_token", {
    app_id: appId,
    app_secret: appSecret,
  });
  logApiResult("获取 tenant_access_token", tenantResult);
  if (!tenantResult.ok) {
    throw new Error("获取 tenant_access_token 失败。");
  }
  const tenantToken = tenantResult.response?.body?.tenant_access_token;
  if (tenantToken) {
    elements.tenantToken.value = tenantToken;
  }

  openOauthWindow(appId, redirectUri, scopeValue);
  appendLog("等待授权码...", { timeout_ms: 300000 });
  const code = await waitForAuthCode();

  appendLog("开始获取 user_access_token", {
    url: "https://open.feishu.cn/open-apis/authen/v2/oauth/token",
    code,
    redirect_uri: redirectUri,
  });

  const userResult = await postJson("/api/user_access_token", {
    client_id: appId,
    client_secret: appSecret,
    code,
    redirect_uri: redirectUri,
    scope: scopeValue || undefined,
  });
  logApiResult("获取 user_access_token", userResult);
  if (!userResult.ok) {
    throw new Error("获取 user_access_token 失败。");
  }
  const userToken =
    userResult.response?.body?.access_token ||
    userResult.response?.body?.user_access_token;
  if (userToken) {
    elements.userToken.value = userToken;
  }

  appendLog("开始获取根目录信息", {
    url: "https://open.feishu.cn/open-apis/drive/explorer/v2/root_folder/meta",
  });
  const rootResult = await postJson("/api/root_folder_meta", {
    user_access_token: userToken,
  });
  logApiResult("根目录元数据", rootResult);
  if (!rootResult.ok) {
    throw new Error("获取根目录失败。");
  }

  setResult(rootResult.response?.body || rootResult);
  const rootToken = rootResult.response?.body?.data?.token;
  const rootName = rootResult.response?.body?.data?.name || "我的空间根目录";
  if (rootToken) {
    addRootNode({
      token: rootToken,
      name: rootName,
      type: "folder",
    });
  }
}

elements.btnQuickAuth.addEventListener("click", async () => {
  if (authState.active) {
    return;
  }
  setAuthActive(true);
  try {
    await runQuickAuthFlow();
  } catch (err) {
    appendLog("一键授权失败", err.message);
  } finally {
    setAuthActive(false);
  }
});

elements.btnAddFolder.addEventListener("click", async () => {
  try {
    const accessToken = requireValue(elements.userToken, "user_access_token");
    const folderToken = requireValue(
      elements.sharedFolderToken,
      "文件夹 Token"
    );

    const result = await postJson("/api/folder_meta", {
      user_access_token: accessToken,
      folder_token: folderToken,
    });

    logApiResult("文件夹元数据", result);
    if (!result.ok) {
      throw new Error("获取共享文件夹信息失败。");
    }
    const name = result.response?.body?.data?.name || "共享文件夹";
    addRootNode({ token: folderToken, name, type: "folder" });
    treeState.sharedTokens.add(folderToken);
  } catch (err) {
    appendLog("添加共享文件夹失败", err.message);
  }
});

elements.btnMigrate.addEventListener("click", async () => {
  resetStepper();
  setResult("");
  resetProgress();
  setMigrationActive(true);
  try {
    const accessToken = requireValue(elements.userToken, "user_access_token");

    setStepState("select", "active");
    const selection = await buildSelectionPlan(accessToken);
    setStepState("select", "done");

    const supportedCandidates = Array.from(selection.fileMap.values()).filter(
      (file) => WIKI_SUPPORTED_TYPES.has(normalizeWikiObjType(file.type))
    );
    if (supportedCandidates.length === 0) {
      throw new Error("所选内容中没有可迁移的文件。");
    }

    setStepState("root-meta", "active");
    const selectedRootMeta = await fetchSelectedRootName(
      accessToken,
      selection.rootToken,
      selection.rootNode?.name
    );
    const rootName = selectedRootMeta.name || selection.rootToken;
    const myDriveRoot = await fetchMyDriveRoot(accessToken);
    setStepState("root-meta", "done");

    setStepState("create-migrate-folder", "active");
    const migrateFolder = await createMigrationFolder(
      accessToken,
      myDriveRoot.token,
      rootName
    );
    setStepState("create-migrate-folder", "done");

    setStepState("copy-files", "active");
    const copyPlan = await copySelectedFiles(
      accessToken,
      migrateFolder.token,
      selection.fileMap
    );
    setStepState("copy-files", "done");

    setStepState("copy-check", "active");
    const copyTasks = await checkCopyTasks(accessToken, copyPlan.tasks);
    const copyCheck = await verifyCopiedFiles(
      accessToken,
      migrateFolder.token,
      copyPlan,
      selection.fileMap,
      selection.rootToken
    );
    if (copyCheck.copiedItems.length < copyPlan.supported.length) {
      throw new Error(
        `复制校验失败：期望 ${copyPlan.supported.length}，实际 ${copyCheck.copiedItems.length}`
      );
    }
    setStepState("copy-check", "done");

    setStepState("create-space", "active");
    const spaceId = await createWikiSpace(accessToken, rootName);
    setStepState("create-space", "done");

    setStepState("create-nodes", "active");
    updateProgress(
      elements.nodeProgressBar,
      elements.nodeProgressText,
      0,
      selection.folderMap.size
    );
    const wikiNodeMap = await createWikiNodes(
      accessToken,
      spaceId,
      selection.folderMap,
      selection.rootToken,
      (current, total) =>
        updateProgress(
          elements.nodeProgressBar,
          elements.nodeProgressText,
          current,
          total
        )
    );
    setStepState("create-nodes", "done");

    setStepState("move-docs", "active");
    const moveTargets = copyCheck.copiedItems.filter((item) =>
      WIKI_SUPPORTED_TYPES.has(normalizeWikiObjType(item.type))
    );
    updateProgress(
      elements.moveProgressBar,
      elements.moveProgressText,
      0,
      moveTargets.length
    );
    const moveReport = await moveDocsToWiki(
      accessToken,
      spaceId,
      moveTargets,
      selection.rootToken,
      wikiNodeMap,
      (current, total) =>
        updateProgress(
          elements.moveProgressBar,
          elements.moveProgressText,
          current,
          total
        )
    );
    setStepState("move-docs", "done");

    setStepState("wiki-task", "active");
    const wikiTasks = await checkWikiTasks(accessToken, moveReport.taskIds);
    setStepState("wiki-task", "done");

    setStepState("delete-migrate-folder", "done");

    setResult({
      root: {
        token: selection.rootToken,
        name: rootName,
      },
      my_drive_root: {
        token: myDriveRoot.token,
        name: myDriveRoot.name,
      },
      selection: {
        folders: selection.folderMap.size,
        files: selection.fileMap.size,
      },
      migrate_folder: {
        name: migrateFolder.name,
        token: migrateFolder.token,
      },
      cleanup: {
        reminder: `请手动删除文件夹 ${migrateFolder.name}`,
      },
      copy: {
        requested: copyPlan.supported.length,
        skipped: copyPlan.skipped.length,
        errors: copyPlan.errors.length,
        listed: copyCheck.listed,
        tasks_checked: copyTasks.length,
      },
      wiki: {
        space_id: spaceId,
        nodes_created: wikiNodeMap.size,
        moved: moveReport.moved.length,
        failed: moveReport.failed.length,
        skipped: moveReport.skipped.length,
        task_ids: moveReport.taskIds,
        task_summaries: wikiTasks,
      },
    });
  } catch (err) {
    appendLog("迁移失败", err.message);
    setResult({ error: err.message });
  } finally {
    setMigrationActive(false);
  }
});

elements.btnStop.addEventListener("click", () => {
  if (!migrationState.active) {
    return;
  }
  migrationState.cancelled = true;
  appendLog("用户已取消迁移");
});

elements.btnClearSelection.addEventListener("click", () => {
  clearSelection();
});

elements.btnClear.addEventListener("click", () => {
  elements.log.textContent = "";
  elements.result.textContent = "";
});

elements.btnRefreshTree.addEventListener("click", async () => {
  try {
    await refreshTreeData();
  } catch (err) {
    appendLog("刷新文件树失败", err.message);
  }
});

elements.btnCollapseAll.addEventListener("click", () => {
  treeState.nodes.forEach((node) => {
    node.expanded = false;
    setToggleState(node);
    if (node.childrenElement) {
      node.childrenElement.classList.remove("is-open");
    }
  });
});

window.addEventListener("message", handleOauthMessage);

renderTree();
updateSelectionSummary();
resetProgress();
