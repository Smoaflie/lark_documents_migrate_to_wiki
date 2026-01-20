const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = path.join(__dirname, "public");

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
};

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload, null, 2);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  res.end(body);
}

function serveFile(res, filePath) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not found");
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || "application/octet-stream";
    res.writeHead(200, {
      "Content-Type": contentType,
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
}

function safeJsonParse(text) {
  try {
    return JSON.parse(text);
  } catch (err) {
    return undefined;
  }
}

async function readJsonBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }
  if (chunks.length === 0) {
    return {};
  }
  const text = Buffer.concat(chunks).toString("utf8");
  if (!text.trim()) {
    return {};
  }
  const parsed = safeJsonParse(text);
  if (parsed === undefined) {
    throw new Error("Invalid JSON body");
  }
  return parsed;
}

async function proxyRequest({ url, method, headers, body }) {
  const response = await fetch(url, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });

  const responseText = await response.text();
  const parsed = safeJsonParse(responseText);
  const responseBody = parsed === undefined ? responseText : parsed;

  const responseHeaders = {};
  response.headers.forEach((value, key) => {
    responseHeaders[key] = value;
  });

  const ok =
    response.ok &&
    (parsed === undefined || parsed.code === undefined || parsed.code === 0);

  return {
    ok,
    request: {
      url,
      method,
      headers,
      body,
    },
    response: {
      status: response.status,
      statusText: response.statusText,
      headers: responseHeaders,
      body: responseBody,
    },
  };
}

function validateRequired(value, name) {
  if (!value || String(value).trim() === "") {
    return `${name} is required`;
  }
  return null;
}

const server = http.createServer(async (req, res) => {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);

  if (requestUrl.pathname === "/api/tenant_access_token" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const appIdError = validateRequired(body.app_id, "app_id");
      const appSecretError = validateRequired(body.app_secret, "app_secret");
      if (appIdError || appSecretError) {
        sendJson(res, 400, {
          ok: false,
          error: appIdError || appSecretError,
        });
        return;
      }

      const result = await proxyRequest({
        url: "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        method: "POST",
        headers: {
          "Content-Type": "application/json; charset=utf-8",
        },
        body: {
          app_id: body.app_id,
          app_secret: body.app_secret,
        },
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/user_access_token" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const missing =
        validateRequired(body.client_id, "client_id") ||
        validateRequired(body.client_secret, "client_secret") ||
        validateRequired(body.code, "code") ||
        validateRequired(body.redirect_uri, "redirect_uri");
      if (missing) {
        sendJson(res, 400, { ok: false, error: missing });
        return;
      }

      const payload = {
        grant_type: "authorization_code",
        client_id: body.client_id,
        client_secret: body.client_secret,
        code: body.code,
        redirect_uri: body.redirect_uri,
      };
      if (body.scope) {
        payload.scope = body.scope;
      }

      const result = await proxyRequest({
        url: "https://open.feishu.cn/open-apis/authen/v2/oauth/token",
        method: "POST",
        headers: {
          "Content-Type": "application/json; charset=utf-8",
        },
        body: payload,
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/root_folder_meta" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      if (tokenError) {
        sendJson(res, 400, { ok: false, error: tokenError });
        return;
      }

      const result = await proxyRequest({
        url: "https://open.feishu.cn/open-apis/drive/explorer/v2/root_folder/meta",
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/folder_meta" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const folderTokenError = validateRequired(body.folder_token, "folder_token");
      if (tokenError || folderTokenError) {
        sendJson(res, 400, { ok: false, error: tokenError || folderTokenError });
        return;
      }

      const result = await proxyRequest({
        url: `https://open.feishu.cn/open-apis/drive/explorer/v2/folder/${encodeURIComponent(
          body.folder_token
        )}/meta`,
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/create_folder" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const nameError = validateRequired(body.name, "name");
      const folderTokenError = validateRequired(body.folder_token, "folder_token");
      if (tokenError || nameError || folderTokenError) {
        sendJson(res, 400, {
          ok: false,
          error: tokenError || nameError || folderTokenError,
        });
        return;
      }

      const result = await proxyRequest({
        url: "https://open.feishu.cn/open-apis/drive/v1/files/create_folder",
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json; charset=utf-8",
        },
        body: {
          name: body.name,
          folder_token: body.folder_token,
        },
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/copy_file" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const fileTokenError = validateRequired(body.file_token, "file_token");
      const nameError = validateRequired(body.name, "name");
      const typeError = validateRequired(body.type, "type");
      const folderTokenError = validateRequired(body.folder_token, "folder_token");
      if (tokenError || fileTokenError || nameError || typeError || folderTokenError) {
        sendJson(res, 400, {
          ok: false,
          error:
            tokenError ||
            fileTokenError ||
            nameError ||
            typeError ||
            folderTokenError,
        });
        return;
      }

      const query = new URLSearchParams();
      if (body.user_id_type) {
        query.set("user_id_type", body.user_id_type);
      }

      const payload = {
        name: body.name,
        type: body.type,
        folder_token: body.folder_token,
      };
      if (body.extra) {
        payload.extra = body.extra;
      }

      const result = await proxyRequest({
        url: `https://open.feishu.cn/open-apis/drive/v1/files/${encodeURIComponent(
          body.file_token
        )}/copy?${query.toString()}`,
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json; charset=utf-8",
        },
        body: payload,
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/drive_task_check" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const taskError = validateRequired(body.task_id, "task_id");
      if (tokenError || taskError) {
        sendJson(res, 400, { ok: false, error: tokenError || taskError });
        return;
      }

      const query = new URLSearchParams();
      query.set("task_id", body.task_id);

      const result = await proxyRequest({
        url: `https://open.feishu.cn/open-apis/drive/v1/files/task_check?${query.toString()}`,
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/delete_file" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const fileTokenError = validateRequired(body.file_token, "file_token");
      if (tokenError || fileTokenError) {
        sendJson(res, 400, { ok: false, error: tokenError || fileTokenError });
        return;
      }

      const query = new URLSearchParams();
      const fileType = body.file_type || body.type;
      if (fileType) {
        query.set("type", fileType);
      }
      if (body.user_id_type) {
        query.set("user_id_type", body.user_id_type);
      }

      const baseUrl = `https://open.feishu.cn/open-apis/drive/v1/files/${encodeURIComponent(
        body.file_token
      )}`;
      const querySuffix = query.toString();
      const queryPart = querySuffix ? `?${querySuffix}` : "";
      const deleteUrl = `${baseUrl}${queryPart}`;

      const deleteResult = await proxyRequest({
        url: deleteUrl,
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      sendJson(res, 200, deleteResult);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/wiki_space_create" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const nameError = validateRequired(body.name, "name");
      if (tokenError || nameError) {
        sendJson(res, 400, { ok: false, error: tokenError || nameError });
        return;
      }

      const payload = {
        name: body.name,
      };
      if (body.description) {
        payload.description = body.description;
      }

      const result = await proxyRequest({
        url: "https://open.feishu.cn/open-apis/wiki/v2/spaces",
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json; charset=utf-8",
        },
        body: payload,
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/wiki_node_create" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const spaceError = validateRequired(body.space_id, "space_id");
      const objTypeError = validateRequired(body.obj_type, "obj_type");
      const nodeTypeError = validateRequired(body.node_type, "node_type");
      if (tokenError || spaceError || objTypeError || nodeTypeError) {
        sendJson(res, 400, {
          ok: false,
          error: tokenError || spaceError || objTypeError || nodeTypeError,
        });
        return;
      }

      const payload = {
        obj_type: body.obj_type,
        node_type: body.node_type,
      };
      if (body.parent_node_token) {
        payload.parent_node_token = body.parent_node_token;
      }
      if (body.origin_node_token) {
        payload.origin_node_token = body.origin_node_token;
      }
      if (body.title) {
        payload.title = body.title;
      }

      const result = await proxyRequest({
        url: `https://open.feishu.cn/open-apis/wiki/v2/spaces/${encodeURIComponent(
          body.space_id
        )}/nodes`,
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json; charset=utf-8",
        },
        body: payload,
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/wiki_move_docs" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const spaceError = validateRequired(body.space_id, "space_id");
      const objTypeError = validateRequired(body.obj_type, "obj_type");
      const objTokenError = validateRequired(body.obj_token, "obj_token");
      if (tokenError || spaceError || objTypeError || objTokenError) {
        sendJson(res, 400, {
          ok: false,
          error: tokenError || spaceError || objTypeError || objTokenError,
        });
        return;
      }

      const payload = {
        obj_type: body.obj_type,
        obj_token: body.obj_token,
      };
      if (body.parent_wiki_token) {
        payload.parent_wiki_token = body.parent_wiki_token;
      }
      if (body.apply !== undefined) {
        payload.apply = Boolean(body.apply);
      }

      const result = await proxyRequest({
        url: `https://open.feishu.cn/open-apis/wiki/v2/spaces/${encodeURIComponent(
          body.space_id
        )}/nodes/move_docs_to_wiki`,
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json; charset=utf-8",
        },
        body: payload,
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/wiki_task_get" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const taskError = validateRequired(body.task_id, "task_id");
      if (tokenError || taskError) {
        sendJson(res, 400, { ok: false, error: tokenError || taskError });
        return;
      }

      const query = new URLSearchParams();
      if (body.task_type) {
        query.set("task_type", body.task_type);
      }

      const baseUrl = `https://open.feishu.cn/open-apis/wiki/v2/tasks/${encodeURIComponent(
        body.task_id
      )}`;
      const querySuffix = query.toString();
      const url = querySuffix ? `${baseUrl}?${querySuffix}` : baseUrl;

      const result = await proxyRequest({
        url,
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/api/folder_list" && req.method === "POST") {
    try {
      const body = await readJsonBody(req);
      const token =
        body.user_access_token || body.access_token || body.userAccessToken;
      const tokenError = validateRequired(token, "user_access_token");
      const folderTokenError = validateRequired(body.folder_token, "folder_token");
      if (tokenError || folderTokenError) {
        sendJson(res, 400, { ok: false, error: tokenError || folderTokenError });
        return;
      }

      const query = new URLSearchParams();
      query.set("folder_token", body.folder_token);
      if (body.page_size) {
        query.set("page_size", String(body.page_size));
      }
      if (body.page_token) {
        query.set("page_token", body.page_token);
      }
      if (body.order_by) {
        query.set("order_by", body.order_by);
      }
      if (body.direction) {
        query.set("direction", body.direction);
      }
      if (body.user_id_type) {
        query.set("user_id_type", body.user_id_type);
      }

      const result = await proxyRequest({
        url: `https://open.feishu.cn/open-apis/drive/v1/files?${query.toString()}`,
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      sendJson(res, 200, result);
    } catch (err) {
      sendJson(res, 500, { ok: false, error: err.message });
    }
    return;
  }

  if (requestUrl.pathname === "/oauth/callback") {
    serveFile(res, path.join(PUBLIC_DIR, "oauth-callback.html"));
    return;
  }

  if (requestUrl.pathname === "/" || requestUrl.pathname === "/index.html") {
    serveFile(res, path.join(PUBLIC_DIR, "index.html"));
    return;
  }

  const safePath = path.normalize(
    path.join(PUBLIC_DIR, requestUrl.pathname.replace(/^\/+/, ""))
  );
  if (!safePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Bad request");
    return;
  }

  fs.stat(safePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not found");
      return;
    }
    serveFile(res, safePath);
  });
});

server.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
