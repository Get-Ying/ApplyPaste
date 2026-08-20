export const FEISHU_FIELD_DEFINITIONS = [
  ["companyName", "公司名称", "公司名称", true],
  ["jobName", "岗位名称", "岗位名称", true],
  ["location", "工作地点-base", "工作地点", true],
  ["jd", "JD", "岗位JD", false],
  ["jobUrl", "投递链接", "岗位链接", false],
  ["notes", "备注", "备注", false],
  ["status", "状态", "状态", true],
  ["applyDate", "投递日期", "投递日期", true],
  ["resumeName", "使用简历", "使用简历", false],
  ["profileVersion", "填写版本", "填写版本", false],
  ["localRecordId", "ApplyPaste记录ID", "记录ID", true]
].map(([key, defaultName, recordKey, required]) => ({ key, defaultName, recordKey, required }));

export const FEISHU_FIELDS = FEISHU_FIELD_DEFINITIONS.map(field => field.key);
export const MINIMAL_FEISHU_FIELDS = ["companyName", "jobName", "location", "status", "applyDate"];

export const DEFAULT_FEISHU_CONFIG = {
  enabled: false, appId: "", appSecret: "", tableUrl: "", appToken: "", tableId: "",
  mapping: Object.fromEntries(FEISHU_FIELD_DEFINITIONS.map(field => [field.key, field.defaultName]))
};

export function normalizeFeishuConfig(config) {
  const sourceMapping = config?.mapping || {}, mapping = { ...DEFAULT_FEISHU_CONFIG.mapping, ...sourceMapping };
  for (const field of FEISHU_FIELD_DEFINITIONS) {
    if (sourceMapping[field.key]) mapping[field.key] = sourceMapping[field.key];
    else if (sourceMapping[field.recordKey] && sourceMapping[field.recordKey] !== field.recordKey) mapping[field.key] = sourceMapping[field.recordKey];
    else mapping[field.key] = field.defaultName;
  }
  return { ...structuredClone(DEFAULT_FEISHU_CONFIG), ...(config || {}), mapping };
}

export function validateFeishuConfig(config) {
  validateFeishuCredentials(config);
  if (!config.appToken.trim()) throw new Error("请填写 app_token");
  if (!config.tableId.trim()) throw new Error("请填写 table_id");
}

export async function testFeishuConnection(config) {
  validateFeishuCredentials(config);
  const token = await getTenantAccessToken(config), resolved = await resolveFeishuTableConfig(config, token);
  validateFeishuConfig(resolved);
  const fields = await listFeishuFields(resolved, token);
  if (!fields.length) throw syncError({ msg: "目标数据表没有可读取字段", unmatchedFields: MINIMAL_FEISHU_FIELDS.map(key => resolved.mapping[key]), sentFields: [] });
  return { config: resolved, fields };
}

export async function listFeishuFields(config, accessToken = "") {
  const token = accessToken || await getTenantAccessToken(config), fields = []; let pageToken = "";
  do {
    const query = new URLSearchParams({ page_size: "100" }); if (pageToken) query.set("page_token", pageToken);
    const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${encodeURIComponent(config.appToken)}/tables/${encodeURIComponent(config.tableId)}/fields?${query}`;
    const result = await requestJson(url, { headers: { Authorization: `Bearer ${token}` } });
    fields.push(...(result.data?.items || []).map(field => ({ field_name: field.field_name, field_id: field.field_id, type: field.type, ui_type: field.ui_type || "", is_primary: Boolean(field.is_primary) })));
    pageToken = result.data?.has_more ? result.data?.page_token || "" : "";
  } while (pageToken);
  return fields;
}

export async function resolveFeishuTableConfig(config, accessToken = "") {
  const resolved = normalizeFeishuConfig(config), link = parseTableLink(resolved.tableUrl);
  if (link.tableId) resolved.tableId = link.tableId;
  if (link.appToken) resolved.appToken = link.appToken;
  if (!link.wikiNodeToken) return resolved;
  const token = accessToken || await getTenantAccessToken(resolved);
  const result = await requestJson(`https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token=${encodeURIComponent(link.wikiNodeToken)}`, { headers: { Authorization: `Bearer ${token}` } });
  const node = result.data?.node;
  if (!node) throw new Error("飞书未返回知识空间节点信息");
  if (node.obj_type !== "bitable") throw new Error(`该 Wiki 节点不是多维表格（类型：${node.obj_type || "未知"}）`);
  if (!node.obj_token) throw new Error("飞书未返回多维表格 app_token");
  resolved.appToken = node.obj_token; return resolved;
}

export async function createFeishuRecord(config, record, options = {}) {
  validateFeishuCredentials(config);
  const token = await getTenantAccessToken(config), resolved = await resolveFeishuTableConfig(config, token);
  validateFeishuConfig(resolved);
  const availableFields = options.fields || await listFeishuFields(resolved, token);
  const payload = buildFeishuFields(resolved, record, availableFields, options.mode === "minimal");
  const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${encodeURIComponent(resolved.appToken)}/tables/${encodeURIComponent(resolved.tableId)}/records`;
  try {
    const result = await requestJson(url, { method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json; charset=utf-8" }, body: JSON.stringify({ fields: payload.fields }) });
    return result.data?.record?.record_id || "";
  } catch (error) {
    throw enrichSyncError(error, payload.unmatchedFields, Object.keys(payload.fields));
  }
}

export async function syncFeishuRecord(config, record, options = {}) {
  validateFeishuCredentials(config);
  const token = await getTenantAccessToken(config), resolved = await resolveFeishuTableConfig(config, token);
  validateFeishuConfig(resolved);
  const availableFields = options.fields || await listFeishuFields(resolved, token);
  const payload = buildFeishuFields(resolved, record, availableFields, false);
  let recordId = String(record["飞书记录ID"] || "").trim();
  try {
    if (!recordId) recordId = await findFeishuRecordId(resolved, record["记录ID"], availableFields, token);
    if (recordId) {
      await updateFeishuRecord(resolved, recordId, payload.fields, token);
      return { recordId, action: "update" };
    }
    const result = await requestJson(recordsUrl(resolved), { method: "POST", headers: apiHeaders(token), body: JSON.stringify({ fields: payload.fields }) });
    return { recordId: result.data?.record?.record_id || "", action: "create" };
  } catch (error) {
    throw enrichSyncError(error, payload.unmatchedFields, Object.keys(payload.fields));
  }
}

async function findFeishuRecordId(config, localRecordId, actualFields, token) {
  if (!localRecordId) throw syncError({ msg: "本地记录缺少 ApplyPaste记录ID，已停止同步以避免重复记录" });
  const fieldName = String(config.mapping?.localRecordId || DEFAULT_FEISHU_CONFIG.mapping.localRecordId).trim();
  const actual = actualFields.find(field => field.field_name === fieldName);
  if (!actual || isReadOnlyField(actual)) throw syncError({ msg: `无法使用唯一标识字段：${fieldName}`, unmatchedFields: [fieldName] });
  const query = new URLSearchParams({ page_size: "1", filter: `CurrentValue.[${fieldName.replaceAll("]", "\\]")}] = "${escapeFilterValue(localRecordId)}"` });
  const result = await requestJson(`${recordsUrl(config)}?${query}`, { headers: { Authorization: `Bearer ${token}` } });
  return result.data?.items?.[0]?.record_id || "";
}

async function updateFeishuRecord(config, recordId, fields, token) {
  await requestJson(`${recordsUrl(config)}/${encodeURIComponent(recordId)}`, { method: "PUT", headers: apiHeaders(token), body: JSON.stringify({ fields }) });
}

function recordsUrl(config) { return `https://open.feishu.cn/open-apis/bitable/v1/apps/${encodeURIComponent(config.appToken)}/tables/${encodeURIComponent(config.tableId)}/records`; }
function apiHeaders(token) { return { Authorization: `Bearer ${token}`, "Content-Type": "application/json; charset=utf-8" }; }
function escapeFilterValue(value) { return String(value).replaceAll("\\", "\\\\").replaceAll('"', '\\"'); }

export function buildFeishuFields(config, record, actualFields, minimal = false) {
  const selectedKeys = minimal ? MINIMAL_FEISHU_FIELDS : FEISHU_FIELDS;
  const byName = new Map(actualFields.filter(field => field?.field_name).map(field => [field.field_name, field]));
  const fields = {}, unmatchedFields = [], missingRequired = [];
  for (const definition of FEISHU_FIELD_DEFINITIONS.filter(field => selectedKeys.includes(field.key))) {
    const targetName = String(config.mapping?.[definition.key] || definition.defaultName).trim();
    const actual = byName.get(targetName);
    if (!actual || isReadOnlyField(actual)) {
      unmatchedFields.push(targetName);
      if (definition.required) missingRequired.push(targetName);
      continue;
    }
    const sourceValue = record[definition.recordKey];
    const raw = definition.key === "status" && sourceValue === "已放弃" ? "放弃" : sourceValue;
    if (raw === undefined || raw === null || raw === "") continue;
    fields[actual.field_name] = convertFieldValue(raw, actual);
  }
  if (missingRequired.length) throw syncError({ msg: `必填字段不存在：${missingRequired.join("、")}`, unmatchedFields, sentFields: Object.keys(fields) });
  return { fields, unmatchedFields };
}

function convertFieldValue(value, field) {
  if (field.type === 5) { const timestamp = typeof value === "number" ? value : new Date(`${value}T00:00:00+08:00`).getTime(); return Number.isFinite(timestamp) ? timestamp : value; }
  if (field.type === 2) { const number = Number(value); return Number.isFinite(number) ? number : value; }
  if (field.type === 7) return value === true || value === "true" || value === "是";
  if (field.type === 15) return { link: String(value), text: String(value) };
  if (field.type === 4) return Array.isArray(value) ? value : String(value).split(/[，,、]/).map(item => item.trim()).filter(Boolean);
  return String(value);
}

function isReadOnlyField(field) { return new Set([19, 20, 1001, 1002, 1003, 1004, 1005]).has(Number(field.type)) || /Formula|Lookup|Created|Modified|AutoNumber/i.test(field.ui_type || ""); }

async function getTenantAccessToken(config) {
  const result = await requestJson("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", { method: "POST", headers: { "Content-Type": "application/json; charset=utf-8" }, body: JSON.stringify({ app_id: config.appId.trim(), app_secret: config.appSecret }) });
  if (!result.tenant_access_token) throw new Error("飞书未返回 tenant_access_token"); return result.tenant_access_token;
}

function validateFeishuCredentials(config) { if (!config.appId.trim()) throw new Error("请填写 App ID"); if (!config.appSecret.trim()) throw new Error("请填写 App Secret"); }
function parseTableLink(value) { if (!value?.trim()) return {}; let url; try { url = new URL(value.trim()); } catch { throw new Error("多维表格链接格式不正确"); } const parts = url.pathname.split("/").filter(Boolean), wikiIndex = parts.indexOf("wiki"), baseIndex = parts.indexOf("base"); return { wikiNodeToken: wikiIndex >= 0 ? parts[wikiIndex + 1] || "" : "", appToken: baseIndex >= 0 ? parts[baseIndex + 1] || "" : "", tableId: url.searchParams.get("table") || "" }; }

async function requestJson(url, options = {}) {
  const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 12000);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal }), text = await response.text(); let payload = {};
    try { payload = text ? JSON.parse(text) : {}; } catch { throw syncError({ httpStatus: response.status, msg: "飞书返回了无法解析的响应" }); }
    if (!response.ok || (payload.code != null && payload.code !== 0)) throw syncError({ httpStatus: response.status, code: payload.code, msg: payload.msg || "飞书请求失败" });
    return payload;
  } catch (error) {
    if (error.name === "AbortError") { const timeout = syncError({ msg: "请求超时，无法确认飞书是否已收到记录" }); timeout.unknown = true; throw timeout; }
    throw error;
  } finally { clearTimeout(timer); }
}

function enrichSyncError(error, unmatchedFields, sentFields) { error.unmatchedFields = unmatchedFields; error.sentFields = sentFields; error.message = formatError(error); return error; }
function syncError({ httpStatus = "-", code = "-", msg, unmatchedFields = [], sentFields = [] }) { const error = new Error(); error.httpStatus = httpStatus; error.code = code; error.feishuMsg = msg || "未知错误"; error.unmatchedFields = unmatchedFields; error.sentFields = sentFields; error.message = formatError(error); return error; }
function formatError(error) { return `HTTP状态码：${error.httpStatus ?? "-"}\n飞书code：${error.code ?? "-"}\nmsg：${error.feishuMsg || error.message || "未知错误"}\n未匹配字段：${error.unmatchedFields?.length ? error.unmatchedFields.join("、") : "无"}\n实际发送字段：${error.sentFields?.length ? error.sentFields.join("、") : "无"}`; }
