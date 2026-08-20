import { createFeishuRecord, normalizeFeishuConfig, syncFeishuRecord, testFeishuConnection } from "../src/feishu.js";

const calls = [], sentRecords = [], remoteRecords = new Map(); let createCount = 0, updateCount = 0, timeoutNextCreate = false;
const actualFields = [
  field("公司名称"), field("岗位名称"), field("工作地点-base"), field("JD"), field("投递链接", 15), field("备注"), field("状态", 3),
  field("投递日期", 5), field("使用简历"), field("填写版本"), field("ApplyPaste记录ID")
];

globalThis.fetch = async (url, options = {}) => {
  calls.push({ url: String(url), options });
  if (String(url).includes("tenant_access_token")) return response({ code: 0, tenant_access_token: "tenant-token" });
  if (String(url).includes("wiki/v2/spaces/get_node")) return response({ code: 0, data: { node: { obj_type: "bitable", obj_token: "app_from_wiki" } } });
  if (String(url).includes("/fields?")) return response({ code: 0, data: { items: actualFields, has_more: false } });
  if ((!options.method || options.method === "GET") && String(url).includes("/records?")) return response({ code: 0, data: { items: [...remoteRecords.entries()].map(([record_id, fields]) => ({ record_id, fields })).slice(0, 1), has_more: false } });
  if (options.method === "PUT" && /\/records\/[^/?]+$/.test(String(url))) { const recordId = decodeURIComponent(String(url).split("/").pop()); remoteRecords.set(recordId, JSON.parse(options.body).fields); updateCount += 1; return response({ code: 0, data: { record: { record_id: recordId } } }); }
  if (options.method === "POST" && String(url).endsWith("/records")) { const fields = JSON.parse(options.body).fields, recordId = `rec-${++createCount}`; sentRecords.push(fields); remoteRecords.set(recordId, fields); if (timeoutNextCreate) { timeoutNextCreate = false; throw new DOMException("timeout", "AbortError"); } return response({ code: 0, data: { record: { record_id: recordId } } }); }
  return response({ code: 0, data: {} });
};

const config = normalizeFeishuConfig({ enabled: true, appId: "cli_test", appSecret: "secret", appToken: "app_test", tableId: "tbl_test", tableUrl: "https://example.feishu.cn/base/app_test?table=tbl_test" });
const originalConfig = structuredClone(config);
const migratedLegacy = normalizeFeishuConfig({ mapping: { "工作地点": "自定义地点", "岗位JD": "自定义JD", "岗位类别": "自定义岗位方向" } });
if (migratedLegacy.mapping.location !== "自定义地点" || migratedLegacy.mapping.jd !== "自定义JD") throw new Error("核心旧字段映射兼容迁移失败");
if ("jobCategory" in migratedLegacy.mapping) throw new Error("公共版仍生成私人化岗位类别映射");

const connection = await testFeishuConnection(config);
if (connection.fields.length !== actualFields.length || !connection.fields.every(item => item.field_name && item.field_id && item.type != null)) throw new Error("真实字段列表读取失败");
console.log("真实字段列表读取通过");

const record = {
  "记录ID": "JP-1", "公司名称": "测试公司", "岗位名称": "材料研发工程师", "工作地点": "上海", "岗位JD": "负责材料研发", "岗位链接": "https://example.com/job",
  "内推码": "", "备注": "测试备注", "状态": "已投递", "行业": "半导体", "岗位类别": "材料研发", "投递优先级": "高", "有效性修正": "有效",
  "实际投递渠道": "官网", "企业性质": "民营", "招聘类型/批次": "公开测试", "投递日期": "2026-08-06", "使用简历": "测试岗位版", "填写版本": "测试岗位版"
};
await createFeishuRecord(config, record, { mode: "minimal", fields: connection.fields });
if (JSON.stringify(Object.keys(sentRecords[0]).sort()) !== JSON.stringify(["公司名称", "岗位名称", "工作地点-base", "投递日期", "状态"].sort())) throw new Error("最小记录没有严格限制为 5 个字段");
console.log("最小记录同步通过");

await createFeishuRecord(config, record, { fields: connection.fields });
const full = sentRecords[1];
if (!full["公司名称"] || !full["JD"] || !full["ApplyPaste记录ID"] || typeof full["投递日期"] !== "number") throw new Error("完整记录字段构建失败");
if (["内推码", "内推人", "行业", "岗位类别/方向", "投递优先级", "有效性修正", "实际投递渠道", "企业性质", "招聘类型/批次"].some(key => key in full) || Object.values(full).some(value => value === undefined || value === null || value === "")) throw new Error("完整记录发送了私人化字段、空值、公式字段或无效值");
console.log("完整记录同步通过");

if (JSON.stringify(config) !== JSON.stringify(originalConfig)) throw new Error("原有飞书配置被调用过程覆盖");
console.log("原有飞书配置保持不变");

remoteRecords.clear(); createCount = 0; updateCount = 0;
const firstSync = await syncFeishuRecord(config, record);
if (firstSync.action !== "create" || createCount !== 1) throw new Error("首次同步未严格 CREATE 一次");
record["飞书记录ID"] = firstSync.recordId; record["备注"] = "更新内容";
const secondSync = await syncFeishuRecord(config, record);
if (secondSync.action !== "update" || createCount !== 1 || updateCount !== 1) throw new Error("已有飞书记录ID时未 UPDATE");
console.log("CREATE/UPDATE 幂等同步通过");

remoteRecords.clear(); createCount = 0; updateCount = 0; delete record["飞书记录ID"]; timeoutNextCreate = true;
let unknown = false; try { await syncFeishuRecord(config, record); } catch (error) { unknown = error.unknown === true; }
if (!unknown || createCount !== 1 || remoteRecords.size !== 1) throw new Error("网络超时场景未标记状态未知");
const retry = await syncFeishuRecord(config, record);
if (retry.action !== "update" || createCount !== 1 || updateCount !== 1) throw new Error("超时重试产生了重复 CREATE");
console.log("网络超时去重重试通过");

function field(field_name, type = 1) { return { field_name, field_id: `fld_${field_name}`, type, ui_type: type === 20 ? "Formula" : type === 5 ? "DateTime" : "Text" }; }
function response(payload) { return { ok: true, status: 200, text: async () => JSON.stringify(payload) }; }
