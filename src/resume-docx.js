import { docxToBlocks } from "./docx.js";

const SECTION_RULES = [
  ["基础信息", /^(基础信息|基本信息|个人信息|个人资料|联系方式|personal\s*(information|details)|contact(\s*information)?)$/i],
  ["教育背景", /^(教育背景|教育经历|学习经历|education(al\s*(background|experience))?)$/i],
  ["实习经历", /^(实习经历|工作经历|职业经历|实践经历|internship(s)?(\s*experience)?|work(ing)?\s*experience|professional\s*experience|employment)$/i],
  ["项目经历", /^(项目经历|项目经验|科研经历|研究经历|课题经历|科研项目|科研项目经历|research(\s*experience|\s*projects?)?|project(s)?(\s*experience)?)$/i],
  ["校园经历", /^(校园经历|学生工作|社会实践|志愿经历|campus\s*experience|leadership|activities)$/i],
  ["论文专利", /^(论文|论文成果|学术成果|科研成果|发表论文|专利|论文[\s/、与和]*专利|publications?|research\s*output|patents?)$/i],
  ["荣誉奖项", /^(获奖经历|荣誉奖项|奖项|竞赛经历|奖励|honou?rs?(\s*(and|&)\s*awards?)?|awards?|competitions?)$/i],
  ["技能", /^(技能|专业技能|语言能力|证书|资格证书|skills?|technical\s*skills?|languages?|certificates?|qualifications?)$/i],
  ["自我评价", /^(自我评价|个人总结|个人简介|自我介绍|summary|profile|personal\s*(summary|statement)|objective)$/i]
];

const FACT_LABELS = /姓名|电话|邮箱|城市|学校|学院|专业|学历|学位|时间|GPA|绩点|排名|公司|岗位|角色|项目名称|论文题目|期刊|作者|状态|DOI|奖项|证书/i;

export async function parseResumeDocx(file) {
  if (!/\.docx$/i.test(file.name || "")) throw new Error("文件无法读取：仅支持 .docx Word 简历");
  let blocks;
  try {
    blocks = await docxToBlocks(file);
  } catch (error) {
    throw new Error(error.message?.startsWith("DOCX 格式异常") ? error.message : `DOCX 格式异常：${error.message}`);
  }
  if (!blocks.length) throw new Error("未识别到文本：Word 正文为空");
  const sections = sectionResumeBlocks(blocks);
  const candidates = sections.flatMap(sectionToCandidates);
  if (!candidates.length) throw new Error("未识别到简历结构：请检查文档是否包含可读取的段落或表格文本");
  return { fileName: file.name, blocks, candidates, summary: summarizeCandidates(candidates) };
}

export function sectionResumeBlocks(blocks) {
  const sections = [];
  let current = { category: "基础信息", heading: "基础信息", confidence: "medium", lines: [] };
  let seenHeading = false;
  const push = () => {
    if (current.lines.length) sections.push(current);
  };
  for (const block of blocks) {
    if (/^(?:ApplyPaste|Job Paste)?\s*(?:信息库|简历|resume)$/i.test(block.text.trim())) continue;
    const category = detectSection(block.text);
    if (category) {
      push();
      current = { category, heading: block.text, confidence: "high", lines: [] };
      seenHeading = true;
      continue;
    }
    if (looksUnknownHeading(block)) {
      push();
      current = { category: "待确认", heading: block.text, confidence: "low", lines: [] };
      seenHeading = true;
      continue;
    }
    current.lines.push({ ...block, text: block.text.trim() });
  }
  push();
  if (!seenHeading && sections.length === 1) sections[0].confidence = "low";
  return sections;
}

export function detectSection(text) {
  const normalized = String(text || "")
    .trim()
    .replace(/[：:｜|·•\s]+$/g, "")
    .replace(/^[一二三四五六七八九十\d.、\s]+/, "")
    .trim();
  for (const [category, pattern] of SECTION_RULES) {
    if (pattern.test(normalized)) return category;
  }
  return "";
}

function looksUnknownHeading(block) {
  const text = block.text.trim();
  if (/^(?:ApplyPaste|Job Paste)?\s*(?:信息库|简历|resume)$/i.test(text)) return false;
  if (text.length > 36 || hasTime(text) || /@|\d{7,}/.test(text)) return false;
  return /heading/i.test(block.style || "")
    || (/^[A-Z][A-Z\s/&-]{3,}$/.test(text) && block.bold)
    || (block.bold && /经历|背景|成果|能力|奖项|评价|总结|信息$/.test(text));
}

export function sectionToCandidates(section) {
  if (!section.lines.length) return [];
  section = truncateLeakedSection(section);
  if (section.category === "基础信息") return [basicCandidate(section)];
  if (section.category === "技能") return [listCandidate(section, "技能")];
  if (section.category === "自我评价") return [expressionCandidate(section, "自我评价", "自我评价")];
  if (section.category === "待确认") {
    return splitAtomicLines(section).map((lines, index) => candidateFromLines(section, lines, `${section.heading || "待确认内容"}${index ? ` ${index + 1}` : ""}`));
  }
  if (["论文专利", "荣誉奖项"].includes(section.category)) {
    return splitAtomicLines(section).map(lines => candidateFromLines(section, lines));
  }
  return splitExperienceEntries(section).flatMap(lines => guardedCandidateFromLines(section, lines));
}

function basicCandidate(section) {
  const source = linesText(section.lines);
  const fields = [];
  const consumed = new Set();
  const add = (label, value, line) => {
    if (!value || fields.some(field => field.label === label)) return;
    fields.push(field(label, value, false));
    if (line) consumed.add(line);
  };
  for (const line of section.lines) {
    const text = line.text;
    const email = text.match(/[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/)?.[0];
    if (email) add("邮箱", email, line);
    const phone = text.match(/(?<!\d)(?:\+?86[-\s]?)?1[3-9]\d[-\s]?\d{4}[-\s]?\d{4}(?!\d)/)?.[0]?.replace(/[^\d+]/g, "").replace(/^\+?86/, "");
    if (phone && /^1[3-9]\d{9}$/.test(phone)) add("手机号码", phone, line);
    const labeled = text.match(/^(姓名|手机(?:号码)?|电话|邮箱|电子邮箱|性别|出生日期|出生年月|所在城市|现居地|所在地)[：:]\s*(.+)$/i);
    if (labeled) {
      const labels = { 手机: "手机号码", 电话: "手机号码", 电子邮箱: "邮箱", 出生年月: "出生日期", 现居地: "所在地", 所在城市: "所在地" };
      add(labels[labeled[1]] || labeled[1], cleanBasicValue(labeled[2]), line);
    }
  }
  const first = section.lines.find(line => !consumed.has(line) && /^[\u4e00-\u9fa5·]{2,8}$/.test(line.text));
  if (first) add("姓名", first.text, first);
  const remaining = section.lines.filter(line => !consumed.has(line)).map(line => lineText(line)).join("\n");
  if (remaining) fields.push(field(fields.length ? "其他信息" : "待确认", remaining, false));
  return makeCandidate("基础信息", fields.find(item => item.label === "姓名")?.value || "简历基础信息", fields, source, fields.length > 1 ? "high" : "medium");
}

function listCandidate(section, title) {
  const fields = section.lines.map((line, index) => field(`${title} ${index + 1}`, lineText(line), false));
  return makeCandidate(section.category, title, fields, linesText(section.lines), section.confidence);
}

function expressionCandidate(section, title, label) {
  return makeCandidate(section.category, title, [field(label, linesText(section.lines), true)], linesText(section.lines), section.confidence);
}

function splitAtomicLines(section) {
  const rows = groupTableRows(section.lines);
  if (rows.length > 1) return rows;
  return section.lines.map(line => [line]);
}

function splitExperienceEntries(section) {
  const lines = section.lines;
  const entries = [];
  const current = [];
  let hasBullet = false;
  let lastTableRow = null;
  const flush = () => {
    if (current.length) entries.push(current.splice(0));
    hasBullet = false;
  };
  for (const line of lines) {
    const tableRow = line.tableIndex != null ? `${line.tableIndex}:${line.rowIndex}` : null;
    const sameTableRow = tableRow != null && tableRow === lastTableRow;
    const tableChanged = tableRow != null && lastTableRow != null && tableRow !== lastTableRow;
    const newAfterBullets = hasBullet && !line.bullet;
    const newStrongHeader = !sameTableRow && current.length && !line.bullet && isNewEntryBoundary(line.text, current, section.category);
    if (tableChanged || newAfterBullets || newStrongHeader) flush();
    current.push(line);
    hasBullet ||= line.bullet;
    if (tableRow != null) lastTableRow = tableRow;
  }
  flush();
  return entries;
}

function guardedCandidateFromLines(section, lines) {
  if (section.category !== "教育背景") return [candidateFromLines(section, lines)];
  const groups = splitMultipleEducationEntities(lines);
  return groups.map(group => candidateFromLines(section, group));
}

function candidateFromLines(section, lines, fallbackTitle = "") {
  const source = linesText(lines);
  const rawTitle = chooseTitle(lines, section.category) || fallbackTitle || section.heading || "待确认内容";
  const title = sanitizeCandidateTitle(rawTitle, section, source);
  const fields = [];
  const time = lines.map(line => line.text).join(" ").match(/(?:19|20)\d{2}[./年-]\d{1,2}(?:月)?\s*(?:[-–—至到]|\s至\s)\s*(?:(?:19|20)\d{2}[./年-]\d{1,2}(?:月)?|至今|现在|Present|Current)/i)?.[0];

  if (section.category === "教育背景") {
    fields.push(field("学校名称", title, false));
    const college = source.match(/(?:学院|院系|School|College)\s*[：:]?\s*([^|｜；;\n]+)/i)?.[1]?.trim();
    if (college) fields.push(field("学院", college, false));
    const major = source.match(/(?:专业|Major)\s*[：:]?\s*([^|｜；;\n]+)/i)?.[1]?.trim();
    if (major) fields.push(field("专业", major, false));
    const degree = source.match(/博士|硕士|本科|学士|Ph\.?D\.?|Master|Bachelor/i)?.[0];
    if (degree) fields.push(field("学历/学位", degree, false));
    const gpa = source.match(/GPA\s*[：:]?\s*[\d.]+(?:\s*\/\s*[\d.]+)?/i)?.[0];
    if (gpa) fields.push(field("GPA", gpa.replace(/^GPA\s*[：:]?\s*/i, ""), false));
    const rank = source.match(/(?:排名|Rank)\s*[：:]?\s*[^|｜；;\n]+/i)?.[0];
    if (rank) fields.push(field("排名", rank.replace(/^(排名|Rank)\s*[：:]?\s*/i, ""), false));
  } else if (section.category === "实习经历") {
    fields.push(field("公司/经历名称", title, false));
  } else if (["项目经历", "校园经历"].includes(section.category)) {
    fields.push(field("项目名称", title, false));
  } else if (section.category === "论文专利") {
    fields.push(field("论文/专利原文", source, false));
  } else if (section.category === "荣誉奖项") {
    fields.push(field("奖项原文", source, false));
  } else {
    fields.push(field("原文", source, false));
  }

  if (time && !["论文专利", "荣誉奖项"].includes(section.category)) fields.push(field("时间", time, false));
  const bullets = lines.filter(line => line.bullet).map(line => `• ${line.text}`).join("\n");
  const detailLines = lines.filter(line => !line.bullet && line.text !== title && (!time || !line.text.includes(time))).map(line => line.text);
  if (section.category === "项目经历") {
    const projectLabels = { 项目概况: "项目概况", 项目背景: "项目背景", 项目职责: "项目职责", 主要职责: "项目职责", 个人贡献: "个人贡献", 项目成果: "项目成果", 成果: "项目成果" };
    for (const text of detailLines) {
      const labeled = text.match(/^(项目概况|项目背景|项目职责|主要职责|个人贡献|项目成果|成果)[：:]\s*(.+)$/);
      if (labeled) fields.push(field(projectLabels[labeled[1]], labeled[2], true));
    }
  }
  const details = detailLines.filter(text => !/^(项目概况|项目背景|项目职责|主要职责|个人贡献|项目成果|成果)[：:]/.test(text)).join("\n");
  const expression = [details, bullets].filter(Boolean).join("\n");
  if (expression && !["教育背景", "论文专利", "荣誉奖项", "待确认"].includes(section.category)) {
    fields.push(field(section.category === "实习经历" ? "职责描述" : "项目描述", expression, true));
  }
  const suspicious = isSuspiciousCandidateTitle(title, source);
  if (suspicious && !fields.some(item => item.label === "待确认")) fields.push(field("待确认", suspicious, false));
  return makeCandidate(section.category, title, fields, source, section.category === "待确认" || suspicious ? "low" : title ? section.confidence : "medium");
}

function chooseTitle(lines, category) {
  const texts = lines.filter(line => !line.bullet).map(line => line.text);
  const source = texts.join(" | ");
  if (category === "教育背景") {
    return source.match(/[A-Za-z\u4e00-\u9fa5· ]{2,40}(?:大学|学院|University|College|Institute)/i)?.[0]?.trim() || cleanTitle(texts[0]);
  }
  if (category === "论文专利" || category === "荣誉奖项") return cleanTitle(lines[0]?.text);
  return cleanTitle(texts.find(text => !hasTime(text)) || texts[0]);
}

function cleanTitle(text = "") {
  return text.replace(/(?:19|20)\d{2}[./年-]\d{1,2}.*$/i, "").replace(/[|｜]+$/g, "").trim();
}

function cleanBasicValue(value) {
  return String(value || "").replace(/^[：:\s]+/, "").trim();
}

function truncateLeakedSection(section) {
  const index = section.lines.findIndex(line => detectSection(line.text));
  if (index <= 0) return section;
  return { ...section, lines: section.lines.slice(0, index), confidence: "low" };
}

function isNewEntryBoundary(text, current, category) {
  if (category === "教育背景") return isEducationSchoolBoundaryLine(text) && current.some(line => isEducationSchoolBoundaryLine(line.text));
  if (category === "项目经历") return isProjectEntryHeader(text) && current.some(line => isProjectEntryHeader(line.text) || line.bullet || /^[^：:]{2,50}[：:]/.test(line.text));
  if (category === "实习经历") return isInternshipEntryHeader(text) && current.some(line => isInternshipEntryHeader(line.text));
  return isLikelyEntryHeader(text, category) && current.some(item => isLikelyEntryHeader(item.text, category));
}

function splitMultipleEducationEntities(lines) {
  const groups = [];
  let current = [];
  for (const line of lines) {
    if (current.length && isEducationSchoolBoundaryLine(line.text) && current.some(item => isEducationSchoolBoundaryLine(item.text))) {
      groups.push(current);
      current = [];
    }
    current.push(line);
  }
  if (current.length) groups.push(current);
  return groups;
}

function isEducationSchoolLine(text) {
  const value = String(text || "").trim();
  if (!value || /^[^：:]{1,12}[：:]/.test(value)) return false;
  return /(?:大学|学院|学校|University|College|Institute)/i.test(value) && value.length <= 80;
}

function isEducationSchoolBoundaryLine(text) {
  const value = String(text || "").trim();
  if (!isEducationSchoolLine(value)) return false;
  return /(?:大学|学校|University|College|Institute)/i.test(value);
}

function isProjectEntryHeader(text) {
  const value = String(text || "").trim();
  if (!value || lineLooksLikeLabeledField(value) || detectSection(value)) return false;
  if (hasTime(value)) return true;
  return value.length <= 60 && !/[。；;，,]/.test(value);
}

function isInternshipEntryHeader(text) {
  const value = String(text || "").trim();
  if (!value || lineLooksLikeLabeledField(value) || detectSection(value)) return false;
  return hasTime(value) || /公司|集团|实验室|研究院|Company|Lab|Institute|Intern/i.test(value);
}

function lineLooksLikeLabeledField(text) {
  return /^(项目名称|项目时间|项目描述|项目概况|项目背景|项目职责|主要职责|个人贡献|项目成果|成果|专业|学院|学历|学位|GPA|排名|职责描述|工作内容|使用工具|量化成果)[：:]/.test(text);
}

function sanitizeCandidateTitle(title, section, source) {
  const clean = cleanTitle(title);
  if (!isSuspiciousCandidateTitle(clean, source)) return clean;
  if (section.category === "教育背景") return source.match(/[A-Za-z\u4e00-\u9fa5· ]{2,40}(?:大学|学校|University|College|Institute)/i)?.[0]?.trim() || "待确认教育经历";
  if (section.category === "项目经历") return source.split(/\n/).map(cleanTitle).find(isProjectEntryHeader) || "待确认项目";
  return "待确认内容";
}

function isSuspiciousCandidateTitle(title, source) {
  const value = String(title || "").trim();
  if (value.length > 48) return "标题过长，请确认是否合并了多段内容";
  const headingHits = String(`${value}\n${source || ""}`).split(/\n/).filter(line => detectSection(line)).length;
  if (headingHits > 0 && detectSection(value)) return "标题包含章节标题，请确认识别边界";
  return "";
}

function isLikelyEntryHeader(text, category) {
  if (hasTime(text)) return true;
  if (category === "教育背景") return isEducationSchoolLine(text);
  if (category === "实习经历") return /公司|实验室|研究院|Company|Lab|Institute|Intern/i.test(text);
  return text.length <= 80 && !lineLooksLikeLabeledField(text) && !/[。；;]/.test(text);
}

function hasTime(text) {
  return /(?:19|20)\d{2}[./年-]\d{1,2}|至今|Present|Current/i.test(text);
}

function lineText(line) {
  return line.bullet ? `${"  ".repeat(line.level || 0)}• ${line.text}` : line.text;
}

function linesText(lines) {
  return lines.map(lineText).join("\n");
}

function groupTableRows(lines) {
  const groups = new Map();
  for (const line of lines) {
    const key = line.tableIndex == null ? `p:${groups.size}` : `t:${line.tableIndex}:${line.rowIndex}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(line);
  }
  return [...groups.values()];
}

function field(label, value, versioned) {
  return { label, value: String(value || "").trim(), versioned };
}

function makeCandidate(category, title, fields, sourceText, confidence) {
  const cleanFields = fields.filter(item => item.value);
  return {
    id: crypto.randomUUID(),
    category,
    title: title || "待确认内容",
    fields: cleanFields,
    content: cleanFields.map(item => `${item.label}：${item.value}`).join("\n"),
    sourceText,
    confidence: confidence === "high" ? "high" : confidence === "low" ? "low" : "medium"
  };
}

function summarizeCandidates(candidates) {
  const counts = {};
  for (const item of candidates) counts[item.category] = (counts[item.category] || 0) + 1;
  return {
    counts,
    high: candidates.filter(item => item.confidence === "high").length,
    confirm: candidates.filter(item => item.confidence !== "high").length,
    total: candidates.length
  };
}

export function findLikelyResumeMatch(candidate, items) {
  let best = null;
  let bestScore = 0;
  for (const item of items || []) {
    const source = `${item.title}\n${item.content || ""}\n${(item.fields || []).map(field => field.value).join("\n")}`;
    const titleScore = Math.max(textSimilarity(candidate.title, item.title), textSimilarity(candidate.title, source));
    const tokenScore = sharedDistinctToken(candidate.sourceText, source) ? 0.62 : 0;
    const categoryBonus = item.category === candidate.category ? 0.15 : 0;
    const score = Math.min(1, Math.max(titleScore, tokenScore) + categoryBonus);
    if (score > bestScore) {
      best = item;
      bestScore = score;
    }
  }
  return bestScore >= 0.56 ? { item: best, score: bestScore } : null;
}

function textSimilarity(left, right) {
  const a = normalizeText(left);
  const b = normalizeText(right);
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.includes(b) || b.includes(a)) return 0.72;
  const grams = value => new Set([...value].slice(0, -1).map((char, index) => char + value[index + 1]));
  const x = grams(a);
  const y = grams(b);
  const common = [...x].filter(value => y.has(value)).length;
  return common / Math.max(1, x.size + y.size - common);
}

function sharedDistinctToken(left, right) {
  const a = new Set(String(left).match(/[A-Za-z][A-Za-z0-9_/-]{3,}/g) || []);
  return [...a].some(token => new RegExp(`\\b${escapeRegExp(token)}\\b`, "i").test(right));
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/硕士|本科|课题|研究|项目|经历|论文|实习|工作/g, "")
    .replace(/[^a-z0-9\u4e00-\u9fa5]/g, "");
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function isResumeFactField(field) {
  return FACT_LABELS.test(field.label);
}
