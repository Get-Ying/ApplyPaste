import { sectionResumeBlocks, sectionToCandidates } from "../src/resume-docx.js";

const block = text => ({ type: "paragraph", text, style: "", bold: false, bullet: false, level: 0 });

function parseBlocks(lines) {
  return sectionResumeBlocks(lines.map(block)).flatMap(sectionToCandidates);
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

{
  const candidates = parseBlocks([
    "教育背景",
    "复合材料大学",
    "材料科学学院",
    "专业：材料科学与工程",
    "2024.09-2027.06",
    "示例师范大学",
    "化学学院",
    "专业：化学",
    "2020.09-2024.06",
    "技能",
    "Python",
    "Origin"
  ]);
  const education = candidates.filter(item => item.category === "教育背景");
  const skills = candidates.filter(item => item.category === "技能");
  assert(education.length === 2, `expected two education cards, got ${education.length}: ${education.map(item => item.title).join(" | ")}`);
  assert(skills.length === 1, "skills section should become its own card");
  assert(!education.some(item => /技能|Python|Origin/.test(item.content)), "skills leaked into education card");
  assert(!education.some(item => item.title.length > 48), "education title is abnormally long");
}

{
  const candidates = parseBlocks([
    "项目经历",
    "界面材料测试平台",
    "项目时间：2025.01-2025.06",
    "项目描述：搭建公开测试流程并输出报告",
    "网申资料整理工具",
    "项目时间：2025.07-2025.08",
    "项目描述：整理字段、版本和投递记录"
  ]);
  const projects = candidates.filter(item => item.category === "项目经历");
  assert(projects.length === 2, `expected two project cards, got ${projects.length}: ${projects.map(item => item.title).join(" | ")}`);
  assert(projects.some(item => item.title.includes("界面材料测试平台")), "first project title missing");
  assert(projects.some(item => item.title.includes("网申资料整理工具")), "second project title missing");
  assert(!projects.some(item => item.title.length > 48), "project title is abnormally long");
}

{
  const email = ["user", "example.test"].join("@");
  const candidates = parseBlocks([
    "基础信息",
    "姓名：测试用户",
    "手机：13800000000",
    `邮箱：${email}`,
    "性别：女",
    "所在地：上海",
    "求职方向：研发工程师"
  ]);
  const basic = candidates.find(item => item.category === "基础信息");
  const labels = new Set(basic.fields.map(field => field.label));
  assert(labels.has("姓名"), "name should be extracted as a field");
  assert(labels.has("手机号码"), "phone should be extracted as a field");
  assert(labels.has("邮箱"), "email should be extracted as a field");
  assert(!labels.has("简历原文"), "basic info should not fall back to resume raw text field");
  assert(basic.fields.find(field => field.label === "手机号码")?.value === "13800000000", "phone value was not normalized");
  assert(basic.fields.find(field => field.label === "邮箱")?.value === email, "email value was not extracted");
}

console.log("ApplyPaste Word parser boundary fixtures verified");
