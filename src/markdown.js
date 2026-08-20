export function parseMarkdownLibrary(markdownText) {
  const categories = [
    "基础信息",
    "教育背景",
    "成绩课程",
    "荣誉奖项",
    "论文专利",
    "项目经历",
    "技能",
    "自我评价",
    "兴趣爱好",
    "职业规划",
    "常见网申问题"
  ];

  const rules = [
    [/个人|快速核对|联系|家庭/, "基础信息"],
    [/教育|学习经历/, "教育背景"],
    [/成绩|课程|绩点|排名/, "成绩课程"],
    [/荣誉|获奖|奖学金|竞赛/, "荣誉奖项"],
    [/论文|专利|成果/, "论文专利"],
    [/项目|实习|实践/, "项目经历"],
    [/技能|语言|证书|软件/, "技能"],
    [/自我|优势|优缺点/, "自我评价"],
    [/兴趣|爱好|特长/, "兴趣爱好"],
    [/职业规划|求职意向|公司期望/, "职业规划"],
    [/开放题|为什么|挑战|团队|附加信息|判断题|常见/, "常见网申问题"]
  ];

  const sections = [];
  const lines = markdownText.split(/\r?\n/);
  let current = null;

  for (const line of lines) {
    const heading = /^(#{2,3})\s+(.+)$/.exec(line);
    if (heading) {
      if (current) sections.push(current);
      current = { title: heading[2].trim(), lines: [] };
      continue;
    }
    if (current) current.lines.push(line);
  }
  if (current) sections.push(current);

  const items = sections
    .map((section, index) => {
      const content = section.lines.join("\n").replace(/^\s+|\s+$/g, "");
      if (!content) return null;
      const category = rules.find(([pattern]) => pattern.test(section.title))?.[1] || "常见网申问题";
      return {
        id: `md-${Date.now()}-${index}`,
        category,
        title: section.title.replace(/^\d+[A-Z]?\.\s*/, ""),
        content,
        tags: ["md导入"]
      };
    })
    .filter(Boolean);

  return { version: 1, updatedAt: new Date().toISOString(), categories, items };
}

export function toMarkdownLibrary(profile) {
  const groups = new Map(profile.categories.map(category => [category, []]));
  for (const item of profile.items) {
    if (!groups.has(item.category)) groups.set(item.category, []);
    groups.get(item.category).push(item);
  }

  const parts = ["# ApplyPaste 信息库", "", `> 更新时间：${new Date().toISOString()}`, ""];
  for (const [category, items] of groups) {
    parts.push(`## ${category}`, "");
    if (!items.length) {
      parts.push("_暂无条目_", "");
      continue;
    }
    for (const item of items) {
      parts.push(`### ${item.title}`, "");
      parts.push(item.content.trim(), "");
      if (item.tags?.length) parts.push(`标签：${item.tags.join("、")}`, "");
    }
  }
  return parts.join("\n");
}
