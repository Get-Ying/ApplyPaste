import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { saveApplicationRecords } from "../src/records.js";

class MemoryDirectory {
  constructor(root, name = "") { this.root = root; this.name = name; }
  async getDirectoryHandle(name) { return new MemoryDirectory(this.root, this.name ? `${this.name}/${name}` : name); }
  async getFileHandle(name) {
    const key = this.name ? `${this.name}/${name}` : name;
    return { createWritable: async () => ({ write: async data => { this.root.set(key, data); }, close: async () => {} }) };
  }
}

const files = new Map();
await saveApplicationRecords(new MemoryDirectory(files), [{
  "公司名称": "测试公司", "岗位名称": "材料研发工程师", "行业": "半导体", "岗位类别": "材料研发",
  "工作地点": "上海", "状态": "已投递", "投递日期": "2026-08-06", "完成时间": "2026-08-06 10:00:00",
  "岗位链接": "https://example.com/job", "使用简历": "测试岗位版", "填写版本": "测试岗位版", "中止原因": "", "备注": "测试"
}]);
for (const [name, data] of files) {
  const target = join(tmpdir(), `job-paste-${name.split("/").pop()}`);
  const bytes = data instanceof Blob ? new Uint8Array(await data.arrayBuffer()) : new TextEncoder().encode(String(data));
  await writeFile(target, bytes);
}
console.log([...files.keys()].join(","));
