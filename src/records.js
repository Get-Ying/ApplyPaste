const HEADERS = [
  "记录ID", "公司名称", "岗位名称", "工作地点", "行业", "岗位类别", "岗位JD", "状态",
  "投递日期", "完成时间", "岗位链接", "页面标题", "使用简历", "填写版本", "填写版本ID",
  "内推码", "实际投递渠道", "招聘类型/批次", "企业性质", "投递优先级", "有效性修正",
  "中止原因", "备注", "创建时间", "最近更新时间", "飞书同步状态", "飞书记录ID"
];

export function createApplicationCsv(records = []) {
  const rows = [HEADERS, ...records.map(record => HEADERS.map(header => record?.[header] || ""))];
  return "\ufeff" + rows.map(csvRow).join("\r\n");
}

export async function saveApplicationRecords(workspaceHandle, records) {
  if (!workspaceHandle) return false;
  const rows = [HEADERS, ...records.map(record => HEADERS.map(header => record[header] || ""))];
  await writeFile(workspaceHandle, "records/applications.csv", "\ufeff" + rows.map(csvRow).join("\r\n"));
  await writeFile(workspaceHandle, "records/applications.xlsx", createXlsx(rows));
  return true;
}

function csvRow(row) {
  return row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(",");
}

async function writeFile(root, path, data) {
  const parts = path.split("/");
  const name = parts.pop();
  let dir = root;
  for (const part of parts) dir = await dir.getDirectoryHandle(part, { create: true });
  const handle = await dir.getFileHandle(name, { create: true });
  const writable = await handle.createWritable();
  await writable.write(data);
  await writable.close();
}

function createXlsx(rows) {
  const files = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="投递记录" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": worksheetXml(rows)
  };
  return zipStored(files);
}

function worksheetXml(rows) {
  const body = rows.map((row, rowIndex) => `<row r="${rowIndex + 1}">${row.map((value, colIndex) => {
    const cell = `${columnName(colIndex)}${rowIndex + 1}`;
    return `<c r="${cell}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
  }).join("")}</row>`).join("");
  const range = `A1:${columnName(Math.max(0, (rows[0]?.length || 1) - 1))}${Math.max(1, rows.length)}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="${range}"/><sheetData>${body}</sheetData></worksheet>`;
}

function columnName(index) {
  let name = "";
  for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

function escapeXml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function zipStored(files) {
  const encoder = new TextEncoder();
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = encoder.encode(name);
    const data = typeof content === "string" ? encoder.encode(content) : content;
    const crc = crc32(data);
    const local = concat(u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(0), u16(0), u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), nameBytes, data);
    locals.push(local);
    const central = concat(u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0), u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), nameBytes);
    centrals.push(central);
    offset += local.length;
  }
  const centralBlock = concat(...centrals);
  const end = concat(u32(0x06054b50), u16(0), u16(0), u16(centrals.length), u16(centrals.length), u32(centralBlock.length), u32(offset), u16(0));
  return new Blob([concat(...locals, centralBlock, end)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(value) {
  return new Uint8Array([value & 255, (value >>> 8) & 255]);
}

function u32(value) {
  return new Uint8Array([value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255]);
}

function concat(...arrays) {
  const output = new Uint8Array(arrays.reduce((sum, array) => sum + array.length, 0));
  let offset = 0;
  for (const array of arrays) {
    output.set(array, offset);
    offset += array.length;
  }
  return output;
}
