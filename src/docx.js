export function createProfileDocx(profile) {
  const paragraphs = [paragraph("ApplyPaste 信息库", "Title")];
  for (const category of profile.categories || []) {
    paragraphs.push(paragraph(category, "Heading1"));
    for (const item of (profile.items || []).filter(candidate => candidate.category === category)) {
      paragraphs.push(paragraph(item.title, "Heading2"));
      for (const line of String(item.content || "").split("\n")) paragraphs.push(paragraph(line));
    }
  }
  const files = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "word/styles.xml": stylesXml(),
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs.join("")}<w:sectPr/></w:body></w:document>`
  };
  return zipStored(files, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
}

export async function docxToMarkdown(file) {
  const xml = new TextDecoder().decode(await readZipEntry(new Uint8Array(await file.arrayBuffer()), "word/document.xml"));
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.querySelector("parsererror")) throw new Error("Word 文档结构无法解析");
  const lines = [];
  for (const node of document.getElementsByTagName("w:p")) {
    const text = [...node.getElementsByTagName("w:t")].map(item => item.textContent || "").join("").trim();
    if (!text) continue;
    const style = node.getElementsByTagName("w:pStyle")[0]?.getAttribute("w:val") || "";
    lines.push(/^Heading1$/i.test(style) ? `## ${text}` : /^Heading2$/i.test(style) ? `### ${text}` : text);
  }
  return lines.join("\n\n");
}

export async function docxToText(file) {
  return (await docxToMarkdown(file)).replace(/^#{2,3}\s+/gm, "").trim();
}

export async function docxToBlocks(file) {
  const xml = new TextDecoder().decode(await readZipEntry(new Uint8Array(await file.arrayBuffer()), "word/document.xml"));
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.querySelector("parsererror")) throw new Error("DOCX 格式异常：document.xml 无法解析");
  const body = [...document.documentElement.children].find(node => node.localName === "body");
  if (!body) throw new Error("DOCX 格式异常：缺少正文");
  const blocks = [];
  let tableIndex = 0;
  for (const node of body.children) {
    if (node.localName === "p") appendParagraphBlock(blocks, node, {});
    if (node.localName !== "tbl") continue;
    let rowIndex = 0;
    for (const row of [...node.children].filter(child => child.localName === "tr")) {
      let cellIndex = 0;
      for (const cell of [...row.children].filter(child => child.localName === "tc")) {
        for (const paragraph of [...cell.children].filter(child => child.localName === "p")) {
          appendParagraphBlock(blocks, paragraph, { tableIndex, rowIndex, cellIndex });
        }
        cellIndex++;
      }
      rowIndex++;
    }
    tableIndex++;
  }
  return blocks;
}

function appendParagraphBlock(blocks, paragraph, position) {
  const text = [...paragraph.getElementsByTagName("w:t")]
    .map(node => node.textContent || "")
    .join("")
    .replace(/\u00a0/g, " ")
    .trim();
  if (!text) return;
  const style = paragraph.getElementsByTagName("w:pStyle")[0]?.getAttribute("w:val") || "";
  const numbered = paragraph.getElementsByTagName("w:numPr").length > 0;
  const level = Number(paragraph.getElementsByTagName("w:ilvl")[0]?.getAttribute("w:val") || 0);
  const bold = paragraph.getElementsByTagName("w:b").length > 0;
  const explicitBullet = /^[•·*●○▪▫\-]\s*/.test(text);
  blocks.push({
    type: "paragraph",
    text: text.replace(/^[•·*●○▪▫\-]\s*/, ""),
    bullet: numbered || explicitBullet,
    level,
    style,
    bold,
    ...position
  });
}

function paragraph(text, style = "") {
  const properties = style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : "";
  return `<w:p>${properties}<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style></w:styles>`;
}

function escapeXml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export async function readZipEntry(bytes, targetName) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let index = bytes.length - 22; index >= Math.max(0, bytes.length - 65557); index--) {
    if (view.getUint32(index, true) === 0x06054b50) {
      eocd = index;
      break;
    }
  }
  if (eocd < 0) throw new Error("不是有效的 DOCX 文件");
  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  for (let index = 0; index < count; index++) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("DOCX 目录损坏");
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = new TextDecoder().decode(bytes.slice(offset + 46, offset + 46 + nameLength));
    if (name === targetName) {
      const localNameLength = view.getUint16(localOffset + 26, true);
      const localExtraLength = view.getUint16(localOffset + 28, true);
      const start = localOffset + 30 + localNameLength + localExtraLength;
      const compressed = bytes.slice(start, start + compressedSize);
      if (method === 0) return compressed;
      if (method === 8) {
        return new Uint8Array(await new Response(new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).arrayBuffer());
      }
      throw new Error(`不支持的 DOCX 压缩方式：${method}`);
    }
    offset += 46 + nameLength + extraLength + commentLength;
  }
  throw new Error("DOCX 缺少 word/document.xml");
}

function zipStored(files, mimeType) {
  const encoder = new TextEncoder();
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = encoder.encode(name);
    const data = encoder.encode(content);
    const crc = crc32(data);
    const local = concat(u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(0), u16(0), u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), nameBytes, data);
    locals.push(local);
    centrals.push(concat(u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0), u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), nameBytes));
    offset += local.length;
  }
  const central = concat(...centrals);
  const end = concat(u32(0x06054b50), u16(0), u16(0), u16(centrals.length), u16(centrals.length), u32(central.length), u32(offset), u16(0));
  return new Blob([concat(...locals, central, end)], { type: mimeType });
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
