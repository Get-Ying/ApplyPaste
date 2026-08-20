# ApplyPaste Feature Parity

基准为只读 `JOBPASTE` 正式自用版；目标为唯一 `APPLYPASTE` 公共目录。功能一致不代表源码逐字一致，空白首装与存储隔离属于必要公共化差异。

| 功能 | Job Paste | ApplyPaste | 状态 | 差异类型 | 结论 |
| -- | -- | -- | -- | -- | -- |
| Manifest V3 / 权限 | 支持 | 支持 | 一致 | - | PASS |
| Side Panel 注册 | 支持 | 支持 | 一致 | - | PASS |
| 数据库与 storage | 个人命名空间 | `ApplyPasteInternal` / `applyPaste...` | 独立 | B | PASS |
| 首次状态 | 个人数据 | 空分类结构、通用版、无内容 | 刻意不同 | B | PASS |
| 分类 / 卡片 / 字段 CRUD | 支持 | 支持 | 一致 | A | PASS |
| 字段结构化 | 支持冒号拆分 | 新增、导入及旧文本归一化 | 已补齐 | A | PASS |
| 字段类型模型 | 仅标签与文本值 | 同 | 一致 | D | 基准未实现，不新增架构 |
| select / multi-select options | 无通用 typed-field | 无 | 一致 | D | 无非法空 options |
| 搜索 / 分类 / 快速定位 | 支持 | 支持 | 一致 | A | PASS |
| 版本新建 / 重命名 / 复制 / 删除 | 支持 | 支持 | 一致 | A | PASS |
| override / 通用继承 / 恢复 / 差异 | 支持 | 支持 | 一致 | A | PASS |
| 用户自主字段版本化 | 事实字段有旧限制 | 普通字段可由用户决定 | ApplyPaste 更新 | C | 保留公共版规则 |
| Side Panel sticky / footer / 滚动记忆 | 支持 | 支持 | 一致 | A | PASS |
| Word 段落 / 表格 / bullet / preview | 支持 | 支持 | 一致 | A/C | PASS |
| Word 字段粒度 | 基准成熟逻辑 | 基础、教育、项目显式字段增强 | ApplyPaste 更新 | C | PASS |
| 附件 Blob / metadata / preview | 支持 | PDF、图片、DOCX 支持 | 一致 | A | PASS |
| 附件替换 / 删除 / 默认 / 版本绑定 | 支持 | 支持 | 一致 | A | PASS |
| 稍后继续 / 放弃 / 完成 | 支持 | 支持 | 一致 | A | PASS |
| 独立 localRecordId | 支持 | 支持 | 一致 | - | PASS |
| 岗位网页识别 | JSON-LD | JSON-LD + 保守 DOM fallback | ApplyPaste 更新 | C | PASS（4 fixtures） |
| 投递记录搜索 / 筛选 / 编辑 / CSV | 支持 | 支持 | 一致 | A | PASS |
| 飞书未配置本地使用 | 支持 | 支持 | 一致 | B | PASS |
| 飞书 mapping / CREATE / UPDATE / retry | 支持 | 支持 | 一致 | A/B | PASS（mock API） |
| JSON / Markdown / Word / backup | 支持 | 支持 | 一致 | A | PASS |
| public-template | 基准有 | 空白公共结构 | 已补齐 | A/B | PASS |
| 隐私扫描 / release gate | 基准有 | 已增强 | 已补齐 | A | PASS |

## 差异分类

- A：ApplyPaste 本应继承的通用能力，已恢复并回归。
- B：空白 profile、空附件/记录、通用版唯一初始版本、独立存储与默认无飞书配置，刻意保留。
- C：用户自主版本化、Word 原子字段与普通网页 DOM fallback 比基准更适合公共用户，保留。
- D：个人数据迁移、个人校园拆分、个人版本/附件推断及未被基准实现的 typed-field 架构不进入公共版。

## 测试层级

- `STRUCTURAL PASS`：manifest、存储命名空间、DOM、逻辑入口、空模板及字段粒度门禁。
- `BEHAVIOR PASS`：隔离、字段指引、版本 14 项、Side Panel 16 项、固定、Word 导入、飞书与记录。
- `E2E PASS`：未打包扩展在全新 Chromium profile 中完成主路径与重启持久化。
- `EDGE MANUAL CHECK`：品牌 Edge 禁止自动化命令行加载扩展，最终侧边栏外壳按人工清单复核。
