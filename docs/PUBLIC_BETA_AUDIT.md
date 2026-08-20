# ApplyPaste Public Beta Audit

## 1. 最终状态

READY WITH WARNINGS

核心 P0 已通过。唯一 warning 是品牌版 Edge 禁止自动化命令行加载未打包扩展，因此 Edge 原生 Side Panel 外壳保留 10 步人工复核；同内核真实扩展运行时、业务行为与持久化均已自动验证。

## 2. Job Paste 与 ApplyPaste 功能差异

ApplyPaste 已继承信息库、版本、Side Panel、Word、附件、投递生命周期、投递记录、飞书可选同步、导入导出与备份。公共版刻意不继承个人 profile、附件、记录、飞书配置、个人版本、个人迁移与推断逻辑。ApplyPaste 额外保留用户自主字段版本化、Word 原子字段和普通岗位页 DOM fallback。

## 3. 本轮发现的 P0

- 字段指引 DOM 未接线，用户无法知道缺失资料应如何拆字段。
- 旧文本/JSON 内容未自动结构化，存在多个网申单元塞入一个字段的风险。
- Word 测试未确认导入，且基础/教育/项目粒度不足。
- 投递弹窗缺少成熟节点，默认简历由 input 改成空 select 后不可用。
- Side Panel 固定只调用 open，未显式启用当前标签路径。
- 普通岗位页没有 JSON-LD 时无法尽量识别公司、岗位、地点与 JD。
- public-template、隐私门禁、功能门禁与 release 构建缺失。
- 遗留测试包含原项目私人 fixture。

## 4. 本轮实际修改

- `src/app.js`：公共分类、字段指引、结构化归一化、Side Panel 固定、投递表单接线、岗位页 fallback。
- `src/resume-docx.js`：基础信息命名、标题忽略、标签拆分、教育专业/日期、项目概况/职责/成果。
- `popup.html`：空状态入口、投递与记录缺失字段、默认简历 datalist。
- `data/`、`public-template/`：空白公共结构。
- `tests/fixtures/`：四种虚构岗位页。
- `tools/`：便携 Playwright 加载、隐私扫描、功能矩阵、扩展 E2E、公开构建及测试 fixture 清洗。
- `README.md`、`docs/USER_GUIDE.md`、`docs/`：分享说明与审计。

## 5. 未修改内容

未修改 Job Paste。未新增账号、云端、AI、自动投递、公司父子模型、typed-field 架构或 IndexedDB schema；这些不属于 v0.1 收敛范围或基准并未实现。

## 6. 隐私扫描

最终门禁要求：空 profile、空附件、空记录、通用版唯一初始版本；无原用户姓名、电话、邮箱、学校、项目、附件、投递记录、飞书 credential 或私人绝对路径。测试只允许 `张三 / 13800000000 / test@example.com` 等明确虚构值。

## 7. 自动测试

- STRUCTURAL PASS：27 项。
- BEHAVIOR PASS：隔离、字段指引、P1 核心、Side Panel 固定、Word 共 5 套浏览器测试。
- 版本管理：14 项 PASS。
- Side Panel：16 项 PASS。
- 飞书：字段读取、最小/完整同步、CREATE/UPDATE 幂等与超时重试 PASS。
- 记录导出：CSV/XLSX PASS。

## 8. E2E

全新临时 profile 中：空白首装、信息库、版本 override、Side Panel 五项、搜索/分类/快速定位、PDF/PNG/DOCX 上传与预览、替换、删除、默认简历、未配置飞书本地投递、记录、关闭并重启持久化均 PASS。四种岗位页 fixture 识别 PASS。

## 9. 已知限制

- 品牌 Edge 拦截自动化命令行加载未打包扩展，Edge 原生 Side Panel 外壳仍需按 `MANUAL_E2E_CHECKLIST.md` 做一次人工表面复核。
- DOCX 预览为文本近似，不复现完整 Word 排版。
- 招聘网页结构差异很大，识别失败时允许用户确认填写。

## 10. 下一步

分享给 1-3 名 Beta 用户前，在目标 Edge 上完成 10 步人工清单并记录结果。
