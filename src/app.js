import { downloadProfile, openProfileFile } from "./store.js";
import {
  readAttachmentBlob as readWorkspaceAttachmentBlob, readWorkspaceRecords, readWorkspaceState
} from "./workspace.js";
import { loadInternalState, saveAttachments, saveCompleteState, saveCoreData, saveRecords, storageErrorMessage } from "./internal-db.js";
import { DEFAULT_FEISHU_CONFIG, FEISHU_FIELDS, normalizeFeishuConfig, syncFeishuRecord, testFeishuConnection, validateFeishuConfig } from "./feishu.js";
import { docxToText } from "./docx.js";
import { findLikelyResumeMatch, isResumeFactField, parseResumeDocx } from "./resume-docx.js";
import { createApplicationCsv } from "./records.js";

const EXTRA_KEYS = ["applyPasteAttachments", "applyPasteSettings", "applyPasteRecords", "applyPasteUiSettings", "applyPasteFeishuConfig", "applyPasteWorkspaceMigrated", "applyPasteActiveRecordId"];
const DISPLAY_MODE = new URLSearchParams(location.search).get("mode") || "popup";
const VIEW_STATE_KEY = DISPLAY_MODE === "sidepanel" ? "applyPasteViewStateSidepanel" : "applyPasteViewStatePopup";
const VIEW_INSTANCE_ID = crypto.randomUUID();
const DEFAULT_CATEGORIES = ["基础信息", "教育背景", "成绩课程", "荣誉奖项", "论文专利", "项目经历", "实习经历", "工作经历", "校园经历", "技能", "自我评价", "开放题"];
const FIELD_GUIDES = {
  "基础信息": { title: "个人基本信息", tip: "适合整理网申常用信息。敏感证件号、家庭信息等建议只保留待填写提示。", fields: ["姓名", "手机号码", "邮箱", "所在城市", "最高学历", "毕业院校", "专业", "预计毕业时间", "求职方向"] },
  "教育背景": { title: "学校 / 学位", tip: "学校、专业和时间建议拆开填写，便于在不同网申表中逐项复制。", fields: ["学校名称", "学院", "专业", "学历/学位", "入学时间", "毕业时间", "GPA", "排名", "主修课程"] },
  "成绩课程": { title: "成绩 / 课程", tip: "课程、成绩与排名分别成字段，不确定的内容可以先留待补充。", fields: ["课程名称", "课程类别", "成绩", "满分", "排名", "学期", "备注"] },
  "实习经历": { title: "公司 / 岗位", tip: "建议按动作、方法和结果拆分，之后可按岗位版本改写表达字段。", fields: ["公司/组织", "岗位名称", "所在部门", "实习时间", "工作地点", "职责描述", "使用工具", "量化成果"] },
  "工作经历": { title: "公司 / 岗位", tip: "正式工作经历建议拆出负责范围、业务场景和成果指标。", fields: ["公司名称", "岗位名称", "任职时间", "工作地点", "业务方向", "职责描述", "关键成果", "离职原因"] },
  "项目经历": { title: "项目名称", tip: "项目经历适合拆成背景、职责、方法与成果，方便按招聘表单逐项填写。", fields: ["项目名称", "项目角色", "项目时间", "项目背景", "项目描述", "技术/方法", "个人贡献", "成果指标"] },
  "校园经历": { title: "组织 / 职务", tip: "突出协作、组织、沟通或领导力，避免只记录活动名称。", fields: ["组织名称", "职务/角色", "经历时间", "活动规模", "职责描述", "成果", "能力关键词"] },
  "荣誉奖项": { title: "奖项名称", tip: "一个奖项一条记录，时间、级别和证明材料分别填写。", fields: ["奖项名称", "颁发机构", "获奖时间", "奖项级别", "排名/比例", "证明材料"] },
  "论文专利": { title: "论文 / 专利名称", tip: "保留正式题名、作者顺序和状态，便于准确粘贴。", fields: ["题名", "作者顺序", "期刊/会议/专利号", "发表/申请时间", "状态", "DOI/链接", "成果简介"] },
  "技能": { title: "技能组", tip: "按工具、语言、实验方法或证书分组，并注明熟练程度或使用场景。", fields: ["技能类别", "技能名称", "熟练程度", "使用场景", "证书/等级"] },
  "自我评价": { title: "自我评价", tip: "用具体经历支撑能力描述，并保留适合不同字数限制的回答。", fields: ["能力关键词", "优势描述", "代表证据", "适配岗位", "短版回答", "长版回答"] },
  "开放题": { title: "题目", tip: "保存题目原文、回答逻辑和可提交答案，之后可按公司版本微调。", fields: ["题目原文", "回答要点", "短版答案", "长版答案", "相关经历", "注意事项"] }
};
const DEFAULT_FIELD_GUIDE = { title: "资料标题", tip: "按“字段名：答案”逐行填写，保存后会拆成可独立复制和编辑的字段。", fields: ["名称", "时间", "地点", "角色", "描述", "成果", "备注"] };
const DEFAULT_SETTINGS = {
  currentVersion: "general",
  celebration: true,
  versions: [createVersion("general", "通用版")]
};

const BLANK_PROFILE = {
  version: 6,
  updatedAt: null,
  categories: DEFAULT_CATEGORIES,
  items: []
};

const state = {
  profile: null, attachments: [], settings: structuredClone(DEFAULT_SETTINGS), records: [], feishuConfig: structuredClone(DEFAULT_FEISHU_CONFIG),
  selectedCategory: "全部", query: "", onlyDifferences: false, viewModeByVersion: {}, categoryScrollPositions: {}, searchRestore: null, editingId: null, editingField: null, editorInitial: null, fieldEditorInitial: null, fieldVersionPreferenceTouched: false, pendingDiscardTarget: null,
  editingAttachment: null, pendingAttachmentFile: null, pendingReplacement: null, pendingUseAttachmentId: null, previewObjectUrl: "", previewState: null, verifiedAttachmentIds: new Set(), confirmedResumeName: "", submissionMode: "complete",
  versionDialogMode: "add", statusField: null, editingRecordId: null, recordDetailMode: "view", syncResultRecordId: null, resumeImport: null,
  activeRecordId: "", recordFilter: "全部", recordQuery: "", pendingResumeRecordId: ""
};

const $ = id => document.getElementById(id);
const els = Object.fromEntries([
  "fileStatus","saveStatus","searchInput","currentVersion","differenceBar","differenceSummary","showAllBtn","showDifferencesBtn","categoryTabs","contentPane","attachmentPanel","attachmentList","list","empty","emptyNewItemBtn","emptyResumeImportBtn","notice",
  "pinSidePanelBtn","openWorkbenchBtn","sidepanelStickyNavigation","sidepanelCategoryControl","sidepanelCategory","quickNavigation","quickNavigationSelect","quickNavigationTop",
  "newItemBtn","manageCategoriesBtn","addAttachmentBtn","addVersionBtn","moveVersionUpBtn","moveVersionDownBtn","renameVersionBtn","copyVersionBtn","deleteVersionBtn",
  "categoryDialog","categoryList","categoryNotice","newCategoryName","addCategoryBtn","closeCategoryBtn",
  "editorDialog","editorForm","editorTitle","editCategory","editTitle","editTags","itemFieldGuide","editContent","itemEditorHint","deleteItemBtn","duplicateItemBtn",
  "fieldEditorDialog","fieldEditorForm","fieldEditorTitle","fieldEditLabel","fieldEditValue","fieldValueLabel","fieldVersioned","fieldEditorHint","fieldEditorCommonBlock","fieldEditorCommonPreview","useCommonFieldAnswerBtn","deleteFieldBtn","duplicateFieldBtn",
  "discardChangesDialog","continueEditingBtn","discardChangesBtn",
  "attachmentDialog","attachmentForm","attachmentDialogTitle","attachmentTitle","attachmentCategory","attachmentTags","attachmentSensitive","attachmentDefault","attachmentDefaultLabel","attachmentVersionOptions","deleteAttachmentBtn",
  "attachmentPreviewDialog","attachmentPreviewTitle","attachmentPreviewMeta","docxPreviewWarning","attachmentPreviewToolbar","attachmentPreviewBody","previewPrevBtn","previewPageLabel","previewNextBtn","previewZoomOutBtn","previewZoomInBtn","previewFitBtn","attachmentVerifiedStatus","markAttachmentVerifiedBtn","closeAttachmentPreviewBtn",
  "attachmentReplaceDialog","replaceOldFile","replaceNewFile","replaceNewMeta","previewReplacementBtn","cancelReplacementBtn","confirmReplacementBtn","closeAttachmentReplaceBtn",
  "attachmentUseDialog","attachmentUseName","attachmentUseMeta","attachmentMismatchWarning","previewBeforeUseBtn","cancelAttachmentUseBtn","confirmAttachmentUseBtn","closeAttachmentUseBtn",
  "resumeMismatchDialog","resumeMismatchSummary","cancelResumeMismatchBtn","confirmResumeMismatchBtn",
  "completeBtn","laterBtn","abortBtn","recordsBtn","pendingRecordBadge","feishuSettingsBtn",
  "completeDialog","completeForm","completeDialogTitle","completeSubmitBtn","completeCompany","completeJob","completeLocation","completeResume","completeIndustry","completeRole","completeChannel","completeBatch","completeCompanyType","completePriority","completeReferralCode","completeReferrer","completeUrl","completeJd","completeNotes",
  "abortDialog","abortForm","abortCompany","abortJob","abortLocation","abortUrl","abortReason","abortNotes",
  "companySuggestions","locationSuggestions","industrySuggestions","roleSuggestions","resumeSuggestions",
  "feishuDialog","feishuForm","feishuEnabled","feishuAppId","feishuAppSecret","feishuTableUrl","feishuAppToken","feishuTableId","mappingFields","feishuTestResult","clearFeishuBtn","testFeishuBtn",
  "resumeImportBtn","dataImportBtn","dataExportBtn","backupExportBtn","backupRestoreBtn","migrateWorkspaceBtn",
  "resumeImportDialog","resumeImportFile","resumeImportSummary","resumeImportVersion","resumeImportDefault","resumeDuplicatePanel","resumeDuplicateAction","resumeImportList","resumeImportError","closeResumeImportBtn","cancelResumeImportBtn","confirmResumeImportBtn",
  "recordsDialog","recordsList","todayRecordCount","recordStatusFilter","recordSearchInput","exportRecordsCsvBtn","closeRecordsBtn",
  "recordDetailDialog","recordDetailForm","recordDetailTitle","detailCompany","detailJob","detailLocation","detailStatus","detailIndustry","detailRole","detailChannel","detailBatch","detailCompanyType","detailPriority","detailReferralCode","detailReferrer","detailUrl","detailJd","detailResume","detailAbortReason","detailNotes","detailMetadata","localEditWarning","saveRecordEditBtn",
  "resumeRecordDialog","resumeRecordSummary","newRecordBtn","continueRecordBtn",
  "syncResultDialog","syncResultTitle","syncResultText","retrySyncBtn","viewFeishuBtn","laterSyncBtn","closeSyncResultBtn",
  "celebration","celebrationText",
  "versionDialog","versionForm","versionDialogTitle","versionName","versionBaseLabel","versionBase","versionDefaultResume",
  "fieldStatusDialog","fieldStatusTitle","fieldGeneralPreview","fieldCurrentBlock","fieldCurrentLabel","fieldCurrentPreview","promoteFieldBtn","restoreFieldBtn","closeFieldStatusBtn"
].map(id => [id, $(id)]));

init().catch(error => showNotice(`载入失败：${error.message}`));

async function init() {
  const [internal, extras] = await Promise.all([loadInternalState(), chrome.storage.local.get(EXTRA_KEYS)]);
  const hasInternalData = Boolean(internal.profile);
  state.profile = normalizeProfile(internal.profile || BLANK_PROFILE);
  const sourceSettings = hasInternalData
    ? { ...DEFAULT_SETTINGS, versions: internal.versions, ...(extras.applyPasteSettings || {}), ...(extras.applyPasteUiSettings || {}) }
    : { ...DEFAULT_SETTINGS, ...(extras.applyPasteSettings || {}), ...(extras.applyPasteUiSettings || {}) };
  state.settings = normalizeSettings(sourceSettings, state.profile);
  state.attachments = normalizeAttachments(hasInternalData ? internal.attachments : []);
  state.records = normalizeRecords(hasInternalData ? internal.records : []);
  if (!hasInternalData) {
    await saveCompleteState({ profile: state.profile, versions: state.settings.versions, records: state.records, attachments: state.attachments });
    await saveUiSettings();
  }
  state.feishuConfig = normalizeFeishuConfig(extras.applyPasteFeishuConfig);
  state.activeRecordId = state.records.some(record => record["记录ID"] === extras.applyPasteActiveRecordId) ? extras.applyPasteActiveRecordId : "";
  const savedView = (await chrome.storage.local.get(VIEW_STATE_KEY))[VIEW_STATE_KEY] || {};
  if (savedView.category && (savedView.category === "全部" || savedView.category === "附件" || state.profile.categories.includes(savedView.category))) state.selectedCategory = savedView.category;
  state.query = String(savedView.query || "").toLowerCase();
  state.viewModeByVersion = { ...(savedView.viewModeByVersion || {}) };
  state.categoryScrollPositions = { ...(savedView.categoryScrollPositions || {}) };
  state.onlyDifferences = state.settings.currentVersion !== "general" && state.viewModeByVersion[state.settings.currentVersion] === "diff";
  document.body.classList.add(DISPLAY_MODE === "sidepanel" ? "sidepanelMode" : DISPLAY_MODE === "workbench" ? "workbenchMode" : "popupMode");
  if (els.openWorkbenchBtn) els.openWorkbenchBtn.hidden = DISPLAY_MODE !== "sidepanel";
  if (els.sidepanelCategoryControl) els.sidepanelCategoryControl.hidden = DISPLAY_MODE !== "sidepanel";
  if (els.pinSidePanelBtn) els.pinSidePanelBtn.hidden = DISPLAY_MODE === "sidepanel";
  setupSidePanelLayout();
  els.searchInput.value = savedView.query || "";
  bindEvents();
  render();
  setTimeout(() => { els.contentPane.scrollTop = Number(savedView.scrollTop ?? state.categoryScrollPositions[state.selectedCategory]) || 0; }, 0);
  detectPendingRecord();
}

function bindEvents() {
  document.addEventListener("click", event => { const menu = event.target.closest(".versionMore, .recordMore, .attachmentMore, .dataImportMenu"); if (menu && event.target.closest("button")) menu.open = false; });
  els.searchInput.addEventListener("input", handleSearchInput);
  els.currentVersion.addEventListener("change", async event => {
    state.viewModeByVersion[state.settings.currentVersion] = state.onlyDifferences ? "diff" : "all";
    state.settings.currentVersion = event.target.value;
    state.onlyDifferences = state.settings.currentVersion !== "general" && state.viewModeByVersion[state.settings.currentVersion] === "diff";
    await saveSettings();
    render();
    saveViewState();
  });
  els.addVersionBtn.addEventListener("click", () => openVersionDialog("add"));
  els.moveVersionUpBtn.addEventListener("click", () => moveCurrentVersion(-1));
  els.moveVersionDownBtn.addEventListener("click", () => moveCurrentVersion(1));
  els.renameVersionBtn.addEventListener("click", () => openVersionDialog("rename"));
  els.copyVersionBtn.addEventListener("click", () => openVersionDialog("copy"));
  els.deleteVersionBtn.addEventListener("click", deleteCurrentVersion);
  els.versionForm.addEventListener("submit", handleVersionSubmit);
  els.showAllBtn.addEventListener("click", () => setDifferenceMode(false));
  els.showDifferencesBtn.addEventListener("click", () => setDifferenceMode(true));
  els.dataImportBtn.addEventListener("click", importProfile);
  els.resumeImportBtn.addEventListener("click", importWordResume);
  els.dataExportBtn.addEventListener("click", chooseProfileExport);
  els.backupExportBtn.addEventListener("click", exportCompleteBackup);
  els.backupRestoreBtn.addEventListener("click", restoreCompleteBackup);
  els.migrateWorkspaceBtn.addEventListener("click", migrateLegacyWorkspace);
  els.closeResumeImportBtn.addEventListener("click", closeResumeImport);
  els.cancelResumeImportBtn.addEventListener("click", closeResumeImport);
  els.confirmResumeImportBtn.addEventListener("click", confirmResumeImport);
  els.newItemBtn.addEventListener("click", () => openEditor());
  els.emptyNewItemBtn?.addEventListener("click", () => openEditor());
  els.emptyResumeImportBtn?.addEventListener("click", importWordResume);
  els.manageCategoriesBtn?.addEventListener("click", openCategoryManager);
  els.closeCategoryBtn.addEventListener("click", () => els.categoryDialog.close());
  els.addCategoryBtn.addEventListener("click", addCategory);
  els.editorForm.addEventListener("submit", handleEditorSubmit);
  els.editCategory.addEventListener("change", updateEditorGuide);
  els.deleteItemBtn.addEventListener("click", deleteEditingItem);
  els.duplicateItemBtn.addEventListener("click", duplicateEditingItem);
  els.fieldEditorForm.addEventListener("submit", handleFieldSubmit);
  els.fieldVersioned.addEventListener("change", handleFieldVersionToggle);
  els.fieldEditLabel.addEventListener("input", () => {
    if (state.editingField?.fieldIndex < 0 && !state.fieldVersionPreferenceTouched) {
      const item = state.profile.items.find(candidate => candidate.id === state.editingField.itemId);
      if (item) els.fieldVersioned.checked = isExpressionField(item, { label: els.fieldEditLabel.value.trim() });
    }
    updateFieldEditorGuidance();
  });
  els.useCommonFieldAnswerBtn.addEventListener("click", useCommonFieldAnswer);
  els.fieldEditorDialog.addEventListener("cancel", event => { if (DISPLAY_MODE === "sidepanel") { event.preventDefault(); closeFieldEditorWithConfirmation(); } });
  els.editorDialog.addEventListener("cancel", event => { if (DISPLAY_MODE === "sidepanel") { event.preventDefault(); closeItemEditorWithConfirmation(); } });
  els.continueEditingBtn.addEventListener("click", () => { state.pendingDiscardTarget = null; els.discardChangesDialog.close(); });
  els.discardChangesBtn.addEventListener("click", discardEditorChanges);
  els.deleteFieldBtn.addEventListener("click", deleteEditingField);
  els.duplicateFieldBtn.addEventListener("click", duplicateEditingField);
  els.closeFieldStatusBtn.addEventListener("click", () => els.fieldStatusDialog.close());
  els.promoteFieldBtn.addEventListener("click", promoteFieldToGeneral);
  els.restoreFieldBtn.addEventListener("click", restoreFieldToGeneral);
  els.addAttachmentBtn.addEventListener("click", addAttachment);
  els.attachmentForm.addEventListener("submit", handleAttachmentSubmit);
  els.deleteAttachmentBtn.addEventListener("click", deleteEditingAttachment);
  els.closeAttachmentPreviewBtn.addEventListener("click", closeAttachmentPreview);
  els.attachmentPreviewDialog.addEventListener("close", cleanupAttachmentPreview);
  els.previewPrevBtn.addEventListener("click", () => changePreviewPage(-1));
  els.previewNextBtn.addEventListener("click", () => changePreviewPage(1));
  els.previewZoomOutBtn.addEventListener("click", () => changePreviewZoom(-0.2));
  els.previewZoomInBtn.addEventListener("click", () => changePreviewZoom(0.2));
  els.previewFitBtn.addEventListener("click", fitAttachmentPreview);
  els.markAttachmentVerifiedBtn.addEventListener("click", markPreviewedAttachmentVerified);
  els.previewReplacementBtn.addEventListener("click", previewPendingReplacement);
  els.cancelReplacementBtn.addEventListener("click", closeAttachmentReplacement);
  els.closeAttachmentReplaceBtn.addEventListener("click", closeAttachmentReplacement);
  els.confirmReplacementBtn.addEventListener("click", confirmAttachmentReplacement);
  els.previewBeforeUseBtn.addEventListener("click", previewPendingUseAttachment);
  els.cancelAttachmentUseBtn.addEventListener("click", closeAttachmentUse);
  els.closeAttachmentUseBtn.addEventListener("click", closeAttachmentUse);
  els.confirmAttachmentUseBtn.addEventListener("click", confirmAttachmentUse);
  els.cancelResumeMismatchBtn.addEventListener("click", () => els.resumeMismatchDialog.close());
  els.confirmResumeMismatchBtn.addEventListener("click", confirmMismatchedResume);
  els.completeBtn.addEventListener("click", openCompleteDialog);
  els.laterBtn.addEventListener("click", openLaterDialog);
  els.abortBtn.addEventListener("click", openAbortDialog);
  els.completeForm.addEventListener("submit", handleCompleteSubmit);
  els.abortForm.addEventListener("submit", handleAbortSubmit);
  els.recordsBtn.addEventListener("click", openRecords);
  els.exportRecordsCsvBtn.addEventListener("click", exportRecordsCsv);
  els.closeRecordsBtn.addEventListener("click", () => els.recordsDialog.close());
  els.recordStatusFilter.addEventListener("change", event => { state.recordFilter = event.target.value; renderRecords(); });
  els.recordSearchInput.addEventListener("input", event => { state.recordQuery = event.target.value.trim().toLowerCase(); renderRecords(); });
  els.continueRecordBtn.addEventListener("click", () => continueRecord(state.pendingResumeRecordId));
  els.newRecordBtn.addEventListener("click", startNewRecordFromPrompt);
  els.feishuSettingsBtn.addEventListener("click", openFeishuSettings);
  els.feishuForm.addEventListener("submit", handleFeishuSettingsSubmit);
  els.testFeishuBtn.addEventListener("click", testFeishuSettings);
  els.clearFeishuBtn.addEventListener("click", clearFeishuSettings);
  els.recordDetailForm.addEventListener("submit", handleRecordDetailSubmit);
  els.closeSyncResultBtn.addEventListener("click", () => els.syncResultDialog.close());
  els.laterSyncBtn.addEventListener("click", () => els.syncResultDialog.close());
  els.retrySyncBtn.addEventListener("click", retryResultRecord);
  els.viewFeishuBtn.addEventListener("click", openFeishuTable);
  els.pinSidePanelBtn?.addEventListener("click", pinSidePanel);
  els.openWorkbenchBtn?.addEventListener("click", openFullWorkbench);
  const handleSidepanelCategory = event => {
    if (event.target.value !== state.selectedCategory) selectCategory(event.target.value);
  };
  els.sidepanelCategory?.addEventListener("input", handleSidepanelCategory);
  els.sidepanelCategory?.addEventListener("change", handleSidepanelCategory);
  els.quickNavigationSelect?.addEventListener("change", navigateToContent);
  els.quickNavigationTop?.addEventListener("click", () => els.contentPane.scrollTo({ top: 0, behavior: "smooth" }));
  els.contentPane?.addEventListener("scroll", scheduleViewStateSave, { passive: true });
  els.contentPane?.addEventListener("mousedown", event => {
    if (event.target.closest(".fieldValue, .copyBtn")) event.preventDefault();
  });
  window.addEventListener("pagehide", saveViewState);
  chrome.runtime.onMessage?.addListener(message => {
    if (message?.type === "job-paste-data-changed" && message.source !== VIEW_INSTANCE_ID) refreshSharedData();
  });
}

function render() {
  renderStorageStatus(); renderVersionOptions(); renderDifferenceBar(); renderTabs(); renderEditorCategories(); renderList(); renderPendingRecordBadge();
}

function renderStorageStatus() { els.fileStatus.textContent = "数据保存在当前 Edge 浏览器中"; }

function renderVersionOptions() {
  els.currentVersion.replaceChildren(...state.settings.versions.map(version => option(version.id, version.name)));
  els.currentVersion.value = state.settings.currentVersion;
  const isGeneral = state.settings.currentVersion === "general";
  const versionIndex = state.settings.versions.findIndex(version => version.id === state.settings.currentVersion);
  els.moveVersionUpBtn.disabled = isGeneral || versionIndex <= 1;
  els.moveVersionDownBtn.disabled = isGeneral || versionIndex < 0 || versionIndex >= state.settings.versions.length - 1;
  els.renameVersionBtn.disabled = isGeneral;
  els.deleteVersionBtn.disabled = isGeneral;
  els.copyVersionBtn.disabled = false;
}

function renderDifferenceBar() {
  const version = currentVersion();
  const count = differenceCount(version);
  els.differenceSummary.textContent = version.id === "general" ? "通用版" : `${version.name} · ${count}项差异`;
  els.showDifferencesBtn.textContent = `本版本差异 ${count}`;
  els.showDifferencesBtn.disabled = version.id === "general" || count === 0;
  els.differenceBar.hidden = version.id === "general";
  if (version.id === "general" || count === 0) {
    state.onlyDifferences = false;
    state.viewModeByVersion[version.id] = "all";
  }
  els.showAllBtn.classList.toggle("active", !state.onlyDifferences);
  els.showDifferencesBtn.classList.toggle("active", state.onlyDifferences);
}

function setDifferenceMode(differencesOnly) {
  const version = currentVersion();
  state.onlyDifferences = Boolean(differencesOnly && version.id !== "general" && differenceCount(version) > 0);
  state.viewModeByVersion[version.id] = state.onlyDifferences ? "diff" : "all";
  renderDifferenceBar(); renderList(); saveViewState();
}

function openVersionDialog(mode) {
  const current = currentVersion();
  if ((mode === "rename" || mode === "delete") && current.id === "general") return;
  state.versionDialogMode = mode;
  els.versionDialogTitle.textContent = mode === "rename" ? "重命名版本" : mode === "copy" ? "复制版本差异" : "新增版本";
  els.versionName.value = mode === "rename" ? current.name : mode === "copy" ? `${current.name}副本` : "";
  els.versionBaseLabel.hidden = mode === "rename";
  els.versionBase.value = mode === "copy" ? "current" : "general";
  const resumes = state.attachments.filter(item => /简历/.test(item.category));
  els.versionDefaultResume.replaceChildren(option("", "不设置"), ...resumes.map(item => option(item.id, item.title)));
  els.versionDefaultResume.value = mode === "rename" ? (current.defaultResumeId || "") : mode === "copy" ? (current.defaultResumeId || "") : "";
  els.versionDialog.showModal();
  els.versionName.focus();
}

async function handleVersionSubmit(event) {
  event.preventDefault();
  if (event.submitter?.value === "cancel") return els.versionDialog.close();
  const name = els.versionName.value.trim();
  if (!name) return;
  if (state.settings.versions.some(version => version.name === name && version.id !== state.settings.currentVersion)) return showNotice("版本名称已存在。");
  const mode = state.versionDialogMode;
  if (mode === "rename") {
    await commit(`重命名版本：${name}`, () => { const version = currentVersion(); version.name = name; version.defaultResumeId = els.versionDefaultResume.value || null; });
  } else {
    const source = els.versionBase.value === "current" ? currentVersion() : null;
    const version = source ? structuredClone(source) : createVersion(uid("version"), name);
    version.id = uid("version"); version.name = name; version.defaultResumeId = els.versionDefaultResume.value || null;
    await commit(`新增版本：${name}`, () => { state.settings.versions.push(version); state.settings.currentVersion = version.id; state.onlyDifferences = false; });
  }
  els.versionDialog.close();
}

async function deleteCurrentVersion() {
  const version = currentVersion();
  if (version.id === "general" || !window.confirm(`确定删除“${version.name}”吗？只会删除该版本的差异、排序和默认附件设置。`)) return;
  await commit(`删除版本：${version.name}`, () => { state.settings.versions = state.settings.versions.filter(candidate => candidate.id !== version.id); state.settings.currentVersion = "general"; state.onlyDifferences = false; });
}

async function moveCurrentVersion(direction) {
  const version = currentVersion(), index = state.settings.versions.findIndex(candidate => candidate.id === version.id), target = index + direction;
  if (version.id === "general" || target <= 0 || target >= state.settings.versions.length) return;
  await commit(`调整版本顺序：${version.name}`, () => {
    const [moved] = state.settings.versions.splice(index, 1);
    state.settings.versions.splice(target, 0, moved);
  });
}

function createVersion(id, name) {
  return { id, name, fieldOverrides: {}, itemOverrides: {}, projectOrder: [], skillOrder: [], skillFieldOrders: {}, attachmentOrder: [], defaultResumeId: null, defaultAttachmentIds: [] };
}

function currentVersion() {
  return state.settings.versions.find(version => version.id === state.settings.currentVersion) || state.settings.versions[0];
}

function differenceCount(version) {
  if (!version || version.id === "general") return 0;
  const activeVersionedFields = new Set(state.profile.items.flatMap(item => (item.fields || []).filter(field => field.versioned).map(field => field.id)));
  return Object.keys(version.fieldOverrides || {}).filter(id => activeVersionedFields.has(id)).length
    + Object.keys(version.itemOverrides || {}).length
    + (version.projectOrder?.length ? 1 : 0)
    + (version.skillOrder?.length ? 1 : 0)
    + Object.keys(version.skillFieldOrders || {}).filter(key => version.skillFieldOrders[key]?.length).length
    + (version.attachmentOrder?.length ? 1 : 0)
    + (version.defaultResumeId ? 1 : 0)
    + (version.defaultAttachmentIds?.length || 0);
}

function renderTabs() {
  const entries = [{ name: "全部", system: true }, ...state.profile.categories.map(name => ({ name })), { name: "附件", system: true }];
  els.sidepanelCategory?.replaceChildren(...entries.map(entry => option(entry.name, entry.name)));
  if (els.sidepanelCategory) els.sidepanelCategory.value = state.selectedCategory;
  els.categoryTabs.replaceChildren(...entries.map(entry => {
    const button = document.createElement("button");
    button.className = `tab${entry.name === state.selectedCategory ? " active" : ""}`;
    button.textContent = entry.name;
    if (!entry.system) makeDraggable(button, entry.name, "category", reorderCategories);
    button.addEventListener("click", () => selectCategory(entry.name));
    button.addEventListener("contextmenu", event => { if (!entry.system) { event.preventDefault(); openCategoryManager(); } });
    return button;
  }));
}

function selectCategory(category) {
  if (DISPLAY_MODE === "sidepanel" && !state.query) state.categoryScrollPositions[state.selectedCategory] = els.contentPane.scrollTop;
  state.selectedCategory = category; renderTabs(); renderList();
  const top = DISPLAY_MODE === "sidepanel" ? Number(state.categoryScrollPositions[category]) || 0 : 0;
  requestAnimationFrame(() => { els.contentPane.scrollTop = top; saveViewState(); });
}

function renderQuickNavigation() {
  if (!els.quickNavigation) return;
  const items = state.selectedCategory === "附件" || state.selectedCategory === "待补充"
    ? []
    : orderedItemsForVersion(state.profile.items).filter(item => state.selectedCategory === "全部" || item.category === state.selectedCategory);
  els.quickNavigation.hidden = DISPLAY_MODE !== "sidepanel" && items.length < 2;
  els.quickNavigationSelect.replaceChildren(option("", "选择内容"), ...items.map(item => option(item.id, item.title)));
}

function navigateToContent(event) {
  const id = event.target.value; if (!id) return;
  const card = document.querySelector(`[data-item-id="${CSS.escape(id)}"]`);
  if (!card) return showNotice("当前搜索条件下未显示该内容");
  const stickyOffset = DISPLAY_MODE === "sidepanel" ? (els.sidepanelStickyNavigation?.offsetHeight || 0) + 12 : 10;
  const top = els.contentPane.scrollTop + card.getBoundingClientRect().top - els.contentPane.getBoundingClientRect().top - stickyOffset;
  els.contentPane.scrollTo({ top, behavior: "smooth" }); card.classList.add("navigationHighlight");
  setTimeout(() => card.classList.remove("navigationHighlight"), 1000); scheduleViewStateSave();
}

async function pinSidePanel() {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const windowId = tab?.windowId ?? (await chrome.windows.getCurrent()).id;
    if (tab?.id != null) await chrome.sidePanel.setOptions?.({ tabId: tab.id, path: "sidepanel.html", enabled: true });
    await chrome.sidePanel.open({ windowId });
    els.pinSidePanelBtn.textContent = "已固定"; els.pinSidePanelBtn.disabled = true;
  } catch (error) { showNotice(`侧边栏打开失败：${error.message}`); }
}

async function openFullWorkbench() {
  const url = chrome.runtime.getURL("popup.html?mode=workbench"), tabs = await chrome.tabs.query({});
  const existing = tabs.find(tab => tab.url === url);
  if (existing?.id) { await chrome.tabs.update(existing.id, { active: true }); if (existing.windowId != null) await chrome.windows.update?.(existing.windowId, { focused: true }); return; }
  await chrome.tabs.create({ url });
}

let viewStateTimer;
function scheduleViewStateSave() {
  if (DISPLAY_MODE === "sidepanel" && !state.query) state.categoryScrollPositions[state.selectedCategory] = els.contentPane.scrollTop;
  clearTimeout(viewStateTimer); viewStateTimer = setTimeout(saveViewState, 120);
}
function saveViewState() {
  chrome.storage.local.set({ [VIEW_STATE_KEY]: { category: state.selectedCategory, query: els.searchInput?.value || "", scrollTop: els.contentPane?.scrollTop || 0, viewModeByVersion: state.viewModeByVersion, categoryScrollPositions: state.categoryScrollPositions } });
}

function setupSidePanelLayout() {
  if (DISPLAY_MODE !== "sidepanel" || !els.sidepanelStickyNavigation) return;
  const topbar = document.querySelector(".topbar"), toolbar = document.querySelector(".toolbar");
  els.sidepanelStickyNavigation.hidden = false;
  els.quickNavigationTop.textContent = "↑"; els.quickNavigationTop.title = "回到顶部"; els.quickNavigationTop.setAttribute("aria-label", "回到顶部");
  els.contentPane.prepend(topbar, els.sidepanelStickyNavigation);
  els.sidepanelStickyNavigation.append(toolbar, els.sidepanelCategoryControl, els.quickNavigation);
}

function handleSearchInput(event) {
  const nextQuery = event.target.value.trim().toLowerCase();
  if (DISPLAY_MODE === "sidepanel" && !state.query && nextQuery) state.searchRestore = { category: state.selectedCategory, scrollTop: els.contentPane.scrollTop };
  state.query = nextQuery;
  if (DISPLAY_MODE === "sidepanel" && !nextQuery && state.searchRestore) {
    const restore = state.searchRestore; state.searchRestore = null; state.selectedCategory = restore.category;
    renderTabs(); renderList(); requestAnimationFrame(() => { els.contentPane.scrollTop = restore.scrollTop; saveViewState(); }); return;
  }
  renderList(); saveViewState();
}

function renderEditorCategories() {
  els.editCategory.replaceChildren(...state.profile.categories.map(name => option(name, name)));
}

function renderList() {
  const attachments = state.selectedCategory === "附件";
  els.attachmentPanel.hidden = !attachments; els.list.hidden = attachments; els.empty.hidden = true;
  if (attachments) { renderQuickNavigation(); return renderAttachments(); }
  const items = filteredItems();
  els.empty.hidden = items.length > 0;
  const nodes = items.map(item => renderCard(item, state.selectedCategory === "待补充"));
  if (state.onlyDifferences && state.selectedCategory === "全部" && hasAttachmentDifferences()) nodes.unshift(renderAttachmentDifferenceSummary());
  els.empty.hidden = nodes.length > 0;
  els.list.replaceChildren(...nodes);
  renderQuickNavigation();
}

function hasAttachmentDifferences() {
  const version = currentVersion();
  return version.id !== "general" && Boolean(version.defaultResumeId || version.defaultAttachmentIds.length || version.attachmentOrder.length);
}

function renderAttachmentDifferenceSummary() {
  const version = currentVersion();
  const ids = [version.defaultResumeId, ...(version.defaultAttachmentIds || [])].filter(Boolean);
  const names = ids.map(id => state.attachments.find(item => item.id === id)?.title).filter(Boolean);
  const card = document.createElement("article"); card.className = "card differenceCard";
  const header = document.createElement("div"); header.className = "cardHeader";
  const title = document.createElement("div"); title.className = "cardTitle"; title.textContent = "默认简历与附件";
  const tag = document.createElement("span"); tag.className = "tag"; tag.textContent = version.name; header.append(title, tag);
  const content = document.createElement("div"); content.className = "cardContent"; content.textContent = names.length ? names.join("\n") : "已设置当前版本附件优先级";
  const actions = document.createElement("div"); actions.className = "cardActions"; actions.append(button("查看附件", "copyBtn", () => { state.selectedCategory = "附件"; renderTabs(); renderList(); }));
  card.append(header, content, actions); return card;
}

function filteredItems() {
  return orderedItemsForVersion(state.profile.items).filter(item => {
    const fields = item.fields || [];
    const pending = fields.some(field => isPendingValue(resolveFieldValue(field)));
    const categoryMatch = state.selectedCategory === "全部" || item.category === state.selectedCategory || (state.selectedCategory === "待补充" && pending);
    const fieldText = fields.map(field => `${field.label} ${resolveFieldValue(field)}`).join(" ");
    const haystack = `${item.category} ${item.title} ${resolveItemContent(item) || ""} ${fieldText} ${(item.tags || []).join(" ")}`.toLowerCase();
    const differenceMatch = !state.onlyDifferences || itemHasDifference(item);
    return categoryMatch && differenceMatch && (!state.query || haystack.includes(state.query));
  });
}

function renderCard(item, pendingOnly) {
  const card = document.createElement("article"); card.className = "card";
  card.id = `content-${item.id}`; card.dataset.itemId = item.id;
  if (itemHasDifference(item)) card.classList.add("differenceCard");
  if (canReorderItems()) makeDraggable(card, item.id, "item", reorderItems);
  const header = document.createElement("div"); header.className = "cardHeader";
  const title = document.createElement("div"); title.className = "cardTitle"; title.textContent = item.title;
  const category = document.createElement("span"); category.className = "tag"; category.textContent = item.category;
  header.append(title, category);
  const content = item.fields?.length ? renderFields(item, pendingOnly) : renderContent(item);
  const tags = document.createElement("div"); tags.className = "tagRow";
  for (const tag of item.tags || []) { const span = document.createElement("span"); span.className = `tag${isSensitiveTag(tag) ? " sensitive" : ""}`; span.textContent = tag; tags.append(span); }
  const actions = document.createElement("div"); actions.className = "cardActions";
  actions.append(
    button(item.fields?.length ? "复制全部" : "复制内容", "copyBtn", () => copyItem(item)),
    button("+ 添加字段", "sidepanelAddField", () => openFieldEditor(item, -1)),
    button("复制条目", "", () => duplicateItem(item)),
    button("编辑条目", "sidepanelItemEdit", () => openEditor(item))
  );
  card.append(header, content, tags, actions); return card;
}

function renderContent(item) {
  const content = document.createElement("div"); content.className = "cardContent"; content.textContent = resolveItemContent(item); content.title = "点击复制"; content.addEventListener("click", () => copyItem(item)); return content;
}

function renderFields(item, pendingOnly) {
  const list = document.createElement("div"); list.className = "fieldList";
  const orderedFields = fieldsForVersion(item);
  for (const field of orderedFields) {
    const index = item.fields.findIndex(candidate => candidate.id === field.id);
    const valueText = resolveFieldValue(field);
    if (pendingOnly && !isPendingValue(valueText)) continue;
    const row = document.createElement("div"); row.className = "fieldRow";
    makeDraggable(row, field.id, `field:${item.id}`, (source, target) => reorderFields(item.id, source, target));
    const handle = dragHandle();
    const label = document.createElement("span"); label.className = "fieldLabel"; label.textContent = field.label;
    const value = button(valueText, "fieldValue", () => copyText(valueText, field.label, value));
    if (isPendingValue(valueText)) value.classList.add("missingValue");
    const overridden = hasFieldOverride(field);
    const scope = button(overridden ? versionName(state.settings.currentVersion) : "通用", overridden ? "fieldScope override" : "fieldScope", () => openFieldStatus(item, field));
    scope.title = overridden ? "当前显示岗位版本差异，点击管理" : "当前继承通用内容，点击查看";
    const edit = button("编辑", "fieldEdit", () => openFieldEditor(item, index));
    row.append(handle, label, value, scope, edit); list.append(row);
  }
  return list;
}

function renderAttachments() {
  const version = currentVersion();
  const items = orderedAttachmentsForVersion(state.attachments).filter(item => {
    const haystack = `${item.title} ${item.fileName} ${item.category} ${(item.tags || []).join(" ")}`.toLowerCase();
    const differenceMatch = !state.onlyDifferences || version.defaultResumeId === item.id || (version.defaultAttachmentIds || []).includes(item.id) || (version.attachmentOrder || []).includes(item.id);
    return differenceMatch && (!state.query || haystack.includes(state.query));
  });
  els.attachmentList.replaceChildren(...items.map(renderAttachment));
}

function renderAttachment(item) {
  const row = document.createElement("article"); row.className = "attachmentRow"; row.dataset.attachmentId = item.id;
  const version = currentVersion();
  const isDefault = version.defaultResumeId === item.id || (version.defaultAttachmentIds || []).includes(item.id);
  if (isDefault) row.classList.add("defaultAttachment");
  makeDraggable(row, item.id, "attachment", reorderAttachments);
  const info = document.createElement("div"); info.className = "attachmentInfo";
  const title = document.createElement("strong"); title.textContent = item.title;
  const file = document.createElement("span"); file.textContent = item.fileName;
  const meta = document.createElement("span"); meta.textContent = `${item.category} · ${attachmentVersionLabel(item)} · ${fileExtension(item.fileName)} · ${formatBytes(item.size)} · 更新于 ${formatAttachmentDate(item.updatedAt)}`;
  info.append(title, file, meta);
  if (isDefault) { const status = document.createElement("span"); status.className = "verifiedAttachment"; status.textContent = "当前版本默认"; info.append(status); }
  if (state.verifiedAttachmentIds.has(item.id)) { const checked = document.createElement("span"); checked.className = "verifiedAttachment"; checked.textContent = "✓ 已核对"; info.append(checked); }
  const actions = document.createElement("div"); actions.className = "attachmentActions";
  const more = document.createElement("details"); more.className = "attachmentMore";
  const summary = document.createElement("summary"); summary.textContent = "更多";
  const menu = document.createElement("div"); menu.className = "attachmentMoreMenu";
  menu.append(button("编辑附件信息", "", () => openAttachmentEditor(item)), button(isDefault ? "取消当前版本默认" : "设置为当前版本默认附件", "", () => setCurrentVersionDefaultAttachment(item, !isDefault)), button("删除附件", "danger", () => deleteAttachment(item)));
  more.append(summary, menu);
  actions.append(button("预览", "", () => previewAttachment(item)), button("替换", "", () => replaceAttachmentFile(item)), button(item.sensitive ? "确认使用" : "使用", item.sensitive ? "uploadBtn sensitiveAction" : "uploadBtn", () => startAttachmentUpload(item)), more);
  row.append(dragHandle(), info, actions);
  return row;
}

async function importWordResume() {
  try {
    const file = await pickFile([{ description: "Word 简历", accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] } }]);
    if (!/\.docx$/i.test(file.name)) throw new Error("文件无法读取：请选择 .docx Word 简历");
    const parsed = await parseResumeDocx(file), fileHash = await sha256(file);
    if (!parsed.candidates.length) throw new Error("识别结果为空，无法确认导入");
    const candidates = parsed.candidates.map(candidate => {
      const match = findLikelyResumeMatch(candidate, state.profile.items);
      return { ...candidate, matchId: match?.item.id || "", matchScore: match?.score || 0 };
    });
    const duplicate = state.attachments.find(item => item.hash === fileHash || item.fileName === file.name) || null;
    state.resumeImport = { file, fileHash, parsed: { ...parsed, candidates }, duplicate };
    renderResumeImportPreview(); els.resumeImportDialog.showModal();
  } catch (error) { if (error.name !== "AbortError") showNotice(resumeImportErrorMessage(error)); }
}

function renderResumeImportPreview() {
  const pending = state.resumeImport; if (!pending) return;
  const { parsed, duplicate } = pending;
  els.resumeImportFile.textContent = `文件：${pending.file.name}`;
  els.resumeImportVersion.replaceChildren(...state.settings.versions.map(version => option(version.id, version.name))); els.resumeImportVersion.value = state.settings.currentVersion;
  els.resumeImportDefault.checked = state.settings.currentVersion !== "general";
  const labels = Object.entries(parsed.summary.counts).map(([category, count]) => `${category} ${count}项`);
  labels.push(`✓ 已识别 ${parsed.summary.high}项`, `△ 请确认 ${parsed.summary.confirm}项`);
  els.resumeImportSummary.replaceChildren(...labels.map(text => Object.assign(document.createElement("span"), { textContent: text })));
  els.resumeDuplicatePanel.hidden = !duplicate; els.resumeDuplicateAction.value = "use";
  if (duplicate) els.resumeDuplicatePanel.querySelector("strong").textContent = `该简历附件可能已经存在：${duplicate.title}（${duplicate.fileName}）`;
  els.resumeImportList.replaceChildren(...parsed.candidates.map(renderResumeImportItem)); els.resumeImportError.hidden = true;
}

function renderResumeImportItem(candidate) {
  const row = document.createElement("article"); row.className = "resumeImportItem"; row.dataset.candidateId = candidate.id;
  const header = document.createElement("div"); header.className = "resumeImportItemHeader";
  const checked = document.createElement("input"); checked.type = "checkbox"; checked.checked = true; checked.className = "resumeImportSelected";
  const title = document.createElement("strong"); title.textContent = `${candidate.category}｜${candidate.title}`;
  const confidence = document.createElement("span"); confidence.textContent = candidate.confidence === "high" ? "✓ 已识别" : "△ 请确认"; confidence.className = candidate.confidence === "high" ? "" : "confirm";
  header.append(checked, title, confidence);
  const source = document.createElement("pre"); source.className = "resumeImportSource"; source.textContent = candidate.sourceText;
  const decisions = document.createElement("div"); decisions.className = "resumeImportDecision";
  const categorySelect = document.createElement("select"); categorySelect.className = "resumeImportCategory"; const categories = [...new Set([...state.profile.categories, "实习经历", "校园经历"])]; categorySelect.replaceChildren(...categories.map(name => option(name, name))); categorySelect.value = categories.includes(candidate.category) ? candidate.category : "待确认";
  const categoryLabel = document.createElement("label"); categoryLabel.textContent = "导入分类"; categoryLabel.append(categorySelect); decisions.append(categoryLabel);
  const action = document.createElement("select"); action.className = "resumeImportAction";
  if (candidate.matchId) action.append(option("merge", "合并到现有"), option("new", "作为新项目/条目"), option("skip", "跳过"));
  else action.append(option("new", "新增"), option("skip", "忽略"));
  const actionLabel = document.createElement("label"); actionLabel.textContent = "导入方式"; actionLabel.append(action); decisions.append(actionLabel);
  if (candidate.matchId) {
    const matched = state.profile.items.find(item => item.id === candidate.matchId), notice = document.createElement("div"); notice.className = "resumeMatchNotice"; notice.textContent = `发现可能对应已有内容：${matched?.title || "未知条目"}`;
    const policy = document.createElement("select"); policy.className = "resumeMergePolicy"; policy.append(option("keep", "保留现有内容"), option("resume", "使用简历内容"), option("version", "作为当前版本内容"));
    const policyLabel = document.createElement("label"); policyLabel.className = "resumeMergePolicyLabel"; policyLabel.textContent = "合并内容"; policyLabel.append(policy); decisions.append(policyLabel);
    action.addEventListener("change", () => policyLabel.hidden = action.value !== "merge"); row.append(header, notice, source, decisions); return row;
  }
  row.append(header, source, decisions); return row;
}

function closeResumeImport() { state.resumeImport = null; els.resumeImportDialog.close(); }

async function confirmResumeImport() {
  const pending = state.resumeImport; if (!pending) return;
  const rows = [...els.resumeImportList.querySelectorAll(".resumeImportItem")].filter(row => row.querySelector(".resumeImportSelected").checked && row.querySelector(".resumeImportAction").value !== "skip");
  if (!rows.length) return showResumeImportError("识别结果为空：请至少选择一项导入内容");
  const nextProfile = structuredClone(state.profile), nextVersions = structuredClone(state.settings.versions), nextAttachments = state.attachments.map(item => ({ ...item }));
  const targetVersionId = els.resumeImportVersion.value, targetVersion = nextVersions.find(version => version.id === targetVersionId);
  if (!targetVersion) return showResumeImportError("目标填写版本不存在");
  try {
    for (const row of rows) {
      const candidate = pending.parsed.candidates.find(item => item.id === row.dataset.candidateId); if (!candidate) continue;
      const category = row.querySelector(".resumeImportCategory").value, action = row.querySelector(".resumeImportAction").value;
      if (!nextProfile.categories.includes(category)) nextProfile.categories.push(category);
      if (action === "new") { const item = createResumeImportItem(candidate, category); nextProfile.items.push(item); applyNewResumeVersionOverrides(item, candidate, targetVersionId, targetVersion); continue; }
      const existing = nextProfile.items.find(item => item.id === candidate.matchId); if (!existing) continue;
      const policy = row.querySelector(".resumeMergePolicy")?.value || "keep";
      if (policy === "resume") mergeResumeCandidateIntoItem(existing, candidate);
      if (policy === "version") applyResumeCandidateAsOverride(existing, candidate, targetVersionId, targetVersion);
    }
    const attachmentId = await applyImportedResumeAttachment(nextAttachments, nextVersions, pending, targetVersionId, els.resumeDuplicateAction.value, els.resumeImportDefault.checked);
    if (!attachmentId) throw new Error("原 Word 简历附件保存失败");
    nextProfile.updatedAt = new Date().toISOString();
    await saveCompleteState({ profile: nextProfile, versions: nextVersions, records: state.records, attachments: nextAttachments });
    state.profile = normalizeProfile(nextProfile); state.settings.versions = nextVersions; state.attachments = normalizeAttachments(nextAttachments);
    await saveUiSettings(); state.resumeImport = null; els.resumeImportDialog.close(); render(); notifyDataChanged(); showNotice(`简历已导入，并保存 ${pending.file.name} 为附件`);
  } catch (error) { showResumeImportError(`导入写入失败：${storageErrorMessage(error)}`); }
}

function createResumeImportItem(candidate, category) {
  const fields = candidate.fields.map(source => ({ id: uid("field"), label: source.label, value: source.value, versioned: Boolean(source.versioned) }));
  return normalizeItem({ id: uid("item"), category, title: candidate.title, fields, content: fieldsToContent(fields), tags: ["Word简历导入", candidate.confidence === "high" ? "已识别" : "待确认"] });
}

function applyNewResumeVersionOverrides(item, candidate, versionId, version) {
  if (versionId === "general") return;
  for (let index = 0; index < item.fields.length; index++) if (item.fields[index].versioned) version.fieldOverrides[item.fields[index].id] = candidate.fields[index].value;
}

function mergeResumeCandidateIntoItem(item, candidate) {
  item.fields ||= [];
  for (const source of candidate.fields) {
    const existing = item.fields.find(field => field.label === source.label);
    if (existing) existing.value = source.value;
    else item.fields.push({ id: uid("field"), label: source.label, value: source.value, versioned: Boolean(source.versioned) });
  }
  item.content = fieldsToContent(item.fields);
}

function applyResumeCandidateAsOverride(item, candidate, versionId, version) {
  if (versionId === "general") return mergeResumeCandidateIntoItem(item, candidate);
  item.fields ||= [];
  for (const source of candidate.fields.filter(field => field.versioned)) {
    let field = item.fields.find(current => current.label === source.label);
    if (!field) { field = { id: uid("field"), label: source.label, value: source.value, versioned: true }; item.fields.push(field); }
    field.versioned = true; version.fieldOverrides[field.id] = source.value;
  }
  item.content = fieldsToContent(item.fields);
}

async function applyImportedResumeAttachment(attachments, versions, pending, targetVersionId, duplicateAction, setDefault) {
  const now = new Date().toISOString(), duplicate = pending.duplicate && attachments.find(item => item.id === pending.duplicate.id); let attachment;
  if (duplicate && duplicateAction === "use") attachment = duplicate;
  else if (duplicate && duplicateAction === "replace") {
    Object.assign(duplicate, { fileName: pending.file.name, mimeType: pending.file.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document", size: pending.file.size, hash: pending.fileHash, updatedAt: now, source: "indexeddb", blob: pending.file }); attachment = duplicate;
  } else {
    attachment = { id: uid("attachment"), title: pending.file.name.replace(/\.docx$/i, ""), displayName: pending.file.name.replace(/\.docx$/i, ""), fileName: pending.file.name, mimeType: pending.file.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document", category: "简历", applicableVersionIds: targetVersionId === "general" ? [] : [targetVersionId], addedAt: now, createdAt: now, updatedAt: now, size: pending.file.size, hash: pending.fileHash, sensitive: false, tags: ["简历", "Word导入"], source: "indexeddb", blob: pending.file }; attachments.push(attachment);
  }
  if (setDefault) { const version = versions.find(item => item.id === targetVersionId); if (version) version.defaultResumeId = attachment.id; }
  return attachment.id;
}

function showResumeImportError(message) { els.resumeImportError.textContent = message; els.resumeImportError.hidden = false; }
function resumeImportErrorMessage(error) { const message = error?.message || "解析过程中发生错误"; if (/文件无法读取|DOCX格式异常|未识别到文本|未识别到简历结构/.test(message)) return message; return `解析过程中发生错误：${message}`; }

async function importProfile() {
  try {
    const opened = await openProfileFile();
    const incoming = normalizeProfile(opened.profile), currentIds = new Set(state.profile.items.map(item => item.id));
    const conflicts = incoming.items.filter(item => currentIds.has(item.id)).length;
    const message = `导入预览\n新增或替换条目：${incoming.items.length}\nID 冲突：${conflicts}\n分类：${incoming.categories.length}\n\n确认后才会更新当前信息库。`;
    if (!window.confirm(message)) return;
    const importedVersions = opened.profile.fillingVersions;
    await commit("导入信息库", () => { state.profile = incoming; state.settings = normalizeSettings(importedVersions ? { ...state.settings, versions: importedVersions } : state.settings, state.profile); state.selectedCategory = "全部"; });
    showNotice("信息库已导入");
  } catch (error) { if (error.name !== "AbortError") showNotice(error.message); }
}

async function handleEditorSubmit(event) {
  event.preventDefault(); if (event.submitter?.value === "cancel") return DISPLAY_MODE === "sidepanel" ? closeItemEditorWithConfirmation() : els.editorDialog.close();
  const scrollState = captureContentPosition();
  const fields = parseFields(els.editContent.value);
  const existing = state.profile.items.find(item => item.id === state.editingId);
  const title = els.editTitle.value.trim(), category = els.editCategory.value, tags = splitTags(els.editTags.value), content = els.editContent.value.trim();
  if (existing?.fields?.length) {
    const version = currentVersion();
    const changedFacts = fields.some(parsed => { const old = existing.fields.find(field => field.label === parsed.label); return old && isFactField(existing, old) && (version.id === "general" || !old.versioned) && old.value !== parsed.value; });
    if (changedFacts && !window.confirm("批量内容中包含事实字段修改，该修改将同步到所有填写版本。确定继续吗？")) return;
  }
  await commit(state.editingId ? `修改条目：${title}` : `新增条目：${title}`, () => {
    if (!existing) {
      state.profile.items.unshift(normalizeItem({ id: uid("item"), category, title, content, tags, ...(fields.length ? { fields } : {}) }));
      return;
    }
    existing.category = category; existing.title = title; existing.tags = tags;
    const version = currentVersion();
    if (!existing.fields?.length) {
      if (version.id !== "general" && isExpressionItem(existing)) {
        if (content && content !== existing.content) version.itemOverrides[existing.id] = content; else delete version.itemOverrides[existing.id];
      } else existing.content = content;
      return;
    }
    for (const parsed of fields) {
      const field = existing.fields.find(candidate => candidate.label === parsed.label);
      if (!field) { existing.fields.push({ ...parsed, versioned: isExpressionField(existing, parsed) }); continue; }
      if (version.id !== "general" && field.versioned) {
        if (parsed.value !== field.value) version.fieldOverrides[field.id] = parsed.value; else delete version.fieldOverrides[field.id];
      } else field.value = parsed.value;
    }
    existing.content = fieldsToContent(existing.fields);
  });
  els.editorDialog.close(); restoreContentPosition(scrollState); showNotice("已保存"); notifyDataChanged();
}

function openEditor(item = null) {
  const category = state.profile.categories.includes(state.selectedCategory) ? state.selectedCategory : "基础信息";
  const defaults = newItemDefaults(category); state.editingId = item?.id || null;
  els.editorTitle.textContent = item ? "编辑条目" : "新增条目"; els.deleteItemBtn.hidden = !item; els.duplicateItemBtn.hidden = !item;
  els.editCategory.value = item?.category || category; els.editTitle.value = item?.title || defaults.title; els.editTags.value = (item?.tags || []).join(" "); els.editContent.value = item?.fields?.length ? fieldsForVersion(item).map(field => `${field.label}：${resolveFieldValue(field)}`).join("\n") : item ? resolveItemContent(item) : defaults.content;
  updateEditorGuide();
  if (DISPLAY_MODE === "sidepanel") {
    const version = currentVersion(), commonEdit = !item || version.id === "general" || !isExpressionItem(item);
    els.itemEditorHint.textContent = `${item ? "字段名称修改将影响所有版本。" : ""}${commonEdit ? "该修改会同步影响所有填写版本。" : `内容将保存为${version.name}差异。`}`;
    state.editorInitial = itemEditorSnapshot();
  }
  els.editorDialog.showModal();
}

function updateEditorGuide() {
  const guide = FIELD_GUIDES[els.editCategory?.value] || DEFAULT_FIELD_GUIDE;
  if (!state.editingId) els.editTitle.placeholder = guide.title;
  els.editContent.placeholder = guide.fields.map(field => `${field}：`).join("\n");
  if (!els.itemFieldGuide) return;
  const intro = document.createElement("p"); intro.textContent = guide.tip;
  const chips = document.createElement("div"); chips.className = "fieldGuideChips";
  chips.replaceChildren(...guide.fields.map(name => button(name, "", () => insertFieldLine(name))));
  const actions = document.createElement("div"); actions.className = "fieldGuideActions";
  actions.append(button("插入推荐字段", "", () => guide.fields.forEach(insertFieldLine)));
  els.itemFieldGuide.replaceChildren(intro, chips, actions);
}

function insertFieldLine(name) {
  const lines = els.editContent.value.split(/\r?\n/).map(value => value.trim());
  if (lines.some(value => value === `${name}：` || value.startsWith(`${name}：`) || value.startsWith(`${name}:`))) return;
  const current = els.editContent.value.trimEnd();
  els.editContent.value = current ? `${current}\n${name}：` : `${name}：`;
  els.editContent.focus();
}

async function deleteEditingItem() {
  const item = state.profile.items.find(candidate => candidate.id === state.editingId); if (!item || !window.confirm(`确定删除条目“${item.title}”吗？会同时删除这个条目下的字段差异。`)) return;
  await commit(`删除条目：${item.title}`, () => { state.profile.items = state.profile.items.filter(candidate => candidate.id !== item.id); for (const version of state.settings.versions) { delete version.itemOverrides[item.id]; version.projectOrder = version.projectOrder.filter(id => id !== item.id); version.skillOrder = version.skillOrder.filter(id => id !== item.id); delete version.skillFieldOrders[item.id]; for (const field of item.fields || []) delete version.fieldOverrides[field.id]; } }); els.editorDialog.close();
}

function duplicateEditingItem() { const item = state.profile.items.find(candidate => candidate.id === state.editingId); if (item) { els.editorDialog.close(); duplicateItem(item); } }
async function duplicateItem(item) { const copy = structuredClone(item); copy.id = uid("item"); copy.title = `${copy.title}（副本）`; copy.fields?.forEach(field => { field.id = uid("field"); }); await commit(`复制条目：${item.title}`, () => state.profile.items.unshift(copy)); }

function openFieldEditor(item, index) {
  const field = index >= 0 ? item.fields[index] : { id: uid("field"), label: "", value: "", versioned: isExpressionField(item, { label: "" }) };
  const version = currentVersion();
  state.editingField = { itemId: item.id, fieldIndex: index };
  state.fieldVersionPreferenceTouched = false;
  els.fieldEditorTitle.textContent = index >= 0 ? "编辑字段" : "新增字段";
  els.fieldEditLabel.value = field.label;
  els.fieldVersioned.checked = Boolean(field.versioned);
  els.fieldVersioned.disabled = false;
  const editOverride = index >= 0 && version.id !== "general" && field.versioned;
  els.fieldEditValue.value = editOverride ? (version.fieldOverrides[field.id] ?? field.value) : field.value;
  els.fieldValueLabel.textContent = editOverride ? `${version.name}答案` : "通用答案";
  els.fieldEditorCommonBlock.hidden = !editOverride;
  els.fieldEditorCommonPreview.textContent = field.value;
  els.deleteFieldBtn.hidden = index < 0; els.duplicateFieldBtn.hidden = index < 0;
  els.fieldEditorDialog.classList.toggle("addingField", index < 0);
  updateFieldEditorGuidance();
  state.fieldEditorInitial = fieldEditorSnapshot();
  els.fieldEditorDialog.showModal(); els.fieldEditLabel.focus();
}

function updateFieldEditorGuidance() {
  const target = state.editingField;
  const item = state.profile.items.find(candidate => candidate.id === target?.itemId);
  if (!item) return;
  const previous = target.fieldIndex >= 0 ? item.fields[target.fieldIndex] : null;
  const field = previous || { label: els.fieldEditLabel.value.trim() };
  const version = currentVersion();
  const fact = isFactField(item, { ...field, label: els.fieldEditLabel.value.trim() });
  const versioned = els.fieldVersioned.checked;
  const editOverride = Boolean(previous && version.id !== "general" && versioned);
  els.fieldValueLabel.textContent = editOverride ? `${version.name}答案` : "通用答案";
  els.fieldEditorCommonBlock.hidden = !editOverride;
  const recommendation = fact ? "推荐：保持通用。" : isExpressionField(item, field) ? "推荐：允许版本差异。" : "";
  const warning = fact && versioned ? "该字段通常属于事实信息。开启后可以在不同填写版本中使用不同内容，请注意保持事实一致性。" : "";
  const scope = editOverride ? "保存后只更新当前岗位版本差异。" : version.id !== "general" && !versioned ? "该修改会同步影响所有填写版本。" : "在通用版中修改会更新所有继承该内容的版本。";
  els.fieldEditorHint.textContent = [recommendation, warning, scope, previous ? "字段名称修改将影响所有版本。" : ""].filter(Boolean).join(" ");
}

function handleFieldVersionToggle() {
  state.fieldVersionPreferenceTouched = true;
  const target = state.editingField;
  const item = state.profile.items.find(candidate => candidate.id === target?.itemId);
  const field = item?.fields?.[target?.fieldIndex];
  const version = currentVersion();
  if (field && els.fieldVersioned.checked && version.id !== "general" && version.fieldOverrides[field.id] != null) {
    els.fieldEditValue.value = version.fieldOverrides[field.id];
  }
  updateFieldEditorGuidance();
}

async function useCommonFieldAnswer() {
  const target = state.editingField;
  const item = state.profile.items.find(candidate => candidate.id === target?.itemId);
  const field = item?.fields?.[target?.fieldIndex];
  const version = currentVersion();
  if (!field || version.id === "general") return;
  const scrollState = captureContentPosition();
  await commit(`恢复通用内容：${field.label}`, () => delete version.fieldOverrides[field.id]);
  els.fieldEditorDialog.close(); restoreContentPosition(scrollState); showNotice("已恢复使用通用答案"); notifyDataChanged();
}

async function handleFieldSubmit(event) {
  event.preventDefault(); if (event.submitter?.value === "cancel") return DISPLAY_MODE === "sidepanel" ? closeFieldEditorWithConfirmation() : els.fieldEditorDialog.close();
  const target = state.editingField; const item = state.profile.items.find(candidate => candidate.id === target?.itemId); if (!item) return;
  const scrollState = captureContentPosition();
  const previous = target.fieldIndex >= 0 ? item.fields[target.fieldIndex] : null;
  const wasVersioned = Boolean(previous?.versioned);
  const version = currentVersion();
  const fact = previous ? isFactField(item, { ...previous, label: els.fieldEditLabel.value.trim() }) : isFactField(item, { label: els.fieldEditLabel.value.trim() });
  const versioned = els.fieldVersioned.checked;
  const value = els.fieldEditValue.value.trim();
  const editOverride = Boolean(previous && version.id !== "general" && versioned);
  const hasStoredOverrides = Boolean(previous && state.settings.versions.some(candidate => candidate.id !== "general" && candidate.fieldOverrides[previous.id] != null));
  if (previous?.versioned && !versioned && hasStoredOverrides && !window.confirm("当前字段已有版本差异，关闭后这些差异将不再生效。\n\n确认关闭版本化吗？已有差异数据会保留。")) return;
  const closingDormantOverride = Boolean(previous && wasVersioned && !versioned && version.id !== "general");
  if (fact && previous && !editOverride && !closingDormantOverride && value !== previous.value && !window.confirm("这是事实字段，该修改将同步到所有填写版本。确定继续吗？")) return;
  await commit(`${previous ? "修改" : "新增"}字段：${els.fieldEditLabel.value.trim()}`, () => {
    if (!previous) {
      item.fields.push({ id: uid("field"), label: els.fieldEditLabel.value.trim(), value, versioned });
    } else {
      previous.label = els.fieldEditLabel.value.trim();
      previous.versioned = versioned;
      if (editOverride) {
        if (value && value !== previous.value) version.fieldOverrides[previous.id] = value;
        else delete version.fieldOverrides[previous.id];
      } else {
        if (!(wasVersioned && !versioned && version.id !== "general")) previous.value = value;
      }
    }
    item.content = fieldsToContent(item.fields);
  });
  els.fieldEditorDialog.close(); restoreContentPosition(scrollState); showNotice("已保存"); notifyDataChanged();
}

function fieldEditorSnapshot() { return JSON.stringify({ label: els.fieldEditLabel.value, value: els.fieldEditValue.value, versioned: els.fieldVersioned.checked }); }
function itemEditorSnapshot() { return JSON.stringify({ category: els.editCategory.value, title: els.editTitle.value, tags: els.editTags.value, content: els.editContent.value }); }
function closeFieldEditorWithConfirmation() { if (state.fieldEditorInitial !== fieldEditorSnapshot()) return openDiscardConfirmation("field"); els.fieldEditorDialog.close(); }
function closeItemEditorWithConfirmation() { if (state.editorInitial !== itemEditorSnapshot()) return openDiscardConfirmation("item"); els.editorDialog.close(); }
function openDiscardConfirmation(target) { state.pendingDiscardTarget = target; if (!els.discardChangesDialog.open) els.discardChangesDialog.showModal(); }
function discardEditorChanges() {
  const target = state.pendingDiscardTarget; state.pendingDiscardTarget = null; els.discardChangesDialog.close();
  if (target === "field") els.fieldEditorDialog.close();
  if (target === "item") els.editorDialog.close();
}

function captureContentPosition() { return { scrollTop: els.contentPane.scrollTop, navigationId: els.quickNavigationSelect?.value || "" }; }
function restoreContentPosition(position) {
  setTimeout(() => {
    els.contentPane.scrollTop = position.scrollTop;
    if (position.navigationId && els.quickNavigationSelect?.querySelector(`option[value="${CSS.escape(position.navigationId)}"]`)) els.quickNavigationSelect.value = position.navigationId;
  }, 0);
}

function notifyDataChanged() {
  const message = { type: "job-paste-data-changed", source: VIEW_INSTANCE_ID };
  try { const sent = chrome.runtime.sendMessage?.(message); sent?.catch?.(() => {}); } catch {}
}

async function refreshSharedData() {
  const position = captureContentPosition(), internal = await loadInternalState();
  if (!internal.profile) return;
  state.profile = normalizeProfile(internal.profile);
  state.settings = normalizeSettings({ ...state.settings, versions: internal.versions }, state.profile);
  render(); restoreContentPosition(position); showNotice("信息已更新");
}

async function deleteEditingField() {
  const target = state.editingField; const item = state.profile.items.find(candidate => candidate.id === target?.itemId); const field = item?.fields?.[target.fieldIndex];
  if (!field || !window.confirm(`确定删除字段“${field.label}”吗？`)) return;
  await commit(`删除字段：${field.label}`, () => { item.fields.splice(target.fieldIndex, 1); for (const version of state.settings.versions) delete version.fieldOverrides[field.id]; item.content = fieldsToContent(item.fields); }); els.fieldEditorDialog.close();
}

async function duplicateEditingField() {
  const target = state.editingField; const item = state.profile.items.find(candidate => candidate.id === target?.itemId); const field = item?.fields?.[target.fieldIndex]; if (!field) return;
  const copy = structuredClone(field); copy.id = uid("field"); copy.label += "（副本）";
  await commit(`复制字段：${field.label}`, () => { item.fields.splice(target.fieldIndex + 1, 0, copy); item.content = fieldsToContent(item.fields); }); els.fieldEditorDialog.close();
}

function openFieldStatus(item, field) {
  const version = currentVersion();
  const override = field.versioned ? version.fieldOverrides[field.id] : null;
  state.statusField = { itemId: item.id, fieldId: field.id };
  els.fieldStatusTitle.textContent = field.label;
  els.fieldGeneralPreview.textContent = field.value;
  els.fieldCurrentBlock.hidden = version.id === "general" || !field.versioned;
  els.fieldCurrentLabel.textContent = `${version.name}内容`;
  els.fieldCurrentPreview.textContent = override ?? "未单独修改，当前继承通用内容";
  els.promoteFieldBtn.hidden = version.id === "general" || !field.versioned || override == null;
  els.restoreFieldBtn.hidden = version.id === "general" || !field.versioned || override == null;
  els.fieldStatusDialog.showModal();
}

async function promoteFieldToGeneral() {
  const target = state.statusField;
  const item = state.profile.items.find(candidate => candidate.id === target?.itemId);
  const field = item?.fields?.find(candidate => candidate.id === target?.fieldId);
  const version = currentVersion();
  const override = field ? version.fieldOverrides[field.id] : null;
  if (!field || override == null || !window.confirm("将当前版本内容设为通用内容，会影响所有继承通用内容的版本。确定继续吗？")) return;
  await commit(`将字段改为通用：${field.label}`, () => { field.value = override; delete version.fieldOverrides[field.id]; item.content = fieldsToContent(item.fields); });
  els.fieldStatusDialog.close();
}

async function restoreFieldToGeneral() {
  const target = state.statusField;
  const item = state.profile.items.find(candidate => candidate.id === target?.itemId);
  const field = item?.fields?.find(candidate => candidate.id === target?.fieldId);
  const version = currentVersion();
  if (!field || version.fieldOverrides[field.id] == null) return;
  await commit(`恢复通用内容：${field.label}`, () => delete version.fieldOverrides[field.id]);
  els.fieldStatusDialog.close();
}

function openCategoryManager() { els.categoryNotice.hidden = true; renderCategoryManager(); els.categoryDialog.showModal(); }
function renderCategoryManager() {
  els.categoryList.replaceChildren(...state.profile.categories.map(category => {
    const row = document.createElement("div"); row.className = "categoryManageRow"; row.dataset.id = category;
    makeDomSortable(row, els.categoryList, "category-manager", async () => { const names = [...els.categoryList.children].map(child => child.dataset.id); await commit("调整分类顺序", () => state.profile.categories = names); });
    const input = document.createElement("input"); input.value = category; input.maxLength = 16;
    row.append(dragHandle(), input, button("保存名称", "", () => renameCategory(category, input.value)), button("删除分类", "danger", () => deleteCategory(category))); return row;
  }));
}
async function addCategory() { const name = els.newCategoryName.value.trim(); if (!name || state.profile.categories.includes(name)) return showCategoryNotice("分类为空或已存在。"); await commit(`新增分类：${name}`, () => state.profile.categories.push(name)); els.newCategoryName.value = ""; renderCategoryManager(); }
async function renameCategory(oldName, value) { const name = value.trim(); if (!name || name === oldName || state.profile.categories.includes(name)) return showCategoryNotice("分类名称为空或已存在。"); await commit(`分类改名：${oldName} → ${name}`, () => { state.profile.categories = state.profile.categories.map(v => v === oldName ? name : v); state.profile.items.forEach(item => { if (item.category === oldName) item.category = name; }); if (state.selectedCategory === oldName) state.selectedCategory = name; }); renderCategoryManager(); }
async function deleteCategory(name) { const count = state.profile.items.filter(item => item.category === name).length; if (count) return showCategoryNotice(`分类内还有 ${count} 个条目，请先移动。`); if (!window.confirm(`确定删除空分类“${name}”吗？`)) return; await commit(`删除分类：${name}`, () => state.profile.categories = state.profile.categories.filter(v => v !== name)); renderCategoryManager(); }
function showCategoryNotice(text) { els.categoryNotice.textContent = text; els.categoryNotice.hidden = false; }

async function addAttachment() {
  try {
    const [handle] = await window.showOpenFilePicker({ multiple: false }); state.pendingAttachmentFile = await handle.getFile(); state.editingAttachment = null;
    els.attachmentDialogTitle.textContent = "新增附件"; els.attachmentTitle.value = state.pendingAttachmentFile.name.replace(/\.[^.]+$/, ""); els.attachmentCategory.value = "其他"; els.attachmentTags.value = ""; els.attachmentSensitive.checked = false; els.attachmentDefault.checked = false; els.attachmentDefaultLabel.hidden = currentVersion().id === "general"; els.deleteAttachmentBtn.hidden = true; renderAttachmentVersionOptions([]); els.attachmentDialog.showModal();
  } catch (error) { if (error.name !== "AbortError") showNotice(error.message); }
}

function openAttachmentEditor(item) { const version = currentVersion(); state.editingAttachment = item.id; state.pendingAttachmentFile = null; els.attachmentDialogTitle.textContent = "修改附件信息"; els.attachmentTitle.value = item.title; ensureSelectValue(els.attachmentCategory, item.category); els.attachmentTags.value = (item.tags || []).join(" "); els.attachmentSensitive.checked = item.sensitive; els.attachmentDefault.checked = version.defaultResumeId === item.id || (version.defaultAttachmentIds || []).includes(item.id); els.attachmentDefaultLabel.hidden = version.id === "general"; els.deleteAttachmentBtn.hidden = false; renderAttachmentVersionOptions(item.applicableVersionIds || []); els.attachmentDialog.showModal(); }

async function handleAttachmentSubmit(event) {
  event.preventDefault(); if (event.submitter?.value === "cancel") return els.attachmentDialog.close();
  if (state.pendingAttachmentFile) return commitNewAttachment();
  const item = state.attachments.find(candidate => candidate.id === state.editingAttachment); if (!item) return;
  const updated = { ...item, title: els.attachmentTitle.value.trim(), displayName: els.attachmentTitle.value.trim(), category: els.attachmentCategory.value, applicableVersionIds: selectedAttachmentVersionIds(), tags: splitTags(els.attachmentTags.value), sensitive: els.attachmentSensitive.checked, updatedAt: new Date().toISOString() };
  await commit(`修改附件：${updated.title}`, () => { state.attachments = state.attachments.map(candidate => candidate.id === item.id ? updated : candidate); updateDefaultAttachment(updated, els.attachmentDefault.checked); }, { attachments: true }); els.attachmentDialog.close();
}

async function commitNewAttachment() {
  const file = state.pendingAttachmentFile; const duplicate = state.attachments.find(item => item.fileName === file.name); let mode = "keep";
  if (duplicate) { const choice = window.prompt(`已存在同名附件“${file.name}”。输入“覆盖”“保留”或“取消”。`, "保留"); if (!choice || choice === "取消") return; mode = choice === "覆盖" ? "overwrite" : "keep"; }
  const category = els.attachmentCategory.value;
  const now = new Date().toISOString();
  const item = { id: mode === "overwrite" && duplicate ? duplicate.id : uid("attachment"), title: els.attachmentTitle.value.trim(), displayName: els.attachmentTitle.value.trim(), fileName: mode === "keep" && duplicate ? uniqueAttachmentName(file.name) : file.name, mimeType: file.type || "application/octet-stream", category, applicableVersionIds: selectedAttachmentVersionIds(), addedAt: duplicate?.addedAt || now, updatedAt: now, size: file.size, hash: await sha256(file), sensitive: els.attachmentSensitive.checked, tags: splitTags(els.attachmentTags.value), source: "indexeddb", blob: file };
  try { await commit(`${mode === "overwrite" ? "覆盖" : "新增"}附件：${item.title}`, () => { if (mode === "overwrite" && duplicate) state.attachments = state.attachments.filter(candidate => candidate.id !== duplicate.id); state.attachments.push(item); updateDefaultAttachment(item, els.attachmentDefault.checked); }, { attachments: true }); }
  catch (error) { return showNotice(storageErrorMessage(error)); }
  state.pendingAttachmentFile = null; els.attachmentDialog.close();
}

async function deleteEditingAttachment() { const item = state.attachments.find(candidate => candidate.id === state.editingAttachment); if (!item || !window.confirm(`确定删除附件“${item.title}”吗？`)) return; await commit(`删除附件：${item.title}`, () => { state.attachments = state.attachments.filter(candidate => candidate.id !== item.id); for (const version of state.settings.versions) { if (version.defaultResumeId === item.id) version.defaultResumeId = null; version.defaultAttachmentIds = version.defaultAttachmentIds.filter(id => id !== item.id); version.attachmentOrder = version.attachmentOrder.filter(id => id !== item.id); } }, { attachments: true }); els.attachmentDialog.close(); }

async function deleteAttachment(item) { state.editingAttachment = item.id; await deleteEditingAttachment(); }

function renderAttachmentVersionOptions(selectedIds) {
  const selected = new Set(selectedIds || []);
  els.attachmentVersionOptions.replaceChildren(...state.settings.versions.filter(version => version.id !== "general").map(version => {
    const label = document.createElement("label"); label.className = "checkboxLabel";
    const input = document.createElement("input"); input.type = "checkbox"; input.value = version.id; input.checked = selected.has(version.id);
    label.append(input, document.createTextNode(version.name)); return label;
  }));
}

function selectedAttachmentVersionIds() { return [...els.attachmentVersionOptions.querySelectorAll('input:checked')].map(input => input.value); }

function ensureSelectValue(select, value) { if (![...select.options].some(item => item.value === value)) select.append(option(value, value)); select.value = value; }

function updateDefaultAttachment(item, enabled) {
  const version = currentVersion();
  if (version.id === "general") return;
  if (version.defaultResumeId === item.id) version.defaultResumeId = null;
  version.defaultAttachmentIds = (version.defaultAttachmentIds || []).filter(id => id !== item.id);
  if (!enabled) return;
  if (/简历/.test(item.category)) {
    version.defaultResumeId = item.id;
    return;
  }
  version.defaultAttachmentIds.push(item.id);
}

async function setCurrentVersionDefaultAttachment(item, enabled) {
  if (currentVersion().id === "general") return showNotice("请先切换到具体填写版本再设置默认附件");
  await commit(`${enabled ? "设置" : "取消"}当前版本默认附件：${item.title}`, () => {
    updateDefaultAttachment(item, enabled);
    if (enabled && item.applicableVersionIds?.length && !item.applicableVersionIds.includes(currentVersion().id)) item.applicableVersionIds.push(currentVersion().id);
  }, { attachments: true });
}

async function startAttachmentUpload(item) {
  state.pendingUseAttachmentId = item.id;
  els.attachmentUseName.textContent = item.fileName;
  els.attachmentUseMeta.textContent = `${item.category} · ${attachmentVersionLabel(item)} · ${formatBytes(item.size)}`;
  els.attachmentMismatchWarning.hidden = !attachmentVersionMismatch(item);
  els.attachmentUseDialog.showModal();
}

function closeAttachmentUse() { state.pendingUseAttachmentId = null; els.attachmentUseDialog.close(); }
async function previewPendingUseAttachment() { const item = state.attachments.find(candidate => candidate.id === state.pendingUseAttachmentId); if (item) await previewAttachment(item); }
async function confirmAttachmentUse() {
  const item = state.attachments.find(candidate => candidate.id === state.pendingUseAttachmentId); if (!item) return closeAttachmentUse();
  els.attachmentUseDialog.close();
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true }); if (!tab?.id || !/^https?:/i.test(tab.url || "")) throw new Error("请先打开网申网页");
    const blob = await readInternalAttachmentBlob(item); const dataUrl = await blobToDataUrl(blob);
    await chrome.scripting.executeScript({ target: { tabId: tab.id }, func: installUploadPicker, args: [{ dataUrl, fileName: item.fileName, mimeType: item.mimeType, title: item.title }] }); window.close();
  } catch (error) { showNotice(`附件选择失败：${error.message}`); }
  finally { state.pendingUseAttachmentId = null; }
}

async function previewAttachment(item, options = {}) {
  const blob = options.blob || await readInternalAttachmentBlob(item);
  cleanupAttachmentPreview();
  state.previewState = { item, blob, page: 1, pages: 1, zoom: 1, temporary: Boolean(options.temporary), scrollTop: els.contentPane.scrollTop };
  els.attachmentPreviewTitle.textContent = item.title || item.fileName;
  els.attachmentPreviewMeta.textContent = `${item.fileName} · ${formatBytes(blob.size)} · ${fileExtension(item.fileName)}`;
  els.docxPreviewWarning.hidden = true; els.attachmentPreviewToolbar.hidden = false; els.attachmentPreviewBody.replaceChildren();
  const type = attachmentPreviewType(item, blob);
  if (type === "pdf") await renderPdfPreview(blob);
  else if (type === "image") renderImagePreview(blob, item);
  else if (type === "docx") await renderDocxPreview(blob);
  else renderUnsupportedPreview();
  els.markAttachmentVerifiedBtn.hidden = Boolean(options.temporary);
  els.attachmentVerifiedStatus.textContent = !options.temporary && state.verifiedAttachmentIds.has(item.id) ? "✓ 已核对" : "";
  els.attachmentPreviewDialog.showModal();
}

async function renderPdfPreview(blob) {
  state.previewObjectUrl = URL.createObjectURL(blob);
  state.previewState.pages = await countPdfPages(blob);
  const frame = document.createElement("iframe"); frame.title = "PDF 附件预览"; frame.dataset.previewType = "pdf"; els.attachmentPreviewBody.append(frame);
  updatePdfPreview();
}

function updatePdfPreview() {
  const preview = state.previewState, frame = els.attachmentPreviewBody.querySelector("iframe"); if (!preview || !frame) return;
  const zoom = preview.fit ? "page-width" : Math.round(preview.zoom * 100);
  frame.src = `${state.previewObjectUrl}#page=${preview.page}&zoom=${zoom}`;
  els.previewPageLabel.textContent = `${preview.page} / ${preview.pages || "?"}`;
  els.previewPrevBtn.disabled = preview.page <= 1; els.previewNextBtn.disabled = Boolean(preview.pages) && preview.page >= preview.pages;
}

function renderImagePreview(blob, item) {
  state.previewObjectUrl = URL.createObjectURL(blob);
  const image = document.createElement("img"); image.alt = item.title || item.fileName; image.src = state.previewObjectUrl; image.dataset.previewType = "image"; els.attachmentPreviewBody.append(image);
  els.previewPrevBtn.hidden = els.previewNextBtn.hidden = els.previewPageLabel.hidden = true;
}

async function renderDocxPreview(blob) {
  els.attachmentPreviewToolbar.hidden = true; els.docxPreviewWarning.hidden = false;
  const pre = document.createElement("pre"); pre.dataset.previewType = "docx";
  try { pre.textContent = await docxToText(blob); }
  catch (error) { pre.textContent = `无法解析正文：${error.message}`; }
  els.attachmentPreviewBody.append(pre);
}

function renderUnsupportedPreview() {
  els.attachmentPreviewToolbar.hidden = true; const message = document.createElement("div"); message.className = "unsupportedPreview"; message.textContent = "当前格式暂不支持内容预览。"; els.attachmentPreviewBody.append(message);
}

function changePreviewPage(offset) { const preview = state.previewState; if (!preview || !els.attachmentPreviewBody.querySelector("iframe")) return; preview.page = Math.max(1, Math.min(preview.pages || Infinity, preview.page + offset)); updatePdfPreview(); }
function changePreviewZoom(offset) { const preview = state.previewState; if (!preview) return; preview.fit = false; preview.zoom = Math.max(.4, Math.min(2.5, preview.zoom + offset)); const image = els.attachmentPreviewBody.querySelector("img"); if (image) image.style.width = `${Math.round(preview.zoom * 100)}%`; else updatePdfPreview(); }
function fitAttachmentPreview() { const preview = state.previewState; if (!preview) return; const image = els.attachmentPreviewBody.querySelector("img"); if (image) { preview.zoom = 1; image.style.width = "100%"; } else { preview.fit = true; updatePdfPreview(); } }
function closeAttachmentPreview() { els.attachmentPreviewDialog.close(); }
function cleanupAttachmentPreview() {
  if (state.previewObjectUrl) URL.revokeObjectURL(state.previewObjectUrl);
  state.previewObjectUrl = "";
  if (state.previewState?.scrollTop != null) restoreContentScroll(state.previewState.scrollTop);
  state.previewState = null;
  els.previewPrevBtn.hidden = els.previewNextBtn.hidden = els.previewPageLabel.hidden = false;
}
function markPreviewedAttachmentVerified() { const item = state.previewState?.item; if (!item) return; state.verifiedAttachmentIds.add(item.id); els.attachmentVerifiedStatus.textContent = "✓ 已核对"; const info = els.attachmentList.querySelector(`[data-attachment-id="${CSS.escape(item.id)}"] .attachmentInfo`); if (info && ![...info.children].some(child => child.textContent === "✓ 已核对")) { const checked = document.createElement("span"); checked.className = "verifiedAttachment"; checked.textContent = "✓ 已核对"; info.append(checked); } showNotice("已标记核对"); }

async function countPdfPages(blob) { const text = new TextDecoder("latin1").decode(await blob.arrayBuffer()); return Math.max(1, (text.match(/\/Type\s*\/Page\b/g) || []).length); }
function attachmentPreviewType(item, blob) { const name = (item.fileName || "").toLowerCase(), type = item.mimeType || blob.type || ""; if (type === "application/pdf" || name.endsWith(".pdf")) return "pdf"; if (/^image\/(png|jpeg|webp)$/.test(type) || /\.(png|jpe?g|webp)$/.test(name)) return "image"; if (type.includes("wordprocessingml") || name.endsWith(".docx")) return "docx"; return "unsupported"; }

async function replaceAttachmentFile(item) {
  try {
    const [handle] = await window.showOpenFilePicker({ multiple: false }), file = await handle.getFile();
    state.pendingReplacement = { attachmentId: item.id, file };
    els.replaceOldFile.textContent = item.fileName; els.replaceNewFile.textContent = file.name; els.replaceNewMeta.textContent = `${fileExtension(file.name)} · ${formatBytes(file.size)}`;
    els.attachmentReplaceDialog.showModal();
  } catch (error) { if (error.name !== "AbortError") showNotice(error.message); }
}
function closeAttachmentReplacement() { state.pendingReplacement = null; els.attachmentReplaceDialog.close(); }
async function previewPendingReplacement() { const pending = state.pendingReplacement, item = state.attachments.find(candidate => candidate.id === pending?.attachmentId); if (pending && item) await previewAttachment({ ...item, fileName: pending.file.name, mimeType: pending.file.type }, { blob: pending.file, temporary: true }); }
async function confirmAttachmentReplacement() {
  const pending = state.pendingReplacement, item = state.attachments.find(candidate => candidate.id === pending?.attachmentId); if (!pending || !item) return closeAttachmentReplacement();
  const file = pending.file, updated = { ...item, fileName: file.name, mimeType: file.type || "application/octet-stream", size: file.size, hash: await sha256(file), updatedAt: new Date().toISOString(), source: "indexeddb", blob: file };
  try { await commit(`替换附件文件：${item.title}`, () => { state.attachments = state.attachments.map(candidate => candidate.id === item.id ? updated : candidate); }, { attachments: true }); els.attachmentReplaceDialog.close(); state.pendingReplacement = null; showNotice("附件文件已替换"); }
  catch (error) { showNotice(storageErrorMessage(error)); }
}

function attachmentVersionMismatch(item) { const ids = item.applicableVersionIds || []; return currentVersion().id !== "general" && ids.length > 0 && !ids.includes("general") && !ids.includes(currentVersion().id); }
function attachmentVersionLabel(item) { const ids = item.applicableVersionIds || []; if (!ids.length) return "全部版本"; return ids.map(id => versionName(id)).join("、"); }
function fileExtension(name = "") { return (name.split(".").pop() || "文件").toUpperCase(); }
function formatAttachmentDate(value) { if (!value) return "未知"; const date = new Date(value); return Number.isNaN(date.getTime()) ? String(value).slice(0, 10) : localDate(date); }

async function activePageContext() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const context = { url: /^https?:/i.test(tab?.url || "") ? tab.url : "", title: String(tab?.title || "").trim(), company: "", job: "", location: "", jd: "" };
  if (!tab?.id || !context.url) return context;
  try { return { ...context, ...((await chrome.scripting.executeScript({ target: { tabId: tab.id }, func: extractStructuredJobPosting }))[0]?.result || {}) }; }
  catch { return context; }
}
async function activePageUrl() { return (await activePageContext()).url; }

function extractStructuredJobPosting() {
  const nodes = [...document.querySelectorAll('script[type="application/ld+json"]')]; let posting = null;
  for (const node of nodes) {
    try {
      const parsed = JSON.parse(node.textContent || "null"), candidates = Array.isArray(parsed) ? parsed : parsed?.["@graph"] || [parsed];
      posting = candidates.find(item => item?.["@type"] === "JobPosting"); if (posting) break;
    } catch {}
  }
  if (posting) {
    const locations = (Array.isArray(posting.jobLocation) ? posting.jobLocation : [posting.jobLocation]).filter(Boolean).map(location => location.address || location).map(address => [address.addressRegion, address.addressLocality, address.streetAddress].filter(Boolean).join(" ")).filter(Boolean);
    const holder = document.createElement("div"); holder.innerHTML = posting.description || "";
    return { company: posting.hiringOrganization?.name || "", job: posting.title || "", location: locations.join("、"), jd: holder.textContent?.trim() || "" };
  }
  const text = selectors => {
    for (const selector of selectors) {
      const value = document.querySelector(selector)?.textContent?.replace(/\s+/g, " ").trim();
      if (value && value.length <= 160) return value;
    }
    return "";
  };
  const job = text(["[data-job-title]", "h1", ".job-title", "[class*='job-title']"]);
  let company = text(["[data-company]", ".company-name", "[class*='company-name']", ".company"]);
  const location = text(["[data-location]", ".job-location", "[class*='job-location']", ".location"]);
  const jdNode = document.querySelector("[data-job-description], .job-description, [class*='job-description'], .job-detail, article, main");
  const jd = jdNode?.textContent?.replace(/\s+\n/g, "\n").replace(/[ \t]+/g, " ").trim().slice(0, 30000) || "";
  if (!company && job && document.title) {
    const titleParts = document.title.split(/\s[-|｜·]\s/).map(value => value.trim()).filter(Boolean);
    company = titleParts.find(value => value !== job && value.length <= 80) || "";
  }
  return { company, job, location, jd };
}

function renderRecordSuggestions() {
  const fill = (element, key) => element.replaceChildren(...[...new Set(state.records.map(record => record[key]).filter(Boolean))].map(value => option(value, value)));
  fill(els.companySuggestions, "公司名称"); fill(els.locationSuggestions, "工作地点"); fill(els.industrySuggestions, "行业"); fill(els.roleSuggestions, "岗位类别");
}

function renderResumeOptions() {
  const resumes = state.attachments.filter(item => /简历/.test(item.category));
  els.resumeSuggestions.replaceChildren(...resumes.map(item => option(item.title, item.title)));
  const preferred = resumes.find(item => item.id === currentVersion().defaultResumeId);
  if (preferred) els.completeResume.value = preferred.title;
}

async function openCompleteDialog() {
  await openApplicationDialog("complete");
}

async function openLaterDialog() {
  await openApplicationDialog("later");
}

async function openApplicationDialog(mode) {
  const record = await ensureActiveRecord(), context = await activePageContext(); state.submissionMode = mode; state.confirmedResumeName = "";
  els.completeForm.reset(); renderRecordSuggestions(); renderResumeOptions();
  els.completeDialogTitle.textContent = mode === "complete" ? "投递完成" : "稍后继续";
  els.completeSubmitBtn.textContent = mode === "complete" ? "保存并同步" : "保存待继续";
  els.completeLocation.required = mode === "complete"; els.completeResume.required = mode === "complete";
  const recommendation = recommendApplicationFields(record, context);
  setValue(els.completeCompany, record["公司名称"]); setValue(els.completeJob, record["岗位名称"]); setValue(els.completeLocation, record["工作地点"]);
  setValue(els.completeIndustry, record["行业"] || recommendation.industry); setValue(els.completeRole, record["岗位类别"] || recommendation.role);
  setValue(els.completeChannel, record["实际投递渠道"] || recommendation.channel); setValue(els.completeBatch, record["招聘类型/批次"] || recommendation.batch);
  setValue(els.completeCompanyType, record["企业性质"]); setValue(els.completePriority, record["投递优先级"]); setValue(els.completeReferralCode, record["内推码"]); setValue(els.completeReferrer, record["内推人"]);
  setValue(els.completeUrl, record["岗位链接"] || context.url); setValue(els.completeJd, record["岗位JD"]); setValue(els.completeNotes, record["备注"]);
  if (record["使用简历"]) els.completeResume.value = record["使用简历"];
  els.completeDialog.showModal();
}

async function openAbortDialog() {
  const record = await ensureActiveRecord(), context = await activePageContext(); els.abortForm.reset(); renderRecordSuggestions();
  setValue(els.abortCompany, record["公司名称"]); setValue(els.abortJob, record["岗位名称"]); setValue(els.abortLocation, record["工作地点"]);
  setValue(els.abortUrl, record["岗位链接"] || context.url); setValue(els.abortReason, record["中止原因"]); setValue(els.abortNotes, record["备注"]); els.abortDialog.showModal();
}

function createRecord(values = {}) {
  const now = new Date(), timestamp = now.toISOString();
  return {
    "记录ID": applicationRecordId(now), "公司名称": values.company || "", "岗位名称": values.job || "", "工作地点": values.location || "",
    "行业": values.industry || "", "岗位类别": values.role || "", "岗位JD": values.jd || "", "状态": values.status || "填写中",
    "投递日期": "", "完成时间": "", "岗位链接": values.url || "", "页面标题": values.pageTitle || "", "使用简历": values.resume || "",
    "填写版本": versionName(state.settings.currentVersion), "填写版本ID": state.settings.currentVersion, "中止原因": values.abortReason || "", "备注": values.notes || "",
    "内推码": values.referralCode || "", "内推人": values.referrer || "", "实际投递渠道": values.channel || "", "招聘类型/批次": values.batch || "", "企业性质": values.companyType || "", "投递优先级": values.priority || "", "有效性修正": "",
    "创建时间": timestamp, "最近更新时间": timestamp, "飞书同步状态": "未同步", "飞书记录ID": "", enrichment: {}
  };
}

async function ensureActiveRecord() {
  const context = await activePageContext(), active = state.records.find(record => record["记录ID"] === state.activeRecordId);
  if (active && ["填写中", "待继续"].includes(active["状态"])) {
    if (!active["岗位链接"] || !context.url || normalizedJobUrl(active["岗位链接"]) === normalizedJobUrl(context.url)) {
      active["岗位链接"] ||= context.url; active["页面标题"] ||= context.title; return active;
    }
    active["状态"] = "待继续"; active["最近更新时间"] = new Date().toISOString(); await persistRecords();
  }
  const record = createRecord({ url: context.url, pageTitle: context.title, company: context.company, job: context.job, location: context.location, jd: context.jd });
  state.records.unshift(record); state.activeRecordId = record["记录ID"]; await persistRecords(); await persistActiveRecordId(); return record;
}

function readApplicationForm() {
  return { company: els.completeCompany.value.trim(), job: els.completeJob.value.trim(), location: els.completeLocation.value.trim(), resume: els.completeResume.value.trim(), industry: els.completeIndustry.value.trim(), role: els.completeRole.value.trim(), jd: els.completeJd.value.trim(), url: els.completeUrl.value.trim(), notes: els.completeNotes.value.trim(), channel: els.completeChannel.value.trim(), batch: els.completeBatch.value.trim(), companyType: els.completeCompanyType.value.trim(), priority: els.completePriority.value, referralCode: els.completeReferralCode.value.trim(), referrer: els.completeReferrer.value.trim() };
}

async function updateActiveRecord(values, status) {
  const record = await ensureActiveRecord(), now = new Date(), completed = status === "已投递" || status === "已放弃";
  Object.assign(record, {
    "公司名称": values.company ?? record["公司名称"], "岗位名称": values.job ?? record["岗位名称"], "工作地点": values.location ?? record["工作地点"],
    "行业": values.industry ?? record["行业"], "岗位类别": values.role ?? record["岗位类别"], "岗位JD": values.jd ?? record["岗位JD"], "岗位链接": values.url ?? record["岗位链接"], "使用简历": values.resume ?? record["使用简历"],
    "内推码": values.referralCode ?? record["内推码"], "内推人": values.referrer ?? record["内推人"], "实际投递渠道": values.channel ?? record["实际投递渠道"], "招聘类型/批次": values.batch ?? record["招聘类型/批次"], "企业性质": values.companyType ?? record["企业性质"], "投递优先级": values.priority ?? record["投递优先级"],
    "中止原因": values.abortReason ?? record["中止原因"], "备注": values.notes ?? record["备注"], "状态": status,
    "填写版本": versionName(state.settings.currentVersion), "填写版本ID": state.settings.currentVersion, "最近更新时间": now.toISOString()
  });
  if (status === "已投递") record["投递日期"] ||= localDate(now);
  if (completed) record["完成时间"] = localDateTime(now);
  await persistRecords(); return record;
}

async function handleCompleteSubmit(event) {
  event.preventDefault(); if (event.submitter?.value === "cancel") return els.completeDialog.close();
  const selectedResume = state.attachments.find(item => item.title === els.completeResume.value);
  if (state.submissionMode === "complete" && selectedResume && attachmentVersionMismatch(selectedResume) && state.confirmedResumeName !== selectedResume.title) {
    els.resumeMismatchSummary.textContent = `当前填写版本：${versionName(state.settings.currentVersion)}；准备使用：${selectedResume.title}`;
    els.resumeMismatchDialog.showModal(); return;
  }
  const status = state.submissionMode === "later" ? "待继续" : "已投递", record = await updateActiveRecord(readApplicationForm(), status);
  els.completeDialog.close(); await clearActiveRecord();
  if (status === "待继续") return showNotice("已保存，可从投递记录继续填写");
  await finishRecordSave(record);
}

function confirmMismatchedResume() { state.confirmedResumeName = els.completeResume.value; els.resumeMismatchDialog.close(); els.completeForm.requestSubmit(els.completeSubmitBtn); }

async function handleAbortSubmit(event) {
  event.preventDefault(); if (event.submitter?.value === "cancel") return els.abortDialog.close();
  const record = await updateActiveRecord({ company: els.abortCompany.value.trim(), job: els.abortJob.value.trim(), location: els.abortLocation.value.trim(), url: els.abortUrl.value.trim(), abortReason: els.abortReason.value, notes: els.abortNotes.value.trim() }, "已放弃");
  els.abortDialog.close(); await clearActiveRecord();
  if (state.feishuConfig.enabled && record["飞书记录ID"]) await finishRecordSave(record); else showNotice("已保存为放弃投递");
}

async function finishRecordSave(record) {
  if (!state.feishuConfig.enabled) return showNotice(record["状态"] === "已投递" ? "投递记录已保存" : "放弃记录已保存");
  const result = await syncRecord(record["记录ID"]);
  if (result.ok) {
    if (record["状态"] === "已投递" && state.settings.celebration) celebrate("已保存并同步至飞书");
    else showNotice("已保存并同步至飞书");
  } else showSyncResult(record, result.error);
}

async function persistRecords() {
  await saveRecords(state.records);
  renderRecordSuggestions(); renderPendingRecordBadge();
}

async function syncRecord(recordId) {
  const record = state.records.find(item => item["记录ID"] === recordId); if (!record) return { ok: false, error: new Error("未找到本地记录") };
  record["飞书同步状态"] = "同步中"; await persistRecords();
  try {
    const result = await syncFeishuRecord(state.feishuConfig, record); record["飞书记录ID"] = result.recordId;
    record["飞书同步状态"] = "已同步"; await persistRecords(); return { ok: true };
  } catch (error) {
    record["飞书同步状态"] = error.unknown ? "状态未知" : "同步失败"; await persistRecords(); return { ok: false, error };
  }
}

function openFeishuSettings() {
  const config = state.feishuConfig; els.feishuEnabled.checked = config.enabled; els.feishuAppId.value = config.appId; els.feishuAppSecret.value = config.appSecret; els.feishuTableUrl.value = config.tableUrl; els.feishuAppToken.value = config.appToken; els.feishuTableId.value = config.tableId;
  els.mappingFields.replaceChildren(...FEISHU_FIELDS.map(source => { const label = document.createElement("label"); label.textContent = source; const input = document.createElement("input"); input.value = config.mapping[source] || source; input.dataset.source = source; label.append(input); return label; }));
  els.feishuTestResult.hidden = true; els.feishuDialog.showModal();
}

function readFeishuForm() {
  const mapping = { ...DEFAULT_FEISHU_CONFIG.mapping }; els.mappingFields.querySelectorAll("input").forEach(input => mapping[input.dataset.source] = input.value.trim() || input.dataset.source);
  return normalizeFeishuConfig({ enabled: els.feishuEnabled.checked, appId: els.feishuAppId.value.trim(), appSecret: els.feishuAppSecret.value, tableUrl: els.feishuTableUrl.value.trim(), appToken: els.feishuAppToken.value.trim(), tableId: els.feishuTableId.value.trim(), mapping });
}

async function handleFeishuSettingsSubmit(event) {
  event.preventDefault(); if (event.submitter?.value === "cancel") return els.feishuDialog.close();
  try {
    let config = readFeishuForm();
    if (config.enabled && /\/wiki\//i.test(config.tableUrl)) config = (await testFeishuConnection(config)).config;
    else if (config.enabled) validateFeishuConfig(config);
    state.feishuConfig = config; await chrome.storage.local.set({ applyPasteFeishuConfig: config }); els.feishuDialog.close(); showNotice("飞书同步设置已保存");
  }
  catch (error) { showFeishuTest(error.message, true); }
}

async function testFeishuSettings() {
  try {
    els.testFeishuBtn.disabled = true; showFeishuTest("正在连接飞书…");
    const result = await testFeishuConnection(readFeishuForm()), resolved = result.config; state.feishuConfig = resolved;
    els.feishuAppToken.value = resolved.appToken; els.feishuTableId.value = resolved.tableId;
    await chrome.storage.local.set({ applyPasteFeishuConfig: resolved });
    const names = result.fields.map(field => `${field.field_name} (${field.field_id}, type ${field.type})`).join("；");
    showFeishuTest(`连接成功，读取 ${result.fields.length} 个真实字段：${names}`);
  }
  catch (error) { showFeishuTest(`连接失败：${error.message}`, true); }
  finally { els.testFeishuBtn.disabled = false; }
}

function showFeishuTest(message, failed = false) { els.feishuTestResult.textContent = message; els.feishuTestResult.hidden = false; els.feishuTestResult.classList.toggle("failed", failed); }

async function clearFeishuSettings() {
  if (!window.confirm("确定清除本机保存的飞书 App ID、App Secret 和表格配置吗？")) return;
  state.feishuConfig = structuredClone(DEFAULT_FEISHU_CONFIG); await chrome.storage.local.remove("applyPasteFeishuConfig");
  els.feishuEnabled.checked = false; els.feishuAppId.value = ""; els.feishuAppSecret.value = ""; els.feishuTableUrl.value = ""; els.feishuAppToken.value = ""; els.feishuTableId.value = "";
  els.mappingFields.querySelectorAll("input").forEach(input => input.value = input.dataset.source); showFeishuTest("配置已清除");
}

function setValue(element, value) { element.value = value == null ? "" : String(value); }
function applicationRecordId(date = new Date()) {
  const parts = [date.getFullYear(), String(date.getMonth() + 1).padStart(2, "0"), String(date.getDate()).padStart(2, "0")].join("");
  const time = [date.getHours(), date.getMinutes(), date.getSeconds()].map(value => String(value).padStart(2, "0")).join("");
  return `JP-${parts}-${time}-${crypto.randomUUID().replaceAll("-", "").slice(0, 4).toUpperCase()}`;
}
function recommendApplicationFields(record, context) {
  const text = `${context.title} ${context.url} ${record["岗位JD"]}`.toLowerCase(), version = currentVersion().name;
  const channel = /nowcoder|牛客/.test(text) ? "牛客" : /zhipin|boss直聘/.test(text) ? "BOSS" : /liepin/.test(text) ? "猎聘" : /zhaopin/.test(text) ? "智联招聘" : /51job/.test(text) ? "前程无忧" : /career|campus|recruit|jobs?\./.test(text) ? "官网" : "";
  const role = /半导体|工艺|量测|可靠性/.test(version) ? "半导体工艺/分析" : /材料研发/.test(version) ? "材料研发" : /教师/.test(version) ? "教师" : "";
  const industry = /半导体|集成电路|晶圆/.test(text) ? "半导体" : /材料|化学|化工/.test(text) ? "新材料" : "";
  const batch = /秋招|校园招聘|校招|campus/.test(text) ? "秋招/校园招聘" : "";
  return { channel, role, industry, batch };
}
async function persistActiveRecordId() { await chrome.storage.local.set({ applyPasteActiveRecordId: state.activeRecordId || "" }); }
async function clearActiveRecord() { state.activeRecordId = ""; await persistActiveRecordId(); }
function normalizedJobUrl(value) {
  try {
    const url = new URL(value); url.hash = "";
    const stable = [...url.searchParams].filter(([key]) => /job|position|requisition|posting|vacancy|adid|submission/i.test(key)).sort(([a], [b]) => a.localeCompare(b));
    const query = new URLSearchParams(stable).toString(); return `${url.origin}${url.pathname.replace(/\/$/, "")}${query ? `?${query}` : ""}`;
  } catch { return value || ""; }
}

async function detectPendingRecord() {
  if (state.activeRecordId || els.resumeRecordDialog.open) return;
  const context = await activePageContext(), url = normalizedJobUrl(context.url); if (!url) return;
  const match = [...state.records].sort((a, b) => String(b["最近更新时间"]).localeCompare(String(a["最近更新时间"]))).find(record => record["状态"] === "待继续" && normalizedJobUrl(record["岗位链接"]) === url);
  if (!match) return;
  state.pendingResumeRecordId = match["记录ID"]; els.resumeRecordSummary.textContent = `${match["公司名称"] || "未填写公司"}｜${match["岗位名称"] || context.title || "未填写岗位"}`; els.resumeRecordDialog.showModal();
}

async function startNewRecordFromPrompt() {
  els.resumeRecordDialog.close(); state.pendingResumeRecordId = ""; await clearActiveRecord(); await ensureActiveRecord(); showNotice("已新建填写记录");
}

async function continueRecord(recordId) {
  const record = state.records.find(item => item["记录ID"] === recordId); if (!record) return;
  record["状态"] = "填写中"; record["最近更新时间"] = new Date().toISOString(); state.activeRecordId = record["记录ID"];
  const versionId = record["填写版本ID"] || state.settings.versions.find(version => version.name === record["填写版本"])?.id;
  if (versionId && state.settings.versions.some(version => version.id === versionId)) state.settings.currentVersion = versionId;
  await persistRecords(); await persistActiveRecordId(); await saveUiSettings(); render();
  if (els.resumeRecordDialog.open) els.resumeRecordDialog.close(); if (els.recordsDialog.open) els.recordsDialog.close();
  const context = await activePageContext(); if (record["岗位链接"] && normalizedJobUrl(record["岗位链接"]) !== normalizedJobUrl(context.url)) openExternal(record["岗位链接"]);
  showNotice("已恢复原投递记录");
}

function renderPendingRecordBadge() {
  const count = state.records.filter(record => record["状态"] === "待继续").length;
  els.pendingRecordBadge.hidden = count < 1; els.pendingRecordBadge.textContent = count > 99 ? "99+" : String(count);
}

function openRecords() { els.recordStatusFilter.value = state.recordFilter; els.recordSearchInput.value = state.recordQuery; renderRecords(); els.recordsDialog.showModal(); }

function exportRecordsCsv() {
  const blob = new Blob([createApplicationCsv(state.records)], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const fileName = `job-paste-applications-${localDate(new Date()).replaceAll("-", "")}.csv`;
  chrome.downloads.download({ url, filename: fileName, saveAs: true }, () => {
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    if (chrome.runtime.lastError) showNotice(`CSV导出失败：${chrome.runtime.lastError.message}`);
    else showNotice(`已导出 ${state.records.length} 条投递记录 CSV`);
  });
}

function renderRecords() {
  const today = localDate(new Date()), completed = state.records.filter(record => record["状态"] === "已投递" && record["投递日期"] === today).length;
  els.todayRecordCount.textContent = `今日已投递 ${completed} 份`;
  const records = [...state.records].sort((a, b) => String(b["最近更新时间"] || b["完成时间"]).localeCompare(String(a["最近更新时间"] || a["完成时间"]))).filter(record => {
    const statusMatch = state.recordFilter === "全部" || (state.recordFilter === "同步失败" ? ["同步失败", "状态未知"].includes(record["飞书同步状态"]) : record["状态"] === state.recordFilter);
    return statusMatch && (!state.recordQuery || `${record["公司名称"]} ${record["岗位名称"]}`.toLowerCase().includes(state.recordQuery));
  });
  els.recordsList.replaceChildren(...records.map(record => {
    const row = document.createElement("article"); row.className = "recordRow";
    const main = document.createElement("div"); main.className = "recordMain"; const title = document.createElement("strong"); title.textContent = `${record["公司名称"] || "未填写公司"}｜${record["岗位名称"] || "未填写岗位"}`; const meta = document.createElement("span"); meta.textContent = `${record["工作地点"] || "地点未填"} · ${formatRecordTime(record["最近更新时间"] || record["完成时间"])}`; const status = document.createElement("span"); status.className = `recordStatus ${recordStatusClass(record)}`; status.textContent = record["状态"]; const sync = document.createElement("span"); sync.className = "recordSync"; sync.textContent = `飞书 · ${record["飞书同步状态"]}`; main.append(title, meta, status, sync);
    const actions = document.createElement("div"); actions.className = "recordActions";
    let primaryAction = "查看";
    if (["待继续", "填写中"].includes(record["状态"])) { primaryAction = "继续填写"; actions.append(button(primaryAction, "primary", () => continueRecord(record["记录ID"]))); }
    else if (record["状态"] === "已放弃") { primaryAction = "恢复填写"; actions.append(button(primaryAction, "primary", () => continueRecord(record["记录ID"]))); }
    else if (["同步失败", "状态未知", "已同步"].includes(record["飞书同步状态"])) { primaryAction = record["飞书同步状态"] === "已同步" ? "同步更新" : "重新同步"; actions.append(button(primaryAction, "primary", () => retryRecord(record))); }
    else actions.append(button("查看", "primary", () => openRecordDetail(record, true)));
    const more = document.createElement("details"); more.className = "recordMore"; const moreSummary = document.createElement("summary"); moreSummary.textContent = "···"; moreSummary.title = "更多操作"; const menu = document.createElement("div"); menu.className = "recordMoreMenu";
    if (primaryAction !== "查看") menu.append(button("查看详情", "", () => openRecordDetail(record, true)));
    if (record["状态"] === "已投递") menu.append(button("打开岗位", "", () => openExternal(record["岗位链接"])));
    menu.append(button("删除本地记录", "danger", () => deleteRecord(record))); more.append(moreSummary, menu); actions.append(more);
    row.append(main, actions); return row;
  }));
  if (!records.length) els.recordsList.textContent = state.records.length ? "没有匹配记录" : "暂无投递记录";
}

function recordStatusClass(record) {
  if (["同步失败", "状态未知"].includes(record["飞书同步状态"])) return "statusFailed";
  if (["待继续", "填写中"].includes(record["状态"])) return "statusPending";
  if (record["状态"] === "已投递") return "statusComplete";
  if (record["状态"] === "已放弃") return "statusAborted";
  return "statusNeutral";
}

function openRecordDetail(record, editable) {
  state.editingRecordId = record["记录ID"]; state.recordDetailMode = editable ? "edit" : "view"; els.recordDetailTitle.textContent = editable ? "编辑本地记录" : "查看记录";
  const values = [[els.detailCompany,"公司名称"],[els.detailJob,"岗位名称"],[els.detailLocation,"工作地点"],[els.detailStatus,"状态"],[els.detailIndustry,"行业"],[els.detailRole,"岗位类别"],[els.detailChannel,"实际投递渠道"],[els.detailBatch,"招聘类型/批次"],[els.detailCompanyType,"企业性质"],[els.detailPriority,"投递优先级"],[els.detailReferralCode,"内推码"],[els.detailReferrer,"内推人"],[els.detailUrl,"岗位链接"],[els.detailJd,"岗位JD"],[els.detailResume,"使用简历"],[els.detailAbortReason,"中止原因"],[els.detailNotes,"备注"]];
  for (const [input, key] of values) { input.value = record[key] || ""; input.disabled = !editable; }
  els.detailMetadata.replaceChildren(...[["ApplyPaste记录ID",record["记录ID"]],["飞书记录ID",record["飞书记录ID"] || "未创建"],["飞书同步状态",record["飞书同步状态"]],["创建时间",formatRecordTime(record["创建时间"])],["最近更新时间",formatRecordTime(record["最近更新时间"])]].map(([label, value]) => { const row = document.createElement("div"); row.append(Object.assign(document.createElement("span"), { textContent: label }), Object.assign(document.createElement("strong"), { textContent: value })); return row; }));
  els.saveRecordEditBtn.hidden = !editable; els.localEditWarning.hidden = !editable || record["飞书同步状态"] !== "已同步"; els.recordDetailDialog.showModal();
}

async function handleRecordDetailSubmit(event) {
  event.preventDefault(); if (event.submitter?.value === "cancel" || state.recordDetailMode !== "edit") return els.recordDetailDialog.close();
  const record = state.records.find(item => item["记录ID"] === state.editingRecordId); if (!record) return;
  Object.assign(record, { "公司名称": els.detailCompany.value.trim(), "岗位名称": els.detailJob.value.trim(), "工作地点": els.detailLocation.value.trim(), "状态": els.detailStatus.value, "行业": els.detailIndustry.value.trim(), "岗位类别": els.detailRole.value.trim(), "实际投递渠道": els.detailChannel.value.trim(), "招聘类型/批次": els.detailBatch.value.trim(), "企业性质": els.detailCompanyType.value.trim(), "投递优先级": els.detailPriority.value.trim(), "内推码": els.detailReferralCode.value.trim(), "内推人": els.detailReferrer.value.trim(), "岗位链接": els.detailUrl.value.trim(), "岗位JD": els.detailJd.value.trim(), "使用简历": els.detailResume.value.trim(), "中止原因": els.detailAbortReason.value.trim(), "备注": els.detailNotes.value.trim(), "最近更新时间": new Date().toISOString() });
  await persistRecords(); els.recordDetailDialog.close(); renderRecords(); showNotice("本地记录已更新");
}

async function retryRecord(record) {
  if (!state.feishuConfig.enabled) return showNotice("请先在设置中启用飞书同步");
  if (record["飞书同步状态"] === "同步中") return showNotice("该记录正在同步");
  const result = await syncRecord(record["记录ID"]); renderRecords();
  if (result.ok) showNotice("已同步至飞书"); else showSyncResult(record, result.error);
}

async function retryResultRecord() { const record = state.records.find(item => item["记录ID"] === state.syncResultRecordId); if (!record) return; els.syncResultDialog.close(); await retryRecord(record); }

function showSyncResult(record, error) {
  state.syncResultRecordId = record["记录ID"]; const unknown = record["飞书同步状态"] === "状态未知"; els.syncResultTitle.textContent = unknown ? "飞书同步状态未知" : "飞书同步失败";
  els.syncResultText.textContent = unknown ? `记录已保存在本地。${error.message}，请先查看飞书表格，再决定是否重新同步。` : `记录已保存在本地，飞书同步失败：${error.message}`; els.syncResultDialog.showModal();
}

async function deleteRecord(record) { if (!window.confirm(`确定删除“${record["公司名称"]} · ${record["岗位名称"]}”的本地记录吗？`)) return; state.records = state.records.filter(item => item["记录ID"] !== record["记录ID"]); await persistRecords(); renderRecords(); }
function openExternal(url) { if (url) chrome.tabs.create({ url }); else showNotice("未保存链接"); }
function openFeishuTable() { openExternal(state.feishuConfig.tableUrl); }
function formatRecordTime(value) { if (!value) return "时间未记录"; const date = new Date(value); return Number.isNaN(date.getTime()) ? value : localDateTime(date); }

function celebrate(message) { const today = localDate(new Date()); const count = state.records.filter(r => r["状态"] === "已投递" && r["投递日期"] === today).length; els.celebrationText.textContent = `${message} · 今日已完成第 ${count} 份投递`; els.celebration.hidden = false; setTimeout(() => els.celebration.hidden = true, 1600); }

async function commit(change, mutate, options = {}) {
  setSaveState("saving");
  mutate(); state.profile.updatedAt = new Date().toISOString();
  try {
    await saveCoreData(state.profile, state.settings.versions);
    if (options.attachments) await saveAttachments(state.attachments);
    await saveUiSettings(); setSaveState("saved");
  } catch (error) { setSaveState("failed"); throw error; }
  render();
}
async function saveUiSettings() { await chrome.storage.local.set({ applyPasteUiSettings: { currentVersion: state.settings.currentVersion, celebration: state.settings.celebration } }); }
async function saveSettings() { setSaveState("saving"); try { await saveCoreData(state.profile, state.settings.versions); await saveUiSettings(); setSaveState("saved"); } catch { setSaveState("failed"); } }

function chooseProfileExport() {
  const choice = window.prompt("导出格式：输入 JSON、MD 或 DOCX", "JSON");
  if (!choice) return;
  const type = choice.trim().toLowerCase();
  if (!['json', 'md', 'docx'].includes(type)) return showNotice("当前可导出 JSON、Markdown 或 Word DOCX");
  exportProfile(type);
}

async function exportCompleteBackup() {
  try {
    setSaveState("saving");
    const attachments = await Promise.all(state.attachments.map(async item => {
      const { blob, ...metadata } = item;
      return { ...metadata, blobData: blob ? await blobToDataUrl(blob) : "" };
    }));
    const backup = {
      format: "job-paste-complete-backup", version: 1, createdAt: new Date().toISOString(),
      profile: state.profile, versions: state.settings.versions, records: state.records, attachments,
      settings: { currentVersion: state.settings.currentVersion, celebration: state.settings.celebration }
    };
    downloadJson(backup, `job-paste-backup-${localDate(new Date()).replaceAll("-", "")}.json`); setSaveState("saved"); showNotice("完整备份已导出");
  } catch (error) { setSaveState("failed"); showNotice(`备份失败：${error.message}`); }
}

async function restoreCompleteBackup() {
  try {
    const file = await pickFile([{ description: "ApplyPaste 完整备份", accept: { "application/json": [".json"] } }]);
    const backup = JSON.parse(await file.text());
    if (backup.format !== "job-paste-complete-backup" || !backup.profile || !Array.isArray(backup.attachments)) throw new Error("不是有效的 ApplyPaste 完整备份");
    if (!window.confirm(`将恢复 ${backup.profile.items?.length || 0} 个信息条目、${backup.attachments.length} 个附件和 ${backup.records?.length || 0} 条投递记录。是否继续？`)) return;
    const attachments = await Promise.all(backup.attachments.map(async item => { const { blobData, ...metadata } = item; return { ...metadata, blob: blobData ? await dataUrlToBlob(blobData) : null, source: "indexeddb" }; }));
    state.profile = normalizeProfile(backup.profile); state.settings = normalizeSettings({ ...state.settings, versions: backup.versions, ...(backup.settings || {}) }, state.profile); state.attachments = normalizeAttachments(attachments); state.records = normalizeRecords(backup.records || []); state.selectedCategory = "全部";
    await saveCompleteState({ profile: state.profile, versions: state.settings.versions, records: state.records, attachments: state.attachments }); await saveUiSettings(); render(); showNotice("完整备份已恢复");
  } catch (error) { if (error.name !== "AbortError") showNotice(`恢复失败：${storageErrorMessage(error)}`); }
}

async function migrateLegacyWorkspace() {
  try {
    const stored = await chrome.storage.local.get("applyPasteWorkspaceMigrated");
    if (stored.applyPasteWorkspaceMigrated && !window.confirm("旧工作区已经迁移过。是否重新选择目录并再次迁移？")) return;
    const handle = await window.showDirectoryPicker({ mode: "read", id: "job-paste-legacy-migration" });
    const [workspace, records] = await Promise.all([readWorkspaceState(handle), readWorkspaceRecords(handle)]);
    if (!workspace?.profile) throw new Error("所选目录中没有 data/profile.json");
    const sourceAttachments = normalizeAttachments(workspace.attachments || []); let missingFiles = 0;
    const attachments = await Promise.all(sourceAttachments.map(async item => {
      try { return { ...item, blob: await readWorkspaceAttachmentBlob(handle, item), source: "indexeddb", localPath: null }; }
      catch { missingFiles++; return { ...item, blob: null, source: "indexeddb", localPath: null }; }
    }));
    const summary = `迁移检查\n信息条目：${workspace.profile.items?.length || 0}\n附件：${attachments.length}\n投递记录：${records.length}\n未读取附件：${missingFiles}\n\n迁移会更新插件内部数据，旧工作区文件保持不变。是否继续？`;
    if (!window.confirm(summary)) return;
    state.profile = normalizeProfile(workspace.profile); state.settings = normalizeSettings({ ...(workspace.settings || state.settings), currentVersion: state.settings.currentVersion }, state.profile); state.attachments = attachments; state.records = normalizeRecords(records); state.selectedCategory = "全部";
    await saveCompleteState({ profile: state.profile, versions: state.settings.versions, records: state.records, attachments: state.attachments }); await saveUiSettings(); await chrome.storage.local.set({ applyPasteWorkspaceMigrated: { completedAt: new Date().toISOString(), sourceName: handle.name } });
    render(); showNotice(`迁移完成：${state.profile.items.length} 个条目、${state.attachments.length} 个附件、${state.records.length} 条记录`);
  } catch (error) { if (error.name !== "AbortError") showNotice(`迁移失败：${storageErrorMessage(error)}`); }
}

async function exportProfile(type) {
  const exported = { ...state.profile, fillingVersions: state.settings.versions };
  downloadProfile(exported, type);
}

function setSaveState(status) {
  if (!els.saveStatus) return;
  els.saveStatus.className = status === "saving" ? "saving" : status === "failed" ? "failed" : "";
  els.saveStatus.textContent = status === "saving" ? "保存中" : status === "failed" ? "保存失败" : "已保存";
}

async function reorderCategories(source, target) { await commit("调整分类顺序", () => moveByValue(state.profile.categories, source, target)); }
async function reorderItems(source, target) {
  const version = currentVersion();
  const versionKey = state.selectedCategory === "项目经历" ? "projectOrder" : state.selectedCategory === "技能" ? "skillOrder" : null;
  if (version.id !== "general" && versionKey) return commit(`调整${state.selectedCategory}版本顺序`, () => { if (!version[versionKey].length) version[versionKey] = state.profile.items.filter(item => item.category === state.selectedCategory).map(item => item.id); moveByValue(version[versionKey], source, target); });
  await commit("调整条目顺序", () => moveObjects(state.profile.items, source, target));
}
async function reorderFields(itemId, source, target) {
  const item = state.profile.items.find(v => v.id === itemId); if (!item) return;
  const version = currentVersion();
  if (version.id !== "general" && item.category === "技能") return commit(`调整技能版本顺序：${item.title}`, () => { const order = version.skillFieldOrders[item.id] ||= item.fields.map(field => field.id); moveByValue(order, source, target); });
  await commit(`调整字段顺序：${item.title}`, () => { moveObjects(item.fields, source, target); item.content = fieldsToContent(item.fields); });
}
async function reorderAttachments(source, target) {
  const version = currentVersion();
  if (version.id !== "general") return commit("调整当前版本附件优先级", () => { if (!version.attachmentOrder.length) version.attachmentOrder = state.attachments.map(item => item.id); moveByValue(version.attachmentOrder, source, target); });
  await commit("调整附件顺序", () => moveObjects(state.attachments, source, target), { attachments: true });
}
function canReorderItems() { return !state.query && state.profile.categories.includes(state.selectedCategory); }

function makeDraggable(element, id, group, onDrop) {
  element.draggable = true; element.dataset.dragId = id; element.dataset.dragGroup = group;
  element.addEventListener("dragstart", event => { event.stopPropagation(); element.classList.add("dragging"); event.dataTransfer.setData("text/plain", JSON.stringify({ id, group })); event.dataTransfer.effectAllowed = "move"; });
  element.addEventListener("dragend", () => element.classList.remove("dragging"));
  element.addEventListener("dragover", event => { event.preventDefault(); event.dataTransfer.dropEffect = "move"; });
  element.addEventListener("drop", event => { event.preventDefault(); event.stopPropagation(); try { const data = JSON.parse(event.dataTransfer.getData("text/plain")); if (data.group === group && data.id !== id) onDrop(data.id, id); } catch {} });
}

function makeDomSortable(row, container, group, afterDrop = null) {
  row.draggable = true; row.addEventListener("dragstart", event => { row.classList.add("dragging"); event.dataTransfer.setData("text/plain", group); }); row.addEventListener("dragend", () => row.classList.remove("dragging")); row.addEventListener("dragover", event => event.preventDefault()); row.addEventListener("drop", async event => { event.preventDefault(); event.stopPropagation(); if (event.dataTransfer.getData("text/plain") !== group) return; const source = container.querySelector(".dragging"); if (source && source !== row) { container.insertBefore(source, row); if (afterDrop) await afterDrop(); } });
}

function orderedItemsForVersion(items) {
  const version = currentVersion();
  if (version.id === "general") return [...items];
  const output = [...items];
  applyCategoryOrder(output, "项目经历", version.projectOrder);
  applyCategoryOrder(output, "技能", version.skillOrder);
  return output;
}

function applyCategoryOrder(items, category, order) {
  if (!order?.length) return;
  const positions = items.map((item, index) => item.category === category ? index : -1).filter(index => index >= 0);
  const group = positions.map(index => items[index]);
  const map = new Map(group.map(item => [item.id, item]));
  const sorted = [...order.map(id => map.get(id)).filter(Boolean), ...group.filter(item => !order.includes(item.id))];
  positions.forEach((position, index) => items[position] = sorted[index]);
}

function fieldsForVersion(item) {
  const fields = [...(item.fields || [])];
  const order = currentVersion().id !== "general" && item.category === "技能" ? currentVersion().skillFieldOrders[item.id] : null;
  if (!order?.length) return fields;
  const map = new Map(fields.map(field => [field.id, field]));
  return [...order.map(id => map.get(id)).filter(Boolean), ...fields.filter(field => !order.includes(field.id))];
}

function orderedAttachmentsForVersion(items) {
  const order = currentVersion().attachmentOrder;
  if (currentVersion().id === "general" || !order?.length) return [...items];
  const map = new Map(items.map(item => [item.id, item]));
  return [...order.map(id => map.get(id)).filter(Boolean), ...items.filter(item => !order.includes(item.id))];
}

function itemHasDifference(item) {
  const version = currentVersion();
  if (version.id === "general") return false;
  if (version.itemOverrides[item.id] != null) return true;
  if ((item.fields || []).some(field => field.versioned && version.fieldOverrides[field.id] != null)) return true;
  if (item.category === "项目经历" && version.projectOrder.length) return true;
  if (item.category === "技能" && (version.skillOrder.length || version.skillFieldOrders[item.id]?.length)) return true;
  return false;
}

function isFactField() { return false; }

function isExpressionField(item, field) {
  if (isFactField(item, field)) return false;
  if (["项目经历", "技能", "自我评价", "职业规划", "常见网申问题"].includes(item.category)) return true;
  return /描述|概况|职责|成果|优势|评价|规划|动机|回答|原因|期望|能力|短版|长版/.test(field.label);
}

function isExpressionItem(item) {
  return ["项目经历", "技能", "自我评价", "职业规划", "常见网申问题"].includes(item.category);
}

function resolveFieldValue(field) { const version = currentVersion(); return version.id === "general" || !field.versioned ? field.value : version.fieldOverrides[field.id] ?? field.value; }
function hasFieldOverride(field) { const version = currentVersion(); return version.id !== "general" && field.versioned && version.fieldOverrides[field.id] != null; }
function versionName(id) { return state.settings.versions.find(v => v.id === id)?.name || "通用版"; }
function normalizeRecords(records) {
  return (records || []).map(record => {
    const status = record["状态"] === "已中止" ? "已放弃" : record["状态"] || "已投递";
    const syncStatus = record["飞书同步状态"] === "待同步" ? "未同步" : record["飞书同步状态"] || "未同步";
    const createdAt = record["创建时间"] || record["完成时间"] || new Date().toISOString();
    return {
      "记录ID": record["记录ID"] || applicationRecordId(), "公司名称": record["公司名称"] || "", "岗位名称": record["岗位名称"] || "", "工作地点": record["工作地点"] || "",
      "行业": record["行业"] || "", "岗位类别": record["岗位类别"] || "", "岗位JD": record["岗位JD"] || "", "状态": status,
      "投递日期": record["投递日期"] || "", "完成时间": record["完成时间"] || "", "岗位链接": record["岗位链接"] || "", "页面标题": record["页面标题"] || "", "使用简历": record["使用简历"] || "",
      "填写版本": record["填写版本"] || "通用版", "填写版本ID": record["填写版本ID"] || "", "中止原因": record["中止原因"] || "", "备注": record["备注"] || "",
      "内推码": record["内推码"] || "", "内推人": record["内推人"] || "", "实际投递渠道": record["实际投递渠道"] || "", "招聘类型/批次": record["招聘类型/批次"] || "", "企业性质": record["企业性质"] || "", "投递优先级": record["投递优先级"] || "", "有效性修正": record["有效性修正"] || "",
      "创建时间": createdAt, "最近更新时间": record["最近更新时间"] || createdAt, "飞书同步状态": syncStatus, "飞书记录ID": record["飞书记录ID"] || "", enrichment: record.enrichment || {}
    };
  });
}
function normalizeProfile(profile) { const next = structuredClone(profile); delete next.fillingVersions; next.version = Math.max(Number(next.version) || 1, 6); next.categories = [...new Set(next.categories || [])]; next.items = (next.items || []).map(normalizeItem); return next; }

function normalizeItem(item) { const next = { ...item, id: item.id || uid("item"), category: item.category === "基本信息" ? "基础信息" : item.category || "基础信息", title: item.title || "未命名", tags: item.tags || [] }; const sourceFields = item.fields?.length ? item.fields : parseStructuredFields(item.content); if (sourceFields.length) { next.fields = sourceFields.map(field => ({ ...field, id: field.id || uid("field"), versioned: field.versioned ?? isExpressionField(next, field) })); next.content = fieldsToContent(next.fields); } return next; }
function normalizeAttachments(items) { return (items || []).map(item => {
  const id = item.id || uid("attachment"), addedAt = item.addedAt || new Date().toISOString();
  const boundVersions = state.settings.versions.filter(version => version.id !== "general" && (version.defaultResumeId === id || (version.defaultAttachmentIds || []).includes(id))).map(version => version.id);
  const tagText = `${item.title || ""} ${(item.tags || []).join(" ")}`;
  const inferredVersions = boundVersions;
  const next = { ...item, id, title: item.title || item.displayName || item.fileName, displayName: item.displayName || item.title || item.fileName, addedAt, updatedAt: item.updatedAt || addedAt, applicableVersionIds: Array.isArray(item.applicableVersionIds) ? [...item.applicableVersionIds] : inferredVersions, localPath: item.localPath || null };
  delete next.version; return next;
}); }

async function hydrateBundledAttachments(items) {
  return Promise.all(items.map(async item => {
    if (item.blob instanceof Blob || !item.path) return item;
    try { const response = await fetch(chrome.runtime.getURL(item.path)); return response.ok ? { ...item, blob: await response.blob(), source: "indexeddb" } : item; }
    catch { return item; }
  }));
}

async function readInternalAttachmentBlob(item) {
  if (item.blob instanceof Blob) return item.blob;
  if (item.path) { const response = await fetch(chrome.runtime.getURL(item.path)); if (response.ok) return response.blob(); }
  throw new Error("附件内容不在内部存储中，请重新上传或迁移旧工作区数据");
}
function normalizeSettings(settings, profile) {
  const source = settings || DEFAULT_SETTINGS;
  const next = { ...structuredClone(DEFAULT_SETTINGS), ...source };
  next.versions = (source.versions?.length ? source.versions : DEFAULT_SETTINGS.versions).map(version => ({ ...createVersion(version.id, version.name), ...version, fieldOverrides: { ...(version.fieldOverrides || {}) }, itemOverrides: { ...(version.itemOverrides || {}) }, skillFieldOrders: { ...(version.skillFieldOrders || {}) }, defaultAttachmentIds: [...(version.defaultAttachmentIds || [])], projectOrder: [...(version.projectOrder || [])], skillOrder: [...(version.skillOrder || [])], attachmentOrder: [...(version.attachmentOrder || [])] }));
  for (const item of profile?.items || []) for (const field of item.fields || []) {
    for (const legacy of field.variants || []) {
      const version = next.versions.find(candidate => candidate.id === legacy.version);
      if (version && legacy.value) version.fieldOverrides[field.id] = legacy.value;
    }
    delete field.variants;
  }
  const general = next.versions.find(version => version.id === "general") || createVersion("general", "通用版");
  general.fieldOverrides = {}; general.itemOverrides = {}; general.projectOrder = []; general.skillOrder = []; general.skillFieldOrders = {}; general.attachmentOrder = []; general.defaultResumeId ||= "resume-general"; general.defaultAttachmentIds ||= [];
  if (!next.versions.some(v => v.id === "general")) next.versions.unshift(general);
  if (!next.versions.some(v => v.id === next.currentVersion)) next.currentVersion = "general";
  return next;
}

function parseFields(content) { return content.split("\n").map(line => { const match = line.match(/^([^：:]+)[：:]\s*(.*)$/); return match && match[2].trim() ? { id: uid("field"), label: match[1].trim(), value: match[2].trim() } : null; }).filter(Boolean); }
function parseStructuredFields(content = "") {
  const lines = String(content).split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  if (lines.length < 2) return [];
  const fields = lines.map(line => {
    const match = /^([^：:]{1,50})[：:]\s*(.+)$/.exec(line.replace(/^[-*•]\s+/, ""));
    if (!match) return null;
    const label = match[1].trim(), value = match[2].trim();
    if (!label || !value || /^https?$/i.test(label) || /[。！？!?]$/.test(label)) return null;
    return { id: uid("field"), label, value };
  });
  return fields.every(Boolean) ? fields : [];
}
function mergeFieldMetadata(fields, oldFields = []) { const map = new Map(oldFields.map(f => [f.label, f])); return fields.map(field => map.has(field.label) ? { ...field, id: map.get(field.label).id, versioned: map.get(field.label).versioned } : field); }
function fieldsToContent(fields) { return fields.map(field => `${field.label}：${field.value}`).join("\n"); }
function splitTags(value) { return value.split(/[,\s，、]+/).map(v => v.trim()).filter(Boolean); }
function moveByValue(array, source, target) { const from = array.indexOf(source), to = array.indexOf(target); if (from < 0 || to < 0) return; array.splice(to, 0, array.splice(from, 1)[0]); }
function moveObjects(array, source, target) { const from = array.findIndex(v => v.id === source), to = array.findIndex(v => v.id === target); if (from < 0 || to < 0) return; array.splice(to, 0, array.splice(from, 1)[0]); }

function resolveItemContent(item) { const version = currentVersion(); return version.id === "general" ? item.content : version.itemOverrides[item.id] ?? item.content; }
async function copyItem(item) { const scrollTop = els.contentPane?.scrollTop; await ensureActiveRecord(); const text = item.fields?.length ? fieldsForVersion(item).map(field => `${field.label}：${resolveFieldValue(field)}`).join("\n") : resolveItemContent(item); await navigator.clipboard.writeText(text); showNotice(`已复制：${item.title}`); restoreContentScroll(scrollTop); }
async function copyText(text, label, feedback) { const scrollTop = els.contentPane?.scrollTop; await ensureActiveRecord(); await navigator.clipboard.writeText(text); feedback?.classList.add("copied"); setTimeout(() => feedback?.classList.remove("copied"), 650); showNotice(`已复制：${label}`); restoreContentScroll(scrollTop); }
function restoreContentScroll(scrollTop) { if (els.contentPane && Number.isFinite(scrollTop)) els.contentPane.scrollTop = scrollTop; }
function showNotice(message) { els.notice.textContent = message; els.notice.hidden = false; clearTimeout(showNotice.timer); showNotice.timer = setTimeout(() => els.notice.hidden = true, 3500); }
function button(text, className, listener) { const el = document.createElement("button"); el.type = "button"; el.textContent = text; if (className) el.className = className; el.addEventListener("click", listener); return el; }
function option(value, text) { const el = document.createElement("option"); el.value = value; el.textContent = text; return el; }
function dragHandle() { const el = document.createElement("span"); el.className = "dragHandle"; el.textContent = "⋮⋮"; el.title = "拖动排序"; return el; }
function uid(prefix) { return `${prefix}-${Date.now().toString(36)}-${crypto.randomUUID().slice(0, 8)}`; }
function formatBytes(size = 0) { return size < 1048576 ? `${Math.round(size / 1024)} KB` : `${(size / 1048576).toFixed(1)} MB`; }
function isSensitiveTag(tag) { return /身份证|手机|隐私|确认|薪资|报到|状态|保密/.test(tag); }
function isPendingValue(value = "") { return /【?待补充|待确认|提交前确认/.test(value); }
function localDate(date) { return new Intl.DateTimeFormat("zh-CN", { year: "numeric", month: "2-digit", day: "2-digit" }).format(date).replaceAll("/", "-"); }
function localDateTime(date) { return `${localDate(date)} ${date.toTimeString().slice(0, 8)}`; }
function blobToDataUrl(blob) { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = () => reject(reader.error); reader.readAsDataURL(blob); }); }
async function dataUrlToBlob(dataUrl) { return (await fetch(dataUrl)).blob(); }
async function pickFile(types) { const [handle] = await window.showOpenFilePicker({ multiple: false, types }); return handle.getFile(); }
function downloadJson(value, fileName) { const blob = new Blob([JSON.stringify(value, null, 2)], { type: "application/json" }); const url = URL.createObjectURL(blob); chrome.downloads.download({ url, filename: fileName, saveAs: true }); setTimeout(() => URL.revokeObjectURL(url), 30000); }
function uniqueAttachmentName(name) { const dot = name.lastIndexOf("."), base = dot > 0 ? name.slice(0, dot) : name, extension = dot > 0 ? name.slice(dot) : ""; return `${base}_${Date.now()}${extension}`; }
async function sha256(file) { const hash = await crypto.subtle.digest("SHA-256", await file.arrayBuffer()); return [...new Uint8Array(hash)].map(v => v.toString(16).padStart(2, "0")).join(""); }
function newItemDefaults(category) { if (category === "项目经历") return { title: "新项目经历", content: ["项目名称：【待补充】","项目角色：【待补充】","开始时间：【待补充】","结束时间：【待补充】","项目背景：【待补充】","主要职责：【待补充】","技术方法：【待补充】","项目成果：【待补充】","网申短版：【待补充】"].join("\n") }; if (category === "论文专利") return { title: "新论文/专利成果", content: ["成果类型：【待补充】","成果名称：【待补充】","期刊/专利号：【待补充】","发表/申请日期：【待补充】","成果状态：【待补充】","本人排序：【待补充】","全部作者/发明人：【待补充】","DOI/公告号：【待补充】"].join("\n") }; return { title: "新信息", content: "字段名：【待补充】" }; }

async function installUploadPicker(attachment) {
  document.getElementById("job-paste-upload-layer")?.remove(); const inputs = [...document.querySelectorAll('input[type="file"]')].filter(input => !input.disabled);
  if (!inputs.length) return window.alert("ApplyPaste：当前页面未找到文件上传框。请展开对应附件区域后重试。");
  const layer = document.createElement("div"); layer.id = "job-paste-upload-layer"; Object.assign(layer.style, { position:"fixed", inset:"0", zIndex:"2147483647", pointerEvents:"none" });
  const banner = document.createElement("div"); banner.textContent = `ApplyPaste：请选择“${attachment.title}”要放入的上传框（Esc 取消）`; Object.assign(banner.style, { position:"fixed", top:"12px", left:"50%", transform:"translateX(-50%)", padding:"10px 14px", borderRadius:"6px", color:"#fff", background:"#155eef", boxShadow:"0 8px 24px rgba(0,0,0,.22)", font:"14px system-ui", pointerEvents:"auto" }); layer.append(banner);
  const cleanup = () => { layer.remove(); document.removeEventListener("keydown", onKeydown, true); }; const onKeydown = event => { if (event.key === "Escape") cleanup(); }; document.addEventListener("keydown", onKeydown, true);
  inputs.forEach((input, index) => { const rect = input.getBoundingClientRect(), anchor = input.closest("label, div, li, section") || input, anchorRect = anchor.getBoundingClientRect(); const badge = document.createElement("button"); badge.type = "button"; badge.textContent = `上传到这里 ${index + 1}`; Object.assign(badge.style, { position:"fixed", left:`${Math.max(8,Math.min(window.innerWidth-130,rect.width?rect.left:anchorRect.left))}px`, top:`${Math.max(52,Math.min(window.innerHeight-42,rect.height?rect.top:anchorRect.top))}px`, zIndex:"2147483647", pointerEvents:"auto", padding:"7px 10px", border:"2px solid #fff", borderRadius:"6px", color:"#fff", background:"#155eef", boxShadow:"0 4px 14px rgba(0,0,0,.28)", font:"600 13px system-ui", cursor:"pointer" }); badge.onclick = async () => { try { const blob = await (await fetch(attachment.dataUrl)).blob(), file = new File([blob], attachment.fileName, { type: attachment.mimeType || blob.type }), transfer = new DataTransfer(); transfer.items.add(file); input.files = transfer.files; input.dispatchEvent(new Event("input", { bubbles:true })); input.dispatchEvent(new Event("change", { bubbles:true })); cleanup(); } catch (error) { window.alert(`ApplyPaste：${error.message}`); } }; layer.append(badge); }); document.documentElement.append(layer);
}
