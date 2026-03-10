const storageKey = "storyboard-app-v1";
const aiConfigKey = "storyboard-ai-config-v1";
const projectStoreKey = "storyboard-project-store-v1";
const exportPrefsKey = "storyboard-export-prefs-v1";
const imageDbName = "storyboard-image-db";
const imageStoreName = "images";
const exportImageMaxEdge = 1920;
const exportJpegQuality = 0.82;
const placeholderImage =
  "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 640 360'%3E%3Crect width='640' height='360' fill='%23ddd5c8'/%3E%3Ctext x='320' y='180' text-anchor='middle' fill='%23766f63' font-size='28' font-family='sans-serif'%3E%E6%97%A0%E5%9B%BE%E7%89%87%3C/text%3E%3C/svg%3E";

const state = {
  cards: [],
  draggingId: null,
  objectUrls: new Map(),
  reportObjectUrls: [],
  printObjectUrls: [],
  imageDbDisabled: false,
  editorVisible: true,
  exportCleanMode: true,
  exportShotsPerPageDesktop: 4,
  exportShotsPerPageMobile: 4,
  currentProjectId: "",
  projectStore: {
    order: [],
    projects: {},
    currentProjectId: "",
  },
  aiConfig: {
    provider: "openai",
    baseUrl: "https://api.openai.com/v1",
    path: "/images/generations",
    model: "gpt-image-1",
    apiKey: "",
    size: "1024x1024",
    customWidth: "",
    customHeight: "",
    metaShotNo: "",
    metaShotType: "",
    metaScene: "",
  },
};

const elements = {
  board: document.getElementById("board"),
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("file-input"),
  pickBtn: document.getElementById("pick-btn"),
  projectSelect: document.getElementById("project-select"),
  newProjectBtn: document.getElementById("new-project-btn"),
  renameProjectBtn: document.getElementById("rename-project-btn"),
  deleteProjectBtn: document.getElementById("delete-project-btn"),
  exportProjectBtn: document.getElementById("export-project-btn"),
  importProjectBtn: document.getElementById("import-project-btn"),
  importProjectInput: document.getElementById("import-project-input"),
  toggleEditorBtn: document.getElementById("toggle-editor-btn"),
  openAiBtn: document.getElementById("open-ai-btn"),
  addEmptyBtn: document.getElementById("add-empty-btn"),
  exportBtn: document.getElementById("export-btn"),
  buildReportBtn: document.getElementById("build-report-btn"),
  pdfBtn: document.getElementById("pdf-btn"),
  pngBtn: document.getElementById("png-btn"),
  exportCleanToggle: document.getElementById("export-clean-toggle"),
  exportShotsDesktopInput: document.getElementById("export-shots-desktop"),
  exportShotsMobileInput: document.getElementById("export-shots-mobile"),
  clearBtn: document.getElementById("clear-btn"),
  status: document.getElementById("status"),
  report: document.getElementById("report"),
  reportMeta: document.getElementById("report-meta"),
  reportBody: document.getElementById("report-body"),
  reportPrint: document.getElementById("report-print"),
  aiModal: document.getElementById("ai-modal"),
  aiProvider: document.getElementById("ai-provider"),
  aiBaseUrl: document.getElementById("ai-base-url"),
  aiPath: document.getElementById("ai-path"),
  aiApiKey: document.getElementById("ai-api-key"),
  aiModel: document.getElementById("ai-model"),
  aiSize: document.getElementById("ai-size"),
  aiCustomSizeWrap: document.getElementById("ai-custom-size-wrap"),
  aiCustomWidth: document.getElementById("ai-custom-width"),
  aiCustomHeight: document.getElementById("ai-custom-height"),
  aiMetaShotNo: document.getElementById("ai-meta-shot-no"),
  aiMetaShotType: document.getElementById("ai-meta-shot-type"),
  aiMetaScene: document.getElementById("ai-meta-scene"),
  aiPrompt: document.getElementById("ai-prompt"),
  aiPromptPreview: document.getElementById("ai-prompt-preview"),
  aiGenerateBtn: document.getElementById("ai-generate-btn"),
  aiStatus: document.getElementById("ai-status"),
  template: document.getElementById("card-template"),
};

let imageDbPromise = null;

initialize();

async function initialize() {
  loadProjectStore();
  loadExportPrefs();
  setupMobileMinimalMode();
  bindGlobalEvents();
  render();
  registerServiceWorker();
  await ensureImageDb();
  await migrateLegacyInlineImages();
  if (!state.editorVisible && state.cards.length) {
    await buildReportTable();
  }
}

function bindGlobalEvents() {
  elements.projectSelect.addEventListener("change", () => {
    switchProject(elements.projectSelect.value);
  });
  elements.newProjectBtn.addEventListener("click", () => {
    createNewProject();
  });
  elements.renameProjectBtn.addEventListener("click", () => {
    renameCurrentProject();
  });
  elements.deleteProjectBtn.addEventListener("click", () => {
    deleteCurrentProject();
  });
  elements.exportProjectBtn.addEventListener("click", async () => {
    await exportCurrentProjectPackage();
  });
  elements.importProjectBtn.addEventListener("click", () => {
    elements.importProjectInput.click();
  });
  elements.importProjectInput.addEventListener("change", async event => {
    const file = event.target.files?.[0];
    if (!file) return;
    await importProjectPackage(file);
    event.target.value = "";
  });

  elements.toggleEditorBtn.addEventListener("click", () => {
    setEditorVisibility(!state.editorVisible);
  });

  elements.openAiBtn.addEventListener("click", () => {
    openAiModal();
  });

  elements.pickBtn.addEventListener("click", () => elements.fileInput.click());
  elements.fileInput.addEventListener("change", async event => {
    await addFiles(Array.from(event.target.files || []));
    event.target.value = "";
  });

  elements.addEmptyBtn.addEventListener("click", () => {
    state.cards.push(createCard());
    persistAndRender();
  });

  elements.exportBtn.addEventListener("click", exportCsv);
  elements.buildReportBtn.addEventListener("click", () => {
    void buildReportTable();
  });
  elements.pdfBtn.addEventListener("click", handlePdfExport);
  elements.pngBtn.addEventListener("click", handlePngExport);
  elements.exportCleanToggle.addEventListener("change", () => {
    state.exportCleanMode = !!elements.exportCleanToggle.checked;
    persistExportPrefs();
  });
  elements.exportShotsDesktopInput.addEventListener("change", () => {
    state.exportShotsPerPageDesktop = normalizeShotsPerPage(elements.exportShotsDesktopInput.value, 4);
    elements.exportShotsDesktopInput.value = String(state.exportShotsPerPageDesktop);
    persistExportPrefs();
  });
  elements.exportShotsMobileInput.addEventListener("change", () => {
    state.exportShotsPerPageMobile = normalizeShotsPerPage(elements.exportShotsMobileInput.value, 4);
    elements.exportShotsMobileInput.value = String(state.exportShotsPerPageMobile);
    persistExportPrefs();
  });
  elements.clearBtn.addEventListener("click", async () => {
    if (!confirm("确认清空全部镜头卡吗？")) return;
    state.cards = [];
    await clearAllImages();
    persistAndRender();
    clearReport();
  });

  elements.aiProvider.addEventListener("change", () => {
    applyProviderPreset(elements.aiProvider.value);
    writeAiConfigFromInputs();
  });
  [elements.aiBaseUrl, elements.aiPath, elements.aiApiKey, elements.aiModel].forEach(input =>
    input.addEventListener("input", writeAiConfigFromInputs)
  );
  elements.aiSize.addEventListener("change", () => {
    syncAiSizeVisibility();
    writeAiConfigFromInputs();
  });
  [elements.aiCustomWidth, elements.aiCustomHeight].forEach(input =>
    input.addEventListener("input", writeAiConfigFromInputs)
  );
  [elements.aiMetaShotNo, elements.aiMetaShotType, elements.aiMetaScene, elements.aiPrompt].forEach(
    input =>
      input.addEventListener("input", () => {
        writeAiConfigFromInputs();
        refreshAiPromptPreview();
      })
  );
  elements.aiGenerateBtn.addEventListener("click", async () => {
    await generateAiImageAndAddCard();
  });

  ["dragenter", "dragover"].forEach(eventName => {
    elements.dropzone.addEventListener(eventName, event => {
      event.preventDefault();
      elements.dropzone.classList.add("active");
    });
  });

  ["dragleave", "drop"].forEach(eventName => {
    elements.dropzone.addEventListener(eventName, event => {
      event.preventDefault();
      elements.dropzone.classList.remove("active");
    });
  });

  elements.dropzone.addEventListener("drop", async event => {
    await addFiles(getImageFilesFromDrop(event));
  });

  elements.board.addEventListener("dragover", event => {
    if (getImageFilesFromDrop(event).length) event.preventDefault();
  });

  elements.board.addEventListener("drop", async event => {
    const files = getImageFilesFromDrop(event);
    if (!files.length) return;
    event.preventDefault();
    await addFiles(files);
  });

  window.addEventListener("afterprint", () => {
    clearPrintPages();
    document.body.classList.remove("export-clean-mode");
  });
}

async function handlePdfExport() {
  const shotsPerPage = getShotsPerPageForCurrentDevice();
  await withExportCleanView(async () => {
    await buildReportTable();
    await buildPrintPagesForExport(shotsPerPage);
    await waitForImagesInContainer(elements.reportPrint);
    window.print();
  });
}

async function handlePngExport() {
  const shotsPerPage = getShotsPerPageForCurrentDevice();
  await withExportCleanView(async () => {
    await exportReportPng(shotsPerPage);
  });
}

async function withExportCleanView(task) {
  const shouldApply = state.exportCleanMode;
  if (shouldApply) document.body.classList.add("export-clean-mode");
  try {
    await task();
  } finally {
    if (shouldApply) {
      // For PDF this gets removed in afterprint; keep fallback for cancelled print.
      setTimeout(() => document.body.classList.remove("export-clean-mode"), 500);
    }
  }
}

function loadAiConfig() {
  // Deprecated: kept for backward compatibility. Project store now owns AI config.
}

function persistAiConfig() {
  persistProjectState();
}

function openAiModal() {
  enforceFixedProviderDefaults();
  syncAiInputsFromState();
  syncAiSizeVisibility();
  updateAiFieldLocks();
  refreshAiPromptPreview();
  setAiStatus("");
  elements.aiModal.showModal();
}

function syncAiInputsFromState() {
  elements.aiProvider.value = state.aiConfig.provider || "openai";
  elements.aiBaseUrl.value = state.aiConfig.baseUrl || "";
  elements.aiPath.value = state.aiConfig.path || "";
  elements.aiApiKey.value = state.aiConfig.apiKey || "";
  elements.aiModel.value = state.aiConfig.model || "";
  elements.aiSize.value = state.aiConfig.size || "1024x1024";
  elements.aiCustomWidth.value = state.aiConfig.customWidth || "";
  elements.aiCustomHeight.value = state.aiConfig.customHeight || "";
  elements.aiMetaShotNo.value = state.aiConfig.metaShotNo || "";
  elements.aiMetaShotType.value = state.aiConfig.metaShotType || "";
  elements.aiMetaScene.value = state.aiConfig.metaScene || "";
}

function writeAiConfigFromInputs() {
  state.aiConfig = {
    provider: elements.aiProvider.value,
    baseUrl: elements.aiBaseUrl.value.trim(),
    path: elements.aiPath.value.trim(),
    apiKey: elements.aiApiKey.value.trim(),
    model: elements.aiModel.value.trim(),
    size: elements.aiSize.value,
    customWidth: elements.aiCustomWidth.value.trim(),
    customHeight: elements.aiCustomHeight.value.trim(),
    metaShotNo: elements.aiMetaShotNo.value.trim(),
    metaShotType: elements.aiMetaShotType.value.trim(),
    metaScene: elements.aiMetaScene.value.trim(),
  };
  persistProjectState();
}

function syncAiSizeVisibility() {
  const isCustom = elements.aiSize.value === "custom";
  elements.aiCustomSizeWrap.classList.toggle("hidden", !isCustom);
}

function applyProviderPreset(provider) {
  const presets = {
    openai: {
      baseUrl: "https://api.openai.com/v1",
      path: "/images/generations",
      model: "gpt-image-1",
      size: "1024x1024",
    },
    doubao: {
      baseUrl: "https://ark.cn-beijing.volces.com/api/v3",
      path: "/images/generations",
      model: "doubao-seedream-3-0-t2i",
      size: "1024x1024",
    },
    doubao_seedream5_fixed: {
      baseUrl: "https://ark.cn-beijing.volces.com/api/v3",
      path: "/images/generations",
      model: "doubao-seedream-5-0-260128",
      size: "2048x2048",
    },
    custom: {},
  };

  const preset = presets[provider] || {};
  if (preset.baseUrl) elements.aiBaseUrl.value = preset.baseUrl;
  if (preset.path) elements.aiPath.value = preset.path;
  if (preset.model) elements.aiModel.value = preset.model;
  if (preset.size) elements.aiSize.value = preset.size;
  syncAiSizeVisibility();
  updateAiFieldLocks();
  refreshAiPromptPreview();
}

function setAiStatus(message) {
  elements.aiStatus.textContent = message || "";
}

function enforceFixedProviderDefaults() {
  if (state.aiConfig.provider !== "doubao_seedream5_fixed") return;
  state.aiConfig.baseUrl = "https://ark.cn-beijing.volces.com/api/v3";
  state.aiConfig.path = "/images/generations";
  state.aiConfig.model = "doubao-seedream-5-0-260128";
  state.aiConfig.size = "2048x2048";
  state.aiConfig.customWidth = "";
  state.aiConfig.customHeight = "";
  persistProjectState();
}

function updateAiFieldLocks() {
  const isFixed = elements.aiProvider.value === "doubao_seedream5_fixed";
  elements.aiBaseUrl.readOnly = isFixed;
  elements.aiPath.readOnly = isFixed;
  elements.aiModel.readOnly = isFixed;
  elements.aiSize.disabled = isFixed;
  elements.aiCustomWidth.readOnly = isFixed;
  elements.aiCustomHeight.readOnly = isFixed;
}

function composeAiPrompt() {
  const parts = [];
  if (state.aiConfig.metaShotNo) parts.push(`镜号:${state.aiConfig.metaShotNo}`);
  if (state.aiConfig.metaShotType) parts.push(`景别:${state.aiConfig.metaShotType}`);
  if (state.aiConfig.metaScene) parts.push(`场景:${state.aiConfig.metaScene}`);
  const freePrompt = elements.aiPrompt.value.trim();
  if (freePrompt) parts.push(`画面描述:${freePrompt}`);
  return parts.join("，");
}

function refreshAiPromptPreview() {
  elements.aiPromptPreview.value = composeAiPrompt();
}

function setupMobileMinimalMode() {
  const isMobile = window.matchMedia("(max-width: 720px)").matches;
  if (!isMobile) {
    setEditorVisibility(true);
    return;
  }
  document.body.classList.add("mobile-minimal");
  setEditorVisibility(state.cards.length === 0);
}

function setEditorVisibility(visible) {
  state.editorVisible = visible;
  const buttonLabel = visible ? "隐藏编辑区" : "显示编辑区";
  if (elements.toggleEditorBtn) elements.toggleEditorBtn.textContent = buttonLabel;

  const method = visible ? "remove" : "add";
  elements.dropzone.classList[method]("hidden");
  elements.board.classList[method]("hidden");
  const legend = document.querySelector(".legend");
  if (legend) legend.classList[method]("hidden");
}

function registerServiceWorker() {
  const isHttp = window.location.protocol === "http:" || window.location.protocol === "https:";
  if (!isHttp || !("serviceWorker" in navigator)) return;
  navigator.serviceWorker.register("./sw.js").catch(() => {});
}

async function addFiles(files) {
  const imageFiles = files.filter(isImageFile);
  if (!imageFiles.length) {
    alert("未识别到可导入的图片文件。");
    return;
  }

  let successCount = 0;
  let failCount = 0;

  for (const file of imageFiles) {
    const card = createCard({ imageName: file.name });
    const attached = await attachImageToCard(card, file);
    if (!attached) {
      failCount += 1;
      continue;
    }
    state.cards.push(card);
    successCount += 1;
  }

  if (successCount) persistAndRender();
  if (failCount) alert(`有 ${failCount} 张图片读取失败，请尝试转为 PNG/JPG 后重试。`);
}

async function generateAiImageAndAddCard() {
  const prompt = composeAiPrompt();
  if (!prompt.trim()) {
    setAiStatus("请输入提示词。");
    return;
  }

  writeAiConfigFromInputs();
  const { baseUrl, path, apiKey, model } = state.aiConfig;
  const sizeResult = resolveAiSize();
  if (!sizeResult.ok) {
    setAiStatus(sizeResult.message || "尺寸参数无效。");
    return;
  }
  const size = sizeResult.size;
  if (!baseUrl || !path || !apiKey || !model) {
    setAiStatus("请填写 Base URL / 接口路径 / API Key / 模型。");
    return;
  }

  setAiStatus("正在生成图片...");

  try {
    const endpoint = joinEndpoint(baseUrl, path);
    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        prompt,
        size,
        n: 1,
        response_format: "b64_json",
      }),
    });

    if (!response.ok) {
      const errorText = await safeReadText(response);
      setAiStatus(`生成失败：${response.status} ${shortenText(errorText, 120)}`);
      return;
    }

    const payload = await response.json();
    const dataItem = Array.isArray(payload?.data) ? payload.data[0] : null;
    if (!dataItem) {
      setAiStatus("生成失败：返回数据为空。");
      return;
    }

    let file = null;
    if (typeof dataItem.b64_json === "string" && dataItem.b64_json) {
      const blob = base64ToImageBlob(dataItem.b64_json);
      file = new File([blob], `ai-${Date.now()}.png`, { type: "image/png" });
    } else if (typeof dataItem.url === "string" && dataItem.url) {
      const fetched = await fetch(dataItem.url);
      if (!fetched.ok) {
        setAiStatus("生成成功，但图片下载失败，请检查图片地址权限。");
        return;
      }
      const blob = await fetched.blob();
      const ext = blob.type.includes("jpeg") ? "jpg" : "png";
      file = new File([blob], `ai-${Date.now()}.${ext}`, { type: blob.type || "image/png" });
    } else {
      setAiStatus("生成失败：未识别到图片字段（b64_json/url）。");
      return;
    }

    const card = createCard({
      shotNo: state.aiConfig.metaShotNo || nextShotNo(),
      shotType: state.aiConfig.metaShotType || "",
      scene: state.aiConfig.metaScene || "AI 生成",
      notes: elements.aiPrompt.value.trim() || prompt,
      imageName: file.name,
      time: currentTimeHHMM(),
    });
    const attached = await attachImageToCard(card, file);
    if (!attached) {
      setAiStatus("图片生成成功，但保存失败。");
      return;
    }
    state.cards.push(card);
    persistAndRender();
    setAiStatus("已生成并加入分镜。");
    setStatus("AI 图片已加入分镜卡。");
  } catch (error) {
    setAiStatus(`生成失败：${shortenText(String(error), 120)}`);
  }
}

function resolveAiSize() {
  const provider = state.aiConfig.provider;
  const model = (state.aiConfig.model || "").toLowerCase();
  const allowedSizes = getAllowedSizes(provider, model);
  const sizeMode = state.aiConfig.size;
  if (!sizeMode || sizeMode === "auto") {
    const adaptive = getAdaptiveSize();
    if (allowedSizes.includes(adaptive)) return { ok: true, size: adaptive };
    return { ok: true, size: allowedSizes[0] || "1024x1024" };
  }
  if (sizeMode === "custom") {
    const width = Number(state.aiConfig.customWidth);
    const height = Number(state.aiConfig.customHeight);
    if (!Number.isFinite(width) || !Number.isFinite(height)) {
      return { ok: false, message: "请填写有效的自定义宽高（整数）。" };
    }
    if (width < 256 || height < 256) {
      return { ok: false, message: "自定义尺寸最小为 256x256。" };
    }
    const custom = `${Math.round(width)}x${Math.round(height)}`;
    if (!allowedSizes.includes(custom)) {
      return {
        ok: false,
        message: `当前模型不支持该尺寸。可用：${allowedSizes.join(" / ")}`,
      };
    }
    return { ok: true, size: custom };
  }
  if (!allowedSizes.includes(sizeMode)) {
    return {
      ok: false,
      message: `当前模型不支持该尺寸。可用：${allowedSizes.join(" / ")}`,
    };
  }
  return { ok: true, size: sizeMode };
}

function getAdaptiveSize() {
  const prompt = String(composeAiPrompt() || "").toLowerCase();
  if (/(竖屏|竖构图|portrait|vertical)/.test(prompt)) return "1024x1536";
  if (/(横屏|横构图|landscape|horizontal|16:9)/.test(prompt)) return "1536x1024";
  return window.innerWidth >= window.innerHeight ? "1536x1024" : "1024x1536";
}

function getAllowedSizes(provider, model) {
  // Most image APIs only accept a fixed allow-list, not arbitrary width/height.
  if (isDoubaoProvider(provider) && model.includes("seedream-5-0-260128")) {
    return ["2048x2048", "1024x1024", "2K"];
  }

  if (isDoubaoProvider(provider) && model.includes("seedream-5")) {
    return ["2048x2048", "1024x1024", "2K"];
  }

  const providerSizes = {
    openai: ["1024x1024", "1536x1024", "1024x1536"],
    doubao: ["1024x1024", "2048x2048", "2K"],
    doubao_seedream5_fixed: ["2048x2048", "1024x1024", "2K"],
    custom: ["1024x1024", "1536x1024", "1024x1536", "2048x2048", "2K"],
  };
  return providerSizes[provider] || providerSizes.custom;
}

function isDoubaoProvider(provider) {
  return provider === "doubao" || provider === "doubao_seedream5_fixed";
}

function joinEndpoint(baseUrl, path) {
  const left = String(baseUrl || "").replace(/\/+$/, "");
  const right = String(path || "").startsWith("/") ? String(path) : `/${String(path || "")}`;
  return `${left}${right}`;
}

function base64ToImageBlob(base64Text) {
  const binary = atob(base64Text);
  const length = binary.length;
  const bytes = new Uint8Array(length);
  for (let index = 0; index < length; index += 1) bytes[index] = binary.charCodeAt(index);
  return new Blob([bytes], { type: "image/png" });
}

function currentTimeHHMM() {
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  return `${hours}:${minutes}`;
}

async function safeReadText(response) {
  try {
    return await response.text();
  } catch {
    return "";
  }
}

function shortenText(value, maxLength) {
  const text = String(value || "").trim();
  if (text.length <= maxLength) return text;
  return `${text.slice(0, maxLength)}...`;
}

function createCard(initial = {}) {
  return {
    id: initial.id || crypto.randomUUID(),
    shotNo: initial.shotNo || nextShotNo(),
    shotType: initial.shotType || "",
    scene: initial.scene || "",
    people: initial.people || "",
    time: initial.time || "",
    notes: initial.notes || "",
    image: initial.image || "",
    imageName: initial.imageName || "",
  };
}

function nextShotNo() {
  return String(state.cards.length + 1).padStart(3, "0");
}

function fileToDataUrl(file) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = () => resolve({ dataUrl: String(reader.result), name: file.name });
    reader.onerror = () => resolve(null);
    reader.onabort = () => resolve(null);
    reader.readAsDataURL(file);
  });
}

function render() {
  revokeAllObjectUrls();
  elements.board.innerHTML = "";
  state.cards.forEach(card => elements.board.appendChild(renderCard(card)));
}

function renderCard(card) {
  const fragment = elements.template.content.cloneNode(true);
  const node = fragment.querySelector(".card");
  node.dataset.id = card.id;

  const imageEl = node.querySelector(".thumb");
  imageEl.src = card.image || placeholderImage;
  hydrateCardImage(card, node, imageEl);

  node.querySelectorAll("[data-field]").forEach(input => {
    const field = input.dataset.field;
    input.value = card[field];
    input.addEventListener("input", event => {
      updateCard(card.id, field, event.target.value);
    });
  });

  node.querySelector(".replace-input").addEventListener("change", async event => {
    const file = event.target.files?.[0];
    if (!file) return;
    const targetCard = state.cards.find(item => item.id === card.id);
    if (!targetCard) return;
    const attached = await attachImageToCard(targetCard, file);
    if (attached) persistAndRender();
    event.target.value = "";
  });

  node.querySelectorAll("[data-action]").forEach(button => {
    button.addEventListener("click", () => {
      void handleCardAction(card.id, button.dataset.action);
    });
  });

  node.addEventListener("dragstart", () => {
    state.draggingId = card.id;
    node.classList.add("dragging");
  });

  node.addEventListener("dragend", () => {
    state.draggingId = null;
    node.classList.remove("dragging");
  });

  node.addEventListener("dragover", event => {
    event.preventDefault();
  });

  node.addEventListener("drop", async event => {
    const files = getImageFilesFromDrop(event);
    if (files.length) {
      event.preventDefault();
      await addFiles(files);
      return;
    }

    event.preventDefault();
    if (!state.draggingId || state.draggingId === card.id) return;

    const fromIndex = state.cards.findIndex(item => item.id === state.draggingId);
    const toIndex = state.cards.findIndex(item => item.id === card.id);
    if (fromIndex === -1 || toIndex === -1) return;

    const [moved] = state.cards.splice(fromIndex, 1);
    state.cards.splice(toIndex, 0, moved);
    persistAndRender();
  });

  return node;
}

async function hydrateCardImage(card, node, imageEl) {
  if (card.image) return;
  const record = await getImageByCardId(card.id);
  if (!record || !node.isConnected) return;
  const previous = state.objectUrls.get(card.id);
  if (previous) URL.revokeObjectURL(previous);
  const objectUrl = URL.createObjectURL(record.blob);
  state.objectUrls.set(card.id, objectUrl);
  imageEl.src = objectUrl;
}

function revokeAllObjectUrls() {
  for (const objectUrl of state.objectUrls.values()) URL.revokeObjectURL(objectUrl);
  state.objectUrls.clear();
}

function revokeReportObjectUrls() {
  state.reportObjectUrls.forEach(url => URL.revokeObjectURL(url));
  state.reportObjectUrls = [];
}

function revokePrintObjectUrls() {
  state.printObjectUrls.forEach(url => URL.revokeObjectURL(url));
  state.printObjectUrls = [];
}

function getImageFilesFromDrop(event) {
  const files = Array.from(event.dataTransfer?.files || []);
  return files.filter(isImageFile);
}

function isImageFile(file) {
  if (!file) return false;
  if (file.type && file.type.startsWith("image/")) return true;
  return /\.(png|jpe?g|webp|gif|bmp|heic|heif|tiff?|avif|svg)$/i.test(file.name || "");
}

async function handleCardAction(cardId, action) {
  const index = state.cards.findIndex(card => card.id === cardId);
  if (index < 0) return;

  if (action === "delete") {
    state.cards.splice(index, 1);
    await deleteImageByCardId(cardId);
  }

  if (action === "duplicate") {
    const source = state.cards[index];
    const duplicate = createCard({
      ...source,
      id: crypto.randomUUID(),
      shotNo: nextShotNo(),
    });
    state.cards.splice(index + 1, 0, duplicate);
    if (source.image) {
      duplicate.image = source.image;
    } else {
      await copyImageByCardId(source.id, duplicate.id);
    }
  }

  if (action === "up" && index > 0) {
    [state.cards[index - 1], state.cards[index]] = [state.cards[index], state.cards[index - 1]];
  }

  if (action === "down" && index < state.cards.length - 1) {
    [state.cards[index + 1], state.cards[index]] = [state.cards[index], state.cards[index + 1]];
  }

  persistAndRender();
}

function updateCard(cardId, field, value, shouldRender = false) {
  const card = state.cards.find(item => item.id === cardId);
  if (!card) return;
  card[field] = value;
  persist();
  if (shouldRender) render();
}

function persist() {
  persistProjectState();
}

function persistAndRender() {
  persist();
  render();
  if (!elements.report.classList.contains("hidden")) {
    setStatus("数据已更新，请点击“整理分镜表”刷新输出表。");
  }
}

function load() {
  // Deprecated: kept for backward compatibility. Project store now owns cards.
}

function createDefaultProject(name) {
  const id = crypto.randomUUID();
  return {
    id,
    name: name || "默认项目",
    cards: [],
    aiConfig: { ...state.aiConfig },
    createdAt: Date.now(),
    updatedAt: Date.now(),
  };
}

function loadProjectStore() {
  const loaded = readProjectStoreFromLocal();
  if (loaded) {
    state.projectStore = loaded;
    state.currentProjectId = loaded.currentProjectId;
    applyCurrentProjectToState();
    refreshProjectSelector();
    return;
  }

  migrateLegacyToProjectStore();
  persistProjectStore();
  refreshProjectSelector();
}

function loadExportPrefs() {
  try {
    const raw = localStorage.getItem(exportPrefsKey);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") {
        if (typeof parsed.exportCleanMode === "boolean") {
          state.exportCleanMode = parsed.exportCleanMode;
        }
        state.exportShotsPerPageDesktop = normalizeShotsPerPage(parsed.exportShotsPerPageDesktop, 4);
        state.exportShotsPerPageMobile = normalizeShotsPerPage(parsed.exportShotsPerPageMobile, 4);
      }
    }
  } catch {}
  elements.exportCleanToggle.checked = state.exportCleanMode;
  elements.exportShotsDesktopInput.value = String(state.exportShotsPerPageDesktop);
  elements.exportShotsMobileInput.value = String(state.exportShotsPerPageMobile);
}

function persistExportPrefs() {
  try {
    localStorage.setItem(
      exportPrefsKey,
      JSON.stringify({
        exportCleanMode: state.exportCleanMode,
        exportShotsPerPageDesktop: state.exportShotsPerPageDesktop,
        exportShotsPerPageMobile: state.exportShotsPerPageMobile,
      })
    );
  } catch {}
}

function normalizeShotsPerPage(value, fallback) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  const rounded = Math.round(parsed);
  return Math.min(10, Math.max(1, rounded));
}

function getShotsPerPageForCurrentDevice() {
  const isMobile = window.matchMedia("(max-width: 720px)").matches;
  return isMobile ? state.exportShotsPerPageMobile : state.exportShotsPerPageDesktop;
}

function readProjectStoreFromLocal() {
  try {
    const raw = localStorage.getItem(projectStoreKey);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    if (!Array.isArray(parsed.order) || !parsed.projects || typeof parsed.projects !== "object") {
      return null;
    }
    const validOrder = parsed.order.filter(id => parsed.projects[id]);
    if (!validOrder.length) return null;
    const currentProjectId = parsed.currentProjectId && parsed.projects[parsed.currentProjectId]
      ? parsed.currentProjectId
      : validOrder[0];
    return {
      order: validOrder,
      projects: parsed.projects,
      currentProjectId,
    };
  } catch {
    return null;
  }
}

function migrateLegacyToProjectStore() {
  const project = createDefaultProject("默认项目");
  project.cards = readLegacyCards();
  project.aiConfig = readLegacyAiConfig();
  state.projectStore = {
    order: [project.id],
    projects: { [project.id]: project },
    currentProjectId: project.id,
  };
  state.currentProjectId = project.id;
  state.cards = project.cards.map(card => createCard(card));
  state.aiConfig = { ...state.aiConfig, ...project.aiConfig };
}

function readLegacyCards() {
  try {
    const raw = localStorage.getItem(storageKey);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function readLegacyAiConfig() {
  try {
    const raw = localStorage.getItem(aiConfigKey);
    if (!raw) return { ...state.aiConfig };
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return { ...state.aiConfig };
    return { ...state.aiConfig, ...parsed };
  } catch {
    return { ...state.aiConfig };
  }
}

function applyCurrentProjectToState() {
  const project = state.projectStore.projects[state.currentProjectId];
  if (!project) return;
  state.cards = (project.cards || []).map(card => createCard(card));
  state.aiConfig = { ...state.aiConfig, ...(project.aiConfig || {}) };
}

function snapshotCurrentProject() {
  const project = state.projectStore.projects[state.currentProjectId];
  if (!project) return;
  project.cards = state.cards.map(({ image, ...card }) => ({ ...card }));
  project.aiConfig = { ...state.aiConfig };
  project.updatedAt = Date.now();
}

function persistProjectStore() {
  try {
    localStorage.setItem(projectStoreKey, JSON.stringify(state.projectStore));
  } catch {
    setStatus("项目保存失败，请先导出 CSV。");
  }
}

function persistProjectState() {
  snapshotCurrentProject();
  persistProjectStore();
}

function refreshProjectSelector() {
  const options = state.projectStore.order
    .map(id => state.projectStore.projects[id])
    .filter(Boolean)
    .map(project => `<option value="${escapeHtml(project.id)}">${escapeHtml(project.name)}</option>`)
    .join("");
  elements.projectSelect.innerHTML = options;
  elements.projectSelect.value = state.currentProjectId;
}

function switchProject(projectId) {
  if (!projectId || projectId === state.currentProjectId) return;
  if (!state.projectStore.projects[projectId]) return;
  persistProjectState();
  state.currentProjectId = projectId;
  state.projectStore.currentProjectId = projectId;
  applyCurrentProjectToState();
  persistProjectStore();
  clearReport();
  render();
  refreshProjectSelector();
  if (window.matchMedia("(max-width: 720px)").matches && document.body.classList.contains("mobile-minimal")) {
    setEditorVisibility(state.cards.length === 0);
    if (!state.editorVisible && state.cards.length) void buildReportTable();
  }
}

function createNewProject() {
  persistProjectState();
  const defaultName = `项目${state.projectStore.order.length + 1}`;
  const name = prompt("新项目名称：", defaultName);
  if (name === null) return;
  const finalName = name.trim() || defaultName;
  const project = createDefaultProject(finalName);
  state.projectStore.order.push(project.id);
  state.projectStore.projects[project.id] = project;
  state.projectStore.currentProjectId = project.id;
  state.currentProjectId = project.id;
  applyCurrentProjectToState();
  persistProjectStore();
  clearReport();
  render();
  refreshProjectSelector();
  if (window.matchMedia("(max-width: 720px)").matches && document.body.classList.contains("mobile-minimal")) {
    setEditorVisibility(true);
  }
}

function renameCurrentProject() {
  const project = state.projectStore.projects[state.currentProjectId];
  if (!project) return;
  const nextName = prompt("重命名项目：", project.name);
  if (nextName === null) return;
  project.name = nextName.trim() || project.name;
  project.updatedAt = Date.now();
  persistProjectStore();
  refreshProjectSelector();
}

function deleteCurrentProject() {
  const count = state.projectStore.order.length;
  if (count <= 1) {
    alert("至少保留一个项目。");
    return;
  }
  const project = state.projectStore.projects[state.currentProjectId];
  if (!project) return;
  const confirmed = confirm(`确认删除项目「${project.name}」吗？`);
  if (!confirmed) return;
  const removedId = state.currentProjectId;
  delete state.projectStore.projects[removedId];
  state.projectStore.order = state.projectStore.order.filter(id => id !== removedId);
  const fallbackId = state.projectStore.order[0];
  state.currentProjectId = fallbackId;
  state.projectStore.currentProjectId = fallbackId;
  applyCurrentProjectToState();
  persistProjectStore();
  clearReport();
  render();
  refreshProjectSelector();
}

async function exportCurrentProjectPackage() {
  const project = state.projectStore.projects[state.currentProjectId];
  if (!project) return;

  setStatus("正在打包项目...");
  const cards = state.cards.map(({ image, ...card }) => ({ ...card }));
  const images = [];

  for (const card of state.cards) {
    if (card.image && card.image.startsWith("data:image/")) {
      const inline = dataUrlToParts(card.image);
      if (inline) {
        const sourceBlob = base64ToBlob(inline.base64, inline.mime);
        const compressed = await compressImageBlob(sourceBlob);
        const base64 = await blobToBase64(compressed);
        if (!base64) continue;
        const targetExt = mimeToExtension(compressed.type || inline.mime);
        images.push({
          cardId: card.id,
          name: normalizeImageFileName(card.imageName || `image-${card.id}.png`, targetExt),
          mime: compressed.type || inline.mime,
          base64,
        });
      }
      continue;
    }

    const record = await getImageByCardId(card.id);
    if (!record?.blob) continue;
    const compressed = await compressImageBlob(record.blob);
    const base64 = await blobToBase64(compressed);
    if (!base64) continue;
    const targetExt = mimeToExtension(compressed.type || record.blob.type || "image/png");
    images.push({
      cardId: card.id,
      name: normalizeImageFileName(record.name || card.imageName || `image-${card.id}.png`, targetExt),
      mime: compressed.type || record.blob.type || "image/png",
      base64,
    });
  }

  const payload = {
    type: "storyboard-project-package",
    version: 1,
    exportedAt: new Date().toISOString(),
    project: {
      name: project.name,
      cards,
      aiConfig: { ...state.aiConfig },
    },
    images,
  };

  const fileNameSafeProject = sanitizeFileName(project.name || "project");
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${fileNameSafeProject}-project-package.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  setStatus("项目包已导出。");
}

async function importProjectPackage(file) {
  setStatus("正在导入项目包...");
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    if (!parsed || parsed.type !== "storyboard-project-package" || !parsed.project) {
      alert("项目包格式无效。");
      setStatus("导入失败：项目包格式无效。");
      return;
    }

    const sourceProject = parsed.project;
    const sourceCards = Array.isArray(sourceProject.cards) ? sourceProject.cards : [];
    const sourceImages = Array.isArray(parsed.images) ? parsed.images : [];
    const sourceAiConfig = sourceProject.aiConfig && typeof sourceProject.aiConfig === "object"
      ? sourceProject.aiConfig
      : {};

    persistProjectState();

    const newProjectName = getImportedProjectName(sourceProject.name || "导入项目");
    const project = createDefaultProject(newProjectName);
    const oldToNewCardId = {};
    const importedCards = sourceCards.map(card => {
      const newCardId = crypto.randomUUID();
      oldToNewCardId[card.id] = newCardId;
      return createCard({
        ...card,
        id: newCardId,
        image: "",
      });
    });

    project.cards = importedCards.map(({ image, ...card }) => ({ ...card }));
    project.aiConfig = { ...state.aiConfig, ...sourceAiConfig };

    state.projectStore.order.push(project.id);
    state.projectStore.projects[project.id] = project;

    for (const imageItem of sourceImages) {
      const oldCardId = imageItem?.cardId;
      const newCardId = oldToNewCardId[oldCardId];
      if (!newCardId) continue;
      if (!imageItem.base64 || !imageItem.mime) continue;
      const dataUrl = `data:${imageItem.mime};base64,${imageItem.base64}`;
      await saveDataUrlImage(newCardId, dataUrl, imageItem.name || "");
    }

    state.currentProjectId = project.id;
    state.projectStore.currentProjectId = project.id;
    applyCurrentProjectToState();
    persistProjectStore();
    clearReport();
    render();
    refreshProjectSelector();
    setStatus(`导入成功：${project.name}`);
  } catch (error) {
    setStatus("导入失败：文件读取或解析错误。");
    alert(`导入失败：${shortenText(String(error), 120)}`);
  }
}

function getImportedProjectName(baseName) {
  const normalized = (baseName || "导入项目").trim();
  const names = state.projectStore.order
    .map(id => state.projectStore.projects[id]?.name || "")
    .filter(Boolean);
  if (!names.includes(normalized)) return normalized;
  let index = 2;
  while (names.includes(`${normalized} (${index})`)) index += 1;
  return `${normalized} (${index})`;
}

function sanitizeFileName(name) {
  return String(name || "project")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, "_")
    .slice(0, 60);
}

function dataUrlToParts(dataUrl) {
  const matched = String(dataUrl || "").match(/^data:(.+?);base64,(.+)$/);
  if (!matched) return null;
  return { mime: matched[1], base64: matched[2] };
}

function blobToBase64(blob) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      const matched = result.match(/^data:.+;base64,(.+)$/);
      resolve(matched ? matched[1] : "");
    };
    reader.onerror = () => resolve("");
    reader.readAsDataURL(blob);
  });
}

function base64ToBlob(base64, mime) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: mime || "application/octet-stream" });
}

async function compressImageBlob(blob) {
  if (!blob?.type?.startsWith("image/")) return blob;
  const image = await loadImageFromBlob(blob);
  if (!image) return blob;

  const sourceWidth = image.naturalWidth || image.width || 0;
  const sourceHeight = image.naturalHeight || image.height || 0;
  if (!sourceWidth || !sourceHeight) return blob;

  const ratio = Math.min(1, exportImageMaxEdge / Math.max(sourceWidth, sourceHeight));
  const targetWidth = Math.max(1, Math.round(sourceWidth * ratio));
  const targetHeight = Math.max(1, Math.round(sourceHeight * ratio));

  const canvas = document.createElement("canvas");
  canvas.width = targetWidth;
  canvas.height = targetHeight;
  const ctx = canvas.getContext("2d", { alpha: true });
  if (!ctx) return blob;
  ctx.drawImage(image, 0, 0, targetWidth, targetHeight);

  const hasAlpha = detectCanvasAlpha(ctx, targetWidth, targetHeight);
  const outputType = hasAlpha ? "image/png" : "image/jpeg";
  const quality = outputType === "image/jpeg" ? exportJpegQuality : undefined;
  const compressed = await canvasToBlob(canvas, outputType, quality);
  if (!compressed) return blob;

  if (compressed.size < blob.size * 0.98) return compressed;
  if (ratio < 1 && compressed.size <= blob.size) return compressed;
  return blob;
}

function detectCanvasAlpha(ctx, width, height) {
  try {
    const data = ctx.getImageData(0, 0, width, height).data;
    for (let index = 3; index < data.length; index += 16) {
      if (data[index] < 255) return true;
    }
    return false;
  } catch {
    return false;
  }
}

function canvasToBlob(canvas, type, quality) {
  return new Promise(resolve => {
    canvas.toBlob(result => resolve(result || null), type, quality);
  });
}

function loadImageFromBlob(blob) {
  return new Promise(resolve => {
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      resolve(null);
    };
    image.src = url;
  });
}

function mimeToExtension(mime) {
  const mapping = {
    "image/jpeg": "jpg",
    "image/png": "png",
    "image/webp": "webp",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/avif": "avif",
  };
  return mapping[mime] || "png";
}

function normalizeImageFileName(name, ext) {
  const base = String(name || "image")
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .trim();
  const safeBase = base || "image";
  return `${safeBase}.${ext}`;
}

async function migrateLegacyInlineImages() {
  const raw = localStorage.getItem(storageKey);
  if (!raw) return;
  let parsed = null;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return;
  }
  if (!Array.isArray(parsed)) return;

  let migrated = 0;
  for (const card of parsed) {
    if (!card || typeof card.image !== "string" || !card.image.startsWith("data:image/")) continue;
    const target = state.cards.find(item => item.id === card.id);
    if (!target) continue;
    const saved = await saveDataUrlImage(target.id, card.image, target.imageName || card.imageName || "");
    if (saved) {
      target.image = "";
      migrated += 1;
    }
  }

  if (migrated) {
    persistAndRender();
    setStatus(`已迁移 ${migrated} 张旧图片到图片存储，后续可继续批量拖图。`);
  }
}

async function attachImageToCard(card, file) {
  const saved = await saveImageForCard(card.id, file);
  if (saved) {
    card.image = "";
    card.imageName = file.name;
    return true;
  }

  if (!state.imageDbDisabled) {
    setStatus("图片存储空间可能已满，请压缩图片后再导入。");
    return false;
  }

  const fallback = await fileToDataUrl(file);
  if (!fallback) return false;
  card.image = fallback.dataUrl;
  card.imageName = file.name;
  setStatus("当前环境不支持图片持久化，图片仅在本次页面会话可用。");
  return true;
}

async function ensureImageDb() {
  const db = await openImageDb();
  if (db) return true;
  state.imageDbDisabled = true;
  setStatus("浏览器限制了图片存储，刷新页面后图片会丢失。");
  return false;
}

function openImageDb() {
  if (state.imageDbDisabled) return Promise.resolve(null);
  if (!("indexedDB" in window)) return Promise.resolve(null);
  if (imageDbPromise) return imageDbPromise;

  imageDbPromise = new Promise(resolve => {
    const request = indexedDB.open(imageDbName, 1);

    request.onupgradeneeded = event => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains(imageStoreName)) {
        db.createObjectStore(imageStoreName, { keyPath: "id" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
    request.onblocked = () => resolve(null);
  });

  return imageDbPromise;
}

async function runImageStore(mode, operation) {
  const db = await openImageDb();
  if (!db) return null;
  return new Promise(resolve => {
    const tx = db.transaction(imageStoreName, mode);
    const store = tx.objectStore(imageStoreName);
    operation(store, resolve);
    tx.onerror = () => resolve(null);
  });
}

async function saveImageForCard(cardId, file) {
  if (!(await ensureImageDb())) return false;
  const result = await runImageStore("readwrite", (store, resolve) => {
    const request = store.put({ id: cardId, blob: file, name: file.name });
    request.onsuccess = () => resolve(true);
    request.onerror = () => resolve(false);
  });
  return result === true;
}

async function saveDataUrlImage(cardId, dataUrl, name) {
  if (!(await ensureImageDb())) return false;
  try {
    const blob = await (await fetch(dataUrl)).blob();
    const result = await runImageStore("readwrite", (store, resolve) => {
      const request = store.put({ id: cardId, blob, name });
      request.onsuccess = () => resolve(true);
      request.onerror = () => resolve(false);
    });
    return result === true;
  } catch {
    return false;
  }
}

async function getImageByCardId(cardId) {
  const result = await runImageStore("readonly", (store, resolve) => {
    const request = store.get(cardId);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => resolve(null);
  });
  return result;
}

async function deleteImageByCardId(cardId) {
  await runImageStore("readwrite", (store, resolve) => {
    const request = store.delete(cardId);
    request.onsuccess = () => resolve(true);
    request.onerror = () => resolve(false);
  });
}

async function clearAllImages() {
  await runImageStore("readwrite", (store, resolve) => {
    const request = store.clear();
    request.onsuccess = () => resolve(true);
    request.onerror = () => resolve(false);
  });
}

async function copyImageByCardId(fromId, toId) {
  const source = await getImageByCardId(fromId);
  if (!source?.blob) return false;
  const result = await runImageStore("readwrite", (store, resolve) => {
    const request = store.put({ id: toId, blob: source.blob, name: source.name || "" });
    request.onsuccess = () => resolve(true);
    request.onerror = () => resolve(false);
  });
  return result === true;
}

function exportCsv() {
  const headers = ["镜号", "景别", "场景", "人员", "时间", "备注", "图片名"];
  const lines = [headers.join(",")];

  state.cards.forEach(card => {
    const row = [
      card.shotNo,
      card.shotType,
      card.scene,
      card.people,
      card.time,
      card.notes,
      card.imageName,
    ].map(csvEscape);
    lines.push(row.join(","));
  });

  const blob = new Blob(["\ufeff" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "storyboard.csv";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const text = String(value || "");
  if (!text.includes(",") && !text.includes('"') && !text.includes("\n")) return text;
  return `"${text.replaceAll('"', '""')}"`;
}

function setStatus(message) {
  if (!elements.status) return;
  elements.status.textContent = message || "";
}

async function buildReportTable() {
  if (!state.cards.length) {
    alert("当前没有可整理的数据。");
    return;
  }

  revokeReportObjectUrls();
  const rows = [];
  for (const card of state.cards) {
    const imageSrc = await getImageSrcForReport(card);
    const shotTypeOptions = getReportShotTypeOptions(card.shotType);
    rows.push(`
      <tr data-card-id="${escapeHtml(card.id)}">
        <td data-label="镜号"><input class="report-input" data-report-field="shotNo" value="${escapeHtml(card.shotNo)}" /></td>
        <td data-label="缩略图">
          <div class="report-thumb-wrap">
            <img class="report-thumb" src="${escapeHtml(imageSrc)}" alt="镜头 ${escapeHtml(card.shotNo)}" />
          </div>
        </td>
        <td data-label="景别">
          <select class="report-select" data-report-field="shotType">
            ${shotTypeOptions}
          </select>
        </td>
        <td data-label="场景"><input class="report-input" data-report-field="scene" value="${escapeHtml(card.scene)}" /></td>
        <td data-label="人员"><input class="report-input" data-report-field="people" value="${escapeHtml(card.people)}" /></td>
        <td data-label="时间"><input class="report-input" data-report-field="time" type="time" step="60" value="${escapeHtml(card.time)}" /></td>
        <td data-label="备注"><textarea class="report-textarea" data-report-field="notes">${escapeHtml(card.notes)}</textarea></td>
        <td data-label="操作">
          <div class="report-actions">
            <button type="button" data-report-action="up">上移</button>
            <button type="button" data-report-action="down">下移</button>
          </div>
        </td>
      </tr>
    `);
  }

  elements.reportBody.innerHTML = rows.join("");
  bindReportTableEvents();
  elements.reportMeta.textContent = `总镜头数：${state.cards.length} | 生成时间：${formatNow()}`;
  elements.report.classList.remove("hidden");
  setStatus("");
}

function clearReport() {
  revokeReportObjectUrls();
  elements.reportBody.innerHTML = "";
  elements.report.classList.add("hidden");
  clearPrintPages();
}

function clearPrintPages() {
  revokePrintObjectUrls();
  elements.reportPrint.innerHTML = "";
}

function bindReportTableEvents() {
  elements.reportBody.querySelectorAll("[data-report-field]").forEach(input => {
    const applyFieldUpdate = (event, shouldRenderBoard) => {
      const row = event.target.closest("tr");
      if (!row) return;
      const cardId = row.dataset.cardId;
      const field = event.target.dataset.reportField;
      const card = state.cards.find(item => item.id === cardId);
      if (!card || !field) return;
      card[field] = event.target.value;
      persist();
      if (shouldRenderBoard) {
        if (field === "shotNo") {
          sortCardsByShotNo();
          persist();
          render();
          void buildReportTable();
          return;
        }
        render();
      }
    };

    input.addEventListener("input", event => applyFieldUpdate(event, false));
    input.addEventListener("change", event => applyFieldUpdate(event, true));
  });

  elements.reportBody.querySelectorAll("[data-report-action]").forEach(button => {
    button.addEventListener("click", event => {
      const row = event.target.closest("tr");
      if (!row) return;
      const cardId = row.dataset.cardId;
      const index = state.cards.findIndex(card => card.id === cardId);
      if (index < 0) return;

      if (button.dataset.reportAction === "up" && index > 0) {
        [state.cards[index - 1], state.cards[index]] = [state.cards[index], state.cards[index - 1]];
      }

      if (button.dataset.reportAction === "down" && index < state.cards.length - 1) {
        [state.cards[index + 1], state.cards[index]] = [state.cards[index], state.cards[index + 1]];
      }

      persist();
      render();
      void buildReportTable();
    });
  });
}

async function getImageSrcForReport(card) {
  if (card.image) return card.image;
  const record = await getImageByCardId(card.id);
  if (!record?.blob) return placeholderImage;
  const objectUrl = URL.createObjectURL(record.blob);
  state.reportObjectUrls.push(objectUrl);
  return objectUrl;
}

function formatNow() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  return `${year}-${month}-${day} ${hours}:${minutes}`;
}

function escapeHtml(value) {
  const text = String(value || "");
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function getReportShotTypeOptions(currentValue) {
  const options = ["", "大全景", "全景", "中景", "近景", "特写"];
  return options
    .map(option => {
      const selected = option === currentValue ? " selected" : "";
      const label = option || "选择景别";
      return `<option value="${escapeHtml(option)}"${selected}>${escapeHtml(label)}</option>`;
    })
    .join("");
}

function sortCardsByShotNo() {
  state.cards.sort((left, right) => compareShotNo(left.shotNo, right.shotNo));
}

function compareShotNo(left, right) {
  const leftText = String(left || "").trim();
  const rightText = String(right || "").trim();
  const leftNumber = parseFirstNumber(leftText);
  const rightNumber = parseFirstNumber(rightText);

  if (leftNumber !== null && rightNumber !== null && leftNumber !== rightNumber) {
    return leftNumber - rightNumber;
  }
  if (leftNumber !== null && rightNumber === null) return -1;
  if (leftNumber === null && rightNumber !== null) return 1;
  return leftText.localeCompare(rightText, "zh-Hans-CN", { numeric: true, sensitivity: "base" });
}

function parseFirstNumber(value) {
  const matched = value.match(/\d+(?:\.\d+)?/);
  if (!matched) return null;
  const parsed = Number(matched[0]);
  return Number.isFinite(parsed) ? parsed : null;
}

async function waitForImagesInContainer(container) {
  const images = Array.from(container.querySelectorAll("img"));
  if (!images.length) return;

  await Promise.all(
    images.map(
      image =>
        new Promise(resolve => {
          if (image.complete && image.naturalWidth > 0) {
            resolve(true);
            return;
          }
          image.onload = () => resolve(true);
          image.onerror = () => resolve(false);
        })
    )
  );
}

async function exportReportPng(shotsPerPage) {
  if (!state.cards.length) {
    alert("当前没有可导出的数据。");
    return;
  }

  setStatus("正在生成 PNG...");
  const snapshotRows = await preparePngRows();
  const pages = chunkRows(snapshotRows, shotsPerPage);
  if (!pages.length) {
    setStatus("PNG 生成失败，请重试。");
    return;
  }

  for (let index = 0; index < pages.length; index += 1) {
    const rows = pages[index];
    const canvas = renderPngCanvas(rows, index + 1, pages.length);
    if (!canvas) {
      setStatus("PNG 生成失败，请重试。");
      return;
    }
    const fileName = pages.length === 1 ? "storyboard-report.png" : `storyboard-report-p${index + 1}.png`;
    await downloadCanvasAsPng(canvas, fileName);
  }
  setStatus(`PNG 导出完成，共 ${pages.length} 张。`);
}

async function preparePngRows() {
  const rows = [];
  for (const card of state.cards) {
    const image = await loadCardImageElement(card);
    rows.push({
      shotNo: card.shotNo || "",
      shotType: card.shotType || "",
      scene: card.scene || "",
      people: card.people || "",
      time: card.time || "",
      notes: card.notes || "",
      image,
    });
  }
  return rows;
}

async function loadCardImageElement(card) {
  if (card.image) {
    return loadImageFromSrc(card.image);
  }

  const record = await getImageByCardId(card.id);
  if (!record?.blob) return null;
  const objectUrl = URL.createObjectURL(record.blob);
  try {
    return await loadImageFromSrc(objectUrl);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function loadImageFromSrc(src) {
  return new Promise(resolve => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => resolve(null);
    image.src = src;
  });
}

function renderPngCanvas(rows, pageIndex, totalPages) {
  const width = 1900;
  const padding = 32;
  const headerHeight = 84;
  const tableHeaderHeight = 44;
  const colWidths = [110, 500, 120, 220, 220, 140, 526];
  const columns = ["镜号", "缩略图", "景别", "场景", "人员", "时间", "备注"];

  const measureCanvas = document.createElement("canvas");
  const measureCtx = measureCanvas.getContext("2d");
  if (!measureCtx) return null;
  measureCtx.font = "16px sans-serif";

  const preparedRows = rows.map(row => {
    const wrapped = {
      shotNo: wrapCanvasText(measureCtx, row.shotNo, colWidths[0] - 16),
      shotType: wrapCanvasText(measureCtx, row.shotType, colWidths[2] - 16),
      scene: wrapCanvasText(measureCtx, row.scene, colWidths[3] - 16),
      people: wrapCanvasText(measureCtx, row.people, colWidths[4] - 16),
      time: wrapCanvasText(measureCtx, row.time, colWidths[5] - 16),
      notes: wrapCanvasText(measureCtx, row.notes, colWidths[6] - 16),
    };
    const lineHeight = 22;
    const textHeight = Math.max(
      wrapped.shotNo.length,
      wrapped.shotType.length,
      wrapped.scene.length,
      wrapped.people.length,
      wrapped.time.length,
      wrapped.notes.length
    );
    const rowHeight = Math.max(320, textHeight * lineHeight + 18);
    return { ...row, wrapped, rowHeight };
  });

  const bodyHeight = preparedRows.reduce((sum, row) => sum + row.rowHeight, 0);
  const height = padding + headerHeight + tableHeaderHeight + bodyHeight + padding;

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.fillStyle = "#1f1c19";
  ctx.font = "700 34px sans-serif";
  ctx.fillText("分镜脚本表", padding, padding + 30);
  ctx.font = "16px sans-serif";
  ctx.fillStyle = "#6a6258";
  ctx.fillText(
    `页码：${pageIndex}/${totalPages}  本页镜头数：${preparedRows.length}  生成时间：${formatNow()}`,
    padding,
    padding + 58
  );

  const tableX = padding;
  const tableY = padding + headerHeight;
  const tableWidth = colWidths.reduce((sum, value) => sum + value, 0);

  ctx.fillStyle = "#f3ecdf";
  ctx.fillRect(tableX, tableY, tableWidth, tableHeaderHeight);
  ctx.strokeStyle = "#c8bba8";
  ctx.lineWidth = 1;

  let x = tableX;
  columns.forEach((title, index) => {
    ctx.strokeRect(x, tableY, colWidths[index], tableHeaderHeight);
    ctx.fillStyle = "#1f1c19";
    ctx.font = "600 16px sans-serif";
    ctx.fillText(title, x + 10, tableY + 27);
    x += colWidths[index];
  });

  let y = tableY + tableHeaderHeight;
  preparedRows.forEach(row => {
    drawPngRow(ctx, row, tableX, y, colWidths);
    y += row.rowHeight;
  });

  return canvas;
}

function chunkRows(rows, pageSize) {
  const chunks = [];
  for (let index = 0; index < rows.length; index += pageSize) {
    chunks.push(rows.slice(index, index + pageSize));
  }
  return chunks;
}

function drawPngRow(ctx, row, tableX, rowY, colWidths) {
  const lineHeight = 22;
  const rowHeight = row.rowHeight;
  let x = tableX;

  colWidths.forEach(width => {
    ctx.strokeStyle = "#c8bba8";
    ctx.strokeRect(x, rowY, width, rowHeight);
    x += width;
  });

  drawLines(ctx, row.wrapped.shotNo, tableX + 8, rowY + 24, lineHeight);
  drawThumbInCell(ctx, row.image, tableX + colWidths[0], rowY, colWidths[1], rowHeight);

  let columnX = tableX + colWidths[0] + colWidths[1];
  drawLines(ctx, row.wrapped.shotType, columnX + 8, rowY + 24, lineHeight);
  columnX += colWidths[2];
  drawLines(ctx, row.wrapped.scene, columnX + 8, rowY + 24, lineHeight);
  columnX += colWidths[3];
  drawLines(ctx, row.wrapped.people, columnX + 8, rowY + 24, lineHeight);
  columnX += colWidths[4];
  drawLines(ctx, row.wrapped.time, columnX + 8, rowY + 24, lineHeight);
  columnX += colWidths[5];
  drawLines(ctx, row.wrapped.notes, columnX + 8, rowY + 24, lineHeight);
}

function drawThumbInCell(ctx, image, cellX, cellY, cellWidth, cellHeight) {
  const innerPad = 8;
  const boxX = cellX + innerPad;
  const boxY = cellY + innerPad;
  const boxW = cellWidth - innerPad * 2;
  const boxH = cellHeight - innerPad * 2;
  ctx.fillStyle = "#ece7df";
  ctx.fillRect(boxX, boxY, boxW, boxH);
  if (!image) return;

  const imageRatio = image.width / image.height;
  const boxRatio = boxW / boxH;
  let drawW = boxW;
  let drawH = boxH;
  if (imageRatio > boxRatio) {
    drawH = boxW / imageRatio;
  } else {
    drawW = boxH * imageRatio;
  }
  const drawX = boxX + (boxW - drawW) / 2;
  const drawY = boxY + (boxH - drawH) / 2;
  ctx.drawImage(image, drawX, drawY, drawW, drawH);
}

function drawLines(ctx, lines, x, startY, lineHeight) {
  ctx.fillStyle = "#1f1c19";
  ctx.font = "16px sans-serif";
  lines.forEach((line, index) => ctx.fillText(line, x, startY + index * lineHeight));
}

function wrapCanvasText(ctx, text, maxWidth) {
  const raw = String(text || "");
  if (!raw) return [""];

  const lines = [];
  raw.split("\n").forEach(paragraph => {
    if (!paragraph) {
      lines.push("");
      return;
    }
    let current = "";
    for (const char of paragraph) {
      const candidate = current + char;
      if (ctx.measureText(candidate).width <= maxWidth || current.length === 0) {
        current = candidate;
      } else {
        lines.push(current);
        current = char;
      }
    }
    lines.push(current);
  });
  return lines;
}

async function downloadCanvasAsPng(canvas, filename) {
  const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/png"));
  if (!blob) return;

  const url = URL.createObjectURL(blob);
  const isDownloadSupported = "download" in document.createElement("a");
  const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);

  if (isDownloadSupported && !isIOS) {
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    return;
  }

  window.open(url, "_blank");
  setStatus("已打开 PNG 预览页，请长按图片保存。");
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

async function buildPrintPagesForExport(shotsPerPage) {
  revokePrintObjectUrls();
  const pageRows = [];
  for (const card of state.cards) {
    const imageSrc = await getImageSrcForPrint(card);
    pageRows.push({
      shotNo: card.shotNo || "",
      shotType: card.shotType || "",
      scene: card.scene || "",
      people: card.people || "",
      time: card.time || "",
      notes: card.notes || "",
      imageSrc,
    });
  }

  const pages = chunkRows(pageRows, shotsPerPage);
  const html = pages
    .map((rows, pageIndex) => {
      const bodyRows = rows
        .map(
          row => `
          <tr>
            <td>${escapeHtml(row.shotNo)}</td>
            <td><div class="report-thumb-wrap"><img class="report-thumb" src="${escapeHtml(row.imageSrc)}" alt="镜头 ${escapeHtml(row.shotNo)}" /></div></td>
            <td>${escapeHtml(row.shotType)}</td>
            <td>${escapeHtml(row.scene)}</td>
            <td>${escapeHtml(row.people)}</td>
            <td>${escapeHtml(row.time)}</td>
            <td>${escapeHtml(row.notes)}</td>
          </tr>
        `
        )
        .join("");

      return `
        <section class="print-page">
          <p class="print-page-head">分镜脚本表 | 第 ${pageIndex + 1}/${pages.length} 页 | 每页最多 ${shotsPerPage} 镜头 | ${formatNow()}</p>
          <table class="report-table">
            <thead>
              <tr>
                <th>镜号</th>
                <th>缩略图</th>
                <th>景别</th>
                <th>场景</th>
                <th>人员</th>
                <th>时间</th>
                <th>备注</th>
              </tr>
            </thead>
            <tbody>${bodyRows}</tbody>
          </table>
        </section>
      `;
    })
    .join("");

  elements.reportPrint.innerHTML = html;
}

async function getImageSrcForPrint(card) {
  if (card.image) return card.image;
  const record = await getImageByCardId(card.id);
  if (!record?.blob) return placeholderImage;
  const objectUrl = URL.createObjectURL(record.blob);
  state.printObjectUrls.push(objectUrl);
  return objectUrl;
}
