// Google Drive の公開フォルダを教材ライブラリとして使うための設定です。
// ここだけ差し替えれば、GitHub Pages 側のコードはそのまま使えます。
const DRIVE_FOLDER_ID = "1a9iaxd2OKBywSuRZvPib1gyGciL22HpM";
const API_KEY = "AIzaSyBvD4bw5zAdaDcsI3cyBRb4q5mbKJTG2J8";

const DRIVE_API_URL = "https://www.googleapis.com/drive/v3/files";
const PAGE_SIZE = 100;

const state = {
  files: [],
  filteredFiles: [],
  activeCategory: "all",
  searchQuery: ""
};

const elements = {
  fileCount: document.getElementById("fileCount"),
  lastUpdated: document.getElementById("lastUpdated"),
  statusPanel: document.getElementById("statusPanel"),
  resultSummary: document.getElementById("resultSummary"),
  loadingGrid: document.getElementById("loadingGrid"),
  libraryGrid: document.getElementById("libraryGrid"),
  emptyState: document.getElementById("emptyState"),
  searchInput: document.getElementById("searchInput"),
  categoryTabs: document.getElementById("categoryTabs")
};

document.addEventListener("DOMContentLoaded", () => {
  renderLoadingSkeleton();
  bindEvents();
  loadDriveFiles();
});

function bindEvents() {
  elements.searchInput.addEventListener("input", (event) => {
    state.searchQuery = event.target.value.trim().toLowerCase();
    filterFiles();
  });

  elements.categoryTabs.addEventListener("click", (event) => {
    const button = event.target.closest("[data-category]");

    if (!button) {
      return;
    }

    state.activeCategory = button.dataset.category;
    updateActiveTab();
    filterFiles();
  });

  elements.libraryGrid.addEventListener("click", (event) => {
    const trigger = event.target.closest("[data-open-url]");

    if (!trigger) {
      return;
    }

    openFile(event, trigger.dataset.openUrl);
  });
}

async function loadDriveFiles() {
  if (!isConfigured()) {
    showStatus("info", "設定が必要です。script.js の DRIVE_FOLDER_ID と API_KEY を入力してください。");
    updateSummary("設定待ち");
    toggleLoading(false);
    return;
  }

  showStatus("info", "Google Drive から教材一覧を読み込んでいます。");
  toggleLoading(true);

  try {
    const files = await fetchAllFiles();

    // 念のため取得後にも更新日時の新しい順へ揃えます。
    state.files = files.sort((a, b) => new Date(b.modifiedTime) - new Date(a.modifiedTime));
    renderCategoryTabs();
    filterFiles();
    updateHero(state.files);
    showStatus("success", `${state.files.length} 件の教材を読み込みました。`);
  } catch (error) {
    console.error(error);
    handleLoadError(error);
    state.files = [];
    state.filteredFiles = [];
    renderCategoryTabs();
    renderFiles();
    updateHero([]);
  } finally {
    toggleLoading(false);
  }
}

async function fetchAllFiles() {
  const allFiles = [];
  let pageToken = "";

  do {
    const requestUrl = new URL(DRIVE_API_URL);
    requestUrl.searchParams.set("key", API_KEY);
    requestUrl.searchParams.set("q", `'${DRIVE_FOLDER_ID}' in parents and trashed=false`);
    requestUrl.searchParams.set("fields", "nextPageToken,files(id,name,mimeType,thumbnailLink,webViewLink,modifiedTime)");
    requestUrl.searchParams.set("orderBy", "modifiedTime desc");
    requestUrl.searchParams.set("pageSize", String(PAGE_SIZE));
    requestUrl.searchParams.set("includeItemsFromAllDrives", "true");
    requestUrl.searchParams.set("supportsAllDrives", "true");

    if (pageToken) {
      requestUrl.searchParams.set("pageToken", pageToken);
    }

    const response = await fetch(requestUrl.toString());

    if (!response.ok) {
      const errorPayload = await response.json().catch(() => ({}));
      const message = errorPayload.error?.message || "Google Drive API request failed.";
      throw new Error(`${response.status}: ${message}`);
    }

    const data = await response.json();
    allFiles.push(...(data.files || []));
    pageToken = data.nextPageToken || "";
  } while (pageToken);

  return allFiles;
}

function renderFiles() {
  const html = state.filteredFiles.map((file) => createCard(file)).join("");
  elements.libraryGrid.innerHTML = html;

  const hasFiles = state.filteredFiles.length > 0;
  elements.emptyState.classList.toggle("is-hidden", hasFiles);
  updateSummary(hasFiles ? `${state.filteredFiles.length} 件を表示中` : "0 件");
}

function filterFiles() {
  state.filteredFiles = state.files.filter((file) => {
    const fileType = getFileType(file.mimeType).category;
    const matchesCategory = state.activeCategory === "all" || state.activeCategory === fileType;
    const matchesSearch = !state.searchQuery || file.name.toLowerCase().includes(state.searchQuery);
    return matchesCategory && matchesSearch;
  });

  renderFiles();
}

function createCard(file) {
  const fileType = getFileType(file.mimeType);
  const thumbnail = file.thumbnailLink
    ? `<img src="${escapeHtml(file.thumbnailLink)}" alt="${escapeHtml(file.name)} のサムネイル" loading="lazy" referrerpolicy="no-referrer">`
    : `
      <div class="icon-fallback">
        <img src="${getFallbackIcon(fileType)}" alt="${escapeHtml(fileType.label)} アイコン" loading="lazy">
      </div>
    `;

  return `
    <article class="card">
      <div class="card-media">
        ${thumbnail}
      </div>
      <div class="card-body">
        <span class="card-type">${escapeHtml(fileType.label)}</span>
        <h3 class="card-title">${escapeHtml(file.name)}</h3>
        <div class="card-meta">
          <span>種別: ${escapeHtml(fileType.description)}</span>
          <span>更新日: ${formatDate(file.modifiedTime)}</span>
        </div>
        <a
          class="card-link"
          href="${escapeHtml(file.webViewLink || "#")}"
          target="_blank"
          rel="noopener noreferrer"
          data-open-url="${escapeHtml(file.webViewLink || "")}"
        >
          開く
        </a>
      </div>
    </article>
  `;
}

function formatDate(dateString) {
  if (!dateString) {
    return "-";
  }

  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  }).format(new Date(dateString));
}

function getFileType(mimeType) {
  const typeMap = [
    {
      match: ["application/pdf"],
      category: "pdf",
      label: "PDF",
      description: "PDFファイル",
      iconKey: "pdf"
    },
    {
      match: ["video/"],
      category: "video",
      label: "動画",
      description: "動画ファイル",
      iconKey: "video"
    },
    {
      match: ["image/"],
      category: "image",
      label: "画像",
      description: "画像ファイル",
      iconKey: "image"
    },
    {
      match: ["application/vnd.google-apps.document"],
      category: "docs",
      label: "Docs",
      description: "Google Docs",
      iconKey: "docs"
    },
    {
      match: ["application/vnd.google-apps.presentation"],
      category: "slides",
      label: "Slides",
      description: "Google Slides",
      iconKey: "slides"
    },
    {
      match: ["application/vnd.google-apps.folder"],
      category: "folder",
      label: "Folder",
      description: "Google Drive フォルダ",
      iconKey: "folder"
    }
  ];

  const matchedType = typeMap.find((type) => type.match.some((pattern) => mimeType?.startsWith(pattern) || mimeType === pattern));
  return matchedType || {
    category: "other",
    label: "その他",
    description: "その他のファイル",
    iconKey: "other"
  };
}

function renderCategoryTabs() {
  const categoryLabels = {
    all: "すべて",
    pdf: "PDF",
    video: "動画",
    image: "画像",
    docs: "Docs",
    slides: "Slides",
    folder: "Folder",
    other: "その他"
  };

  const counts = state.files.reduce((accumulator, file) => {
    const category = getFileType(file.mimeType).category;
    accumulator[category] = (accumulator[category] || 0) + 1;
    return accumulator;
  }, {});

  const categories = ["all", ...Object.keys(categoryLabels).filter((key) => counts[key])];

  elements.categoryTabs.innerHTML = categories.map((category) => {
    const count = category === "all" ? state.files.length : counts[category];
    const isActive = category === state.activeCategory ? " is-active" : "";

    return `
      <button class="category-tab${isActive}" data-category="${category}" type="button">
        ${categoryLabels[category]} (${count})
      </button>
    `;
  }).join("");
}

function renderLoadingSkeleton() {
  const skeletonCards = Array.from({ length: 8 }, () => `
    <div class="skeleton-card">
      <div class="skeleton-media"></div>
      <div class="skeleton-body">
        <div class="skeleton-pill"></div>
        <div class="skeleton-line long"></div>
        <div class="skeleton-line mid"></div>
        <div class="skeleton-line short"></div>
      </div>
    </div>
  `).join("");

  elements.loadingGrid.innerHTML = skeletonCards;
}

function updateHero(files) {
  elements.fileCount.textContent = String(files.length);
  elements.lastUpdated.textContent = files.length ? formatDate(files[0].modifiedTime) : "-";
}

function updateSummary(text) {
  elements.resultSummary.textContent = text;
}

function updateActiveTab() {
  const tabButtons = elements.categoryTabs.querySelectorAll("[data-category]");
  tabButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.category === state.activeCategory);
  });
}

function toggleLoading(isLoading) {
  elements.loadingGrid.hidden = !isLoading;
  elements.libraryGrid.hidden = isLoading;
}

function showStatus(type, message) {
  elements.statusPanel.className = `status-panel is-visible is-${type}`;
  elements.statusPanel.innerHTML = `
    <div class="status-message">
      <strong>${getStatusTitle(type)}</strong>
      <span>${escapeHtml(message)}</span>
    </div>
  `;
}

function getStatusTitle(type) {
  if (type === "error") {
    return "読み込みに失敗しました。";
  }

  if (type === "success") {
    return "読み込み完了。";
  }

  return "お知らせ";
}

function handleLoadError(error) {
  const message = error.message || "";

  if (message.includes("403")) {
    showStatus("error", "API Key の設定、または対象フォルダの公開権限を確認してください。");
    updateSummary("アクセスエラー");
    return;
  }

  if (message.includes("400")) {
    showStatus("error", "フォルダID、API Key、または Drive API のクエリ設定を確認してください。");
    updateSummary("設定エラー");
    return;
  }

  showStatus("error", "通信に失敗しました。ネットワーク接続と Google Drive API の状態を確認してください。");
  updateSummary("通信エラー");
}

function getFallbackIcon(fileType) {
  const iconMap = {
    pdf: createIconDataUri("#ef4444", "PDF"),
    video: createIconDataUri("#f59e0b", "MOV"),
    image: createIconDataUri("#10b981", "IMG"),
    docs: createIconDataUri("#2563eb", "DOC"),
    slides: createIconDataUri("#ea580c", "SLD"),
    folder: createIconDataUri("#7c3aed", "DIR"),
    other: createIconDataUri("#475569", "FILE")
  };

  return iconMap[fileType.iconKey] || iconMap.other;
}

function createIconDataUri(color, label) {
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="640" height="400" viewBox="0 0 640 400">
      <defs>
        <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stop-color="#ffffff"/>
          <stop offset="100%" stop-color="#eef2ff"/>
        </linearGradient>
      </defs>
      <rect width="640" height="400" rx="36" fill="url(#g)"/>
      <rect x="48" y="48" width="544" height="304" rx="28" fill="${color}" opacity="0.14"/>
      <rect x="72" y="86" width="132" height="160" rx="20" fill="${color}" opacity="0.92"/>
      <rect x="230" y="108" width="230" height="26" rx="13" fill="#cbd5e1"/>
      <rect x="230" y="156" width="168" height="22" rx="11" fill="#e2e8f0"/>
      <rect x="230" y="196" width="210" height="22" rx="11" fill="#e2e8f0"/>
      <text x="138" y="184" text-anchor="middle" font-family="Arial, sans-serif" font-size="48" font-weight="700" fill="#ffffff">${label}</text>
    </svg>
  `;

  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`;
}

function isConfigured() {
  return DRIVE_FOLDER_ID !== "ここにフォルダID" && API_KEY !== "ここにAPIキー";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function openFile(event, url) {
  event.preventDefault();

  if (!url) {
    return;
  }

  window.open(url, "_blank", "noopener");
}
