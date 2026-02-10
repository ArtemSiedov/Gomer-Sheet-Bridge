function appendFields(sheetUrl, fields) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(
      { type: "APPEND_FIELDS", sheetUrl, fields },
      (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        if (!response || !response.ok) {
          reject(new Error(response && response.error ? response.error : "Ошибка добавления"));
          return;
        }
        resolve(response.result);
      }
    );
  });
}

function getSourceIdFromUrl(url) {
  if (!url) url = window.location.href;
  const u = typeof url === "string" ? new URL(url) : url;
  const pathMatch = u.pathname.match(/\/source\/(\d+)/);
  let id = (pathMatch ? pathMatch[1] : null) ||
    u.searchParams.get("params[syncSourceId]") ||
    u.searchParams.get("syncSourceId");
  if (!id && u.hash && u.hash.includes("?")) {
    const hashQuery = u.hash.replace(/^#/, "").split("?")[1] || "";
    const hashParams = new URLSearchParams(hashQuery);
    id = hashParams.get("params[syncSourceId]") || hashParams.get("syncSourceId");
  }
  return id || "";
}

function getCategoryIdFromCurrentPage() {
  const u = new URL(window.location.href);
  let id = u.searchParams.get("rz_category_id") || u.searchParams.get("sync_source_category_id") || "";
  if (!id && u.hash && u.hash.includes("?")) {
    const hashParams = new URLSearchParams(u.hash.replace(/^#/, "").split("?")[1] || "");
    id = hashParams.get("rz_category_id") || hashParams.get("sync_source_category_id") || "";
  }
  if (id) return String(id).trim();
  if (!window.location.href.includes("/gomer/items/item-details/source/")) return "";
  return getCategoryIdForOnModerationFromItemDetails();
}

/** Категория для сбора offer IDs на on-moderation: из ссылки td:nth-child(6) > a:nth-child(1) (категория розетки). */
function getCategoryIdForOnModerationFromItemDetails() {
  if (!window.location.href.includes("/gomer/items/item-details/source/")) return "";
  const linkSelectors = [
    "#itemsDetailsPjaxContainer > div.box > div.box-header.with-border > table > tbody > tr > td:nth-child(6) > a:nth-child(1)",
    "#itemsDetailsPjaxContainer .box-header.with-border table td:nth-child(6) a",
    "#itemsDetailsPjaxContainer .box-header table td a[href*='rz_category_id']"
  ];
  for (const sel of linkSelectors) {
    const el = document.querySelector(sel);
    if (el && el.href) {
      try {
        const linkUrl = new URL(el.href, window.location.origin);
        const rzId = linkUrl.searchParams.get("rz_category_id");
        if (rzId) return String(rzId).trim();
      } catch (_) {}
      const fromText = extractIdFromText(el.textContent || "");
      if (fromText) return fromText;
    }
  }
  return "";
}

/** Категория продавца только для поиска в прайсе: из span td:nth-child(5) > span (title с id в скобках). Если span нет — пусто. */
function getCategoryIdForPriceFromItemDetails() {
  if (!window.location.href.includes("/gomer/items/item-details/source/")) return "";
  const pinSpan = document.querySelector("#itemsDetailsPjaxContainer > div.box > div.box-header.with-border > table > tbody > tr > td:nth-child(5) > span");
  if (!pinSpan) return "";
  const title = (pinSpan.getAttribute("title") || "").trim();
  const fromTitle = extractIdFromText(title);
  return fromTitle || "";
}

function getTaskIdFromCurrentPage() {
  const u = new URL(window.location.href);
  let id = u.searchParams.get("bpm_number") || u.searchParams.get("ItemSearch[bpm_number]") || "";
  if (!id && u.hash && u.hash.includes("?")) {
    const hashParams = new URLSearchParams(u.hash.replace(/^#/, "").split("?")[1] || "");
    id = hashParams.get("bpm_number") || hashParams.get("ItemSearch[bpm_number]") || "";
  }
  if (id) return String(id).trim();
  
  // Для страницы active используем специальный селектор
  if (window.location.href.includes("/gomer/items/active/source/")) {
    const taskEl = document.querySelector("#sync-sources-container > table > tbody > tr > td:nth-child(10) > a:nth-child(1)");
    if (taskEl && taskEl.textContent) {
      return (taskEl.textContent || "").trim().replace(/,/g, "");
    }
    return "";
  }
  
  // Для страницы changes используем специальный селектор
  if (window.location.href.includes("/gomer/items/changes/source/")) {
    const taskEl = document.querySelector("#sync-sources-container > table > tbody > tr > td:nth-child(14) > a:nth-child(1)");
    if (taskEl && taskEl.textContent) {
      return (taskEl.textContent || "").trim().replace(/,/g, "");
    }
    return "";
  }
  
  if (!window.location.href.includes("/gomer/items/item-details/source/")) return "";
  const taskSelectors = [
    "#itemsDetailsPjaxContainer .box-header.with-border table td:nth-child(9) a",
    "#itemsDetailsPjaxContainer .box-header table td a[href*='request']",
    ".box-header.with-border table tbody tr td:nth-child(9) a",
    "[class*='box-header'] table td a[href*='lisa']"
  ];
  for (const sel of taskSelectors) {
    const el = document.querySelector(sel);
    if (el && (el.textContent || "").trim()) return (el.textContent || "").trim().replace(/,/g, "");
  }
  return "";
}

function buildValueWithLink(value) {
  const url = new URL(window.location.href);
  let baseUrl = null;
  let productUrl = null;

  if (window.location.href.includes("/gomer/sellers/attributes/binding-attribute-page/source/")) {
    const id = url.searchParams.get("id");
    if (id) {
      baseUrl = `${url.origin}${url.pathname}?id=${id}` +
        `&rw_source_value=${encodeURIComponent(value)}&binding_type=0`;
    }
  } else if (window.location.href.includes("/gomer/items/")) {
    if (window.location.href.includes("/gomer/items/item-details/source/")) {
      productUrl = `${url.origin}${url.pathname}${url.search}`;
    }
    const hash = url.hash.startsWith("#") ? url.hash.slice(1) : url.hash;
    const sourceId = getSourceIdFromUrl(url);
    let modalId = null;
    if (hash.includes("?")) {
      const [, hashQuery] = hash.split("?");
      const hashParams = new URLSearchParams(hashQuery);
      modalId = hashParams.get("id");
    }
    if (sourceId && modalId) {
      baseUrl = `${url.origin}/gomer/sellers/attributes/binding-attribute-page/source/${sourceId}` +
        `?id=${modalId}&rw_source_value=${encodeURIComponent(value)}&binding_type=0`;
    }
  }

  if (!baseUrl) {
    return value;
  }

  // Возвращаем объект с URL и текстом для создания ссылок через Google Sheets API
  // Формат: { text: "текст", url: "url" } - будет обработан в appendFields
  return {
    valueLink: { text: value, url: baseUrl },
    productLink: productUrl ? { text: "Ссылка", url: productUrl } : ""
  };
}

// На странице со списком товаров (on-moderation и т.п.): запоминаем товар, с которого открыли модалку со значениями.
let lastModalProductKey = null;

function trackModalOpener() {
  const container = document.querySelector("#sync-sources-container");
  if (!container || container.getAttribute("data-rw-modal-track") === "1") return;
  container.setAttribute("data-rw-modal-track", "1");
  container.addEventListener(
    "click",
    function (e) {
      const link = e.target.closest('a[href*="get-binding-modal"]');
      if (!link) return;
      const row = link.closest("[data-key]");
      if (row) lastModalProductKey = row.getAttribute("data-key");
    },
    true
  );
}

const BINDING_TASK_ID_KEY = "bindingPageTaskId";
const BINDING_CATEGORY_ID_KEY = "bindingPageCategoryId";
const BINDING_SOURCE_PAGE_KEY = "bindingPageSourceType";

/** Перед переходом по ссылке на binding-attribute-page сохраняем номер заявки (делегированный клик на document). */
function setupBindingLinkTaskSave() {
  if (document.documentElement.getAttribute("data-rw-binding-track") === "1") return;
  document.documentElement.setAttribute("data-rw-binding-track", "1");
  document.addEventListener(
    "click",
    function (e) {
      if (window.location.href.includes("/gomer/sellers/attributes/binding-attribute-page/source/")) return;
      const link = e.target.closest('a[href*="binding-attribute-page"]') || e.target.closest("a.glyphicon.glyphicon-th-list");
      if (!link || !link.href || !link.href.includes("binding-attribute-page")) return;
      e.preventDefault();
      e.stopPropagation();
      let taskId = "";
      let sourceType = "on-moderation";
      if (window.location.href.includes("/gomer/items/on-moderation/source/")) {
        sourceType = "on-moderation";
        let row = link.closest("tr.sync-sources") || link.closest("tr[data-key]");
        if (!row) {
          const anyTr = link.closest("tr");
          if (anyTr && anyTr.previousElementSibling && (anyTr.previousElementSibling.matches("tr.sync-sources") || anyTr.previousElementSibling.hasAttribute("data-key"))) {
            row = anyTr.previousElementSibling;
          } else if (anyTr) {
            row = anyTr;
          }
        }
        if (row) {
          const requestCell = row.querySelector("td:nth-child(11)");
          let taskLink = requestCell ? requestCell.querySelector("a") : null;
          if (!taskLink) taskLink = row.querySelector('a[href*="lisa/#/request/view"]');
          taskId = taskLink ? (taskLink.textContent || "").trim().replace(/,/g, "") : "";
        }
      } else if (window.location.href.includes("/gomer/items/active/source/")) {
        sourceType = "active";
      } else if (window.location.href.includes("/gomer/items/changes/source/")) {
        sourceType = "changes";
      }
      if (!taskId) {
        const taskEl = document.querySelector(
          "#itemsDetailsPjaxContainer > div.box > div.box-header.with-border > table > tbody > tr > td:nth-child(9) > a"
        ) || document.querySelector('a[href*="lisa/#/request/view"]');
        taskId = taskEl ? (taskEl.textContent || "").trim().replace(/,/g, "") : "";
      }
      
      // Сохраняем категорию, если есть селектор #w1 > tbody > tr:nth-child(5) > td > div > table > tbody > tr:nth-child(2) > td:nth-child(4) > a.glyphicon.glyphicon-th-list
      let categoryId = "";
      const w1Selector = document.querySelector("#w1 > tbody > tr:nth-child(5) > td > div > table > tbody > tr:nth-child(2) > td:nth-child(4) > a.glyphicon.glyphicon-th-list");
      if (w1Selector) {
        // Пытаемся получить категорию из текущей страницы
        categoryId = getCategoryIdFromCurrentPage();
        console.log("[RW] binding: найден селектор w1, категория=", categoryId);
      }
      
      console.log("[RW] binding: сохранён номер заявки перед переходом, taskId=", taskId, "categoryId=", categoryId);
      const storageData = { [BINDING_TASK_ID_KEY]: taskId, [BINDING_SOURCE_PAGE_KEY]: sourceType };
      if (categoryId) {
        storageData[BINDING_CATEGORY_ID_KEY] = categoryId;
      }
      chrome.storage.sync.set(storageData, () => {
        window.open(link.href, "_blank");
      });
    },
    true
  );
}

function setupBindingLinkTaskSaveFromOnModeration() {
  /* Логика перенесена в setupBindingLinkTaskSave (один делегированный обработчик). */
}

/**
 * Возвращает строку товара (tr.sync-sources) по data-key.
 * nth-child не нужен: строка однозначно находится по tr.sync-sources[data-key="<itemId>"].
 */
function getCurrentProductRow(productKey) {
  const key = productKey != null ? productKey : lastModalProductKey;
  if (!key) return null;
  const tbody = document.querySelector("#sync-sources-container table tbody");
  if (!tbody) return null;
  return tbody.querySelector('tr.sync-sources[data-key="' + key + '"]');
}

/**
 * Данные текущего товара на странице on-moderation.
 * Строка берётся по data-key (не по nth-child), затем:
 * - ID заявки: td:nth-child(11) > a (fallback: ссылка с lisa/#/request/view)
 * - Категория магазина (для записи в файл): td:nth-child(8) > a (fallback: td[data-col-seq="priceCategoryId"] > a)
 */
function getCurrentProductData() {
  const row = getCurrentProductRow();
  if (!row) return null;
  const requestCell = row.querySelector("td:nth-child(11)");
  let taskLink = requestCell ? requestCell.querySelector("a") : null;
  if (!taskLink) taskLink = row.querySelector('td a[href*="lisa/#/request/view"]');
  const taskId = taskLink ? (taskLink.textContent || "").trim().replace(/,/g, "") : "";
  const categoryCell = row.querySelector("td:nth-child(8)") || row.querySelector('td[data-col-seq="priceCategoryId"]');
  const categoryLink = categoryCell ? categoryCell.querySelector("a:nth-child(1)") || categoryCell.querySelector("a") : null;
  let categoryId = "";
  let categoryText = "";
  if (categoryLink) {
    categoryText = (categoryLink.textContent || "").trim();
    const match = categoryText.match(/\((\d+)\)/);
    if (match) categoryId = match[1];
  }
  const itemId = row.getAttribute("data-key") || "";
  return { itemId, taskId, categoryId, categoryText, row };
}

function extractIdFromText(value) {
  const match = (value || "").match(/\((\d+)\)/);
  return match ? match[1] : "";
}

function extractRuValue(str) {
  const trimmed = (str || "").trim();
  const ruMatch = trimmed.match(/ru:\s*(.+?)(?=\n|uk:|$)/i);
  let result = ruMatch ? ruMatch[1].trim() : trimmed;
  // Убираем кавычки из начала и конца, если они есть
  while ((result.startsWith('"') && result.endsWith('"')) ||
         (result.startsWith("'") && result.endsWith("'"))) {
    result = result.slice(1, -1).trim();
  }
  return result;
}

/** Получает выделенный текст из поля ввода/textarea или из любого элемента через window.getSelection(). Если выделения нет, возвращает null. */
function getSelectedTextFromInput(input) {
  if (!input) return null;
  
  // Для полей ввода/textarea используем selectionStart/selectionEnd
  if (input.tagName === "TEXTAREA" || input.tagName === "INPUT") {
    const start = input.selectionStart;
    const end = input.selectionEnd;
    if (start !== null && end !== null && start !== end && start >= 0 && end > start) {
      const selected = (input.value || "").slice(start, end).trim();
      if (selected) return selected;
    }
  }
  
  // Для других элементов (например, ячеек таблицы) используем window.getSelection()
  const selection = window.getSelection();
  if (selection && selection.rangeCount > 0) {
    const selectedText = selection.toString().trim();
    if (selectedText) {
      // Проверяем, что выделение находится внутри нужного элемента
      const range = selection.getRangeAt(0);
      if (input.contains(range.commonAncestorContainer) || input === range.commonAncestorContainer || input.contains(range.startContainer)) {
        return selectedText;
      }
    }
  }
  
  return null;
}

function stripIdFromText(value) {
  return (value || "").replace(/\s*\(\d+\)\s*$/g, "").trim();
}

/** Для поиска в прайсе: убираем суффикс в скобках (например "Назначение (ListValues)" → "Назначение"). */
function stripParamNameForPrice(value) {
  const s = (value || "").trim();
  const withoutParen = s.replace(/\s*\([^)]*\)\s*$/g, "").trim();
  return withoutParen || s;
}

const ON_MODERATION_BASE = "https://gomer.rozetka.company/gomer/items/on-moderation/source";
const ACTIVE_BASE = "https://gomer.rozetka.company/gomer/items/active/source";
const CHANGES_BASE = "https://gomer.rozetka.company/gomer/items/changes/source";
const ITEMS_PER_PAGE = 500;

/** ID текущего товара на item-details (для поиска в прайсе, когда on-moderation пустой). */
function getCurrentItemOfferIdFromItemDetails() {
  if (!window.location.href.includes("/gomer/items/item-details/source/")) return "";
  const prefix = "Offer ID: ";
  
  const firstCell = document.querySelector("#itemsDetailsPjaxContainer > div.box > div.box-header.with-border > table > tbody > tr > td:nth-child(1)");
  if (!firstCell) return "";
  
  const text = (firstCell.textContent || "").trim();
  const idx = text.indexOf(prefix);
  if (idx === -1) return "";
  
  const afterPrefix = text.slice(idx + prefix.length).trim();
  // Извлекаем только цифры с начала (например "2906294461<br>RZ" -> "2906294461")
  const match = afterPrefix.match(/^\d+/);
  return match ? match[0] : "";
}

/**
 * Собирает offerId со всех страниц on-moderation по sourceId, rz_category_id и номеру заявки (bpm_number).
 * URL: .../source/{sourceId}?page=...&per-page=10&size=500&ItemSearch[bpm_number]=...&rz_category_id=...
 * 500 товаров на странице — totalPages = ceil(totalCount / 500).
 */
async function collectOfferIdsFromOnModeration(sourceId, categoryId, taskId, anchor) {
  const offerIds = [];
  const pageUrl = (page) => {
    const params = new URLSearchParams({
      page: String(page),
      "per-page": "10",
      size: "500",
      "ItemSearch[id]": "",
      "ItemSearch[name]": "",
      "ItemSearch[sync_source_vendors_id]": "",
      "ItemSearch[available]": "",
      "ItemSearch[bpm_number]": taskId || "", // Используем номер заявки для сбора offerIds с модерации
      "ItemSearch[moderation_type]": "",
      sort: "",
      sync_source_category_id: "",
      rz_category_id: categoryId || "",
      reason_id: ""
    });
    return `${ON_MODERATION_BASE}/${sourceId}?${params.toString()}`;
  };

  function parseOfferIdsFromDoc(document) {
    const ids = [];
    const prefix = "Offer ID: ";
    const rowSelectors = [
      "#sync-sources-container > table > tbody > tr",
      "#sync-sources-container table tbody tr",
      "table.items-table tbody tr",
      "table tbody tr"
    ];
    let rows = [];
    for (const sel of rowSelectors) {
      rows = document.querySelectorAll(sel);
      if (rows.length) break;
    }
    rows.forEach((tr) => {
      const td = tr.querySelector("td:nth-child(2)");
      if (!td) return;
      let id = "";
      const elWithTitle = td.querySelector("[title*='Offer ID']");
      if (elWithTitle) {
        const title = (elWithTitle.getAttribute("title") || "").trim();
        if (title.indexOf(prefix) !== -1) id = title.slice(title.indexOf(prefix) + prefix.length).trim();
      }
      if (!id) {
        const text = (td.textContent || "").trim();
        const idx = text.indexOf(prefix);
        if (idx !== -1) id = text.slice(idx + prefix.length).trim();
      }
      if (id) id = id.split(/\s/)[0] || id;
      if (id) ids.push(id);
    });
    return ids;
  }

  const firstUrl = pageUrl(1);
  const res = await fetch(firstUrl, { credentials: "include" });
  if (!res.ok) {
    console.error("[RW] collectOfferIds: 404 или ошибка, url=", firstUrl, "status=", res.status);
    throw new Error("Не удалось загрузить страницу on-moderation: " + res.status + " " + firstUrl);
  }
  const html = await res.text();
  const doc = new DOMParser().parseFromString(html, "text/html");

  offerIds.push(...parseOfferIdsFromDoc(doc));
  console.log("[RW] collectOfferIds: страница 1, собрано строк=", offerIds.length, "примеры id=", offerIds.slice(0, 3));

  const totalCountSelectors = [
    "#items_summary_container > div.summary_container_content > div.selected_summary > strong.selected_rows_total",
    "strong.selected_rows_total",
    ".selected_summary strong",
    "[class*='selected_rows_total']"
  ];
  let totalCount = 0;
  for (const sel of totalCountSelectors) {
    const totalEl = doc.querySelector(sel);
    if (totalEl) {
      totalCount = parseInt(totalEl.textContent.trim(), 10) || 0;
      if (totalCount > 0) break;
    }
  }
  const totalPages = totalCount > 0 ? Math.ceil(totalCount / 500) || 1 : 1;
  console.log("[RW] collectOfferIds: url=", pageUrl(1), "totalCount=", totalCount, "totalPages=", totalPages, "уже собрано=", offerIds.length);

  for (let page = 2; page <= totalPages; page += 1) {
    const pageU = pageUrl(page);
    const r = await fetch(pageU, { credentials: "include" });
    if (!r.ok) {
      console.warn("[RW] collectOfferIds: страница", page, "ошибка", r.status, "url=", pageU);
      break;
    }
    const h = await r.text();
    const d = new DOMParser().parseFromString(h, "text/html");
    const pageIds = parseOfferIdsFromDoc(d);
    offerIds.push(...pageIds);
    console.log("[RW] collectOfferIds: страница", page, "добавлено=", pageIds.length, "всего=", offerIds.length);
  }

  console.log("[RW] collectOfferIds: итого offerIds=", offerIds.length);
  return offerIds;
}

/**
 * Собирает offerId со страницы active по sourceId, rz_category_id и номеру заявки (bpm_number).
 * URL: .../active/source/{sourceId}?ItemSearch[bpm_number]=...&rz_category_id=...
 */
async function collectOfferIdsFromActive(sourceId, categoryId, taskId, anchor) {
  const offerIds = [];
  const pageUrl = () => {
    const params = new URLSearchParams({
      "ItemSearch[id]": "",
      "ItemSearch[name]": "",
      "ItemSearch[sync_source_vendors_id]": "",
      sort: "",
      sync_source_category_id: "",
      rz_category_id: categoryId || "",
      "ItemSearch[available]": "",
      "ItemSearch[bpm_number]": taskId || "", // Используем номер заявки для сбора offerIds с модерации
      "ItemSearch[moderation_type]": ""
    });
    return `${ACTIVE_BASE}/${sourceId}?${params.toString()}`;
  };

  function parseOfferIdsFromDoc(document) {
    const ids = [];
    const prefix = "Offer ID: ";
    const rowSelectors = [
      "#sync-sources-container > table > tbody > tr",
      "#sync-sources-container table tbody tr",
      "table.items-table tbody tr",
      "table tbody tr"
    ];
    let rows = [];
    for (const sel of rowSelectors) {
      rows = document.querySelectorAll(sel);
      if (rows.length) break;
    }
    rows.forEach((tr) => {
      const td = tr.querySelector("td:nth-child(2)");
      if (!td) return;
      let id = "";
      const elWithTitle = td.querySelector("[title*='Offer ID']");
      if (elWithTitle) {
        const title = (elWithTitle.getAttribute("title") || "").trim();
        if (title.indexOf(prefix) !== -1) id = title.slice(title.indexOf(prefix) + prefix.length).trim();
      }
      if (!id) {
        const text = (td.textContent || "").trim();
        const idx = text.indexOf(prefix);
        if (idx !== -1) id = text.slice(idx + prefix.length).trim();
      }
      if (id) id = id.split(/\s/)[0] || id;
      if (id) ids.push(id);
    });
    return ids;
  }

  const url = pageUrl();
  const res = await fetch(url, { credentials: "include" });
  if (!res.ok) {
    console.error("[RW] collectOfferIdsFromActive: ошибка, url=", url, "status=", res.status);
    throw new Error("Не удалось загрузить страницу active: " + res.status + " " + url);
  }
  const html = await res.text();
  const doc = new DOMParser().parseFromString(html, "text/html");

  offerIds.push(...parseOfferIdsFromDoc(doc));
  console.log("[RW] collectOfferIdsFromActive: url=", url, "собрано строк=", offerIds.length, "примеры id=", offerIds.slice(0, 3));

  return offerIds;
}

/**
 * Собирает offerId со страницы changes по sourceId, rz_category_id и номеру заявки (bpm_number).
 * URL: .../changes/source/{sourceId}?ItemSearch[bpm_number]=...&rz_category_id=...
 */
async function collectOfferIdsFromChanges(sourceId, categoryId, taskId, anchor, simpleMode = false) {
  const offerIds = [];
  const pageUrl = () => {
    if (simpleMode) {
      const params = new URLSearchParams({
        "ItemSearch[bpm_number]": taskId || "",
        size: "500",
        rz_category_id: categoryId || ""
      });
      return `${CHANGES_BASE}/${sourceId}?${params.toString()}`;
    }
    const params = new URLSearchParams({
      "ItemSearch[id]": "",
      "ItemSearch[name]": "",
      "ItemSearch[sync_source_vendors_id]": "",
      sort: "",
      sync_source_category_id: "",
      rz_category_id: categoryId || "",
      "ItemSearch[upload_status]": "",
      "ItemSearch[change_type]": "",
      "ItemSearch[change_status]": "",
      reason_id: "",
      "ItemSearch[available]": "",
      "ItemSearch[bpm_number]": taskId || "",
      change_date: ""
    });
    return `${CHANGES_BASE}/${sourceId}?${params.toString()}`;
  };

  function parseOfferIdsFromDoc(document) {
    const ids = [];
    const prefix = "Offer ID: ";
    const rowSelectors = [
      "#sync-sources-container > table > tbody > tr",
      "#sync-sources-container table tbody tr",
      "table.items-table tbody tr",
      "table tbody tr"
    ];
    let rows = [];
    for (const sel of rowSelectors) {
      rows = document.querySelectorAll(sel);
      if (rows.length) break;
    }
    rows.forEach((tr) => {
      const td = tr.querySelector("td:nth-child(2)");
      if (!td) return;
      let id = "";
      const elWithTitle = td.querySelector("[title*='Offer ID']");
      if (elWithTitle) {
        const title = (elWithTitle.getAttribute("title") || "").trim();
        if (title.indexOf(prefix) !== -1) id = title.slice(title.indexOf(prefix) + prefix.length).trim();
      }
      if (!id) {
        const text = (td.textContent || "").trim();
        const idx = text.indexOf(prefix);
        if (idx !== -1) id = text.slice(idx + prefix.length).trim();
      }
      if (id) id = id.split(/\s/)[0] || id;
      if (id) ids.push(id);
    });
    return ids;
  }

  const url = pageUrl();
  const res = await fetch(url, { credentials: "include" });
  if (!res.ok) {
    console.error("[RW] collectOfferIdsFromChanges: ошибка, url=", url, "status=", res.status);
    throw new Error("Не удалось загрузить страницу changes: " + res.status + " " + url);
  }
  const html = await res.text();
  const doc = new DOMParser().parseFromString(html, "text/html");

  offerIds.push(...parseOfferIdsFromDoc(doc));
  console.log("[RW] collectOfferIdsFromChanges: url=", url, "собрано строк=", offerIds.length, "примеры id=", offerIds.slice(0, 3));

  return offerIds;
}

function findOfferInPriceFeed(priceUrl, paramName, paramValue, offerIds, categoryId) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(
      {
        type: "FIND_OFFER_IN_PRICE",
        url: priceUrl,
        paramName,
        paramValue,
        offerIds: offerIds || null,
        categoryId: categoryId || null
      },
      (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        if (!response || !response.ok) {
          reject(new Error(response && response.error ? response.error : "Ошибка поиска в прайсе"));
          return;
        }
        if (response && response.debug) {
          console.log("[RW] findOfferInPriceFeed: debug", response.debug);
          if (response.debug.sampleOfferId) {
            console.log("[RW] findOfferInPriceFeed: sampleOfferId=", response.debug.sampleOfferId, "sampleParamNames=", response.debug.sampleParamNames);
          }
          if (response.debug.sampleMismatch) {
            console.log("[RW] findOfferInPriceFeed: sampleMismatch=", response.debug.sampleMismatch);
          }
        }
        resolve(response.offerId || "");
      }
    );
  });
}

async function buildProductLinkFromPrice({ categoryId, taskId, paramName, value, anchor }) {
  const isItemDetails = window.location.href.includes("/gomer/items/item-details/source/");
  const isActivePage = window.location.href.includes("/gomer/items/active/source/");
  const isChangesPage = window.location.href.includes("/gomer/items/changes/source/");
  const isListPage = window.location.href.includes("/gomer/items/") && !isItemDetails && !isActivePage && !isChangesPage;
  const isBindingPage = window.location.href.includes("/gomer/sellers/attributes/binding-attribute-page/source/");
  if (!isItemDetails && !isListPage && !isActivePage && !isChangesPage && !isBindingPage) {
    return { url: "", toast: null, found: false };
  }
  const sourceId = getSourceIdFromUrl(window.location.href);
  let taskIdToUse = (taskId || "").trim();
  if (!taskIdToUse) taskIdToUse = getTaskIdFromCurrentPage();
  const paramNameForPrice = stripParamNameForPrice(paramName || "");
  let categoryIdForOnModeration = (categoryId || "").trim();
  let categoryIdForPrice = "";
  if (isItemDetails) {
    categoryIdForOnModeration = getCategoryIdForOnModerationFromItemDetails() || categoryIdForOnModeration;
    categoryIdForPrice = getCategoryIdForPriceFromItemDetails();
  } else {
    categoryIdForOnModeration = categoryIdForOnModeration || getCategoryIdFromCurrentPage();
  }
  const needOnModeration = !categoryIdForPrice;
  if (needOnModeration && !categoryIdForOnModeration) {
    const errorToast = showToast("Укажите категорию (rz_category_id) или откройте карточку товара с категорией", anchor, "error", 0);
    return { url: "", toast: errorToast, found: false };
  }
  if (categoryIdForPrice && !categoryIdForOnModeration) categoryIdForOnModeration = categoryIdForPrice;
  console.log("[RW] buildProductLink: url=", window.location.href, "sourceId=", sourceId, "rz_category_on_moderation=", categoryIdForOnModeration, "rz_category_for_price=", categoryIdForPrice || "(нет)", "bpm_number=", taskIdToUse, "paramName=", paramNameForPrice, "value=", JSON.stringify((value || "").slice(0, 80)));
  if (!sourceId) {
    const errorToast = showToast("Не удалось найти sourceId", anchor, "error", 0);
    return { url: "", toast: errorToast, found: false };
  }

  const nav = document.querySelector("ul.nav-groupsourceHref");
  let priceLinkEl = null;
  if (nav) {
    const activeLi = nav.querySelector("li.active");
    const link = activeLi ? activeLi.querySelector("a[href]") : null;
    if (link && link.href) priceLinkEl = link;
    if (!priceLinkEl) {
      const bySource = Array.from(nav.querySelectorAll("a[href]")).find((a) => a.href && String(a.href).includes(sourceId));
      if (bySource) priceLinkEl = bySource;
    }
    if (!priceLinkEl) priceLinkEl = nav.querySelector("a[href]");
  }
  if (!priceLinkEl) {
    priceLinkEl = document.querySelector(
      "body > div.wrapper > header > nav > ul.nav.navbar-nav.btn-group.nav-group.nav-groupsourceHref > a"
    );
  }
  const priceUrl = priceLinkEl ? priceLinkEl.href : "";
  if (!priceUrl) {
    const errorToast = showToast("Не удалось найти ссылку на прайс", anchor, "error", 0);
    return { url: "", toast: errorToast, found: false };
  }
  
  // Проверка на недоступность прайса (Api virtual source)
  const priceLinkTitle = priceLinkEl ? priceLinkEl.getAttribute("title") || "" : "";
  const isPriceUnavailable = priceUrl === "javascript:void(0)" || 
                            priceUrl.includes("javascript:void(0)") ||
                            priceLinkTitle.includes("Api virtual source");
  
  if (isPriceUnavailable) {
    const errorToast = showToast("Прайс недоступен", anchor, "error", 0);
    // Если есть taskId - возвращаем его для записи в файл, но без ссылки на товар
    const taskIdForFields = taskIdToUse || "";
    return { url: "", toast: errorToast, found: false, taskId: taskIdForFields };
  }

  let loadingToast = showToast("Поиск товара", anchor, "success", 0);
    let offerIds = [];
  let categoryIdForPriceSearch = categoryIdForPrice || null;
  try {
    let sourceTypeOverride = "";
    if (isBindingPage) {
      const storageData = await new Promise(resolve => {
        chrome.storage.sync.get([BINDING_SOURCE_PAGE_KEY], resolve);
      });
      sourceTypeOverride = (storageData[BINDING_SOURCE_PAGE_KEY] || "").trim();
    }
    if (sourceTypeOverride === "active" || isActivePage) {
      offerIds = await collectOfferIdsFromActive(sourceId, categoryIdForOnModeration, taskIdToUse, anchor);
    } else if (sourceTypeOverride === "changes" || isChangesPage) {
      const useSimpleChanges = isBindingPage && sourceTypeOverride === "changes";
      offerIds = await collectOfferIdsFromChanges(sourceId, categoryIdForOnModeration, taskIdToUse, anchor, useSimpleChanges);
    } else {
      offerIds = await collectOfferIdsFromOnModeration(sourceId, categoryIdForOnModeration, taskIdToUse, anchor);
    }
    if (loadingToast && loadingToast.parentNode) loadingToast.remove();
    if (!offerIds.length && isItemDetails) {
      const currentItemId = getCurrentItemOfferIdFromItemDetails();
      if (currentItemId) {
        offerIds = [currentItemId];
        console.log("[RW] buildProductLink: on-moderation пустой, используем ID текущего товара", currentItemId);
      }
    }
    if (!offerIds.length) {
      if (categoryIdForPrice) {
        categoryIdForPriceSearch = categoryIdForPrice;
        console.log("[RW] buildProductLink: нет offerIds — ищем в прайсе по categoryId (span)=", categoryIdForPrice);
      } else {
        console.log("[RW] buildProductLink: нет offerIds после сбора");
        let pageName = "on-moderation";
        if (isActivePage) pageName = "active";
        else if (isChangesPage) pageName = "changes";
        showToast(`Нет товаров в категории на ${pageName}`, anchor, "error");
        return "";
      }
    } else {
      if (categoryIdForPrice) console.log("[RW] buildProductLink: категория продавца (span)=", categoryIdForPrice, ", ищем по offerIds + param");
      console.log("[RW] buildProductLink: priceUrl=", priceUrl, "offerIds=", offerIds, "paramName=", paramNameForPrice, "value=", JSON.stringify(value));
    }
    
    // Специальная обработка для a99.com.ua/rozetka_feed.xml - добавляем offerIds в product_ids
    let finalPriceUrl = priceUrl;
    if (priceUrl && priceUrl.includes("a99.com.ua/rozetka_feed.xml") && offerIds.length) {
      try {
        const urlObj = new URL(priceUrl);
        urlObj.searchParams.set("product_ids", offerIds.join(","));
        finalPriceUrl = urlObj.toString();
        console.log("[RW] buildProductLink: модифицирован URL для a99.com.ua, добавлены product_ids=", offerIds.join(","));
      } catch (e) {
        console.error("[RW] buildProductLink: ошибка модификации URL", e);
      }
    }
    
    loadingToast = showToast("Поиск товара", anchor, "success", 0);
    console.log("[RW] buildProductLink: перед поиском value=", JSON.stringify(value), "тип=", typeof value, "длина=", value ? value.length : 0, "первый символ=", value ? value.charCodeAt(0) : "нет", "последний символ=", value && value.length > 0 ? value.charCodeAt(value.length - 1) : "нет");
    const offerId = await findOfferInPriceFeed(finalPriceUrl, paramNameForPrice, value, offerIds.length ? offerIds : null, categoryIdForPriceSearch);
    if (!offerId) {
      console.log("[RW] buildProductLink: товар не найден в прайсе (offerId пустой)");
      if (loadingToast && loadingToast.parentNode) {
        loadingToast.textContent = "Товара с таким оффер ид и значением параметра нет в прайсе";
        loadingToast.style.background = "#c0392b";
        // Не закрываем тост — он будет переиспользован для "Добавление в файл"
      } else {
        loadingToast = showToast("Товара с таким ид заявки и значением нет в прайсе", anchor, "error", 0);
      }
      return { url: "", toast: loadingToast, found: false };
    }
    if (loadingToast && loadingToast.parentNode) loadingToast.remove();
    console.log("[RW] buildProductLink: найден offerId=", offerId);

    const params = new URLSearchParams({
      "ItemSearch[id]": offerId,
      "ItemSearch[name]": "",
      "ItemSearch[sync_source_vendors_id]": "",
      sort: "",
      sync_source_category_id: "",
      rz_category_id: "",
      "ItemSearch[upload_status]": "",
      "ItemSearch[available]": "",
      "ItemSearch[bpm_number]": "",
      "ItemSearch[moderation_type]": ""
    });

    return { url: `${window.location.origin}/gomer/items/source/${sourceId}?${params.toString()}`, toast: null, found: true };
  } catch (error) {
    if (loadingToast && loadingToast.parentNode) loadingToast.remove();
    const errorToast = showToast(error.message || "Ошибка загрузки прайса", anchor, "error", 0);
    return { url: "", toast: errorToast, found: false };
  }
}

function showToast(message, anchor, type = "success", duration = 1500) {
  const toast = document.createElement("div");
  toast.textContent = message;
  toast.style.position = "fixed";
  toast.style.zIndex = 999999;
  toast.style.background = type === "error" ? "#c0392b" : "#1f8f4d";
  toast.style.color = "#ffffff";
  toast.style.padding = "6px 10px";
  toast.style.borderRadius = "6px";
  toast.style.fontSize = "12px";
  toast.style.boxShadow = "0 2px 6px rgba(0,0,0,0.2)";

  const rect = anchor.getBoundingClientRect();
  const top = Math.max(8, Math.min(rect.top - 32, window.innerHeight - 44));
  const left = Math.max(8, Math.min(rect.left, window.innerWidth - 16));
  toast.style.top = `${top}px`;
  toast.style.left = `${left}px`;

  document.body.appendChild(toast);
  if (duration > 0) {
    setTimeout(() => toast.remove(), duration);
  }
  return toast;
}

function createBlockButton(row) {
  if (row.querySelector(".rw-sheets-btn")) {
    return;
  }

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "rw-sheets-btn";
  btn.title = "Записать в таблицу";
  btn.setAttribute("aria-label", "Записать в таблицу");
  btn.style.marginLeft = "8px";
  btn.style.width = "28px";
  btn.style.height = "28px";
  btn.style.padding = "0";
  btn.style.cursor = "pointer";
  btn.style.border = "1px solid #d0d0d0";
  btn.style.borderRadius = "4px";
  btn.style.background = "#ffffff";
  btn.innerHTML =
    '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">' +
    '<path fill="#1F7244" d="M4 3h10l6 6v12H4z"/>' +
    '<path fill="#fff" d="M14 3v6h6"/>' +
    '<path fill="#fff" d="M7 9h3l1.5 3L13 9h3l-3 6 3 6h-3l-1.5-3-1.5 3H7l3-6z"/>' +
    "</svg>";

  btn.onclick = async () => {
    const isItemDetails = window.location.href.includes("/gomer/items/item-details/source/");
    const productData = !isItemDetails ? getCurrentProductData() : null;
    const categoryEl = isItemDetails
      ? document.querySelector(
          "#itemsDetailsPjaxContainer > div.box > div.box-header.with-border > table > tbody > tr > td:nth-child(6) > a:nth-child(1)"
        ) || document.querySelector(
          "#itemsDetailsPjaxContainer > div.box > div.box-header.with-border > table > tbody > tr > td:nth-child(5) > a"
        )
      : document.querySelector("#category-title");
    const attributeEl = document.querySelector(
      "#select2-syncsourceattributebindingform-attribute_id-container"
    );
    const valueEl =
      row.querySelector("textarea") ||
      row.querySelector('input[name^="value["]') ||
      row.querySelector('input[type="text"][name*="value"]');

    if (!categoryEl && !(productData && productData.categoryText)) {
      if (!valueEl) {
        showToast("Нет поля: категория и значение", btn, "error");
        return;
      }
    }
    if (!valueEl) {
      showToast("Нет поля: textarea/input значения", btn, "error");
      return;
    }

    const rawCategory = isItemDetails && categoryEl
      ? (categoryEl.textContent || "")
      : (productData && productData.categoryText)
        ? productData.categoryText
        : (categoryEl ? (categoryEl.value || categoryEl.getAttribute("value") || "") : "");
    const category = rawCategory.trim();
    let categoryId = "";
    if (isItemDetails) {
      const categoryLink = document.querySelector("#itemsDetailsPjaxContainer > div.box > div.box-header.with-border > table > tbody > tr > td:nth-child(6) > a:nth-child(1)");
      if (categoryLink && categoryLink.href) {
        try {
          const linkUrl = new URL(categoryLink.href, window.location.origin);
          categoryId = linkUrl.searchParams.get("rz_category_id") || "";
        } catch (_) {}
        if (!categoryId) categoryId = extractIdFromText(categoryLink.textContent || "");
      }
    }
    if (!categoryId && isItemDetails && categoryEl && categoryEl.href) {
      try {
        const linkUrl = new URL(categoryEl.href, window.location.origin);
        categoryId = linkUrl.searchParams.get("rz_category_id") || "";
      } catch (_) {}
    }
    if (!categoryId) categoryId = isItemDetails ? extractIdFromText(rawCategory) : (productData ? productData.categoryId : "");
    if (!categoryId) categoryId = getCategoryIdFromCurrentPage();
    const attributeText = attributeEl ? attributeEl.textContent || "" : "";
    const attribute = attributeText.trim();
    const rawValue = valueEl.value || valueEl.getAttribute("value") || "";
    const value = extractRuValue(rawValue);
    
    // Проверяем, есть ли выделенный текст в поле значения
    const selectedValue = getSelectedTextFromInput(valueEl);
    // Для записи в файл и ссылки используем выделенный текст, если есть, иначе все значение
    const valueForFile = selectedValue || value;
    // Для поиска в прайсе всегда используем все значение (lowercase)
    const valueForPrice = (value || "").toLowerCase();

    if (!category && !attribute && !value) {
      showToast("Пустые значения — нечего добавлять", btn, "error");
      return;
    }

    if (typeof chrome === "undefined" || !chrome.storage || !chrome.storage.sync) {
      showToast("Расширение недоступно. Перезагрузите страницу.", btn, "error");
      return;
    }

    chrome.storage.sync.get("sheetUrl", async (data) => {
      if (!data.sheetUrl) {
        showToast("Сначала вставь ссылку на таблицу", btn, "error");
        return;
      }

      try {
        const timestamp = new Date().toLocaleString("ru-RU");
        const links = buildValueWithLink(valueForFile);
        const valueLink = typeof links === "string" ? links : links.valueLink;
        let productLink = typeof links === "string" ? "" : links.productLink;

        let taskId = isItemDetails
          ? (() => {
              const taskLinkEl = document.querySelector(
                "#itemsDetailsPjaxContainer > div.box > div.box-header.with-border > table > tbody > tr > td:nth-child(9) > a"
              );
              return taskLinkEl ? (taskLinkEl.textContent || "").trim().replace(/,/g, "") : "";
            })()
          : (productData ? productData.taskId : "");
        if (!taskId) taskId = getTaskIdFromCurrentPage();
        
        // Если на странице параметров и нет taskId из DOM, читаем из storage
        const isBindingPage = window.location.href.includes("/gomer/sellers/attributes/binding-attribute-page/source/");
        if (isBindingPage && !taskId) {
          const storageData = await new Promise(resolve => {
            chrome.storage.sync.get([BINDING_TASK_ID_KEY, BINDING_CATEGORY_ID_KEY], resolve);
          });
          taskId = (storageData[BINDING_TASK_ID_KEY] || "").trim();
          // Если категория не найдена в DOM, используем сохраненную из storage
          if (!categoryId && storageData[BINDING_CATEGORY_ID_KEY]) {
            categoryId = (storageData[BINDING_CATEGORY_ID_KEY] || "").trim();
            console.log("[RW] createBlockButton: на странице параметров, прочитана категория из storage, categoryId=", categoryId);
          }
          console.log("[RW] createBlockButton: на странице параметров, прочитан номер заявки из storage, taskId=", taskId);
        }

        const hasCategoryId = Boolean(categoryId);
        let taskIdForFields = taskId;
        let priceResult = null;
        if (hasCategoryId) {
          const attrTitleEl = document.querySelector("#syncsourceattribute-title");
          const attrTitleRaw = attrTitleEl
            ? (attrTitleEl.value || attrTitleEl.getAttribute("value") || "")
            : "";
          const paramName = stripIdFromText(attrTitleRaw);
          if (!paramName) {
            showToast("Не найдено название параметра (#syncsourceattribute-title)", btn, "error");
            return;
          }
          priceResult = await buildProductLinkFromPrice({
            categoryId,
            taskId,
            paramName,
            value: valueForPrice,
            anchor: btn
          });
          if (!priceResult.found) {
            // Товар не найден в прайсе или прайс недоступен
            // Если есть taskId из результата (например, когда прайс недоступен) - используем его
            // Иначе сохраняем taskId, который был установлен выше (не сбрасываем в пустую строку)
            if (priceResult.taskId) {
              taskIdForFields = priceResult.taskId;
            }
            // taskIdForFields уже установлен из taskId выше, не сбрасываем его
            productLink = ""; // Явно очищаем ссылку на товар
          } else if (priceResult.url) {
            productLink = { text: "Ссылка", url: priceResult.url };
          }
        }

        const fields = {
          "Дата добавления": timestamp,
          "Категория товара": category,
          "Атрибут": attribute,
          "Значение параметра": valueLink,
          "Номер задачи": taskIdForFields,
          "Ссылка на товар": productLink
        };
        let addToast = null;
        if (hasCategoryId && priceResult && priceResult.toast && priceResult.toast.parentNode) {
          // Переиспользуем тост из buildProductLinkFromPrice
          addToast = priceResult.toast;
          // Если товар не найден, тост уже показывает сообщение об ошибке (красный)
          // Теперь меняем на этап добавления в файл
          if (!priceResult.found) {
            // Небольшая задержка перед переходом к следующему этапу
            await new Promise(resolve => setTimeout(resolve, 500));
          }
          addToast.textContent = "Добавление значения в файл";
          addToast.style.background = "#1f8f4d";
        } else {
          addToast = showToast("Добавление значения в файл", btn, "success", 0);
        }
        const result = await appendFields(data.sheetUrl, fields);
        if (!result || !result.updates) {
          if (addToast && addToast.parentNode) {
            addToast.textContent = "Ошибка: ответ без updates";
            addToast.style.background = "#c0392b";
            setTimeout(() => {
              if (addToast.parentNode) addToast.remove();
            }, 2000);
          } else {
            showToast("Ответ без updates", btn, "error");
          }
          return;
        }

        if (addToast && addToast.parentNode) {
          addToast.textContent = "Добавление значения в файл — Добавленно";
          addToast.style.background = "#1f8f4d";
          setTimeout(() => {
            if (addToast.parentNode) addToast.remove();
          }, 1500);
        } else {
          showToast("Добавление значения в файл — Добавленно", btn);
        }
      } catch (error) {
        showToast(error.message || "Ошибка авторизации", btn, "error");
      }
    });
  };

  const valueContainer = row.querySelector(".col-md-6");
  if (valueContainer) {
    valueContainer.style.display = "flex";
    valueContainer.style.alignItems = "center";
    const input = valueContainer.querySelector("textarea") || valueContainer.querySelector('input[name^="value["]') || valueContainer.querySelector('input[type="text"]');
    if (input) {
      input.style.width = "320px";
    }
    valueContainer.appendChild(btn);
  } else {
    row.appendChild(btn);
  }
}

function addButtonsToBlocks() {
  const rows = document.querySelectorAll("#w0 > div.row.form-group");
  rows.forEach((row) => createBlockButton(row));
}

function createTableButton(cell) {
  if (cell.querySelector(".rw-sheets-btn")) {
    return;
  }

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "rw-sheets-btn";
  btn.title = "Записать в таблицу";
  btn.setAttribute("aria-label", "Записать в таблицу");
  btn.style.marginLeft = "8px";
  btn.style.width = "28px";
  btn.style.height = "28px";
  btn.style.padding = "0";
  btn.style.cursor = "pointer";
  btn.style.border = "1px solid #d0d0d0";
  btn.style.borderRadius = "4px";
  btn.style.background = "#ffffff";
  btn.innerHTML =
    '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">' +
    '<path fill="#1F7244" d="M4 3h10l6 6v12H4z"/>' +
    '<path fill="#fff" d="M14 3v6h6"/>' +
    '<path fill="#fff" d="M7 9h3l1.5 3L13 9h3l-3 6 3 6h-3l-1.5-3-1.5 3H7l3-6z"/>' +
    "</svg>";

  btn.onclick = async () => {
    const categoryEl = document.querySelector(
      "#pv_id_6_content > div > div > div:nth-child(2) > div:nth-child(1) > span"
    );
    const attributeEl = document.querySelector("#pv_id_7 > span > div");
    const attributeFileEl = document.querySelector("#pv_id_8 > span > div");
    const valueEl = cell.querySelector("div > div");

    if (!categoryEl || !attributeEl || !attributeFileEl || !valueEl) {
      const missing = [];
      if (!categoryEl) missing.push("#pv_id_6_content ... span");
      if (!attributeEl) missing.push("#pv_id_7 > span > div");
      if (!attributeFileEl) missing.push("#pv_id_8 > span > div");
      if (!valueEl) missing.push("td:nth-child(2) > div > div");
      showToast(`Нет поля: ${missing.join(", ")}`, btn, "error");
      return;
    }

    const category = (categoryEl.textContent || "").trim();
    const categoryId = extractIdFromText(category);
    const attribute = (attributeFileEl.textContent || "").trim();
    let paramName = stripParamNameForPrice(stripIdFromText(attribute));
    // Берем значение из value или textContent и обрабатываем через extractRuValue
    const rawValue = valueEl.value || valueEl.textContent || valueEl.getAttribute("value") || "";
    const value = extractRuValue(rawValue);
    console.log("[RW] createTableButton: rawValue=", JSON.stringify(rawValue), "value=", JSON.stringify(value));
    
    // Проверяем, есть ли выделенный текст в ячейке значения
    const selectedValue = getSelectedTextFromInput(valueEl);
    // Для записи в файл и ссылки используем выделенный текст, если есть, иначе все значение
    const valueForFile = selectedValue || value;
    // Для поиска в прайсе всегда используем все значение (lowercase)
    const valueForPrice = (value || "").toLowerCase();

    if (!category && !attribute && !value) {
      showToast("Пустые значения — нечего добавлять", btn, "error");
      return;
    }

    if (typeof chrome === "undefined" || !chrome.storage || !chrome.storage.sync) {
      showToast("Расширение недоступно. Перезагрузите страницу.", btn, "error");
      return;
    }

    chrome.storage.sync.get(["sheetUrl", BINDING_TASK_ID_KEY, BINDING_CATEGORY_ID_KEY], async (data) => {
      if (!data.sheetUrl) {
        showToast("Сначала вставь ссылку на таблицу", btn, "error");
        return;
      }

      try {
        const timestamp = new Date().toLocaleString("ru-RU");
        const links = buildValueWithLink(valueForFile);
        const valueLink = typeof links === "string" ? links : links.valueLink;
        let productLink = typeof links === "string" ? "" : links.productLink;
        let taskIdForFields = (data[BINDING_TASK_ID_KEY] || "").trim();
        console.log("[RW] createTableButton: прочитан номер заявки из storage, taskId=", taskIdForFields, "data[BINDING_TASK_ID_KEY]=", data[BINDING_TASK_ID_KEY]);
        
        // Если категория не найдена в DOM, используем сохраненную из storage
        let finalCategoryId = categoryId;
        if (!finalCategoryId && data[BINDING_CATEGORY_ID_KEY]) {
          finalCategoryId = (data[BINDING_CATEGORY_ID_KEY] || "").trim();
          console.log("[RW] createTableButton: прочитана категория из storage, categoryId=", finalCategoryId);
        }
        
        let priceResult = null;
        if (finalCategoryId) {
          // На странице binding-attribute-page иногда нужен "реальный" paramName
          if (window.location.href.includes("/gomer/sellers/attributes/binding-attribute-page/source/")) {
            const attrFromBinding = document.querySelector("#pv_id_7 > span > div");
            const attrText = attrFromBinding ? (attrFromBinding.textContent || "").trim() : "";
            const cleaned = stripIdFromText(attrText);
            if (cleaned) {
              paramName = stripParamNameForPrice(cleaned);
            }
          }
          console.log("[RW] createTableButton: paramNameForPrice=", JSON.stringify(paramName));
          // valueForPrice уже обработан через extractRuValue выше (кавычки уже убраны)
          console.log("[RW] createTableButton: valueForPrice=", JSON.stringify(valueForPrice), "length=", valueForPrice ? valueForPrice.length : 0);
          priceResult = await buildProductLinkFromPrice({
            categoryId: finalCategoryId,
            taskId: taskIdForFields,
            paramName,
            value: valueForPrice,
            anchor: btn
          });
          if (!priceResult.found) {
            // Товар не найден в прайсе или прайс недоступен
            // Если есть taskId из результата (например, когда прайс недоступен) - используем его
            // Иначе сохраняем taskId, который был прочитан из storage (не сбрасываем в пустую строку)
            if (priceResult.taskId) {
              taskIdForFields = priceResult.taskId;
            }
            // taskIdForFields уже установлен из storage выше, не сбрасываем его
            productLink = ""; // Явно очищаем ссылку на товар
          } else if (priceResult.url) {
            productLink = { text: "Ссылка", url: priceResult.url };
          }
        }

        const fields = {
          "Дата добавления": timestamp,
          "Категория товара": category,
          "Атрибут": attribute,
          "Значение параметра": valueLink,
          "Номер задачи": taskIdForFields,
          "Ссылка на товар": productLink
        };
        let addToast = null;
        if (priceResult && priceResult.toast && priceResult.toast.parentNode) {
          // Переиспользуем тост из buildProductLinkFromPrice
          addToast = priceResult.toast;
          // Если товар не найден, тост уже показывает сообщение об ошибке (красный)
          // Теперь меняем на этап добавления в файл
          if (!priceResult.found) {
            // Небольшая задержка перед переходом к следующему этапу
            await new Promise(resolve => setTimeout(resolve, 500));
          }
          addToast.textContent = "Добавление значения в файл";
          addToast.style.background = "#1f8f4d";
        } else {
          addToast = showToast("Добавление значения в файл", btn, "success", 0);
        }
        const result = await appendFields(data.sheetUrl, fields);
        if (!result || !result.updates) {
          if (addToast && addToast.parentNode) {
            addToast.textContent = "Ошибка: ответ без updates";
            addToast.style.background = "#c0392b";
            setTimeout(() => {
              if (addToast.parentNode) addToast.remove();
            }, 2000);
          } else {
            showToast("Ответ без updates", btn, "error");
          }
          return;
        }

        if (addToast && addToast.parentNode) {
          addToast.textContent = "Добавление значения в файл — Добавленно";
          addToast.style.background = "#1f8f4d";
          setTimeout(() => {
            if (addToast.parentNode) addToast.remove();
          }, 1500);
        } else {
          showToast("Добавление значения в файл — Добавленно", btn);
        }
      } catch (error) {
        showToast(error.message || "Ошибка авторизации", btn, "error");
      }
    });
  };

  cell.style.display = "flex";
  cell.style.alignItems = "center";
  cell.appendChild(btn);
}

function addButtonsToTable() {
  const cells = document.querySelectorAll(
    "#gomer-app > div > div.p-datatable.p-component.p-datatable-responsive-scroll.p-datatable-gridlines > div.p-datatable-wrapper > table > tbody > tr > td:nth-child(2)"
  );
  cells.forEach((cell) => createTableButton(cell));
}

function observeTableAjax() {
  const wrapper = document.querySelector(
    "#gomer-app > div > div.p-datatable.p-component.p-datatable-responsive-scroll.p-datatable-gridlines > div.p-datatable-wrapper"
  );
  if (!wrapper) {
    return;
  }
  let scheduled = false;
  const observer = new MutationObserver(() => {
    if (scheduled) {
      return;
    }
    scheduled = true;
    setTimeout(() => {
      scheduled = false;
      addButtonsToTable();
    }, 200);
  });
  observer.observe(wrapper, { childList: true, subtree: true });
}


function applyFilterFromUrl() {
  if (!window.location.href.includes("/gomer/sellers/attributes/binding-attribute-page/source/")) {
    return;
  }
  const filterApplied = document.body.getAttribute("data-rw-filter-applied") === "1";
  if (filterApplied) {
    return;
  }

  const url = new URL(window.location.href);
  const filterValue = url.searchParams.get("rw_source_value");
  const applyValueFilter = () => {
    if (!filterValue) {
      return false;
    }
    const input = document.querySelector(
      "#gomer-app > div > div.p-datatable.p-component.p-datatable-responsive-scroll.p-datatable-gridlines > div.p-datatable-wrapper > table > thead > tr:nth-child(2) > th:nth-child(2) > div > div > input"
    );
    if (!input) {
      return false;
    }
    input.focus();
    input.value = filterValue;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    input.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", bubbles: true }));
    return true;
  };

  const appliedFilter = applyValueFilter();
  if (appliedFilter) {
    document.body.setAttribute("data-rw-filter-applied", "1");
  }
}

function observePagination() {
  const container = document.querySelector("#bindingValuesPage");
  if (!container) {
    return;
  }
  let scheduled = false;
  const observer = new MutationObserver(() => {
    if (scheduled) {
      return;
    }
    scheduled = true;
    setTimeout(() => {
      scheduled = false;
      addButtonsToBlocks();
    }, 200);
  });
  observer.observe(container, { childList: true, subtree: true });
}

function initButtons() {
  addButtonsToBlocks();
  observePagination();
  trackModalOpener();
  setupBindingLinkTaskSave();
  setupBindingLinkTaskSaveFromOnModeration();
  if (window.location.href.includes("/gomer/sellers/attributes/binding-attribute-page/source/")) {
    addButtonsToTable();
    observeTableAjax();
  }
}

const initObserver = new MutationObserver(() => {
  initButtons();
  applyFilterFromUrl();
});
initObserver.observe(document.documentElement, { childList: true, subtree: true });

setTimeout(() => {
  initButtons();
  applyFilterFromUrl();
  trackModalOpener();
}, 1000);
