// Счетчики для логирования (используются в offerMatchesIdsAndParam)
let _offerMatchLogCount = 0;
let _paramNotFoundLogCount = 0;
let _lastOfferDebug = null;

function initOfferDebug() {
  return {
    offerIdMatched: 0,
    paramFound: 0,
    paramNotFound: 0,
    valueMatched: 0,
    valueMismatched: 0,
    sampleOfferId: "",
    sampleParamNames: [],
    sampleMismatch: null
  };
}

function getSpreadsheetId(url) {
  const match = url.match(/\/d\/(.*?)\//);
  return match ? match[1] : null;
}

function getSheetGid(url) {
  const match = url.match(/[?#]gid=(\d+)/);
  return match ? Number(match[1]) : null;
}

function getAuthToken() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve(token);
    });
  });
}

async function getSheetInfo(token, spreadsheetId, gid) {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties`,
    {
      headers: {
        Authorization: "Bearer " + token
      }
    }
  );
  if (!response.ok) {
    return null;
  }
  const data = await response.json();
  const sheets = data.sheets || [];
  let sheet = null;
  if (gid !== null && !Number.isNaN(gid)) {
    sheet = sheets.find((item) => item.properties && item.properties.sheetId === gid) || null;
  }
  if (!sheet) {
    sheet = sheets.length > 0 ? sheets[0] : null;
  }
  return sheet && sheet.properties
    ? { title: sheet.properties.title, sheetId: sheet.properties.sheetId }
    : null;
}

function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i += 1) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

function parseRangeToGrid(updatedRange) {
  const match = updatedRange.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  if (!match) {
    return null;
  }
  const startCol = columnLetterToIndex(match[1]);
  const startRow = Number(match[2]) - 1;
  const endCol = columnLetterToIndex(match[3]) + 1;
  const endRow = Number(match[4]);
  return {
    startRowIndex: startRow,
    endRowIndex: endRow,
    startColumnIndex: startCol,
    endColumnIndex: endCol
  };
}

async function clearRowFill(token, spreadsheetId, sheetId, updatedRange) {
  const gridRange = parseRangeToGrid(updatedRange);
  if (!gridRange) {
    return;
  }
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        requests: [
          {
            repeatCell: {
              range: {
                sheetId,
                ...gridRange
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 1, green: 1, blue: 1 }
                }
              },
              fields: "userEnteredFormat.backgroundColor"
            }
          }
        ]
      })
    }
  );
  if (!response.ok) {
    // ignore formatting errors
  }
}

const LINK_COLUMN_NAMES = ["Значение параметра", "Ссылка на товар"];
const BLUE = { red: 0.0, green: 0.48, blue: 1.0 };

async function formatLinkCellsInRow(token, spreadsheetId, sheetId, updatedRange, headerMap, linkData = new Map()) {
  const grid = parseRangeToGrid(updatedRange);
  if (!grid) return;
  const requests = [];
  for (const name of LINK_COLUMN_NAMES) {
    const colIndex = headerMap.get(name);
    if (colIndex == null) continue;
    // Форматируем ячейки со ссылками (синий цвет, подчеркивание)
    // Формула уже вставлена через values API, поэтому только форматируем
    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: grid.startRowIndex,
          endRowIndex: grid.endRowIndex,
          startColumnIndex: colIndex,
          endColumnIndex: colIndex + 1
        },
        cell: {
          userEnteredFormat: {
            textFormat: {
              foregroundColor: BLUE,
              underline: true
            }
          }
        },
        fields: "userEnteredFormat.textFormat(foregroundColor,underline)"
      }
    });
  }
  if (requests.length === 0) return;
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ requests })
    }
  );
  if (!response.ok) {
    // ignore formatting errors
  }
}

function stripCdata(s) {
  if (typeof s !== "string") return s;
  const raw = s.trim();
  const cdataStart = raw.indexOf("<![CDATA[");
  if (cdataStart === -1) return raw;
  const innerStart = cdataStart + 9;
  const cdataEnd = raw.indexOf("]]>", innerStart);
  if (cdataEnd === -1) return raw;
  return raw.slice(innerStart, cdataEnd).trim();
}

function decodeXmlEntities(value) {
  if (typeof value !== "string") return value;
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ")
    .replace(/&#(\d+);/g, (_, code) => {
      const n = Number(code);
      return Number.isFinite(n) ? String.fromCharCode(n) : _;
    })
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => {
      const n = Number.parseInt(hex, 16);
      return Number.isFinite(n) ? String.fromCharCode(n) : _;
    });
}

function extractComparableParamValue(rawValue) {
  const src = String(rawValue || "");
  if (!src) return "";

  const pickTaggedValue = (lang) => {
    const re = new RegExp(`<value\\s+[^>]*lang=(["'])${lang}\\1[^>]*>([\\s\\S]*?)<\\/value>`, "i");
    const m = src.match(re);
    return m ? m[2] : "";
  };

  const ru = pickTaggedValue("ru");
  if (ru) return ru;
  const uk = pickTaggedValue("uk");
  if (uk) return uk;
  const anyValueMatch = src.match(/<value(?:\s+[^>]*)?>([\s\S]*?)<\/value>/i);
  if (anyValueMatch && anyValueMatch[1]) return anyValueMatch[1];

  // plain <param>text</param> fallback
  return src.replace(/<[^>]+>/g, " ");
}

function normalizeParamValue(value) {
  let v = extractComparableParamValue(value);
  v = stripCdata(v);
  v = decodeXmlEntities(v);
  v = v
    .replace(/\u00A0/g, " ")
    .replace(/[‐‑‒–—]/g, "-")
    .replace(/℃/g, "°c")
    .replace(/°\s*c/gi, "°c")
    .replace(/\s*-\s*/g, " - ")
    .replace(/\s+/g, " ")
    .trim();
  if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) {
    v = v.slice(1, -1).trim();
  }
  return v.toLowerCase();
}

// Поиск оффера по categoryId и параметру (имя + значение). Потоковое чтение XML.
function offerMatchesCategoryAndParam(offerXml, categoryId, paramName, paramValue) {
  const catStr = String(categoryId);
  const hasCategory =
    offerXml.includes("categoryId=\"" + catStr + "\"") ||
    offerXml.includes("<categoryId>" + catStr + "</categoryId>");
  if (!hasCategory) return false;

  const nameIdx = offerXml.indexOf("name=\"" + paramName + "\"");
  if (nameIdx === -1) return false;
  const afterName = offerXml.slice(nameIdx);
  const startVal = afterName.indexOf(">");
  if (startVal === -1) return false;
  const endTag = afterName.indexOf("</param>");
  if (endTag === -1) return false;
  const value = afterName.slice(startVal + 1, endTag).trim();
  return normalizeParamValue(value) === normalizeParamValue(paramValue);
}

function extractOfferIdFromXmlTag(offerXml) {
  const m = offerXml.match(/<offer\b[^>]*\bid=(["'])([^"']+)\1/i);
  return m ? String(m[2] || "").trim() : "";
}

// Проверка: оффер в списке offerIds и имеет параметр paramName = paramValue.
function offerMatchesIdsAndParam(offerXml, offerIdsSet, paramName, paramValue) {
  // Сначала быстро проверяем ID (это быстрее чем парсить параметры)
  const offerId = extractOfferIdFromXmlTag(offerXml);
  if (!offerId) return false;
  const idNorm = offerId.toLowerCase();
  if (!offerIdsSet.has(idNorm)) return false; // Ранний выход если ID не в списке

  if (_lastOfferDebug) _lastOfferDebug.offerIdMatched += 1;

  // Логируем, что товар найден по offerId (первые 10 раз)
  if (_offerMatchLogCount < 10) {
    _offerMatchLogCount++;
    console.log("[RW] ТОВАР НАЙДЕН по offerId:", offerId);
  }
  
  // Только если ID совпал, ищем параметр
  const nameKey = (paramName || "").trim();
  if (!nameKey) return false;

  const normalizeName = (s) =>
    (s || "")
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  
  // Ищем параметр с нужным именем (оптимизированный поиск)
  const namePattern1 = `name="${nameKey}"`;
  const namePattern2 = `name='${nameKey}'`;
  let nameIdx = offerXml.indexOf(namePattern1);
  if (nameIdx === -1) nameIdx = offerXml.indexOf(namePattern2);
  if (nameIdx === -1) {
    // Фоллбек: пробуем найти параметр по нормализованному имени
    const wanted = normalizeName(nameKey);
    const paramRegex = /<param\s+name=(["'])(.*?)\1[^>]*>([\s\S]*?)<\/param>/g;
    let match;
    let foundParamValue = null;
    let foundParamName = null;
    while ((match = paramRegex.exec(offerXml)) !== null) {
      const candidateName = match[2];
      if (normalizeName(candidateName) === wanted) {
        foundParamName = candidateName;
        foundParamValue = match[3];
        break;
      }
    }
    if (foundParamValue != null) {
      if (_paramNotFoundLogCount < 10) {
        _paramNotFoundLogCount++;
        console.log("[RW] ПАРАМЕТР НАЙДЕН (норм.): offerId=", offerId, "paramName=", nameKey, "matchedName=", foundParamName);
      }
      if (_lastOfferDebug) _lastOfferDebug.paramFound += 1;

      const value = normalizeParamValue(foundParamValue);
      const searchValue = normalizeParamValue(paramValue);
      const matches = value === searchValue;
      if (matches) {
        console.log("[RW] ЗНАЧЕНИЕ СОВПАЛО:", "offerId=", offerId, "paramName=", nameKey, "value=", JSON.stringify(value));
        if (_lastOfferDebug) _lastOfferDebug.valueMatched += 1;
      } else {
        if (_offerMatchLogCount < 20) {
          _offerMatchLogCount++;
          console.log("[RW] ЗНАЧЕНИЕ НЕ СОВПАЛО:", "offerId=", offerId, "paramName=", nameKey, "ищем=", JSON.stringify(searchValue), "найдено=", JSON.stringify(value));
        }
        if (_lastOfferDebug) _lastOfferDebug.valueMismatched += 1;
      }
      return matches;
    }

    // Логируем первые несколько случаев для диагностики
    if (_paramNotFoundLogCount < 5) {
      _paramNotFoundLogCount++;
      // Извлекаем все параметры из XML для диагностики
      const paramMatches = offerXml.matchAll(/<param\s+name=(["'])(.*?)\1[^>]*>/g);
      const paramNames = Array.from(paramMatches, m => m[2]);
      console.log("[RW] ПАРАМЕТР НЕ НАЙДЕН в товаре:", offerId, "paramName=", nameKey, "найденные параметры:", paramNames);
    }
    if (_lastOfferDebug && !_lastOfferDebug.sampleOfferId) {
      const paramMatches = offerXml.matchAll(/<param\s+name=(["'])(.*?)\1[^>]*>/g);
      const paramNames = Array.from(paramMatches, m => m[2]);
      _lastOfferDebug.sampleOfferId = offerId;
      _lastOfferDebug.sampleParamNames = paramNames.slice(0, 30);
    }
    if (_lastOfferDebug) _lastOfferDebug.paramNotFound += 1;
    return false;
  }

  // Логируем, что параметр найден (первые 10 раз)
  if (_paramNotFoundLogCount < 10) {
    _paramNotFoundLogCount++;
    console.log("[RW] ПАРАМЕТР НАЙДЕН в товаре:", offerId, "paramName=", nameKey);
  }
  if (_lastOfferDebug) _lastOfferDebug.paramFound += 1;
  
  // Извлекаем значение параметра
  const afterName = offerXml.slice(nameIdx);
  const startVal = afterName.indexOf(">");
  if (startVal === -1) return false;
  const endTag = afterName.indexOf("</param>", startVal);
  if (endTag === -1) return false;
  
  const rawValue = afterName.slice(startVal + 1, endTag).trim();
  const value = normalizeParamValue(rawValue);
  const searchValue = normalizeParamValue(paramValue);
  const matches = value === searchValue;
  if (matches) {
    console.log("[RW] ЗНАЧЕНИЕ СОВПАЛО:", "offerId=", offerId, "paramName=", nameKey, "value=", JSON.stringify(value));
    if (_lastOfferDebug) _lastOfferDebug.valueMatched += 1;
  }
  // Логируем несовпадения для диагностики (первые 20 случаев)
  if (!matches) {
    if (_offerMatchLogCount < 20) {
      _offerMatchLogCount++;
      console.log("[RW] ЗНАЧЕНИЕ НЕ СОВПАЛО:", "offerId=", offerId, "paramName=", nameKey, "ищем=", JSON.stringify(searchValue), "найдено=", JSON.stringify(value));
    }
    if (_lastOfferDebug && !_lastOfferDebug.sampleMismatch) {
      _lastOfferDebug.sampleMismatch = {
        offerId,
        paramName: nameKey,
        expected: String(searchValue),
        found: String(value)
      };
    }
    if (_lastOfferDebug) _lastOfferDebug.valueMismatched += 1;
  }
  
  return matches;
}

// Fallback: проверяем только совпадение значения в любом <param> у оффера из списка offerIds.
function offerMatchesIdsAndAnyParamValue(offerXml, offerIdsSet, paramValue) {
  const offerId = extractOfferIdFromXmlTag(offerXml);
  if (!offerId) return false;
  const idNorm = offerId.toLowerCase();
  if (!offerIdsSet.has(idNorm)) return false;

  const searchValue = normalizeParamValue(paramValue);
  const paramRegex = /<param\s+name=(["'])(.*?)\1[^>]*>([\s\S]*?)<\/param>/g;
  let match = null;
  while ((match = paramRegex.exec(offerXml)) !== null) {
    const raw = match[3] || "";
    const val = normalizeParamValue(raw);
    if (val === searchValue) {
      return true;
    }
  }
  return false;
}

function getOfferIdFromXml(offerXml) {
  const id = extractOfferIdFromXmlTag(offerXml);
  return id || null;
}

/** Поиск в указанном диапазоне байт. */
async function findOfferInRange(url, offerIdsSet, paramName, paramValue, startByte, endByte, signal) {
  if (endByte < startByte) return null;
  const response = await fetch(url, {
    headers: {
      Range: `bytes=${startByte}-${endByte}`
    },
    signal
  });
  if (!response.ok && response.status !== 206) {
    return null; // Range не поддерживается или ошибка
  }
  if (!response.body) return null;

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  const MAX_TAIL = 50000;
  let isFirstChunk = true; // Флаг для пропуска неполного тега в начале диапазона
  let offersChecked = 0;
  let offersWithMatchingId = 0;

  try {
    let eof = false;
    while (true) {
      const { done, value } = await reader.read();
      if (done) {
        eof = true;
        break;
      }
      buffer += decoder.decode(value, { stream: true });

      // Если это первый чанк и мы начали с середины файла, пропускаем неполный тег
      if (isFirstChunk && startByte > 0) {
        // Ищем первый полный тег <offer> (начинается с <offer и заканчивается </offer>)
        const firstOfferStart = buffer.indexOf("<offer");
        if (firstOfferStart !== -1) {
          // Проверяем, есть ли закрывающий тег после этого
          const firstOfferEnd = buffer.indexOf("</offer>", firstOfferStart);
          if (firstOfferEnd === -1) {
            // Если нет закрывающего тега, ищем начало следующего полного тега
            // Пропускаем все до следующего <offer
            const nextOffer = buffer.indexOf("<offer", firstOfferStart + 1);
            if (nextOffer !== -1) {
              buffer = buffer.slice(nextOffer);
            } else {
              // Если следующего <offer нет, оставляем буфер как есть для следующей итерации
            }
          }
        }
        isFirstChunk = false;
      }

      for (;;) {
        const start = buffer.indexOf("<offer");
        if (start === -1) break;
        const end = buffer.indexOf("</offer>", start);
        if (end === -1) {
          if (eof) break;
          break;
        }

        const offerXml = buffer.slice(start, end + 8);
        buffer = buffer.slice(end + 8);
        offersChecked++;

        // Проверяем ID
        const parsedOfferId = extractOfferIdFromXmlTag(offerXml);
        if (parsedOfferId) {
          const idNorm = parsedOfferId.toLowerCase();
          if (offerIdsSet.has(idNorm)) {
            offersWithMatchingId++;
            // Если ID совпал, проверяем параметр
            if (offerMatchesIdsAndParam(offerXml, offerIdsSet, paramName, paramValue)) {
              const foundId = getOfferIdFromXml(offerXml);
              console.log("[RW] findOfferInRange: найден offerId=", foundId, "диапазон", startByte, "-", endByte, "проверено офферов=", offersChecked, "с совпадающим ID=", offersWithMatchingId);
              reader.cancel();
              return foundId;
            }
          }
        }
      }

      if (buffer.length > MAX_TAIL) buffer = buffer.slice(-MAX_TAIL);
      if (eof) break;
    }
  } finally {
    try {
      reader.releaseLock();
    } catch (_) {}
  }

  // Логируем только если были совпадения по ID, но не по параметру
  if (offersWithMatchingId > 0) {
    console.log("[RW] findOfferInRange: диапазон", startByte, "-", endByte, "проверено офферов=", offersChecked, "с совпадающим ID=", offersWithMatchingId, "но значение параметра не совпало");
  }
  
  return null;
}

/** Поиск с центра отключен (используем обычный последовательный поиск). */
async function findOfferByOfferIdsAndParamFromCenter(url, offerIdsSet, paramName, paramValue) {
  return await findOfferByOfferIdsAndParamSequential(url, offerIdsSet, paramName, paramValue);
}

/** Обычный последовательный поиск (fallback). */
async function findOfferByOfferIdsAndParamSequential(url, offerIdsSet, paramName, paramValue) {
  const response = await fetch(url);
  if (!response.ok || !response.body) {
    if (!response.ok) console.error("[RW] findOfferByOfferIds: прайс вернул ошибку, url=", url, "status=", response.status);
    throw new Error(response.ok ? "Streaming not supported" : "HTTP " + response.status);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  const MAX_TAIL = 50000;
  const MIN_BUFFER = 10485760; // Буфер 10 МБ

  try {
    let eof = false;
    while (true) {
      while (!eof) {
        const { done, value } = await reader.read();
        if (done) {
          eof = true;
          break;
        }
        buffer += decoder.decode(value, { stream: true });
        
        if (buffer.length >= MIN_BUFFER || buffer.includes("</offer>")) {
          break;
        }
      }

      for (;;) {
        const start = buffer.indexOf("<offer");
        if (start === -1) break;
        const end = buffer.indexOf("</offer>", start);
        if (end === -1) {
          if (eof) break;
          break;
        }

        const offerXml = buffer.slice(start, end + 8);
        buffer = buffer.slice(end + 8);

        if (offerMatchesIdsAndParam(offerXml, offerIdsSet, paramName, paramValue)) {
          const foundId = getOfferIdFromXml(offerXml);
          console.log("[RW] findOfferByOfferIds: найден offerId=", foundId);
          reader.cancel();
          return foundId;
        }
      }

      if (buffer.length > MAX_TAIL) buffer = buffer.slice(-MAX_TAIL);
      if (eof) break;
    }
  } finally {
    try {
      reader.releaseLock();
    } catch (_) {}
  }

  return null;
}

async function findOfferByCategoryAndParam(url, categoryId, paramName, paramValue, taskSignal) {
  const response = await fetch(url, { signal: taskSignal });
  if (!response.ok && response.status !== 206) {
    throw new Error("HTTP " + response.status + ": " + response.statusText);
  }
  if (!response.body) {
    throw new Error("Streaming not supported");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  const MAX_TAIL = 50000;
  const CHUNK_SIZE = 10485760; // Читаем порциями по 10 МБ

  try {
    let eof = false;
    while (true) {
      // Читаем данные порциями для лучшей производительности
      while (!eof && buffer.length < CHUNK_SIZE) {
        const { done, value } = await reader.read();
        if (done) {
          eof = true;
          break;
        }
        buffer += decoder.decode(value, { stream: true });
      }

      // Обрабатываем все полные офферы в буфере
      for (;;) {
        const start = buffer.indexOf("<offer");
        if (start === -1) break;
        const end = buffer.indexOf("</offer>", start);
        if (end === -1) {
          if (eof) break; // Если конец файла, обрабатываем и неполный
          break; // Неполный оффер - оставляем для следующей итерации
        }

        const offerXml = buffer.slice(start, end + 8);
        buffer = buffer.slice(end + 8);

        if (offerMatchesCategoryAndParam(offerXml, categoryId, paramName, paramValue)) {
          reader.cancel();
          const offerId = getOfferIdFromXml(offerXml);
          return offerId;
        }
      }

      if (buffer.length > MAX_TAIL) {
        buffer = buffer.slice(-MAX_TAIL);
      }
      if (eof) break;
    }
  } finally {
    try {
      reader.releaseLock();
    } catch (_) {}
  }

  return null;
}

async function findOfferByOfferIdsAndParam(url, offerIds, paramName, paramValue, taskSignal) {
  if (!offerIds || !offerIds.length) return null;
  _lastOfferDebug = initOfferDebug();
  const offerIdsSet = new Set(
    offerIds.map((id) => String(id).trim().toLowerCase()).filter(Boolean)
  );
  console.log("[RW] findOfferByOfferIds: url=", url?.slice(0, 80), "offerIdsCount=", offerIds.length, "paramName=", paramName, "paramValue=", JSON.stringify((paramValue || "").slice(0, 60)));

  // Обычный последовательный поиск (без Range/параллелизма)
  const response = await fetch(url, { signal: taskSignal });
  if (!response.ok || !response.body) {
    if (!response.ok) console.error("[RW] findOfferByOfferIds: прайс вернул ошибку, url=", url, "status=", response.status);
    throw new Error(response.ok ? "Streaming not supported" : "HTTP " + response.status);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  const MAX_TAIL = 50000;
  const MIN_BUFFER = 10485760; // Буфер 10 МБ
  let offersChecked = 0;
  let offersWithMatchingId = 0;

  try {
    let eof = false;
    while (true) {
      // Читаем данные порциями, но начинаем обработку как только есть хотя бы один полный <offer>
      while (!eof) {
        const { done, value } = await reader.read();
        if (done) {
          eof = true;
          break;
        }
        buffer += decoder.decode(value, { stream: true });
        
        // Если буфер достаточно большой или есть полные офферы, начинаем обработку
        if (buffer.length >= MIN_BUFFER || buffer.includes("</offer>")) {
          break;
        }
      }

      // Обрабатываем все полные офферы в буфере
      for (;;) {
        const start = buffer.indexOf("<offer");
        if (start === -1) break;
        const end = buffer.indexOf("</offer>", start);
        if (end === -1) {
          // Неполный оффер - оставляем в буфере для следующей итерации
          if (eof) break; // Если конец файла, обрабатываем и неполный
          break;
        }

        const offerXml = buffer.slice(start, end + 8);
        buffer = buffer.slice(end + 8);
        offersChecked++;

        // Проверяем ID
        const parsedOfferId = extractOfferIdFromXmlTag(offerXml);
        if (parsedOfferId) {
          const idNorm = parsedOfferId.toLowerCase();
          if (offerIdsSet.has(idNorm)) {
            offersWithMatchingId++;
            // Если ID совпал, проверяем параметр
            if (offerMatchesIdsAndParam(offerXml, offerIdsSet, paramName, paramValue)) {
              const foundId = getOfferIdFromXml(offerXml);
              console.log("[RW] findOfferByOfferIds: найден offerId=", foundId, "проверено офферов=", offersChecked, "с совпадающим ID=", offersWithMatchingId);
              reader.cancel();
              return foundId;
            }
          }
        }
      }

      if (buffer.length > MAX_TAIL) buffer = buffer.slice(-MAX_TAIL);
      if (eof) break;
    }
  } finally {
    try {
      reader.releaseLock();
    } catch (_) {}
  }

  console.log("[RW] findOfferByOfferIds: товар не найден в прайсе. Проверено офферов=", offersChecked, "с совпадающим ID=", offersWithMatchingId, "paramName=", paramName, "paramValue=", JSON.stringify(paramValue));

  // Fallback-проход: если имя параметра не совпадает в источниках, ищем по значению
  // среди тех же offerIds.
  console.warn("[RW] findOfferByOfferIds: fallback to value-only search for matched offerIds");
  const response2 = await fetch(url, { signal: taskSignal });
  if (!response2.ok || !response2.body) {
    return null;
  }
  const reader2 = response2.body.getReader();
  const decoder2 = new TextDecoder();
  let buffer2 = "";
  const MAX_TAIL2 = 50000;
  const MIN_BUFFER2 = 10485760;
  try {
    let eof2 = false;
    while (true) {
      while (!eof2) {
        const { done, value } = await reader2.read();
        if (done) {
          eof2 = true;
          break;
        }
        buffer2 += decoder2.decode(value, { stream: true });
        if (buffer2.length >= MIN_BUFFER2 || buffer2.includes("</offer>")) {
          break;
        }
      }

      for (;;) {
        const start = buffer2.indexOf("<offer");
        if (start === -1) break;
        const end = buffer2.indexOf("</offer>", start);
        if (end === -1) {
          if (eof2) break;
          break;
        }

        const offerXml = buffer2.slice(start, end + 8);
        buffer2 = buffer2.slice(end + 8);
        if (offerMatchesIdsAndAnyParamValue(offerXml, offerIdsSet, paramValue)) {
          const foundId = getOfferIdFromXml(offerXml);
          console.log("[RW] findOfferByOfferIds: fallback matched by value-only, offerId=", foundId);
          reader2.cancel();
          return foundId;
        }
      }

      if (buffer2.length > MAX_TAIL2) buffer2 = buffer2.slice(-MAX_TAIL2);
      if (eof2) break;
    }
  } finally {
    try {
      reader2.releaseLock();
    } catch (_) {}
  }

  return null;
}

async function getHeaderMap(token, spreadsheetId, rangePrefix) {
  const headerRange = rangePrefix ? `${rangePrefix}!1:1` : "1:1";
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${headerRange}`,
    {
      headers: {
        Authorization: "Bearer " + token
      }
    }
  );
  if (!response.ok) {
    throw new Error(await response.text());
  }
  const data = await response.json();
  const headers = data.values && data.values[0] ? data.values[0] : [];
  const map = new Map();
  headers.forEach((value, index) => {
    if (value && !map.has(value)) {
      map.set(value, index);
    }
  });
  return { map, headers };
}

async function appendFields(sheetUrl, fields) {
  const spreadsheetId = getSpreadsheetId(sheetUrl);
  if (!spreadsheetId) {
    return { ok: false, error: "Ссылка на таблицу неверная" };
  }

  const token = await getAuthToken();
  const sheetGid = getSheetGid(sheetUrl);
  const sheetInfo = await getSheetInfo(token, spreadsheetId, sheetGid);
  const sheetTitle = sheetInfo ? sheetInfo.title : null;
  const safeTitle = sheetTitle ? sheetTitle.replace(/'/g, "''") : null;
  const rangePrefix = safeTitle ? `'${safeTitle}'` : null;
  const { map, headers } = await getHeaderMap(token, spreadsheetId, rangePrefix);
  const missing = Object.keys(fields).filter((name) => !map.has(name));
  if (missing.length > 0) {
    return { ok: false, error: `Не найдены колонки: ${missing.join(", ")}` };
  }

  let maxIndex = 0;
  map.forEach((index) => {
    if (index > maxIndex) {
      maxIndex = index;
    }
  });
  const row = new Array(maxIndex + 1).fill("");
  const linkData = new Map(); // colIndex -> { url, text }
  Object.entries(fields).forEach(([name, value]) => {
    const colIndex = map.get(name);
    // Проверяем, что это объект ссылки (не null, не массив, есть url и text)
    if (value && typeof value === "object" && !Array.isArray(value) && value.url && value.text && typeof value.url === "string" && typeof value.text === "string") {
      // Это ссылка в формате { text, url }
      // Создаем формулу HYPERLINK с точкой с запятой
      const safeUrl = value.url.replace(/"/g, '""');
      const safeText = value.text.replace(/"/g, '""');
      row[colIndex] = `=HYPERLINK("${safeUrl}";"${safeText}")`; // Вставляем формулу сразу
      linkData.set(colIndex, { url: value.url, text: value.text });
    } else {
      row[colIndex] = value;
    }
  });

  const range = rangePrefix ? `${encodeURIComponent(rangePrefix)}!A1:append` : "A1:append";

  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        values: [row]
      })
    }
  );

  if (!response.ok) {
    return { ok: false, error: await response.text(), status: response.status };
  }

  const result = await response.json();
  if (sheetInfo && result && result.updates && result.updates.updatedRange) {
    await clearRowFill(token, spreadsheetId, sheetInfo.sheetId, result.updates.updatedRange);
    await formatLinkCellsInRow(token, spreadsheetId, sheetInfo.sheetId, result.updates.updatedRange, map, linkData);
  }
  return { ok: true, result };
}

const EXPORT_QUEUE_STATE_KEY = "rwExportQueueState";
const EXPORT_QUEUE_DATA_KEY = "rwExportQueueData";
const EXPORT_QUEUE_MAX_ITEMS = 200;
const EXPORT_QUEUE_HISTORY_TTL_MS = 5 * 60 * 60 * 1000; // 5 часов
const OFFER_IDS_CACHE_TTL_MS = 10 * 60 * 1000; // 10 минут
let _exportQueue = [];
let _exportIsProcessing = false;
let _queueInitialized = false;
let _queueInitPromise = null;
let _activeTaskId = "";
let _activeTaskAbortController = null;
const _offerIdsCache = new Map();

function makeTaskId() {
  return `task_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function makeTaskTitle(task) {
  const attr = task?.fields?.["Атрибут"] || "Без атрибута";
  const value = task?.fields?.["Значение параметра"];
  const text = typeof value === "object" && value && value.text ? value.text : value || "";
  return `${attr}${text ? `: ${String(text).slice(0, 40)}` : ""}`;
}

function createQueueStateSnapshot() {
  pruneQueueHistory();
  const pendingCount = _exportQueue.filter((task) => task.status === "pending").length;
  const processingCount = _exportQueue.filter((task) => task.status === "processing").length;
  const doneCount = _exportQueue.filter((task) => task.status === "done").length;
  const errorCount = _exportQueue.filter((task) => task.status === "error").length;
  const items = _exportQueue.slice(-EXPORT_QUEUE_MAX_ITEMS).map((task) => ({
    id: task.id,
    status: task.status,
    title: task.title,
    message: task.message || "",
    createdAt: task.createdAt,
    updatedAt: task.updatedAt
  }));
  return {
    updatedAt: Date.now(),
    processing: _exportIsProcessing,
    pendingCount,
    processingCount,
    doneCount,
    errorCount,
    totalCount: _exportQueue.length,
    items
  };
}

function saveQueueState() {
  chrome.storage.local.set({ [EXPORT_QUEUE_STATE_KEY]: createQueueStateSnapshot() });
}

function saveQueueData() {
  pruneQueueHistory();
  chrome.storage.local.set({ [EXPORT_QUEUE_DATA_KEY]: _exportQueue.slice(-EXPORT_QUEUE_MAX_ITEMS) });
}

function pruneQueueHistory() {
  const cutoff = Date.now() - EXPORT_QUEUE_HISTORY_TTL_MS;
  _exportQueue = _exportQueue.filter((task) => {
    if (!task) return false;
    if (task.status === "pending" || task.status === "processing") return true;
    const ts = Number(task.updatedAt || task.createdAt || 0);
    return ts >= cutoff;
  });
}

function setTaskState(task, status, message) {
  task.status = status;
  if (message) task.message = message;
  task.updatedAt = Date.now();
  saveQueueState();
  saveQueueData();
}

function isTaskAborted(taskSignal) {
  return Boolean(taskSignal && taskSignal.aborted);
}

function throwIfTaskAborted(taskSignal) {
  if (!isTaskAborted(taskSignal)) return;
  throw new Error("Остановлено пользователем");
}

function parseOfferIdsFromHtml(html) {
  if (!html || typeof html !== "string") return [];
  const seen = new Set();
  let match = null;

  // Берем только значения вида "Offer ID: ...".
  // Не парсим data-key/data-id и другие числа со страницы, чтобы не ловить лишние ID.
  const offerIdLabelRegex = /Offer ID:\s*([A-Za-z0-9_-]+)/g;
  while ((match = offerIdLabelRegex.exec(html)) !== null) {
    const id = String(match[1] || "").trim();
    if (id) seen.add(id);
  }

  return Array.from(seen).filter((id) => {
    if (/^\d{5,}$/.test(id)) return true;
    return /^[A-Za-z0-9_-]{5,}$/.test(id);
  });
}

function debugOfferIdsLog(url, offerIds) {
  const count = Array.isArray(offerIds) ? offerIds.length : 0;
  const sample = Array.isArray(offerIds) ? offerIds.slice(0, 12) : [];
  console.log(
    "[RW][QUEUE] offerIds parsed:",
    "count=",
    count,
    "sample=",
    sample,
    "url=",
    String(url || "").slice(0, 220)
  );
}

function sendOfferIdsToContentConsole(task, offerIds, url) {
  const tabId = Number(task && task.senderTabId);
  if (!Number.isInteger(tabId)) return;
  try {
    chrome.tabs.sendMessage(tabId, {
      type: "RW_QUEUE_DEBUG_OFFER_IDS",
      count: Array.isArray(offerIds) ? offerIds.length : 0,
      offerIds: Array.isArray(offerIds) ? offerIds : [],
      url: String(url || "")
    });
  } catch (_) {}
}

async function fetchText(url, taskSignal) {
  const response = await fetch(url, { credentials: "include", signal: taskSignal });
  if (!response.ok) {
    throw new Error(`Ошибка загрузки: ${response.status} ${url}`);
  }
  return response.text();
}

async function fetchTextWithMeta(url, taskSignal) {
  const response = await fetch(url, { credentials: "include", signal: taskSignal });
  if (!response.ok) {
    throw new Error(`Ошибка загрузки: ${response.status} ${url}`);
  }
  const pageCountHeader = Number(response.headers.get("X-Pagination-Page-Count") || 0);
  const totalCountHeader = Number(response.headers.get("X-Pagination-Total-Count") || 0);
  const perPageHeader = Number(response.headers.get("X-Pagination-Per-Page") || 0);
  return {
    text: await response.text(),
    pageCount: Number.isFinite(pageCountHeader) && pageCountHeader > 0 ? pageCountHeader : 0,
    totalCount: Number.isFinite(totalCountHeader) && totalCountHeader > 0 ? totalCountHeader : 0,
    perPage: Number.isFinite(perPageHeader) && perPageHeader > 0 ? perPageHeader : 0
  };
}

async function pingUrl(url, taskSignal) {
  const timeoutController = new AbortController();
  const timeoutId = setTimeout(() => timeoutController.abort(), 10000);
  const composite = new AbortController();
  const abortComposite = () => composite.abort();
  timeoutController.signal.addEventListener("abort", abortComposite, { once: true });
  if (taskSignal) taskSignal.addEventListener("abort", abortComposite, { once: true });
  try {
    await fetch(url, { credentials: "include", signal: composite.signal });
  } finally {
    clearTimeout(timeoutId);
  }
}

async function restorePageSizeIfNeeded(task, taskSignal) {
  const restoreUrl = String(task && task.restorePageSizeUrl ? task.restorePageSizeUrl : "").trim();
  if (!restoreUrl) return;
  try {
    await pingUrl(restoreUrl, taskSignal);
  } catch (_) {}
}

async function withTimeout(promise, timeoutMs, timeoutMessage) {
  let timeoutId = null;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = setTimeout(() => reject(new Error(timeoutMessage)), timeoutMs);
  });
  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
  }
}

function shouldSkipPriceSearchOnError(error) {
  const msg = String(error && error.message ? error.message : error || "").toLowerCase();
  return (
    msg.includes("504") ||
    msg.includes("gateway time-out") ||
    msg.includes("gateway timeout") ||
    msg.includes("signal is aborted") ||
    msg.includes("aborted") ||
    msg.includes("timeout") ||
    msg.includes("failed to fetch") ||
    msg.includes("networkerror")
  );
}

function getMaxPageFromHtml(html) {
  if (!html || typeof html !== "string") return 1;
  let maxPage = 1;
  const pageRegex = /[?&]page=(\d+)/g;
  let match = null;
  while ((match = pageRegex.exec(html)) !== null) {
    const page = Number(match[1]);
    if (Number.isFinite(page) && page > maxPage) maxPage = page;
  }
  return maxPage;
}

function getTotalRowsFromHtml(html) {
  if (!html || typeof html !== "string") return 0;
  // Берем значение из strong.selected_rows_total (по аналогии с DOM-селектором)
  const classMatch = html.match(/<strong[^>]*class=["'][^"']*selected_rows_total[^"']*["'][^>]*>\s*([\d\s.,]+)\s*<\/strong>/i);
  if (!classMatch || !classMatch[1]) return 0;
  const normalized = String(classMatch[1]).replace(/[^\d]/g, "");
  const total = Number(normalized);
  return Number.isFinite(total) ? total : 0;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function makeOfferIdsCacheKey(task, sourceType, sourceId, categoryId, taskId) {
  const explicit = String(task && task.offerIdsCacheKey ? task.offerIdsCacheKey : "").trim();
  if (explicit) return explicit;
  return `${sourceType}|${sourceId}|${categoryId}|${taskId}`;
}

function getOfferIdsFromCache(cacheKey) {
  if (!cacheKey) return null;
  const item = _offerIdsCache.get(cacheKey);
  if (!item) return null;
  if (Date.now() - item.ts > OFFER_IDS_CACHE_TTL_MS) {
    _offerIdsCache.delete(cacheKey);
    return null;
  }
  return Array.isArray(item.offerIds) ? item.offerIds : null;
}

function setOfferIdsCache(cacheKey, offerIds) {
  if (!cacheKey || !Array.isArray(offerIds)) return;
  _offerIdsCache.set(cacheKey, { ts: Date.now(), offerIds: offerIds.slice() });
}

async function collectOfferIdsForTask(task, taskSignal) {
  const sourceId = (task.sourceId || "").trim();
  if (!sourceId) return [];
  const categoryId = (task.categoryIdForOnModeration || "").trim();
  const taskId = task.useTaskIdForOfferSelection ? (task.taskIdForOfferSelection || "").trim() : "";
  const sourceType = (task.sourceType || "on-moderation").trim();
  const cacheKey = makeOfferIdsCacheKey(task, sourceType, sourceId, categoryId, taskId);
  const cachedOfferIds = getOfferIdsFromCache(cacheKey);
  if (cachedOfferIds) {
    console.log("[RW][QUEUE] offerIds cache hit:", cacheKey, "count=", cachedOfferIds.length);
    sendOfferIdsToContentConsole(task, cachedOfferIds, "cache");
    return cachedOfferIds;
  }
  const perPage = "500";
  const makeUrl = (page) => {
    let base = "";
    const p = new URLSearchParams();
    if (categoryId) p.set("rz_category_id", categoryId);
    if (taskId) p.set("ItemSearch[bpm_number]", taskId);

    if (sourceType === "active") {
      p.set("page", String(page));
      p.set("per-page", perPage);
      base = `https://gomer.rozetka.company/gomer/items/active/source/${sourceId}`;
    } else if (sourceType === "changes") {
      p.set("page", String(page));
      p.set("per-page", perPage);
      base = `https://gomer.rozetka.company/gomer/items/changes/source/${sourceId}`;
    } else {
      // Для /items/source используется size, а не per-page (иначе часто отдает только 20 строк)
      p.set("size", perPage);
      p.set("page", String(page));
      base = `https://gomer.rozetka.company/gomer/items/source/${sourceId}`;
    }

    return `${base}?${p.toString()}`;
  };

  const logOfferIdsRequest = (page, url) => {
    console.log(
      "[RW][OFFER_IDS][REQUEST]",
      "sourceId=",
      sourceId || "(empty)",
      "rz_category_id=",
      categoryId || "(empty)",
      "bpm_number=",
      taskId || "(empty)",
      "page=",
      page,
      "perPage=",
      perPage,
      "sourceType=",
      sourceType,
      "url=",
      url
    );
  };

  const allIdsSet = new Set();
  const perPageNum = Number(perPage) || 500;
  const parallelPages = 3;
  const restoreDelayMs = 500;
  const firstUrl = makeUrl(1);
  logOfferIdsRequest(1, firstUrl);
  throwIfTaskAborted(taskSignal);
  // После небольшой задержки возвращаем пользовательский size/per-page.
  const firstFetchPromise = fetchTextWithMeta(firstUrl, taskSignal);
  await delay(restoreDelayMs);
  await restorePageSizeIfNeeded(task, taskSignal);
  const firstMeta = await firstFetchPromise;
  const firstHtml = firstMeta.text;
  const firstIds = parseOfferIdsFromHtml(firstHtml);
  firstIds.forEach((id) => allIdsSet.add(id));
  const totalRows = firstMeta.totalCount || getTotalRowsFromHtml(firstHtml);
  const maxPageByHeader = firstMeta.pageCount || 0;
  const maxPageByTotal = totalRows > 0 ? Math.ceil(totalRows / perPageNum) : 0;
  const maxPageByLinks = getMaxPageFromHtml(firstHtml);
  const maxPage = Math.max(1, maxPageByHeader || 0, maxPageByTotal || 0, maxPageByLinks || 0);

  console.log(
    "[RW][QUEUE] offerIds pages:",
    "totalRows=",
    totalRows,
    "pageCountHeader=",
    maxPageByHeader,
    "totalPages=",
    maxPage,
    "per-page=",
    firstMeta.perPage || perPage,
    "sourceType=",
    sourceType
  );
  console.log("[RW][QUEUE] offerIds page 1/", maxPage, "ids=", firstIds.length);

  for (let startPage = 2; startPage <= maxPage; startPage += parallelPages) {
    const endPage = Math.min(maxPage, startPage + parallelPages - 1);
    const pages = [];
    for (let page = startPage; page <= endPage; page += 1) {
      pages.push(page);
    }

    const results = await Promise.all(
      pages.map(async (page) => {
        const pageUrl = makeUrl(page);
        logOfferIdsRequest(page, pageUrl);
        try {
          throwIfTaskAborted(taskSignal);
          // Не ждем ответа выборки: после небольшой задержки отправляем restore пользовательского размера страницы.
          const pageFetchPromise = fetchText(pageUrl, taskSignal);
          await delay(restoreDelayMs);
          await restorePageSizeIfNeeded(task, taskSignal);
          const html = await pageFetchPromise;
          const pageIds = parseOfferIdsFromHtml(html);
          return { ok: true, page, pageIds };
        } catch (error) {
          return { ok: false, page, error: error?.message || String(error) };
        }
      })
    );
    throwIfTaskAborted(taskSignal);

    let hasError = false;
    for (const result of results) {
      if (!result.ok) {
        hasError = true;
        console.warn("[RW][QUEUE] offerIds page error", `${result.page}/${maxPage}`, result.error);
        continue;
      }
      result.pageIds.forEach((id) => allIdsSet.add(id));
      console.log("[RW][QUEUE] offerIds page", `${result.page}/${maxPage}`, "ids=", result.pageIds.length);
    }
    if (hasError) break;
    throwIfTaskAborted(taskSignal);
  }

  const offerIds = Array.from(allIdsSet);
  setOfferIdsCache(cacheKey, offerIds);
  debugOfferIdsLog(firstUrl, offerIds);
  sendOfferIdsToContentConsole(task, offerIds, firstUrl);
  return offerIds;
}

function buildProductLink(origin, sourceId, offerId) {
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
  return `${origin}/gomer/items/source/${sourceId}?${params.toString()}`;
}

async function processExportTask(task, taskSignal) {
  task.skipPriceReason = "";
  throwIfTaskAborted(taskSignal);
  setTaskState(task, "processing", "Подготовка");

  const fields = { ...(task.fields || {}) };
  const generateProductLink = typeof task.generateProductLink === "boolean" ? task.generateProductLink : true;
  const sourceId = (task.sourceId || "").trim();
  const sourceOrigin = (task.sourceOrigin || "https://gomer.rozetka.company").trim();
  const paramName = (task.paramNameForPrice || "").trim();
  const paramValue = String(task.valueForPrice || "").trim().toLowerCase();
  const priceUrl = (task.priceUrl || "").trim();
  const priceLinkTitle = String(task.priceLinkTitle || "");
  const isPriceUnavailable =
    !priceUrl ||
    priceUrl === "javascript:void(0)" ||
    priceUrl.includes("javascript:void(0)") ||
    priceLinkTitle.includes("Api virtual source");

  let productLink = "";

  if (!generateProductLink) {
    setTaskState(task, "processing", "Генерация ссылки отключена");
    productLink = "";
  } else if (!isPriceUnavailable && sourceId && paramName && paramValue) {
    try {
      throwIfTaskAborted(taskSignal);
      setTaskState(task, "processing", "Загрузка списка товаров");
      const offerIds = await collectOfferIdsForTask(task, taskSignal);
      throwIfTaskAborted(taskSignal);
      setTaskState(task, "processing", `Собрано offerId: ${offerIds.length}`);

      let foundOfferId = "";
      if (offerIds.length) {
        throwIfTaskAborted(taskSignal);
        setTaskState(task, "processing", "Поиск в прайсе по offerId");
        foundOfferId = await findOfferByOfferIdsAndParam(priceUrl, offerIds, paramName, paramValue, taskSignal);
      } else if (task.categoryIdForPrice) {
        throwIfTaskAborted(taskSignal);
        setTaskState(task, "processing", "Поиск в прайсе по категории");
        foundOfferId = await findOfferByCategoryAndParam(
          priceUrl,
          task.categoryIdForPrice,
          paramName,
          paramValue,
          taskSignal
        );
      }

      if (foundOfferId) {
        productLink = { text: "Ссылка", url: buildProductLink(sourceOrigin, sourceId, foundOfferId) };
        setTaskState(task, "processing", `Найден offerId: ${foundOfferId}`);
      } else {
        productLink = "";
        setTaskState(task, "processing", "offerId не найден, запись без ссылки");
      }
    } catch (error) {
      if (isTaskAborted(taskSignal)) {
        throw new Error("Остановлено пользователем");
      }
      if (shouldSkipPriceSearchOnError(error)) {
        productLink = "";
        task.skipPriceReason = "504/timeout, запись без ссылки";
        setTaskState(task, "processing", "Прайс/поиск недоступен (504/timeout), запись без ссылки");
      } else {
        throw error;
      }
    }
  } else {
    setTaskState(task, "processing", "Прайс недоступен, запись без ссылки");
  }

  fields["Ссылка на товар"] = productLink;

  throwIfTaskAborted(taskSignal);
  setTaskState(task, "processing", "Запись в Google Sheets");
  const appendResult = await withTimeout(
    appendFields(task.sheetUrl, fields),
    45000,
    "Таймаут записи в Google Sheets"
  );
  if (!appendResult || !appendResult.ok) {
    throw new Error(appendResult?.error || "Ошибка записи в Google Sheets");
  }

  // Важно: запрос с увеличенным size/per-page может менять серверное состояние пагинации в сессии.
  // После обработки задачи возвращаем пользовательский размер страницы.
  const restoreUrl = String(task.restorePageSizeUrl || "").trim();
  if (restoreUrl) {
    try {
      await restorePageSizeIfNeeded(task, taskSignal);
      console.log("[RW][QUEUE] page-size restored:", restoreUrl);
    } catch (e) {
      console.warn("[RW][QUEUE] page-size restore failed:", e?.message || e);
    }
  }
}

function restoreQueueForProcessing(items) {
  const normalized = Array.isArray(items) ? items : [];
  _exportQueue = normalized
    .slice(-EXPORT_QUEUE_MAX_ITEMS)
    .map((task) => ({
      ...task,
      status: task.status === "done" || task.status === "error" ? task.status : "pending",
      updatedAt: Number(task.updatedAt || task.createdAt || Date.now())
    }));
  pruneQueueHistory();
}

function initializeQueue() {
  if (_queueInitialized) return Promise.resolve();
  if (_queueInitPromise) return _queueInitPromise;
  _queueInitPromise = new Promise((resolve) => {
    chrome.storage.local.get([EXPORT_QUEUE_DATA_KEY], (data) => {
      restoreQueueForProcessing(data && data[EXPORT_QUEUE_DATA_KEY]);
      _queueInitialized = true;
      saveQueueState();
      saveQueueData();
      processExportQueue();
      resolve();
      _queueInitPromise = null;
    });
  });
  return _queueInitPromise;
}

async function processExportQueue() {
  await initializeQueue();
  if (_exportIsProcessing) return;
  _exportIsProcessing = true;
  saveQueueState();
  try {
    while (true) {
      const nextTask = _exportQueue.find((task) => task.status === "pending");
      if (!nextTask) break;
      const taskController = new AbortController();
      _activeTaskId = nextTask.id;
      _activeTaskAbortController = taskController;
      try {
        setTaskState(nextTask, "processing", "Запуск обработки");
        await processExportTask(nextTask, taskController.signal);
        const doneMessage = nextTask.skipPriceReason
          ? `Готово • ${nextTask.skipPriceReason}`
          : "Готово";
        setTaskState(nextTask, "done", doneMessage);
      } catch (error) {
        setTaskState(nextTask, "error", error?.message || "Ошибка обработки");
      } finally {
        if (_activeTaskId === nextTask.id) {
          _activeTaskId = "";
          _activeTaskAbortController = null;
        }
      }
    }
  } finally {
    _exportIsProcessing = false;
    saveQueueState();
  }
}

function enqueueExportTask(payload) {
  pruneQueueHistory();
  const task = {
    id: makeTaskId(),
    title: makeTaskTitle(payload),
    status: "pending",
    message: "В очереди",
    createdAt: Date.now(),
    updatedAt: Date.now(),
    ...payload
  };
  _exportQueue.push(task);
  if (_exportQueue.length > EXPORT_QUEUE_MAX_ITEMS) {
    _exportQueue = _exportQueue.slice(-EXPORT_QUEUE_MAX_ITEMS);
  }
  saveQueueState();
  saveQueueData();
  processExportQueue();
  return task.id;
}

function killExportTask(taskId) {
  const id = String(taskId || "").trim();
  if (!id) return { ok: false, error: "Не передан id задачи" };
  const task = _exportQueue.find((item) => item.id === id);
  if (!task) return { ok: false, error: "Задача не найдена" };

  if (_activeTaskId === id && _activeTaskAbortController) {
    try {
      _activeTaskAbortController.abort();
    } catch (_) {}
  }

  _exportQueue = _exportQueue.filter((item) => item.id !== id);
  saveQueueState();
  saveQueueData();
  processExportQueue();
  return { ok: true };
}

chrome.runtime.onStartup.addListener(() => {
  initializeQueue();
});

chrome.runtime.onInstalled.addListener(() => {
  initializeQueue();
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message && message.type === "ENQUEUE_EXPORT_TASK") {
    initializeQueue()
      .then(() => {
        const payload = {
          ...(message.task || {}),
          senderTabId: sender && sender.tab ? sender.tab.id : null
        };
        const taskId = enqueueExportTask(payload);
        sendResponse({ ok: true, taskId });
      })
      .catch((error) => {
        sendResponse({ ok: false, error: error?.message || "Не удалось добавить задачу в очередь" });
      });
    return true;
  }

  if (message && message.type === "KILL_EXPORT_TASK") {
    initializeQueue()
      .then(() => {
        const result = killExportTask(message.taskId);
        sendResponse(result);
      })
      .catch((error) => {
        sendResponse({ ok: false, error: error?.message || "Не удалось удалить задачу" });
      });
    return true;
  }

  if (message && message.type === "GET_AUTH_TOKEN") {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) {
        sendResponse({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      sendResponse({ ok: true, token });
    });
    return true;
  }

  if (message && message.type === "APPEND_FIELDS") {
    appendFields(message.sheetUrl, message.fields)
      .then((data) => sendResponse(data))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message && message.type === "FIND_OFFER_IN_PRICE") {
    let searchUrl = message.url;
    
    // Специальная обработка для a99.com.ua/rozetka_feed.xml
    // Если URL содержит этот домен и есть offerIds, добавляем их в product_ids
    if (searchUrl && searchUrl.includes("a99.com.ua/rozetka_feed.xml") && message.offerIds && message.offerIds.length) {
      try {
        const urlObj = new URL(searchUrl);
        // Добавляем оффер ид через запятую в параметр product_ids
        urlObj.searchParams.set("product_ids", message.offerIds.join(","));
        searchUrl = urlObj.toString();
        console.log("[RW] FIND_OFFER_IN_PRICE: модифицирован URL для a99.com.ua, добавлены product_ids=", message.offerIds.length, "офферов");
      } catch (e) {
        console.error("[RW] FIND_OFFER_IN_PRICE: ошибка модификации URL", e);
      }
    }
    
    let promise;
    if (message.offerIds && message.offerIds.length) {
      promise = findOfferByOfferIdsAndParam(
        searchUrl,
        message.offerIds,
        message.paramName,
        message.paramValue
      );
    } else {
      _lastOfferDebug = null;
      promise = findOfferByCategoryAndParam(
        searchUrl,
        message.categoryId,
        message.paramName,
        message.paramValue
      );
    }
    promise
      .then((offerId) => sendResponse({ ok: true, offerId: offerId || "", debug: _lastOfferDebug }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }
});
