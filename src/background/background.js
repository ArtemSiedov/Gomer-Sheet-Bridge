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
  let value = afterName.slice(startVal + 1, endTag).trim();
  value = stripCdata(value);
  return value.toLowerCase() === String(paramValue || "").trim().toLowerCase();
}

// Проверка: оффер в списке offerIds и имеет параметр paramName = paramValue.
function offerMatchesIdsAndParam(offerXml, offerIdsSet, paramName, paramValue) {
  // Сначала быстро проверяем ID (это быстрее чем парсить параметры)
  const idMatch = offerXml.match(/<offer\s+[^>]*\bid="([^"]+)"/) || offerXml.match(/<offer\s+id="([^"]+)"/);
  if (!idMatch) return false;
  const idNorm = idMatch[1].trim().toLowerCase();
  const offerId = idMatch[1].trim();
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

      let value = String(foundParamValue).trim();
      value = stripCdata(value);
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1).trim();
      }

      const paramValNorm = (paramValue || "").trim();
      let searchValue = paramValNorm;
      if ((searchValue.startsWith('"') && searchValue.endsWith('"')) || (searchValue.startsWith("'") && searchValue.endsWith("'"))) {
        searchValue = searchValue.slice(1, -1).trim();
      }

      value = value.toLowerCase();
      searchValue = searchValue.toLowerCase();
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
  
  let value = afterName.slice(startVal + 1, endTag).trim();
  value = stripCdata(value);
  
  // Убираем кавычки из начала и конца значения, если они есть
  if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
    value = value.slice(1, -1).trim();
  }
  
  const paramValNorm = (paramValue || "").trim();
  // Также убираем кавычки из искомого значения
  let searchValue = paramValNorm;
  if ((searchValue.startsWith('"') && searchValue.endsWith('"')) || (searchValue.startsWith("'") && searchValue.endsWith("'"))) {
    searchValue = searchValue.slice(1, -1).trim();
  }
  
  value = value.toLowerCase();
  searchValue = searchValue.toLowerCase();
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

function getOfferIdFromXml(offerXml) {
  const m = offerXml.match(/<offer\s+[^>]*\bid="([^"]+)"/) || offerXml.match(/<offer\s+id="([^"]+)"/);
  return m ? m[1] : null;
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
        const idMatch = offerXml.match(/<offer\s+[^>]*\bid="([^"]+)"/) || offerXml.match(/<offer\s+id="([^"]+)"/);
        if (idMatch) {
          const idNorm = idMatch[1].trim().toLowerCase();
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

async function findOfferByCategoryAndParam(url, categoryId, paramName, paramValue) {
  const response = await fetch(url);
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

async function findOfferByOfferIdsAndParam(url, offerIds, paramName, paramValue) {
  if (!offerIds || !offerIds.length) return null;
  _lastOfferDebug = initOfferDebug();
  const offerIdsSet = new Set(
    offerIds.map((id) => String(id).trim().toLowerCase()).filter(Boolean)
  );
  console.log("[RW] findOfferByOfferIds: url=", url?.slice(0, 80), "offerIdsCount=", offerIds.length, "paramName=", paramName, "paramValue=", JSON.stringify((paramValue || "").slice(0, 60)));

  // Обычный последовательный поиск (без Range/параллелизма)
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
        const idMatch = offerXml.match(/<offer\s+[^>]*\bid="([^"]+)"/) || offerXml.match(/<offer\s+id="([^"]+)"/);
        if (idMatch) {
          const idNorm = idMatch[1].trim().toLowerCase();
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

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
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
