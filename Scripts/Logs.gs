const Logs = (function () {
  const LOG_SHEET_NAME = "APILogs";
  const HEADERS = ["Timestamp", "Method", "Request Payload", "Response", "Status"];
  const MAX_RESPONSE_LENGTH = 50000;

  function isLoggingEnabled() {  
    const settingsSheet = sheetHelper.Get(Settings.SHEETS.SETTINGS.NAME);  
    const logCell = Settings.SHEETS.SETTINGS.CELLS.LOGS;  
    const value = settingsSheet.getRange(logCell).getValue();  
    return value === "ДА";  
  }

  // Функция для построения конфигурации фильтрации из листа Settings
  function buildFilterConfigFromSettings() {
    const settingsSheet = sheetHelper.Get(Settings.SHEETS.SETTINGS.NAME);
    if (!settingsSheet) throw new Error(`Лист "${Settings.SHEETS.SETTINGS.NAME}" не найден`);
    const dictRange = settingsSheet.getRange("D12:D17").getValues().flat();
    const stateRange = settingsSheet.getRange("E12:E17").getValues().flat();
    const maxArrayItems = settingsSheet.getRange("E18").getValues().flat();

    const excludeRefs = [];
    const showCount = [];

    for (let i = 0; i < dictRange.length; i++) {
      const dictName = dictRange[i];
      const state = (stateRange[i] || "").trim().toLowerCase();

      if (!dictName) continue;

      if (state === "исключить") {
        excludeRefs.push(dictName);
      } else if (state === "сократить") {
        showCount.push(dictName);
      }
      // "отобразить" — ничего не добавляем
    }

    return {
      excludeRefs,
      showCount,
      maxArrayItems,
      excludeNullFields: [
        'payee',
        'originalPayee',
        'opIncome',
        'opOutcome',
        'opIncomeInstrument',
        'opOutcomeInstrument',
        'latitude',
        'longitude',
        'merchant',
        'incomeBankID',
        'outcomeBankID',
        'reminderMarker'
      ]
    };
  }

  let FILTER_CONFIG = null;
  function getFilterConfig() {
    if (!FILTER_CONFIG) {
      FILTER_CONFIG = buildFilterConfigFromSettings();
    }
    return FILTER_CONFIG;
  }

  // Кастомный форматтер JSON с переводами строк для массивов объектов
  function customJSONStringify(obj, indent = 2) {
    const filterConfig = getFilterConfig();
    return JSON.stringify(obj, function (key, value) {
      if (value && typeof value === 'object' && !Array.isArray(value)) {
        const filtered = {};
        for (let k in value) {
          if (filterConfig.excludeNullFields.includes(k) && value[k] === null) {
            continue;
          }
          filtered[k] = value[k];
        }
        return filtered;
      }
      return value;
    }, indent).replace(/},\s*{/g, '},\n{');
  }

  // Фильтрация и форматирование ответа
  function filterResponse(response) {
    try {
      const filterConfig = getFilterConfig();
      const responseObj = typeof response === "string"
        ? JSON.parse(response)
        : response;

      const filteredResponse = {};

      Object.keys(responseObj).forEach(key => {
        if (filterConfig.excludeRefs.includes(key)) return;

        if (key === 'transaction' && Array.isArray(responseObj[key])) {
          filteredResponse[key] = responseObj[key].filter(tag => !tag.deleted);
          if (filterConfig.showCount.includes(key)) {
            filteredResponse[key] = `[${filteredResponse[key].length} items]`;
          } else if (filteredResponse[key].length > filterConfig.maxArrayItems) {
            filteredResponse[key] = [
              ...filteredResponse[key].slice(0, filterConfig.maxArrayItems),
              `... and ${filteredResponse[key].length - filterConfig.maxArrayItems} more`
            ];
          }
        } else if (filterConfig.showCount.includes(key) && Array.isArray(responseObj[key])) {
          filteredResponse[key] = `[${responseObj[key].length} items]`;
        } else if (Array.isArray(responseObj[key])) {
          if (responseObj[key].length > filterConfig.maxArrayItems) {
            filteredResponse[key] = [
              ...responseObj[key].slice(0, filterConfig.maxArrayItems),
              `... and ${responseObj[key].length - filterConfig.maxArrayItems} more`
            ];
          } else {
            filteredResponse[key] = responseObj[key];
          }
        } else {
          filteredResponse[key] = responseObj[key];
        }
      });

      return customJSONStringify(filteredResponse, 2);
    } catch (e) {
      return typeof response === "string" ? response : JSON.stringify(response);
    }
  }

  function formatJSON(data) {
    if (typeof data === "string") {
      try {
        return JSON.stringify(JSON.parse(data), null, 2);
      } catch (e) {
        return data;
      }
    }
    return JSON.stringify(data, null, 2);
  }

  function getStatus(response) {
    try {
      const responseObj = typeof response === "string"
        ? JSON.parse(response)
        : response;

      if (responseObj.error) {
        return "Error";
      }

      const keys = Object.keys(responseObj);
      if (keys.length === 0 || (keys.length === 1 && keys[0] === "serverTimestamp")) {
        return "Empty Response";
      }

      return "Success";
    } catch (e) {
      return "Error";
    }
  }

  function initLogSheet(sheet) {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setValues([HEADERS])
        .setFontWeight('bold');
      sheet.getRange("C:D").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.setFrozenRows(1);
    }
  }
  
  function logApiCall(method, requestPayload, responseContent) {
    if (!isLoggingEnabled()) {  
      return; // если чекбокс выключен, не логируем  
    }

    const logSheet = sheetHelper.Get(LOG_SHEET_NAME);
    initLogSheet(logSheet);

    logSheet.setColumnWidth(3, 400); // столбец Request Payload
    logSheet.setColumnWidth(4, 400); // столбец Response

    const formattedRequest = formatJSON(requestPayload);
    let formattedResponse = filterResponse(responseContent);

    if (formattedResponse.length > MAX_RESPONSE_LENGTH) {
      formattedResponse = formattedResponse.substring(0, MAX_RESPONSE_LENGTH) + "... [truncated]";
    }

    const rowData = [
      new Date().toISOString(),
      method,
      formattedRequest,
      formattedResponse,
      getStatus(responseContent)
    ];

    logSheet.appendRow(rowData);
    logSheet.autoResizeColumns(1, HEADERS.length);
  }
  
  return {
    logApiCall
  };
})();
