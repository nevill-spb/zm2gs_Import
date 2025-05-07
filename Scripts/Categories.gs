const SetupCategories = (function () {
  const sheet = sheetHelper.GetSheetFromSettings('TAGS_SHEET');

  // Парсеры для булевых значений
  function parseBool(value) {
    return value === true || value === "TRUE";
  }

  function parseBoolRequired(value) {
    return value || value === null ? "TRUE" : "FALSE";
  }

  // Индексы полей по id для удобства доступа
  const fieldIndex = {};
  Settings.CATEGORY_FIELDS.forEach((f, i) => fieldIndex[f.id] = i);

  // Построение строки данных по полям из Settings
  function buildRow(tag) {
    return Settings.CATEGORY_FIELDS.map(field => {
      switch (field.id) {
        case "showIncome":
        case "showOutcome":
        case "budgetIncome":
        case "budgetOutcome":
          return parseBool(tag[field.id]) ? "TRUE" : "FALSE";
        case "required":
          return parseBoolRequired(tag[field.id]);
        case "delete":
          return false; // чекбокс пустой
        case "color":
          return tag.color != null ? tag.color : "";
        case "user":
        case "changed":
          return tag[field.id] !== undefined ? tag[field.id] : "";
        default:
          return tag[field.id] || "";
      }
    });
  }

  // Подготовка данных категорий для записи в лист
  const prepareData = (json) => {
    if (!('tag' in json)) return;

    // Сортируем категории по названию и id для стабильности
    const sortedTags = json.tag.slice().sort((a, b) => {
      if (a.title === b.title) return a.id.localeCompare(b.id);
      return a.title.localeCompare(b.title);
    });

    // Создаем словарь для быстрого поиска дочерних категорий по parent id
    const childrenMap = {};
    sortedTags.forEach(tag => {
      if (tag.parent) {
        if (!childrenMap[tag.parent]) childrenMap[tag.parent] = [];
        childrenMap[tag.parent].push(tag);
      }
    });

    // Формируем массив данных с группировкой: родительские + дочерние
    const data = [];

    sortedTags.forEach(tag => {
      if (!tag.parent) {
        data.push(buildRow(tag));
        const children = childrenMap[tag.id] || [];
        children.forEach(child => data.push(buildRow(child)));
      }
    });

    // Очистка листа перед записью
    sheet.clearContents();
    sheet.clearFormats();

    // Заголовки из Settings
    const headers = Settings.CATEGORY_FIELDS.map(f => f.title);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

      // Список id булевых полей, для которых нужны чекбоксы
      const boolFields = ["showIncome", "showOutcome", "budgetIncome", "budgetOutcome", "required", "delete"];

      // Вставляем чекбоксы по каждой колонке
      boolFields.forEach(fieldId => {
        const colIndex = fieldIndex[fieldId];
        if (colIndex === undefined) return;
        sheet.getRange(2, colIndex + 1, data.length, 1).insertCheckboxes();
      });

      // Очищаем лишние чекбоксы и валидации одним вызовом на весь диапазон
      const boolCols = boolFields.map(id => fieldIndex[id]).filter(i => i !== undefined);
      if (boolCols.length > 0) {
        const minCol = Math.min(...boolCols);
        const maxCol = Math.max(...boolCols);
        const lastRow = sheet.getMaxRows();
        const startRow = data.length + 2; // строка после данных + заголовок
        const numRows = lastRow - startRow + 1;
        if (numRows > 0) {
          sheet.getRange(startRow, minCol + 1, numRows, maxCol - minCol + 1)
            .clearContent()
            .clearDataValidations();
        }
      }

      // Устанавливаем цвет фона для колонки "Цвет"
      for (let i = 0; i < data.length; i++) {
        const colorNum = data[i][fieldIndex.color];
        if (colorNum && colorNum > 0) {
          let num = colorNum >>> 0;
          const b = num & 0xFF;
          const g = (num & 0xFF00) >>> 8;
          const r = (num & 0xFF0000) >>> 16;
          const hexColor = "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
          const cell = sheet.getRange(i + 2, fieldIndex.color + 1);
          cell.setBackground(hexColor);
          cell.setValue("");
        } else {
          sheet.getRange(i + 2, fieldIndex.color + 1).setBackground(null);
        }
      }
    }
  };

  // Преобразует hex-цвет в формат ZenMoney (число)
  function hexToZenColor(hex) {
    if (!hex) return null;
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    if (!result) return null;
    const r = parseInt(result[1], 16);
    const g = parseInt(result[2], 16);
    const b = parseInt(result[3], 16);
    return (r << 16) + (g << 8) + b;
  }

  // Загрузка категорий с сервера и отображение в листе
  const doLoad = function () {
    const json = zmData.RequestForceFetch(['tag']);
    Logs.logApiCall("Fetch Categories", { tag: [] }, json);
    prepareData(json);
  };

  // Сохранение изменений категорий из листа обратно на сервер
  const doUpdate = function () {
    const json = zmData.RequestForceFetch(['tag']);
    const tags = json['tag'] || [];
    const ts = Math.floor(Date.now() / 1000);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("Нет данных для обновления категорий");
      return;
    }

    Dictionaries.loadDictionariesFromSheet();

    const range = sheet.getRange(2, 1, lastRow - 1, Settings.CATEGORY_FIELDS.length);
    const values = range.getValues();

    const newTags = [];
    const deletionRequests = [];
    const idMap = {};

    // Массив для новых ID, чтобы записать пакетно
    const newIds = Array(values.length).fill(null);

    // Первый проход: определяем новые ID для всех категорий
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const oldTagId = row[fieldIndex.id];
      const title = row[fieldIndex.title];
      if (!title || typeof title !== 'string' || title.trim() === '') continue;
      const user = row[fieldIndex.user] || Number(Settings.DefaultUserId);
      const deleteFlag = row[fieldIndex.delete] === true;

      if (deleteFlag && oldTagId) {
        deletionRequests.push({
          id: oldTagId,
          object: 'tag',
          stamp: ts,
          user: user
        });
        continue;
      }

      if (!oldTagId || !tags.find(t => t.id === oldTagId)) {
        const newId = Utilities.getUuid().toLowerCase();
        idMap[oldTagId || `new_${i}`] = newId;
        newIds[i] = newId;  // Запоминаем новый ID для записи
      } else {
        idMap[oldTagId] = oldTagId;
      }
    }

    // Пакетная запись новых ID в колонку id
    const idColumnRange = sheet.getRange(2, fieldIndex.id + 1, values.length, 1);
    const idColumnValues = idColumnRange.getValues();
    for (let i = 0; i < values.length; i++) {
      if (newIds[i]) {
        idColumnValues[i][0] = newIds[i];
      }
    }
    idColumnRange.setValues(idColumnValues);

    // Второй проход: формируем объекты категорий с обновлёнными parentId
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const oldTagId = row[fieldIndex.id];
      const oldParentId = row[fieldIndex.parent] || null;
      const title = row[fieldIndex.title];
      if (!title || typeof title !== 'string' || title.trim() === '') continue;
      const colorHex = sheet.getRange(i + 2, fieldIndex.color + 1).getBackground();
      const icon = row[fieldIndex.icon];
      const showIncome = parseBool(row[fieldIndex.showIncome]);
      const showOutcome = parseBool(row[fieldIndex.showOutcome]);
      const budgetIncome = parseBool(row[fieldIndex.budgetIncome]);
      const budgetOutcome = parseBool(row[fieldIndex.budgetOutcome]);
      const required = parseBool(row[fieldIndex.required]);
      const deleteFlag = row[fieldIndex.delete] === true;
      const user = row[fieldIndex.user] || Number(Settings.DefaultUserId);
      const changed = row[fieldIndex.changed];

      if (!title || title.trim() === '') continue;
      if (deleteFlag && oldTagId) continue;

      const newTagId = idMap[oldTagId || `new_${i}`];
      if (!newTagId) continue;

      const newParentId = (oldParentId && idMap[oldParentId]) ? idMap[oldParentId] : null;

      let tag = tags.find(t => t.id === newTagId);
      if (!tag) {
        tag = {
          id: newTagId,
          user: user,
          changed: ts
        };
      }

      tag.parent = newParentId;
      tag.title = title;
      tag.color = hexToZenColor(colorHex);
      tag.icon = icon;
      tag.showIncome = showIncome;
      tag.showOutcome = showOutcome;
      tag.budgetIncome = budgetIncome;
      tag.budgetOutcome = budgetOutcome;
      tag.required = required;
      tag.user = user;
      tag.changed = ts;

      newTags.push(tag);
    }

    const data = {
      currentClientTimestamp: ts,
      serverTimestamp: ts
    };

    if (newTags.length > 0) data.tag = newTags;
    if (deletionRequests.length > 0) data.deletion = deletionRequests;

    if (newTags.length === 0 && deletionRequests.length === 0) {
      Logger.log("Нет изменений или удалений для обработки.");
      return;
    }

    const result = zmData.Request(data);
    Logs.logApiCall("Update Tags", data, result);
    Logger.log("Результат обновления/удаления категорий: " + JSON.stringify(result));

    doLoad();
  };

  // Добавляем обработчик полной синхронизации для загрузки категорий
  fullSyncHandlers.push(prepareData);

  // Добавляем пункты меню для управления категориями
  function createMenu() {
    const ui = SpreadsheetApp.getUi();
    const subMenu = ui.createMenu("Setup categories")
      .addItem("Load", "SetupCategories.doLoad")
      .addItem("Save", "SetupCategories.doUpdate");
    gsMenu.addSubMenu(subMenu);
  }

  // Вызываем создание меню при инициализации модуля
  createMenu();

  return {
    doLoad,
    doUpdate,
  };
})();
