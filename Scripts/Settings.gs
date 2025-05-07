/**
 * Глобальные настройки приложения
 */
const Settings = {
  // Настройки листов
  SHEETS: {
    SETTINGS: {
      NAME: "Settings",
      CELLS: {
        TOKEN: "B1",
        TIMESTAMP: "B2",
        DEFAULT_USER_ID: "B3",
        DEFAULT_CURRENCY_ID: "B4",
        LOGS_ENABLED: "B5",
        EXPORT_SHEET: "B6",  
        IMPORT_SHEET: "B7",  
        DICTIONARIES_SHEET: "B8",  
        TAGS_SHEET: "B9",  
        ACCOUNTS_SHEET: "B10",  
        MERCHANTS_SHEET: "B11",
        TAG_MODE: "B12"
      }
    },
    LOGS: "Логи",
    CHANGES: "Изменения",
    ERRORS: "Ошибки импорта",
  },

  // Фильтры для словарей и их значения
  LOG_FILTER: {  
    DICT_RANGE: "D12:D17",  // account, tag, merchant, instrument, user, transaction
    STATE_RANGE: "E12:E17", // Отобразить, Сократить, Исключить
    MAX_ARRAY_ITEMS_CELL: "E18"  // Максимальное количество объектов словаря в логах
  },

  // Базовая структура колонок в листе данных
  TRANSACTION_FIELDS: [
    { id: "id", title: "ID" },
    { id: "date", title: "Дата" },
    { id: "tag", title: "Категория" },
    { id: "tag1", title: "Тег1" },
    { id: "tag2", title: "Тег2" },
    { id: "merchant", title: "Место" },
    { id: "comment", title: "Комментарий" },
    { id: "outcomeAccount", title: "Счёт расход" },
    { id: "outcome", title: "Расход" },
    { id: "outcomeInstrument", title: "Валюта расхода" },
    { id: "incomeAccount", title: "Счёт дохода" },
    { id: "income", title: "Доход" },
    { id: "incomeInstrument", title: "Валюта дохода" },
    { id: "created", title: "Дата создания" },
    { id: "changed", title: "Дата изменения" },
    { id: "user", title: "Пользователь" },
    { id: "deleted", title: "Удалить" },
    { id: "modified", title: "Изменить" }, // для чекбокса, не из API 
    { id: "viewed", title: "Просмотрено" },
    { id: "hold", title: "Холд" },
    { id: "payee", title: "Получатель" },
    { id: "originalPayee", title: "Исходный получатель" },
    { id: "qrCode", title: "QR-код" },
    { id: "source", title: "Источник" },
    { id: "opIncome", title: "Доход в валюте операции" },
    { id: "opOutcome", title: "Расход в валюте операции" },
    { id: "opIncomeInstrument", title: "Валюта операции дохода" },
    { id: "opOutcomeInstrument", title: "Валюта операции расхода" },
    { id: "incomeBankID", title: "ID банка дохода" },
    { id: "outcomeBankID", title: "ID банка расхода" },
    { id: "latitude", title: "Широта" },
    { id: "longitude", title: "Долгота" },
    { id: "reminderMarker", title: "Маркер напоминания" },
  ],

  CATEGORY_FIELDS: [  
    { id: "id", title: "ID" },  
    { id: "parent", title: "Parent ID" },  
    { id: "title", title: "Название" },  
    { id: "color", title: "Цвет" },  
    { id: "icon", title: "Иконка" },  
    { id: "showIncome", title: "В доходе" },  
    { id: "showOutcome", title: "В расходе" },  
    { id: "budgetIncome", title: "В бюджете доходов" },  
    { id: "budgetOutcome", title: "В бюджете расходов" },  
    { id: "required", title: "Обязательная" },
    { id: "delete", title: "Удалить" }, // для чекбокса, не из API  
    { id: "modify", title: "Изменить" }, // для чекбокса, не из API  
    { id: "user", title: "Пользователь" },
    { id: "changed", title: "Дата изменения" },
  ],

  // Основные валюты, добавьте свои при желании
  ALLOWED_CURRENCY_CODES: ['USD', 'EUR', 'RUB', 'UAH'], 

  TAG_MODES: {  
    SINGLE_COLUMN: "Одной строкой",  
    MULTIPLE_COLUMNS: "Разделить"  
  },

  EXPORT: {
      COLORS: {
      NEW: "#d9ead3",     // green
      WAS_NEW: "#e6ffe1",   // jade (логика временно отключена)
      MODIFIED: "#fff2cc", // yellow
      TO_DELETE: "#fad3ab", // orange
      DELETED: "#f4cccc"   // red
    }
  },

  IMPORT: {
    BATCH_SIZE: 100,
    PROGRESS_INTERVAL: 100,
    UPDATE_COLUMNS_LIST: ['id', 'deleted', 'modified', 'created', 'changed', 'user'],
  },

  get DefaultUserId() { 
    const userName = sheetHelper.GetCellValue(this.SHEETS.SETTINGS.NAME, this.SHEETS.SETTINGS.CELLS.DEFAULT_USER_ID);  
    if (!userName || userName.trim() === "") {  
      throw new Error("Имя пользователя в настройках не задано");  
    }  
    const userId = Dictionaries.getUserId(userName.trim());  
    if (!userId) {  
      throw new Error(`Пользователь с именем "${userName}" не найден в справочниках`);  
    }  
    return Number(userId);
  },  
  
  get DefaultCurrencyId() { 
    const currencyShortTitle = sheetHelper.GetCellValue(this.SHEETS.SETTINGS.NAME, this.SHEETS.SETTINGS.CELLS.DEFAULT_CURRENCY_ID);  
    if (!currencyShortTitle || currencyShortTitle.trim() === "") {  
      throw new Error("Валюта в настройках не задана");  
    }  
    const currencyId = Dictionaries.getInstrumentId(currencyShortTitle.trim());  
    if (!currencyId) {  
      throw new Error(`Валюта "${currencyShortTitle}" не найдена в справочниках`);  
    }  
    return Number(currencyId);
  },

  get TagMode() {  
    const tagMode = sheetHelper.GetCellValue(this.SHEETS.SETTINGS.NAME, this.SHEETS.SETTINGS.CELLS.TAG_MODE);  
    return tagMode === this.TAG_MODES.MULTIPLE_COLUMNS ?   
      this.TAG_MODES.MULTIPLE_COLUMNS :   
      this.TAG_MODES.SINGLE_COLUMN;  
  }

};
