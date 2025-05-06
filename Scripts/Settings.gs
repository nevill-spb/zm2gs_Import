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
        LOGS: "B5",
        EXPORT_SHEET: "B6",  
        IMPORT_SHEET: "B7",  
        DICTIONARIES_SHEET: "B8",  
        TAGS_SHEET: "B9",  
        ACCOUNTS_SHEET: "B10",  
        MERCHANTS_SHEET: "B11"
      }
    },
    CHANGES: "Изменения",
    ERRORS: "Ошибки импорта",
  },

  // Базовая структура колонок в листе данных
  TRANSACTION_FIELDS: [
    { id: "id", title: "ID" },
    { id: "date", title: "Дата" },
    { id: "tag", title: "Категория" },
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
    { id: "deleted", title: "Удалено" },
    { id: "modified", title: "Изменено" },
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
    { id: "reminderMarker", title: "Маркер напоминания" }
  ],

  EXPORT: {
    // HEADERS генерируются из TRANSACTION_FIELDS (используем title)
    get HEADERS() {
      return Settings.TRANSACTION_FIELDS.map(field => field.title);
    },
    
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

};
