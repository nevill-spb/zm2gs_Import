/**
 * Глобальные настройки приложения
 */

//═══════════════════════════════════════════════════════════════════════════
// НАСТРОЙКИ ЛИСТОВ
// Определяет имена листов и адреса ячеек с ключевыми настройками
//═══════════════════════════════════════════════════════════════════════════
const Settings = {

  SHEETS: {
    SETTINGS: {
      NAME: "Settings", // Имя листа с настройками
      CELLS: {
        TOKEN: "B1",                // Токен для API Zen Money
        TIMESTAMP: "B2",            // Временная метка последней синхронизации
        DEFAULT_USER_ID: "B3",      // ID пользователя по умолчанию
        DEFAULT_CURRENCY_ID: "B4",  // ID валюты по умолчанию
        LOGS_ENABLED: "B5",         // Включение/выключение логирования
        EXPORT_SHEET: "B6",         // Лист для экспорта данных
        IMPORT_SHEET: "B7",         // Лист для импорта данных
        DICTIONARIES_SHEET: "B8",   // Лист со справочниками
        TAGS_SHEET: "B9",           // Лист с тегами
        ACCOUNTS_SHEET: "B10",      // Лист со счетами
        MERCHANTS_SHEET: "B11",     // Лист с получателями
        TAG_MODE: "B12",            // Режим отображения тегов
        ALLOW_DUPLICATES: "B13"     // Разрешить создание дубликатов категорий, счетов и мест
      }
    },
    LOGS: "Логи",                   // Лист для логов
    CHANGES: "Изменения",           // Лист для отслеживания изменений
    ERRORS: "Ошибки импорта",       // Лист для ошибок импорта
  },

  //═══════════════════════════════════════════════════════════════════════════
  // НАСТРОЙКИ ФИЛЬТРАЦИИ ЛОГОВ
  // Диапазоны ячеек для фильтрации логируемых данных
  //═══════════════════════════════════════════════════════════════════════════
  LOG_FILTER: {  
    DICT_RANGE: "D12:D17",        // Типы данных (account, tag, merchant и т.д.)
    STATE_RANGE: "E12:E17",       // Состояния (Отобразить, Сократить, Исключить)
    MAX_ARRAY_ITEMS_CELL: "E18"   // Максимальное количество элементов в массиве
  },

  //═══════════════════════════════════════════════════════════════════════════
  // СТРУКТУРА ПОЛЕЙ ДЛЯ ТРАНЗАКЦИЙ
  // Описывает все возможные поля транзакции
  //═══════════════════════════════════════════════════════════════════════════
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
    { id: "deleted", title: "Удалить" },
    { id: "modified", title: "Изменить" }, // для чекбокса, не из API 
    { id: "viewed", title: "Просмотрено" },
    { id: "created", title: "Дата создания" },
    { id: "changed", title: "Дата изменения" },
    { id: "user", title: "Пользователь" },
    { id: "payee", title: "Получатель" },
    { id: "originalPayee", title: "Исходный получатель" },
    { id: "qrCode", title: "QR-код" },
    { id: "hold", title: "Холд" },
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

  //═══════════════════════════════════════════════════════════════════════════
  // СТРУКТУРА ПОЛЕЙ ДЛЯ КАТЕГОРИЙ
  // Описывает все возможные поля категории
  //═══════════════════════════════════════════════════════════════════════════
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

  //═══════════════════════════════════════════════════════════════════════════
  // СТРУКТУРА ПОЛЕЙ ДЛЯ СЧЕТОВ
  // Описывает все возможные поля счета
  //═══════════════════════════════════════════════════════════════════════════
  ACCOUNT_FIELDS: [
    { id: "id", title: "ID" },
    { id: "title", title: "Название" },
    { id: "instrument", title: "Валюта" },
    { id: "type", title: "Тип" },
    { id: "balance", title: "Баланс" },
    { id: "startBalance", title: "Смещение" },
    { id: "inBalance", title: "Балансовый" },
    { id: "private", title: "Личный" },
    { id: "savings", title: "Накопительный" },
    { id: "archive", title: "Архивный" },
    { id: "creditLimit", title: "Кредитный лимит" },
    { id: "role", title: "Роль" },
    { id: "company", title: "Компания" },
    { id: "enableCorrection", title: "Коррекция баланса" },
    { id: "balanceCorrectionType", title: "Тип коррекции" },
    { id: "startDate", title: "Вклад с даты" },
    { id: "capitalization", title: "Капитализация" },
    { id: "percent", title: "Процент" },
    { id: "changed", title: "Изменён" },
    { id: "syncID", title: "SyncID" },
    { id: "enableSMS", title: "Включить SMS" },
    { id: "endDateOffset", title: "Срок" },
    { id: "endDateOffsetInterval", title: "Ед. изм. срока" },
    { id: "payoffStep", title: "Шаг погашения" },
    { id: "payoffInterval", title: "Ед. изм. погашения" },
    { id: "user", title: "Пользователь" },
    { id: "delete", title: "Удалить" }, // для чекбокса, не из API  
    { id: "modify", title: "Изменить" } // для чекбокса, не из API  
  ],

  //═══════════════════════════════════════════════════════════════════════════
  // СТРУКТУРА ПОЛЕЙ ДЛЯ МЕСТ
  // Описывает все возможные поля места
  //═══════════════════════════════════════════════════════════════════════════
  MERCHANT_FIELDS: [  
    { id: "id", title: "ID" },  
    { id: "title", title: "Название" },  
    { id: "delete", title: "Удалить" }, // для чекбокса, не из API  
    { id: "modify", title: "Изменить" }, // для чекбокса, не из API  
    { id: "user", title: "Пользователь" },
    { id: "changed", title: "Дата изменения" },
  ],

  //═══════════════════════════════════════════════════════════════════════════
  // ОСНОВНЫЕ ВАЛЮТЫ
  // Список поддерживаемых валют (можно расширять)
  //═══════════════════════════════════════════════════════════════════════════
  ALLOWED_CURRENCY_CODES: ['USD', 'EUR', 'RUB', 'UAH'],

  //═══════════════════════════════════════════════════════════════════════════
  // РЕЖИМЫ ОТОБРАЖЕНИЯ ТЕГОВ
  // Определяет, как отображаются теги: одной строкой или в нескольких колонках
  //═══════════════════════════════════════════════════════════════════════════
  TAG_MODES: {  
    SINGLE_COLUMN: "Одной строкой",   // Все теги в одной колонке
    MULTIPLE_COLUMNS: "Разделить"     // Теги в разных колонках
  },

  //═══════════════════════════════════════════════════════════════════════════
  // НАСТРОЙКИ ЭКСПОРТА
  // Цвета для различных состояний строк при экспорте изменений
  //═══════════════════════════════════════════════════════════════════════════
  EXPORT: {
    COLORS: {
      NEW: "#d9ead3",      // Зеленый - для новых записей
      WAS_NEW: "#e6ffe1",  // Нефритовый - для записей,бывших новыми на момент изменения, логика временно отключена
      MODIFIED: "#fff2cc", // Желтый - для измененных записей
      TO_DELETE: "#fad3ab",// Оранжевый - для записей на удаление
      DELETED: "#f4cccc"   // Красный - для удаленных записей
    }
  },

  //═══════════════════════════════════════════════════════════════════════════
  // НАСТРОЙКИ ИМПОРТА
  // Параметры пакетной обработки и обновления данных при импорте
  //═══════════════════════════════════════════════════════════════════════════
  IMPORT: {
    BATCH_SIZE: 100,       // Размер пакета для импорта, операций
    PROGRESS_INTERVAL: 100,// Интервал обновления прогресса, операций
    UPDATE_COLUMNS_LIST: [ // Список колонок для визуального обновления данных при импорте (не заменяет экспорт)
      'id', 'deleted', 'modified', 'created', 'changed', 'user'
    ],
  },

  //═══════════════════════════════════════════════════════════════════════════
  // ГЕТТЕРЫ ДЛЯ ПОЛУЧЕНИЯ ДАННЫХ ИЗ НАСТРОЕК
  // Автоматически получают значения из листа Settings и справочников
  //═══════════════════════════════════════════════════════════════════════════

  /**
   * Получает ID пользователя по умолчанию из настроек
   * @throws {Error} Если пользователь не задан или не найден
   * @returns {number} ID пользователя
   */
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
  
  /**
   * Получает ID валюты по умолчанию из настроек
   * @throws {Error} Если валюта не задана или не найдена
   * @returns {number} ID валюты
   */
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

  /**
   * Получает режим отображения тегов из настроек
   * @returns {string} Режим отображения тегов
   */
  get TagMode() {  
    const tagMode = sheetHelper.GetCellValue(this.SHEETS.SETTINGS.NAME, this.SHEETS.SETTINGS.CELLS.TAG_MODE);  
    return tagMode === this.TAG_MODES.MULTIPLE_COLUMNS ?   
      this.TAG_MODES.MULTIPLE_COLUMNS :   
      this.TAG_MODES.SINGLE_COLUMN;  
  }
};
