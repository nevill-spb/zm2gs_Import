const Dictionaries = (function () {
  const sheet = sheetHelper.GetSheetFromSettings('DICTIONARIES_SHEET');
  if (!sheet) {
    Logger.log("Лист справочников не найден");
    return;
  }

  // Внутренние объекты справочников
  let accounts = {};
  let merchants = {};
  let instruments = {};
  let users = {};
  let tags = {};

  // Обратные справочники для быстрого поиска ID по названию
  let accountsRev = {};
  let merchantsRev = {};
  let instrumentsRev = {};
  let usersRev = {};
  let tagsRev = {};

  // Получение обратных справочников
  function invertDictionary(dict) {
    const rev = {};
    for (const key in dict) {
      if (dict.hasOwnProperty(key)) {
        rev[dict[key]] = key;
      }
    }
    return rev;
  }

  // Сохранение справочников в лист
  function saveDictionariesToSheet() {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 3).setValues([["type", "id", "title"]]);

    // Вспомогательная функция для подготовки массива данных справочника
    const allowedCodes = new Set(Settings.ALLOWED_CURRENCY_CODES); // Валюты берём из листа настроек
    function prepareDictData(type, dict) {  
      if (type === "instruments") {
        return Object.entries(dict)  
          .filter(([id, title]) => allowedCodes.has(title))  
          .map(([id, title]) => [type, id, title]);  
      }  
      return Object.entries(dict).map(([id, title]) => [type, id, title]);  
    }

    // Собираем все данные в один массив
    const allData = [
      ...prepareDictData("accounts", accounts),
      ...prepareDictData("merchants", merchants),
      ...prepareDictData("instruments", instruments),
      ...prepareDictData("users", users),
      ...prepareDictData("tags", tags)
    ];

    if (allData.length > 0) {
      sheet.getRange(2, 1, allData.length, 3).setValues(allData);
    }
    /* 
    //Добавляем заголовки словарей в E1-H1  
    const dictTypes = ["accounts", "tags", "merchants", "instruments"];  
    const headerRange = sheet.getRange(1, 5, 1, dictTypes.length); // E1:H1  
    headerRange.setValues([dictTypes]);  
    
    // Добавляем формулы фильтрации в E2-H2  
    const formulas = dictTypes.map(type =>   
      `=IFERROR(FILTER(INDIRECT("'" & Settings!B8 & "'!C:C");INDIRECT("'" & Settings!B8 & "'!A:A") = "${type}");"Обновите справочники")`  
    );  
    const formulaRange = sheet.getRange(2, 5, 1, formulas.length); // E2:H2  
    formulaRange.setFormulas([formulas]);
    */
  }

  // Загрузка справочников из листа "Справочники"
  function loadDictionariesFromSheet() {
    try {
      const data = sheet.getDataRange().getValues();

      // Предполагается, что в листе есть заголовки и данные в формате Тип, ID, Название
      // Собираем справочники по типу
      accounts = {};
      merchants = {};
      instruments = {};
      users = {};
      tags = {};

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const type = row[0];
        const id = String(row[1]);
        const title = String(row[2]);
        if (!type || !id || !title) continue;

        switch (type.toLowerCase()) {
          case "accounts":
            accounts[id] = title;
            break;
          case "merchants":
            merchants[id] = title;
            break;
          case "instruments":
            instruments[id] = title;
            break;
          case "users":
            users[id] = title;
            break;
          case "tags":
            tags[id] = title;
            break;
          default:
            // Игнорируем неизвестные типы
            break;
        }
      }

      updateReverseDictionaries();
      return getAllDictionaries();  
    } catch (e) {  
      Logger.log("Ошибка загрузки справочников: " + e.message);  
      return null;  
    }
  }

  // Загрузка справочников из JSON (например, с API)
  function updateDictionaries(json) {
    if (!json) return;

    if (json.account) {
      accounts = {};
      json.account.forEach(item => {
        if (item.id && item.title) accounts[item.id] = item.title;
      });
    }

    if (json.merchant) {
      merchants = {};
      json.merchant.forEach(item => {
        if (item.id && item.title) merchants[item.id] = item.title;
      });
    }

    if (json.instrument) {
      instruments = {};
      json.instrument.forEach(item => {
        if (item.id && item.shortTitle) instruments[item.id] = item.shortTitle;
      });
    }

    if (json.user) {
      users = {};
      json.user.forEach(item => {
        if (item.id && item.login) users[item.id] = item.login;
      });    }

    if (json.tag) {
      const tagObjects = {};
      json.tag.forEach(({id, title, parent}) => {
        if (id && title) tagObjects[id] = {title, parent: parent || null};
      });

      const buildTagPath = (id) => {
        const tag = tagObjects[id];
        if (!tag) return "";
        return tag.parent ? buildTagPath(tag.parent) + " / " + tag.title : tag.title;
      };

      tags = Object.fromEntries(
        Object.keys(tagObjects).map(id => [id, buildTagPath(id)])
      );
    }
    updateReverseDictionaries();  
    return getAllDictionaries();
  }

  // Обновление обратных справочников
  function updateReverseDictionaries() {
    accountsRev = invertDictionary(accounts);
    merchantsRev = invertDictionary(merchants);
    instrumentsRev = invertDictionary(instruments);
    usersRev = invertDictionary(users);
    tagsRev = invertDictionary(tags);
  }

  // Получение ID по названию
  function getAccountId(title) {
    return accountsRev[title] || null;
  }

  function getMerchantId(title) {
    return merchantsRev[title] || null;
  }

  function getInstrumentId(title) {
    return Number(instrumentsRev[title]) || null;
  }

  function getUserId(login) {
    return Number(usersRev[login]) || null;
  }

  function getTagId(title) {
    return tagsRev[title] || null;
  }

  // Получение названия по ID
  function getAccountTitle(id) {
    return accounts[id] || null;
  }

  function getMerchantTitle(id) {
    return merchants[id] || null;
  }

  function getInstrumentShortTitle(id) {
    return instruments[id] || null;
  }

  function getUserLogin(id) {
    return users[id] || null;
  }

  function getTagTitle(id) {
    return tags[id] || null;
  }

  // Возвращает все справочники (для загрузки в Import.DICTIONARIES)
  function getAllDictionaries() {
    return {
      accounts,
      merchants,
      instruments,
      users,
      tags
    };
  }

  // Возвращает все обратные справочники (для загрузки в Import.DICTIONARIES)
  function getAllReverseDictionaries () {  
    return {  
      accountsRev,  
      merchantsRev,  
      instrumentsRev,  
      usersRev,  
      tagsRev  
    };  
  }

  return {
    updateDictionaries,
    saveDictionariesToSheet,
    loadDictionariesFromSheet,
    getUserId,
    getInstrumentId,
    getAccountId,
    getMerchantId,
    getTagId,
    getUserLogin,
    getInstrumentShortTitle,
    getAccountTitle,
    getMerchantTitle,
    getTagTitle,
    getAllDictionaries,
    getAllReverseDictionaries,
  };
})();
