const Tracking = (function() {
    // Кэшируем часто используемые значения при инициализации
    const SHEET_NAME = sheetHelper.GetSheetFromSettings('IMPORT_SHEET').getName();
    
    // Получаем индексы отслеживаемых столбцов из TRANSACTION_FIELDS
    const WATCHED_FIELDS = [
        'date',
        'tag',
        'tag1',
        'tag2',
        'merchant',
        'comment',
        'outcomeAccount',
        'outcome',
        'incomeAccount',
        'income'
    ];

    const WATCHED_COLUMNS = WATCHED_FIELDS.map(fieldId => {
        const index = Settings.TRANSACTION_FIELDS.findIndex(f => f.id === fieldId);
        if (index === -1) {
            Logger.log(`Поле ${fieldId} не найдено в TRANSACTION_FIELDS`);
            return null;
        }
        return index + 1;
    }).filter(col => col !== null);

    // Создаем Set для быстрой проверки колонок
    const WATCHED_COLUMNS_SET = new Set(WATCHED_COLUMNS);

    // Кэшируем индексы важных колонок
    const DATE_COLUMN = Settings.TRANSACTION_FIELDS.findIndex(f => f.id === 'date') + 1;
    const MODIFIED_COLUMN = Settings.TRANSACTION_FIELDS.findIndex(f => f.id === 'modified') + 1;

    const PROP_KEY = 'TRACKING_ENABLED';
    let trackingEnabled = false;

    // Оптимизированная проверка колонок
    function isWatchedColumn(col) {
        return WATCHED_COLUMNS_SET.has(col);
    }

    function logDebug(message, data = null) {
        Logger.log(`[Tracking] ${message}`);
        if (data) Logger.log(JSON.stringify(data));
    }

    function isTrackingEnabled() {
        return trackingEnabled;
    }

    function createTriggers() {
        try {
            deleteTriggers(true);

            ScriptApp.newTrigger('Tracking.onEdit')
                .forSpreadsheet(SpreadsheetApp.getActive())
                .onEdit()
                .create();

            ScriptApp.newTrigger('Tracking.onChange')
                .forSpreadsheet(SpreadsheetApp.getActive())
                .onChange()
                .create();

            trackingEnabled = true;
            PropertiesService.getScriptProperties().setProperty(PROP_KEY, 'true');
            SpreadsheetApp.getActive().toast('Триггеры успешно установлены!', 'Успех');
            
            logDebug('Триггеры установлены');
        } catch (error) {
            logDebug('Ошибка при установке триггеров', error);
            SpreadsheetApp.getActive().toast(`Ошибка при установке триггеров: ${error.toString()}`, 'Ошибка');
        }
    }

    function deleteTriggers(silent = false) {  
        try {  
            const triggers = ScriptApp.getProjectTriggers();  
            triggers.forEach(trigger => {  
                if (trigger.getHandlerFunction().startsWith('Tracking.')) {  
                    ScriptApp.deleteTrigger(trigger);  
                }  
            });  
      
            trackingEnabled = false;  
            PropertiesService.getScriptProperties().setProperty(PROP_KEY, 'false');  
              
            if (!silent) {  
                SpreadsheetApp.getActive().toast('Триггеры успешно отключены!', 'Успех');  
            }  
              
            logDebug('Триггеры удалены');  
        } catch (error) {  
            logDebug('Ошибка при удалении триггеров', error);  
            if (!silent) {  
                SpreadsheetApp.getActive().toast(`Ошибка при удалении триггеров: ${error.toString()}`, 'Ошибка');  
            }  
        }  
    }

    function onEdit(e) {
        try {
            if (!isTrackingEnabled()) return;
            if (!e || !e.range) return;
            
            logDebug('onEdit вызван', {
                sheet: e.range.getSheet().getName(),
                range: e.range.getA1Notation()
            });
            
            handleChange(e.range);
        } catch (error) {
            logDebug('Ошибка в onEdit', error);
        }
    }

    function onChange(e) {
        try {
            if (!isTrackingEnabled()) return;
            if (!e) return;
            
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            if (sheet.getName() !== SHEET_NAME) return;

            const range = sheet.getActiveRange();
            if (range) {
                logDebug('onChange вызван', {
                    sheet: sheet.getName(),
                    range: range.getA1Notation()
                });
                
                handleChange(range);
            }
        } catch (error) {
            logDebug('Ошибка в onChange', error);
        }
    }

    function handleChange(range) {
        try {
            const sheet = range.getSheet();
            if (sheet.getName() !== SHEET_NAME) return;

            const startRow = range.getRow();
            const endRow = range.getLastRow();
            const startCol = range.getColumn();
            const endCol = range.getLastColumn();

            // Быстрая проверка на отслеживаемые столбцы
            let needsModification = false;
            for (let col = startCol; col <= endCol; col++) {
                if (isWatchedColumn(col)) {
                    needsModification = true;
                    break;
                }
            }
            if (!needsModification) return;

            // Получаем только колонку с датами вместо всех данных
            const dates = sheet.getRange(startRow, DATE_COLUMN, endRow - startRow + 1, 1).getValues();

            // Собираем все строки для модификации в один массив
            const rowsToModify = [];
            const valuesToSet = [];
            
            dates.forEach((dateRow, index) => {
                const date = dateRow[0];
                if (date !== "" && date !== null && date !== undefined) {
                    rowsToModify.push(startRow + index);
                    valuesToSet.push([true]);
                }
            });

            // Если есть строки для модификации, устанавливаем флаги одним запросом
            if (rowsToModify.length > 0) {
                // Группируем последовательные строки для оптимизации
                let currentGroup = {
                    start: rowsToModify[0],
                    count: 1,
                    values: [valuesToSet[0]]
                };
                const groups = [currentGroup];

                for (let i = 1; i < rowsToModify.length; i++) {
                    if (rowsToModify[i] === currentGroup.start + currentGroup.count) {
                        // Строка последовательная, добавляем в текущую группу
                        currentGroup.count++;
                        currentGroup.values.push(valuesToSet[i]);
                    } else {
                        // Строка не последовательная, создаем новую группу
                        currentGroup = {
                            start: rowsToModify[i],
                            count: 1,
                            values: [valuesToSet[i]]
                        };
                        groups.push(currentGroup);
                    }
                }

                // Устанавливаем значения для каждой группы одним запросом
                groups.forEach(group => {
                    sheet.getRange(group.start, MODIFIED_COLUMN, group.count, 1)
                         .setValues(group.values);
                });
            }
        } catch (error) {
            logDebug('Ошибка в handleChange', error);
        }
    }

    function addTrackingMenuItems(importMenu) {    
        return importMenu    
            .addSeparator()  
            .addItem("Установить триггеры отслеживания", "Tracking.createTriggers")    
            .addItem("Отключить триггеры отслеживания", "Tracking.deleteTriggers");    
    }

    // Инициализация при загрузке модуля
    trackingEnabled = PropertiesService.getScriptProperties().getProperty(PROP_KEY) === 'true';

    return {    
        createTriggers,  
        deleteTriggers,  
        onEdit,  
        onChange,  
        addTrackingMenuItems  
    };
})();
