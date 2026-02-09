# Gomer Sheet Bridge
### Chrome extension for Gomer -> Google Sheets

Расширение добавляет кнопку рядом с параметрами товара в Gomer, ищет товар в прайсе и записывает строку в Google Sheets.

## Что умеет
- Добавляет `Excel`-кнопки на страницах Gomer.
- Собирает данные товара и параметра из разных контекстов:
`on-moderation`, `active`, `changes`, `item-details`, `binding-attribute-page`.
- Ищет товар в XML-прайсе:
по `offerId + paramName + paramValue` (строгий матч), либо по `categoryId + param`.
- Записывает строку в Google Sheets через API.
- Поддерживает ссылки в ячейках через `HYPERLINK(...)` и автоформат (синий/подчеркнутый текст).
- Сохраняет номер задачи и контекст между страницами binding flow.
- Корректно работает, когда прайс недоступен:
строка добавляется без ссылки на товар.

## Архитектура
- `manifest.json`:
MV3 manifest, permissions, popup, background, content script.
- `src/content/content.js`:
инъекция кнопок, сбор данных из DOM, UI/тосты, отправка сообщений в background.
- `src/background/background.js`:
OAuth (Chrome Identity), поиск офферов в XML, запись и форматирование строк в Google Sheets.
- `src/popup/popup.html` и `src/popup/popup.js`:
сохранение URL таблицы в `chrome.storage.sync`.

## Требования
- Google Chrome (Developer mode).
- Google Cloud проект с включенным `Google Sheets API`.
- OAuth Client ID для Chrome Extension.
- Доступ к Gomer и прайсам поставщика.

## Установка
1. Открой `chrome://extensions`.
2. Включи `Developer mode`.
3. Нажми `Load unpacked`.
4. Выбери папку:
`Gomer Sheet Bridge/`.
5. Скопируй `Extension ID` (нужен для OAuth в Google Cloud).

## Как получить `extension.pem` (стабильный Extension ID)
`extension.pem` нужен, чтобы ID расширения не менялся между перезагрузками и машинами.

### Вариант 1: через Chrome (рекомендуется)
1. Открой `chrome://extensions`.
2. Нажми `Pack extension`.
3. В `Extension root directory` укажи папку проекта `Gomer Sheet Bridge/`.
4. Поле `Private key file` оставь пустым при первой генерации.
5. Chrome создаст `.pem` и `.crx`.
6. Положи `.pem` в корень проекта как `extension.pem` и не коммить его в Git.

### Вариант 2: через `openssl`
```bash
cd "/Users/artem/Rozetka/Params/Gomer Sheet Bridge"
openssl genrsa -out extension.pem 2048
openssl rsa -in extension.pem -pubout -outform DER | openssl base64 -A
```
Вторая команда выведет публичный ключ в base64. Его нужно вставить в `manifest.json` в поле `key`.

### Важно
- Если `extension.pem` или `manifest.json -> key` поменяются, изменится `Extension ID`.
- После смены ID обнови OAuth-настройки в Google Cloud (разрешенный ID расширения).

## Настройка Google Cloud
1. Создай или открой проект в Google Cloud.
2. Включи `Google Sheets API`.
3. Настрой `OAuth consent screen`.
4. Создай OAuth клиент для Chrome Extension и привяжи `Extension ID`.
5. Проверь, что scope включает:
`https://www.googleapis.com/auth/spreadsheets`.

## Первый запуск
1. Открой popup расширения.
2. Вставь ссылку на Google Sheet.
3. Нажми `Сохранить`.
4. Открой страницу Gomer с параметрами и нажми кнопку `Excel` рядом со значением.

## Как работает запись
1. Контент-скрипт читает категорию, атрибут, значение, номер задачи.
2. Фоновый скрипт ищет `offerId` в прайсе.
3. Формируется ссылка на товар (если найдена).
4. В Google Sheets добавляется новая строка с данными.

### Колонки, которые ожидаются в таблице
Названия колонок в первой строке должны совпадать:
- `Дата добавления`
- `Категория товара`
- `Атрибут`
- `Значение параметра`
- `Номер задачи`
- `Ссылка на товар`

## Особенности и ограничения
- Поиск в прайсе чувствителен к имени параметра и значению после нормализации.
- Если у источника `Api virtual source` или `javascript:void(0)`, товарный URL не строится.
- Для стабильного `Extension ID` локально используется `extension.pem`.
