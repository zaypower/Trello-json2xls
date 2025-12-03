# Trello JSON → Excel Exporter

Утилита для конвертации Trello-экспорта (`todo.json`) в Excel-файл, где  
**каждый Trello-список = отдельный лист**, с форматированием, нормализацией дат и автошириной колонок.

## Возможности
- Парсинг Trello `*.json` (экспорт board)
- Группировка карточек по спискам
- Создание отдельного листа Excel на каждый список
- Автоматическая ширина колонок
- Чередование строк (зебра)
- Заголовок с выделением
- Нормализация дат (`YYYY-MM-DD HH:MM:SS`)
- Очистка названий листов от запрещённых Excel-символов

## Пример структуры XLSX
To Do
shortId | name | dateLastActivity
Doing
shortId | name | dateLastActivity
Done
...
To Sell
...

## Требования
- Python 3.10+
- Библиотеки:
  ```bash
  pip install openpyxl

Использование

1 - Скачайте JSON-экспорт доски Trello:
Board → More → Print and Export → Export as JSON

2 - Поместите файл в папку с проектом под именем todo.json

3- Запустите:
python main.py

4 - Результат появится в файле:
todo_export.xlsx
