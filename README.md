## Script для работы с таблицами Excel

Что можно делать:
  1) Проверка строки на заглавную первую букву
  2) Поиск текста или слова в указаном столбце
  3) Добавление фрагмента текста в начало ячейки (если уже ячейка не пустая)
  4) Добавление фрагмента текста в каждую ячейку в конец
  5) Удалить фрагмент текста со всех ячейк в столбце
  6) Удаяем весь текст с одной ячееки и добавляем в другую ячейку
  7) Удаяем текст поиска с ячейки и добавляем в другую ячейку
  8) Обьяденяем данные столбцов в один столбец
  9) Сохранять файла с новым названием


Проверка строки на заглавную первую букву, если текст начинается с маленькой буквы или с символа то скрипт покажет  
```python
search.checking_is_title()
```

Поиск текста или слова в указаном столбце  
search_text = текст что ищем указываем списком ['мАч']  
```python
search.serch_text(text=["NFC"])
```

Добавление фрагмента текста в начало ячейки (если уже ячейка не пустая)  
cell_past - куда вставляем текст  
```python
change.add_text_to_cell_stert(cell_past='AD', text="мАч")
```

Добавление фрагмента текста в каждую ячейку в конец  
cell_past - куда вставляем текст  
```python
change.add_text_to_column(cell_past='AD', text="мАч")
```

Удалить фрагмент текста со всех ячейк в столбце  
cell_remove - откуда удаляем текст  
```python
change.remove_text_from_cell(cell_remove='AD', text="мАч")
```

Удаяем весь текст с одной ячееки и добавляем в другую ячейку  
cell_move - откуда вырезаем текст  
cell_past - куда вставляем текст  
```python
change.move_text_to_another_cell(cell_move='AE', cell_past='AD')
```

Удаяем текст поиска с ячейки и добавляем в другую ячейку
search_text = ['мАч'] или ['мАч', 'str'] текст что ищем указываем списком.  
В список можно указывать нужное количество аргументов  
cell_move - откуда вырезаем текст  
cell_past - куда вставляем текст  
method_remove - как вырезаем текст. если указать "str" то метод будет отрабатывать для строки (строка это каждый разделитьтель текста ";")
```python
change.serch_move_text_to_another_cell(cell_move="AD", cell_past="AG", method_remove='str', search=search_text)
```

Обьяденяем данные столбцов в один столбец  
save_column - указываем буквенную колонку куда хотим сохранить результат  
join_columns - колонки что нужно соеденить  
join_separator - если нужно соеденить текст между собой  
end_text = текст в конце  
```python
change.join_columns_text(save_column='AH', join_columns=['AE', 'AF', 'AG'], join_separator=' x ', end_text='см')
```

Сохранение файла с новым названием  
new_name - указываем имя дял нового файла  
```python
save.save(new_name="new file name")
```


Python 3.11.3

## Настройка  
1) Установить библиотеку poetry
```python
pip install poetry
```

2) Создаем virtualenv и устанавливаем зависимости
```python
poetry install
```

3) Активируем virtualenv
```python
poetry shell
```

## Запуск  
4) Для запуска, необходимо разкомментировать нужную функцию в файле "main.py"
    <b>строка 5</b>  
    read_document="указать свой файл - название"  
    sheet_name="название листа с которым работаете, обычно это - "Worksheet" как у примере ниже  
    ```python
    document = OpenDocument(read_document='smart-chasy.xlsx', sheet_name='Worksheet')
    ```
    <b>строка 10</b>  
    column_name="указываем название колонки" (для чтения)  
    ```python
    read_document = ReadDocument(document=document, column_name='Особенности-25849')
    ```

5) Выполнить команду для запуска скрипта
```python
python main.py
```
