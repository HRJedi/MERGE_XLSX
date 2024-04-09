# MERGE_XLSX
## Скрипт для объединения однородных файлов XLSX
На вход принимает папку с набором XLSX файлов (предполагается одинаковая структура)\
На выходе один объединенный XLSX файл.\
Предназначался для автоматизации рабочей задачи - сведения учётных форм водителей-экспедиторов (Использовать PQ было не целесообразно по ряду причин).
___
**Реализовано:**
1. Удаление дубликатов строк и пустых строк;
2. Автоматический поиск и проверка заголовка документа (в т.ч. для многоуровневых заголовков);
3. Проверка соответствия заголовков эталонному и выявление ошибкок;
4. Возможность быстрой замены / исправления некорректных файлов - не прерывая проверку (проверка хеш.сумм);
5. Автоматический поиск позиции столбца-нумератора (если предусмотрен);
6. Работа с многостраничными документами.
