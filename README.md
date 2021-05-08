# OLE2Viewer
==========================================================================
                     Плагин FAR 3.0 OLE2Viewer 1.0.10

                   Авторские права (с) 2016-2019 rivitna
==========================================================================

Просмотр и анализ структуры составных файлов OLE2 (compound files OLE2): doc, xls, ppt, vsd, msi и др.
Для просмотра составных файлов OLE2 используется их обработка на уровне бинарного формата файлов, что позволяет обрабатывать и запароленные файлы.

Плагин позволяет:
- извлекать потоки, каталоги;
- извлекать текст документов (для doc-файлов из потока "WordDocument");
- извлекать исходные коды макросов VBA ("_VBA_PROJECT_CUR\VBA" - MS Excel, "Macros\VBA" - MS Word);
- извлекать метаданные (из потоков "♣DocumentSummaryInformation", "♣SummaryInformation");
- извлекать вложенные объекты (из потока "☺Ole10Native");
- просматривать "потерянные" (не имеющие родительской записи) записи каталога;
- отображать следующую информацию о файле в информационной панели:
  • версия файла;
  • тип файла ("OLE2 Unknown", "Microsoft Word", "Microsoft Excel", "Microsoft PowerPoint", "Microsoft Visio", "Microsoft Excel 5", "Ichitaro", "MSI/MSP/MSM");
  • размер файла (размер вычисляется исходя из структуры файла, так как возможен оверлей);
  • общее количество потоков:
  • общее количество каталогов;
  • общее количество макросов VBA.

Для открываемых потоков (потоков-контейнеров) может быть использована раскраска по атрибутам "Сжатый" ("Compressed") и "Точка повторного анализа" ("Reparse point"). Для этого необходимо создать соответствующую группу раскраски файлов. Для однообразия в отображении файлов-контейнеров целесообразно дублировать группу для маски "<arch>", удалить в новой группе маску и добавить указанные выше атрибуты включения.

