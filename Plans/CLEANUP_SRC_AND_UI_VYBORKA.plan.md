---
name: Очистка Src и UI Выборка
overview: Очистка папки Src от лишних файлов и палитр, переименование палитры «Таблица выделенного» в «Выборка», перенос кнопки «Отправить в Excel» с тулбара в палитру Выборка и замена HelpPalette на открытие URL в браузере.
---

# План: очистка Src, переименование в Выборка, перенос кнопки Excel

## 1. Удаление ненужных файлов из Src

**Оставляем только:**
- Ядро: Main.cpp, BrowserRepl.cpp/hpp, ResourceIDs.hpp, LicenseManager.cpp/hpp, APICommon.c/h, APIEnvir.h
- Палитра «Таблица выделенного» → «Выборка»: SelectionDetailsPalette.cpp/hpp
- Хелперы выборки: SelectionHelper, SelectionPropertyHelper, SelectionMetricsHelper
- Расчёт распила и GDL: GDLHelper.cpp/hpp

**Удаляем из Src:** HelpPalette, SendXlsPalette, BuildHelper, ColumnOrientHelper, GroundHelper, LandscapeHelper, LayerHelper, MarkupHelper, MeshIntersectionHelper, PropertyUtils, RandomizerHelper, RoadHelper, RotateHelper, ShellHelper, UserItemDialog (все .cpp/.hpp).

## 2. Удаление палитры Help и замена кнопки «Поддержка»

- Main.cpp: убрать HelpPalette include и RegisterPaletteControlCallBack.
- BrowserRepl.cpp: убрать HelpPalette; вместо ShowWithURL — открывать URL в браузере (ShellExecuteW).
- RINT/Browser_Repl.grc: удалить GDLG/DLGH 32510.
- ResourceIDs.hpp: удалить HelpPaletteResId, HelpBrowserCtrlId.

## 3. Удаление палитры SendXls и кнопки «Отправить в Excel» с тулбара

- Main.cpp: убрать SendXlsPalette.
- BrowserRepl.cpp: убрать OpenSendXlsPalette из JS, оставить SaveSendXls; убрать case ToolbarButtonExcelId.
- RINT/Browser_Repl.grc: три кнопки тулбара (Close, Выборка, Support); удалить GDLG/DLGH 32620.
- RFIX/Browser_ReplFix.grc: удалить DATA 300.
- ResourceIDs.hpp: удалить SendXls* ID, ToolbarButtonExcelId при необходимости.

## 4. Переименование палитры «Таблица выделенного» в «Выборка»

- RINT/Browser_Repl.grc: STR# 32501 и GDLG 32610 — заголовок «Выборка».
- RFIX/Selection_Details_Palette.html: title и h1 — «Выборка».

## 5. Кнопка «Отправить в Excel» в палитре Выборка

- Selection_Details_Palette.html: добавить кнопку «Отправить в Excel», по нажатию собрать CSV из таблицы и вызвать ACAPI.SaveSendXls(csv).

## 6. Порядок выполнения

1. Удалить лишние файлы из Src.
2. Убрать HelpPalette, Support → ShellExecute.
3. Убрать SendXlsPalette и кнопку Excel с тулбара.
4. Переименовать палитру в «Выборка».
5. Добавить кнопку «Отправить в Excel» в палитру Выборка.
6. Сборка и проверка.
