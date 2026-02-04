




'STR#' 32000 "Add-on Name and Description" {
		"ArchPlanks AC29"
		"План распила и пиломатериалы"
}

'STR#' 32501 "Strings for the Menu" {
		"227.info"
		"ArchPlanks"
		"Выборка"
		"Поддержка"
		"Toolbar"
}



'GDLG'  32500    Palette | leftCaption | noGrow   0   0  100  32  ""  {
 IconButton          3    1   30   30    32111									
 IconButton         33    1   30   30    32100									
 IconButton         63    1   30   30    32110									
}

'DLGH'  32500  DLG_32500_Browser_Repl {
1	"Close"					IconButton_0
2	"Выборка"				IconButton_1
3	"Support"				IconButton_2
}



'GDLG'  32610    Palette | topCaption | close | grow   0   0  400  450  "Выборка"  {
 Browser         0   0  400  450
}

'DLGH'  32610  DLG_32610_SelectionDetails_Browser {
1   "Выборка Browser Control"     Browser_0
}
