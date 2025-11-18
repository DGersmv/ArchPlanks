#ifndef RESOURCEIDS_HPP
#define RESOURCEIDS_HPP

#define BrowserReplResId        32500
#define BrowserReplMenuResId    32501
#define BrowserReplMenuItemIndex 1  // "Toolbar" item in menu (index 1 in STR# 32501)

// Icon button IDs for native toolbar (IDs start from 1 in .grc)
#define ToolbarButtonCloseId    1
#define ToolbarButtonTableId    2
#define ToolbarButtonSplineId   3
#define ToolbarButtonRotateId   4
#define ToolbarButtonRotSurfId  5
#define ToolbarButtonLandId     6
#define ToolbarButtonDimsId     7
#define ToolbarButtonLayersId   8
#define ToolbarButtonContourId  9
#define ToolbarButtonMeshId     10
#define ToolbarButtonCSVId      11
#define ToolbarButtonSupportId  12

// Icon resource IDs
#define IconTableResId          32100
#define IconSplineResId         32101
#define IconRotateResId         32102
#define IconRotSurfResId        32103
#define IconLandResId           32104
#define IconDimsResId           32105
#define IconLayersResId         32106
#define IconContourResId        32107
#define IconMeshResId           32108
#define IconCSVResId            32109
#define IconSupportResId        32110
#define IconCloseResId          32111

constexpr short HelpPaletteResId = 32510; // 'GDLG' новой палитры
constexpr short HelpBrowserCtrlId = 1;     // ID Browser-контрола внутри неё

constexpr short DistributionPaletteResId = 32520;
constexpr short DistributionBrowserCtrlId = 1;
constexpr short DistributionHtmlResId = 110;

constexpr short OrientationPaletteResId = 32530;
constexpr short OrientationBrowserCtrlId = 1;
constexpr short OrientationHtmlResId = 120;

constexpr short GroundPaletteResId = 32540;
constexpr short GroundBrowserCtrlId = 1;
constexpr short GroundHtmlResId = 130;

constexpr short MarkupPaletteResId = 32550;
constexpr short MarkupBrowserCtrlId = 1;
constexpr short MarkupHtmlResId = 140;

constexpr short ContourPaletteResId = 32560;
constexpr short ContourBrowserCtrlId = 1;
constexpr short ContourHtmlResId = 150;

constexpr short MeshPaletteResId = 32570;
constexpr short MeshBrowserCtrlId = 1;
constexpr short MeshHtmlResId = 160;

constexpr short IdLayersPaletteResId = 32580;
constexpr short IdLayersBrowserCtrlId = 1;
constexpr short IdLayersHtmlResId = 170;

constexpr short AnglePaletteResId = 32590;
constexpr short AngleBrowserCtrlId = 1;
constexpr short AngleHtmlResId = 180;

constexpr short SendXlsPaletteResId = 32600;
constexpr short SendXlsBrowserCtrlId = 1;
constexpr short SendXlsHtmlResId = 190;

constexpr short SelectionDetailsPaletteResId = 32610;
constexpr short SelectionDetailsBrowserCtrlId = 1;
constexpr short SelectionDetailsHtmlResId = 200;

#endif
