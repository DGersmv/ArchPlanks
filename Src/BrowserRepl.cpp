#include "ACAPinc.h"
#include "APICommon.h"
#include "BrowserRepl.hpp"
#include "DGBrowser.hpp"
#include "LicenseManager.hpp"

#include <Windows.h>
#include "SelectionHelper.hpp"

// Внешние функции для проверки состояния лицензии
extern "C" {
	bool IsLicenseValid();
	bool IsDemoExpired();
}
#include "SelectionPropertyHelper.hpp"
#include "SelectionMetricsHelper.hpp"
#include "SelectionDetailsPalette.hpp"

#include <commdlg.h>
#include <cmath>
#include <cstdio>
#include <cstring>
#include <vector>

// --------------------- Palette GUID / Instance ---------------------
static const GS::Guid paletteGuid("{22a3b4c5-6d7e-8f9a-0b1c-2d3e4f5a6b7c}");
GS::Ref<BrowserRepl> BrowserRepl::instance;

// HTML/JavaScript helpers removed - using native buttons now

// --- Extract double from JS::Base (supports 123 / "123.4" / "123,4") ---
// Общий хелпер: вытащить double из JS::Base (поддерживает number, string "3,5"/"3.5", bool)
static double GetDoubleFromJs(GS::Ref<JS::Base> p, double def = 0.0)
{
	if (p == nullptr) {
		return def;
	}

	if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(p)) {
		const auto t = v->GetType();

		if (t == JS::Value::DOUBLE) {
			return v->GetDouble();
		}
		
		if (t == JS::Value::INTEGER) {
			// Для INTEGER используем GetDouble(), но также можно попробовать GetInteger() если есть
			return static_cast<double>(v->GetInteger());
		}

		if (t == JS::Value::STRING) {
			GS::UniString s = v->GetString();
			for (UIndex i = 0; i < s.GetLength(); ++i) if (s[i] == ',') s[i] = '.';
			double out = def;
			std::sscanf(s.ToCStr().Get(), "%lf", &out);
			return out;
		}
	}

	return def;
}

// --- Extract integer from JS::Base (supports 123 / "123") ---
static Int32 GetIntFromJs(GS::Ref<JS::Base> p, Int32 def = 0)
{
	if (p == nullptr) {
		return def;
	}

	if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(p)) {
		const auto t = v->GetType();

		if (t == JS::Value::INTEGER) {
			return v->GetInteger();
		}
		
		if (t == JS::Value::DOUBLE) {
			return static_cast<Int32>(v->GetDouble());
		}

		if (t == JS::Value::STRING) {
			GS::UniString s = v->GetString();
			Int32 out = def;
			std::sscanf(s.ToCStr().Get(), "%d", &out);
			return out;
		}
	}

	return def;
}


static GS::UniString GetStringFromJavaScriptVariable(GS::Ref<JS::Base> jsVariable)
{
	GS::Ref<JS::Value> jsValue = GS::DynamicCast<JS::Value>(jsVariable);
	if (DBVERIFY(jsValue != nullptr && jsValue->GetType() == JS::Value::STRING))
		return jsValue->GetString();
	return GS::EmptyUniString;
}

// --- Save CSV with Windows Save dialog (UTF-8 with BOM) ---
static bool SaveCsvWithDialog(const GS::UniString& csvContent)
{
	wchar_t pathBuf[MAX_PATH] = L"";
	OPENFILENAMEW ofn = {};
	ofn.lStructSize = sizeof(ofn);
	ofn.lpstrFilter = L"CSV files (*.csv)\0*.csv\0All files (*.*)\0*.*\0";
	ofn.lpstrFile = pathBuf;
	ofn.nMaxFile = MAX_PATH;
	ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
	ofn.lpstrDefExt = L"csv";
	if (!GetSaveFileNameW(&ofn))
		return false;

	FILE* fp = _wfopen(pathBuf, L"wb");
	if (!fp)
		return false;

	// UTF-8 BOM
	const unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
	if (fwrite(bom, 1, 3, fp) != 3) {
		fclose(fp);
		return false;
	}

	const int wideLen = (int)csvContent.GetLength();
	if (wideLen > 0) {
		const wchar_t* wstr = csvContent.ToUStr().Get();
		int utf8Size = WideCharToMultiByte(CP_UTF8, 0, wstr, wideLen, NULL, 0, NULL, NULL);
		if (utf8Size > 0) {
			std::vector<char> buf((size_t)utf8Size);
			WideCharToMultiByte(CP_UTF8, 0, wstr, wideLen, buf.data(), utf8Size, NULL, NULL);
			if (fwrite(buf.data(), 1, (size_t)utf8Size, fp) != (size_t)utf8Size) {
				fclose(fp);
				return false;
			}
		}
	}
	fclose(fp);
	return true;
}

// --- Extract array of strings (GUIDs) from JS::Base ---
static GS::Array<GS::UniString> GetStringArrayFromJavaScriptVariable(GS::Ref<JS::Base> jsVariable)
{
	GS::Array<GS::UniString> result;
	GS::Ref<JS::Array> jsArray = GS::DynamicCast<JS::Array>(jsVariable);
	if (jsArray == nullptr)
		return result;
	
	// JS::Array имеет метод GetItemArray(), который возвращает const GS::Array<GS::Ref<Base>>&
	const GS::Array<GS::Ref<JS::Base>>& items = jsArray->GetItemArray();
	for (UIndex i = 0; i < items.GetSize(); ++i) {
		GS::Ref<JS::Base> item = items[i];
		if (item != nullptr) {
			GS::Ref<JS::Value> jsValue = GS::DynamicCast<JS::Value>(item);
			if (jsValue != nullptr && jsValue->GetType() == JS::Value::STRING) {
				result.Push(jsValue->GetString());
			}
		}
	}
	return result;
}

template<class Type>
static GS::Ref<JS::Base> ConvertToJavaScriptVariable(const Type& cppVariable)
{
	return new JS::Value(cppVariable);
}

template<>
GS::Ref<JS::Base> ConvertToJavaScriptVariable(const SelectionHelper::ElementInfo& elemInfo)
{
	GS::Ref<JS::Array> js = new JS::Array();
	js->AddItem(ConvertToJavaScriptVariable(elemInfo.guidStr));
	js->AddItem(ConvertToJavaScriptVariable(elemInfo.typeName));
	js->AddItem(ConvertToJavaScriptVariable(elemInfo.elemID));
	js->AddItem(ConvertToJavaScriptVariable(elemInfo.layerName));
	return js;
}

template<class Type>
static GS::Ref<JS::Base> ConvertToJavaScriptVariable(const GS::Array<Type>& cppArray)
{
	GS::Ref<JS::Array> newArray = new JS::Array();
	for (const Type& item : cppArray) {
		newArray->AddItem(ConvertToJavaScriptVariable(item));
	}
	return newArray;
}

static void EnsureModelWindowIsActive()
{
	API_WindowInfo windowInfo = {};
	const GSErrCode getDbErr = ACAPI_Database_GetCurrentDatabase(&windowInfo);
	if (DBERROR(getDbErr != NoError))
		return;

	// Не пытаемся активировать собственные кастомные окна
	if (windowInfo.typeID == APIWind_MyDrawID || windowInfo.typeID == APIWind_MyTextID)
		return;

	const GSErrCode changeErr = ACAPI_Window_ChangeWindow(&windowInfo);
#ifdef DEBUG_UI_LOGS
	if (changeErr != NoError) {
		ACAPI_WriteReport("[BrowserRepl] ACAPI_Window_ChangeWindow failed, err=%d", false, (int)changeErr);
	}
#else
	(void)changeErr;
#endif
}

// --------------------- Project event handler ---------------------
static GSErrCode NotificationHandler(API_NotifyEventID notifID, Int32 /*param*/)
{
	if (notifID == APINotify_Quit) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[BrowserRepl] APINotify_Quit to DestroyInstance", false);
#endif
		BrowserRepl::DestroyInstance();
	}
	return NoError;
}

// --------------------- BrowserRepl impl ---------------------
BrowserRepl::BrowserRepl() :
	DG::Palette(ACAPI_GetOwnResModule(), BrowserReplResId, ACAPI_GetOwnResModule(), paletteGuid),
	buttonClose(GetReference(), ToolbarButtonCloseId),
	buttonTable(GetReference(), ToolbarButtonTableId),
	buttonSupport(GetReference(), ToolbarButtonSupportId)
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] ctor with native buttons", false);
#endif
	ACAPI_ProjectOperation_CatchProjectEvent(APINotify_Quit, NotificationHandler);

	Attach(*this);
	AttachToAllItems(*this);
	BeginEventProcessing();
}

BrowserRepl::~BrowserRepl()
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] dtor", false);
#endif
	EndEventProcessing();
}

bool BrowserRepl::HasInstance() { return instance != nullptr; }

void BrowserRepl::CreateInstance()
{
	DBASSERT(!HasInstance());
	instance = new BrowserRepl();
	ACAPI_KeepInMemory(true);
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] CreateInstance", false);
#endif
}

BrowserRepl& BrowserRepl::GetInstance()
{
	DBASSERT(HasInstance());
	return *instance;
}

void BrowserRepl::DestroyInstance() { instance = nullptr; }

void BrowserRepl::Show()
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] Show", false);
#endif
	DG::Palette::Show();
	SetMenuItemCheckedState(true);
}

void BrowserRepl::Hide()
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] Hide", false);
#endif
	DG::Palette::Hide();
	SetMenuItemCheckedState(false);
}

// Browser-related methods removed - using native buttons now

// JavaScript API registration - still needed for other palettes (DistributionPalette, etc.)
// but not used by the main toolbar anymore
void BrowserRepl::RegisterACAPIJavaScriptObject(DG::Browser& targetBrowser)
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] RegisterACAPIJavaScriptObject", false);
#endif

	JS::Object* jsACAPI = new JS::Object("ACAPI");

	// --- Selection API ---
	jsACAPI->AddItem(new JS::Function("GetSelectedElements", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] GetSelectedElements()");
		GS::Array<SelectionHelper::ElementInfo> elements = SelectionHelper::GetSelectedElements();
		// if (BrowserRepl::HasInstance()) {
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[C++] GetSelectedElements вернул %d элементов", (int)elements.GetSize()));
		//	for (UIndex i = 0; i < elements.GetSize(); ++i) {
		//		BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[C++] Элемент %d: GUID=%s, Type=%s, ID=%s, Layer=%s", 
		//			(int)i, elements[i].guidStr.ToCStr().Get(), elements[i].typeName.ToCStr().Get(), 
		//			elements[i].elemID.ToCStr().Get(), elements[i].layerName.ToCStr().Get()));
		//	}
		// }
		return ConvertToJavaScriptVariable(elements);
		}));

	jsACAPI->AddItem(new JS::Function("AddElementToSelection", [](GS::Ref<JS::Base> param) {
		const GS::UniString id = GetStringFromJavaScriptVariable(param);
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] AddElementToSelection " + id);
		SelectionHelper::ModifySelection(id, SelectionHelper::AddToSelection);
		return ConvertToJavaScriptVariable(true);
		}));

	jsACAPI->AddItem(new JS::Function("RemoveElementFromSelection", [](GS::Ref<JS::Base> param) {
		const GS::UniString id = GetStringFromJavaScriptVariable(param);
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] RemoveElementFromSelection " + id);
		SelectionHelper::ModifySelection(id, SelectionHelper::RemoveFromSelection);
		return ConvertToJavaScriptVariable(true);
		}));

	jsACAPI->AddItem(new JS::Function("ChangeSelectedElementsID", [](GS::Ref<JS::Base> param) {
		const GS::UniString baseID = GetStringFromJavaScriptVariable(param);
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] ChangeSelectedElementsID " + baseID);
		const bool success = SelectionHelper::ChangeSelectedElementsID(baseID);
		return ConvertToJavaScriptVariable(success);
		}));

	jsACAPI->AddItem(new JS::Function("ApplyCheckedSelection", [](GS::Ref<JS::Base> param) {
		GS::Array<GS::UniString> guidStrings = GetStringArrayFromJavaScriptVariable(param);
		// if (BrowserRepl::HasInstance()) {
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyCheckedSelection: %d GUIDs", (int)guidStrings.GetSize()));
		// }
		
		GS::Array<API_Guid> guids;
		for (UIndex i = 0; i < guidStrings.GetSize(); ++i) {
			API_Guid guid = APIGuidFromString(guidStrings[i].ToCStr().Get());
			if (guid != APINULLGuid) {
				guids.Push(guid);
			}
		}
		
		SelectionHelper::ApplyCheckedSelectionResult result = SelectionHelper::ApplyCheckedSelection(guids);
		
		// Возвращаем объект { applied: N, requested: M }
		GS::Ref<JS::Object> jsResult = new JS::Object();
		jsResult->AddItem("applied", ConvertToJavaScriptVariable((Int32)result.applied));
		jsResult->AddItem("requested", ConvertToJavaScriptVariable((Int32)result.requested));

		EnsureModelWindowIsActive();

		return jsResult;
		}));

	jsACAPI->AddItem(new JS::Function("GetSelectedProperties", [](GS::Ref<JS::Base> param) {
		API_Guid requestedGuid = APINULLGuid;
		if (param != nullptr) {
			GS::UniString guidStr = GetStringFromJavaScriptVariable(param);
			if (!guidStr.IsEmpty()) {
				requestedGuid = APIGuidFromString(guidStr.ToCStr().Get());
			}
		}

		GS::Array<SelectionPropertyHelper::PropertyInfo> props = (requestedGuid == APINULLGuid)
			? SelectionPropertyHelper::CollectForFirstSelected()
			: SelectionPropertyHelper::CollectForGuid(requestedGuid);
		GS::Ref<JS::Array> jsProps = new JS::Array();
		for (const auto& info : props) {
			GS::Ref<JS::Object> obj = new JS::Object();
			obj->AddItem("guid", new JS::Value(APIGuidToString(info.propertyGuid)));
			obj->AddItem("name", new JS::Value(info.propertyName));
			obj->AddItem("value", new JS::Value(info.valueString));
			jsProps->AddItem(obj);
		}
		return jsProps;
	}));

	jsACAPI->AddItem(new JS::Function("GetSelectionSeoMetrics", [](GS::Ref<JS::Base> param) {
		API_Guid requestedGuid = APINULLGuid;
		if (param != nullptr) {
			GS::UniString guidStr = GetStringFromJavaScriptVariable(param);
			if (!guidStr.IsEmpty()) {
				requestedGuid = APIGuidFromString(guidStr.ToCStr().Get());
			}
		}

		GS::Array<SelectionMetricsHelper::Metric> metrics = (requestedGuid == APINULLGuid)
			? SelectionMetricsHelper::CollectForFirstSelected()
			: SelectionMetricsHelper::CollectForGuid(requestedGuid);
		GS::Ref<JS::Array> jsMetrics = new JS::Array();
		for (const auto& metric : metrics) {
			GS::Ref<JS::Object> obj = new JS::Object();
			obj->AddItem("key", new JS::Value(metric.key));
			obj->AddItem("name", new JS::Value(metric.name));
			obj->AddItem("grossValue", new JS::Value(metric.grossValue));
			obj->AddItem("netValue", new JS::Value(metric.netValue));
			obj->AddItem("diffValue", new JS::Value(metric.diffValue));
			jsMetrics->AddItem(obj);
		}
		return jsMetrics;
	}));

	// --- Help / Palettes ---
	jsACAPI->AddItem(new JS::Function("OpenHelp", [](GS::Ref<JS::Base> param) {
		GS::UniString url;
		if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
			if (v->GetType() == JS::Value::STRING) url = v->GetString();
		}
		if (url.IsEmpty()) url = "https://landscape.227.info/help/start";
		ShellExecuteW(NULL, L"open", url.ToUStr().Get(), NULL, NULL, SW_SHOWNORMAL);
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenSelectionDetailsPalette", [](GS::Ref<JS::Base>) {
		SelectionDetailsPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("SaveSendXls", [](GS::Ref<JS::Base> param) -> GS::Ref<JS::Base> {
		GS::UniString csv = GetStringFromJavaScriptVariable(param);
		bool ok = SaveCsvWithDialog(csv);
		if (ok)
			ACAPI_WriteReport("CSV exported successfully.", false);
		return new JS::Value(ok);
		}));

	jsACAPI->AddItem(new JS::Function("ClosePalette", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance() && BrowserRepl::GetInstance().IsVisible())
			BrowserRepl::GetInstance().Hide();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("LogMessage", [](GS::Ref<JS::Base> param) {
		if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
			if (v->GetType() == JS::Value::STRING) {
				(void)v->GetString();
			}
		}
		return new JS::Value(true);
		}));

	// --- Register ---
	targetBrowser.RegisterAsynchJSObject(jsACAPI);
}

// ------------------- Palette and Events ----------------------

void BrowserRepl::ButtonClicked(const DG::ButtonClickEvent& ev)
{
	const DG::ButtonItem* clickedButton = ev.GetSource();
	if (clickedButton == nullptr) return;

	const short buttonId = clickedButton->GetId();
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] ButtonClicked: id=%d", false, (int)buttonId);
#endif

	// Проверяем лицензию/демо перед выполнением команд (кроме Close и Support)
	if (!IsLicenseValid() && IsDemoExpired()) {
		if (buttonId != ToolbarButtonCloseId && buttonId != ToolbarButtonSupportId) {
			ACAPI_WriteReport("Demo period expired. Please purchase a license.", true);
			return;
		}
	}

	switch (buttonId) {
		case ToolbarButtonCloseId:
			Hide();
			break;
		case ToolbarButtonTableId:
			SelectionDetailsPalette::ShowPalette();
			break;
		case ToolbarButtonSupportId:
			{
				GS::UniString url = LicenseManager::BuildLicenseUrl();
				ShellExecuteW(NULL, L"open", url.ToUStr().Get(), NULL, NULL, SW_SHOWNORMAL);
			}
			break;
		default:
			break;
	}
}

void BrowserRepl::SetMenuItemCheckedState(bool isChecked)
{
	API_MenuItemRef itemRef = {};
	GSFlags itemFlags = {};

	itemRef.menuResID = BrowserReplMenuResId;
	itemRef.itemIndex = BrowserReplMenuItemIndex;

	ACAPI_MenuItem_GetMenuItemFlags(&itemRef, &itemFlags);
	if (isChecked) itemFlags |= API_MenuItemChecked;
	else           itemFlags &= ~API_MenuItemChecked;
	ACAPI_MenuItem_SetMenuItemFlags(&itemRef, &itemFlags);
}

void BrowserRepl::PanelResized(const DG::PanelResizeEvent& ev)
{
	// Toolbar buttons don't need resizing
	(void)ev;
}

void BrowserRepl::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] PanelCloseRequested will Hide", false);
#endif
	Hide();
	if (accepted) *accepted = true;
}

// SelectionChangeHandler removed - not needed for toolbar

GSErrCode BrowserRepl::PaletteControlCallBack(Int32, API_PaletteMessageID messageID, GS::IntPtr param)
{
#ifdef DEBUG_UI_LOGS
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		ACAPI_WriteReport("[BrowserRepl] PalMsg: OpenPalette", false);
		break;
	case APIPalMsg_ClosePalette:
		ACAPI_WriteReport("[BrowserRepl] PalMsg: ClosePalette", false);
		break;
	case APIPalMsg_HidePalette_Begin:
		ACAPI_WriteReport("[BrowserRepl] PalMsg: HidePalette_Begin", false);
		break;
	case APIPalMsg_HidePalette_End:
		ACAPI_WriteReport("[BrowserRepl] PalMsg: HidePalette_End", false);
		break;
	case APIPalMsg_DisableItems_Begin:
		ACAPI_WriteReport("[BrowserRepl] PalMsg: DisableItems_Begin", false);
		break;
	case APIPalMsg_DisableItems_End:
		ACAPI_WriteReport("[BrowserRepl] PalMsg: DisableItems_End", false);
		break;
	case APIPalMsg_IsPaletteVisible:
		*(reinterpret_cast<bool*> (param)) = HasInstance() && GetInstance().IsVisible();
		ACAPI_WriteReport("[BrowserRepl] PalMsg: IsPaletteVisible this %d", false, (int)*(reinterpret_cast<bool*> (param)));
		break;
	default:
		ACAPI_WriteReport("[BrowserRepl] PalMsg: %d", false, (int)messageID);
		break;
	}
#endif

	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!HasInstance()) CreateInstance();
		GetInstance().Show();
		break;
	case APIPalMsg_ClosePalette:
		if (!HasInstance()) break;
		GetInstance().Hide();
		break;
	case APIPalMsg_HidePalette_Begin:
		if (HasInstance() && GetInstance().IsVisible()) GetInstance().Hide();
		break;
	case APIPalMsg_HidePalette_End:
		if (HasInstance() && !GetInstance().IsVisible()) GetInstance().Show();
		break;
	case APIPalMsg_DisableItems_Begin:
		if (HasInstance() && GetInstance().IsVisible()) GetInstance().DisableItems();
		break;
	case APIPalMsg_DisableItems_End:
		if (HasInstance() && GetInstance().IsVisible()) GetInstance().EnableItems();
		break;
	case APIPalMsg_IsPaletteVisible:
		*(reinterpret_cast<bool*> (param)) = HasInstance() && GetInstance().IsVisible();
		break;
	default:
		break;
	}
	return NoError;
}

GSErrCode BrowserRepl::RegisterPaletteControlCallBack()
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] RegisterPaletteControlCallBack()", false);
#endif
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(paletteGuid),
		PaletteControlCallBack,
		API_PalEnabled_FloorPlan |
		API_PalEnabled_Section |
		API_PalEnabled_Elevation |
		API_PalEnabled_InteriorElevation |
		API_PalEnabled_3D |
		API_PalEnabled_Detail |
		API_PalEnabled_Worksheet |
		API_PalEnabled_Layout |
		API_PalEnabled_DocumentFrom3D,
		GSGuid2APIGuid(paletteGuid)
	);
}
