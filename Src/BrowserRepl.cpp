#include "ACAPinc.h"
#include "APICommon.h"
#include "BrowserRepl.hpp"
#include "DGBrowser.hpp"

#include <Windows.h>
#include "SelectionHelper.hpp"
#include "RotateHelper.hpp"
#include "LandscapeHelper.hpp"
#include "GroundHelper.hpp"
#include "BuildHelper.hpp"
#include "GDLHelper.hpp"
#include "MarkupHelper.hpp"
#include "RoadHelper.hpp"
#include "ShellHelper.hpp"
#include "HelpPalette.hpp"
#include "LayerHelper.hpp"
#include "DistributionPalette.hpp"
#include "OrientationPalette.hpp"
#include "GroundPalette.hpp"
#include "MarkupPalette.hpp"
#include "ContourPalette.hpp"
#include "MeshPalette.hpp"
#include "IdLayersPalette.hpp"
#include "AnglePalette.hpp"
#include "ColumnOrientHelper.hpp"
#include "SelectionPropertyHelper.hpp"
#include "SelectionMetricsHelper.hpp"
#include "SendXlsPalette.hpp"
#include "SelectionDetailsPalette.hpp"



#include <cmath>
#include <cstdio>
#include <cstring>

// --------------------- Palette GUID / Instance ---------------------
static const GS::Guid paletteGuid("{11bd981d-f772-4a57-8709-42e18733a0cc}");
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

template<>
GS::Ref<JS::Base> ConvertToJavaScriptVariable(const LayerHelper::LayerInfo& layerInfo)
{
	GS::Ref<JS::Array> js = new JS::Array();
	js->AddItem(ConvertToJavaScriptVariable(layerInfo.name));
	js->AddItem(ConvertToJavaScriptVariable(layerInfo.folder));
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

// Буфер последнего ΔZ (м) — используется, если ApplyZDelta вызвали без аргумента
static double g_lastZDeltaMeters = 0.0;

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
static GSErrCode __ACENV_CALL NotificationHandler(API_NotifyEventID notifID, Int32 /*param*/)
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
	buttonSpline(GetReference(), ToolbarButtonSplineId),
	buttonRotate(GetReference(), ToolbarButtonRotateId),
	buttonRotSurf(GetReference(), ToolbarButtonRotSurfId),
	buttonLand(GetReference(), ToolbarButtonLandId),
	buttonDims(GetReference(), ToolbarButtonDimsId),
	buttonLayers(GetReference(), ToolbarButtonLayersId),
	buttonContour(GetReference(), ToolbarButtonContourId),
	buttonMesh(GetReference(), ToolbarButtonMeshId),
	buttonCSV(GetReference(), ToolbarButtonCSVId),
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

	// --- Send to Excel / CSV export ---
	jsACAPI->AddItem(new JS::Function("SaveSendXls", [](GS::Ref<JS::Base> param) {
		// param: строка CSV, которую нужно сохранить
		GS::UniString csvData = GetStringFromJavaScriptVariable(param);

		// Используем стандартный Windows диалог "Сохранить как"
		OPENFILENAMEW ofn;
		wchar_t fileName[MAX_PATH] = L"selection_export.csv";
		ZeroMemory(&ofn, sizeof(ofn));
		ofn.lStructSize = sizeof(ofn);
		ofn.hwndOwner = nullptr;
		ofn.lpstrFilter = L"CSV Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0";
		ofn.lpstrFile = fileName;
		ofn.nMaxFile = MAX_PATH;
		ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
		ofn.lpstrDefExt = L"csv";

		if (!GetSaveFileNameW(&ofn)) {
			// Пользователь мог просто отменить диалог
			return ConvertToJavaScriptVariable(false);
		}

		FILE* fp = _wfopen(fileName, L"wb");
		if (fp == nullptr) {
			return ConvertToJavaScriptVariable(false);
		}

		// Пишем BOM UTF‑8 для правильного распознавания Excel
		const unsigned char bom[3] = { 0xEF, 0xBB, 0xBF };
		fwrite(bom, 1, 3, fp);

		// Конвертируем UniString в UTF-8 через Windows API
		// UniString внутренне хранит UTF-16, получаем через GetLength и итерацию
		GS::UniString dataToWrite = csvData;
		const USize len = dataToWrite.GetLength();
		if (len > 0) {
			// Выделяем буфер для UTF-16
			wchar_t* wideBuffer = new wchar_t[len + 1];
			for (USize i = 0; i < len; ++i) {
				wideBuffer[i] = static_cast<wchar_t>(dataToWrite[i]);
			}
			wideBuffer[len] = L'\0';

			// Конвертируем UTF-16 в UTF-8
			int utf8Len = WideCharToMultiByte(CP_UTF8, 0, wideBuffer, -1, nullptr, 0, nullptr, nullptr);
			if (utf8Len > 0) {
				char* utf8Buffer = new char[utf8Len];
				if (WideCharToMultiByte(CP_UTF8, 0, wideBuffer, -1, utf8Buffer, utf8Len, nullptr, nullptr) > 0) {
					// Записываем без завершающего нуля (utf8Len включает нуль)
					fwrite(utf8Buffer, 1, utf8Len - 1, fp);
				}
				delete[] utf8Buffer;
			}
			delete[] wideBuffer;
		}

		fclose(fp);
		return ConvertToJavaScriptVariable(true);
		}));

	jsACAPI->AddItem(new JS::Function("CreateLayerAndMoveElements", [](GS::Ref<JS::Base> param) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateLayerAndMoveElements");
		
		// Парсим параметры из JavaScript строки (передаем как JSON строку)
		LayerHelper::LayerCreationParams params;
		
		GS::UniString jsonStr = GetStringFromJavaScriptVariable(param);
		// if (BrowserRepl::HasInstance()) {
		//	BrowserRepl::GetInstance().LogToBrowser("[C++] JSON строка: " + jsonStr);
		// }
		
		// Простой парсинг JSON (папка|слой)
		GS::Array<GS::UniString> parts;
		jsonStr.Split(GS::UniString("|"), [&parts](const GS::UniString& part) {
			parts.Push(part);
		});
		
		if (parts.GetSize() >= 2) {
			params.folderPath = parts[0];
			params.layerName = parts[1];
			params.baseID = GS::UniString(""); // ID не меняем
			params.hideLayer = false; // По умолчанию не скрываем
			
			// Если есть третий параметр, это флаг скрытия
			if (parts.GetSize() >= 3) {
				params.hideLayer = (parts[2] == "1" || parts[2] == "true" || parts[2] == "True");
			}
		}
		
		// if (BrowserRepl::HasInstance()) {
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[C++] Создание слоя: папка='%s', слой='%s', ID='%s', hide=%s", 
		//		params.folderPath.ToCStr().Get(), 
		//		params.layerName.ToCStr().Get(), 
		//		params.baseID.ToCStr().Get(),
		//		params.hideLayer ? "true" : "false"));
		// }
		
		const bool success = LayerHelper::CreateLayerAndMoveElements(params);
		return ConvertToJavaScriptVariable(success);
		}));

	// --- Get Layers List API ---
	jsACAPI->AddItem(new JS::Function("GetLayersList", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] GetLayersList()");
		const GS::Array<LayerHelper::LayerInfo> layers = LayerHelper::GetLayersList();
		return ConvertToJavaScriptVariable(layers);
		}));

	// --- ΔZ API (двухшаговый буфер + совместимость со старым мостом) ---
	jsACAPI->AddItem(new JS::Function("SetZDelta", [](GS::Ref<JS::Base> param) {
		g_lastZDeltaMeters = GetDoubleFromJs(param, 0.0);
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetZDelta=%.3f m", g_lastZDeltaMeters));
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("ApplyZDelta", [](GS::Ref<JS::Base> param) {
		// Диагностика: проверяем что пришло
		// if (BrowserRepl::HasInstance()) {
		//	if (param == nullptr) {
		//		BrowserRepl::GetInstance().LogToBrowser("[JS] ApplyZDelta: param is nullptr!");
		//	} else if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
		//		const auto t = v->GetType();
		//		if (t == JS::Value::DOUBLE || t == JS::Value::INTEGER) {
		//			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyZDelta: got number %.6f", v->GetDouble()));
		//		} else if (t == JS::Value::STRING) {
		//			BrowserRepl::GetInstance().LogToBrowser("[JS] ApplyZDelta: got string '" + v->GetString() + "'");
		//		} else {
		//			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyZDelta: got type %d", (int)t));
		//		}
		//	} else {
		//		BrowserRepl::GetInstance().LogToBrowser("[JS] ApplyZDelta: param is not JS::Value");
		//	}
		// }
		
		const double val = (param != nullptr) ? GetDoubleFromJs(param, g_lastZDeltaMeters) : g_lastZDeltaMeters;
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyZDelta(%.3f m)", val));
		const bool ok = GroundHelper::ApplyZDelta(val);
		return new JS::Value(ok);
		}));

	// --- Ground API (посадка на Mesh) ---
	jsACAPI->AddItem(new JS::Function("SetGroundSurface", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetGroundSurface()");
		const bool ok = GroundHelper::SetGroundSurface();
		return new JS::Value(ok);
		}));

	jsACAPI->AddItem(new JS::Function("SetGroundObjects", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetGroundObjects()");
		const bool ok = GroundHelper::SetGroundObjects();
		return new JS::Value(ok);
		}));

	jsACAPI->AddItem(new JS::Function("ApplyGroundOffset", [](GS::Ref<JS::Base> param) {
		const double offset = GetDoubleFromJs(param, 0.0); // offset на C++ стороне сейчас игнорируется, но логируем
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyGroundOffset(%.3f m)", offset));
		const bool ok = GroundHelper::ApplyGroundOffset(offset);
		return new JS::Value(ok);
		}));

	// --- Rotate API ---
	jsACAPI->AddItem(new JS::Function("RotateSelected", [](GS::Ref<JS::Base> param) {
		const double angle = GetDoubleFromJs(param, 0.0);
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelected angle=%.2f", angle));
		return new JS::Value(RotateHelper::RotateSelected(angle));
		}));

	jsACAPI->AddItem(new JS::Function("AlignSelectedX", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] AlignSelectedX()");
		return new JS::Value(RotateHelper::AlignSelectedX());
		}));

	jsACAPI->AddItem(new JS::Function("RandomizeSelectedAngles", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] RandomizeSelectedAngles()");
		return new JS::Value(RotateHelper::RandomizeSelectedAngles());
		}));

	jsACAPI->AddItem(new JS::Function("OrientObjectsToPoint", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] OrientObjectsToPoint()");
		return new JS::Value(RotateHelper::OrientObjectsToPoint());
		}));

	// --- GDL Generator ---
	jsACAPI->AddItem(new JS::Function("GenerateGDLFromSelection", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] GenerateGDLFromSelection()");
		return new JS::Value(GDLHelper::GenerateGDLFromSelection());
		}));

	// --- Landscape API (заглушки с логом) ---
	jsACAPI->AddItem(new JS::Function("SetDistributionLine", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetDistributionLine()");
		return new JS::Value(LandscapeHelper::SetDistributionLine());
		}));
	jsACAPI->AddItem(new JS::Function("SetDistributionObject", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetDistributionObject()");
		return new JS::Value(LandscapeHelper::SetDistributionObject());
		}));
	jsACAPI->AddItem(new JS::Function("SetDistributionStep", [](GS::Ref<JS::Base> param) {
		const double step = GetDoubleFromJs(param, 0.0);
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetDistributionStep step=%.3f", step));
		return new JS::Value(LandscapeHelper::SetDistributionStep(step));
		}));
	jsACAPI->AddItem(new JS::Function("SetDistributionCount", [](GS::Ref<JS::Base> param) {
		const int count = (int)std::llround(GetDoubleFromJs(param, 0.0));
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetDistributionCount count=%d", count));
		return new JS::Value(LandscapeHelper::SetDistributionCount(count));
		}));
	jsACAPI->AddItem(new JS::Function("DistributeNow", [](GS::Ref<JS::Base> param) {
		// принимаем step/count в строке "step:.."/"count:.." или просто число
		double step = 0.0; int count = 0;
		if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
			switch (v->GetType()) {
			case JS::Value::DOUBLE:
			case JS::Value::INTEGER: step = v->GetDouble(); break;
			case JS::Value::STRING: {
				GS::UniString s = v->GetString();
				for (UIndex i = 0; i < s.GetLength(); ++i) if (s[i] == ',') s[i] = '.';
				const char* c = s.ToCStr().Get();
				if (std::strncmp(c, "step:", 5) == 0) { std::sscanf(c + 5, "%lf", &step); }
				else if (std::strncmp(c, "count:", 6) == 0) { std::sscanf(c + 6, "%d", &count); }
				else { std::sscanf(c, "%lf", &step); }
				break;
			}
			default: break;
			}
		}
		// if (BrowserRepl::HasInstance()) {
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] DistributeNow parsed: step=%.6f, count=%d", step, count));
		// }
		return new JS::Value(LandscapeHelper::DistributeSelected(step, count));
		}));

	// --- Beam Orientation API ---
	jsACAPI->AddItem(new JS::Function("SetBeams", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetBeams()");
		return new JS::Value(ColumnOrientHelper::SetBeams());
		}));
	jsACAPI->AddItem(new JS::Function("SetMeshForColumns", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetMeshForColumns()");
		return new JS::Value(ColumnOrientHelper::SetMesh());
		}));
	jsACAPI->AddItem(new JS::Function("OrientBeamsToSurface", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] OrientBeamsToSurface()");
		return new JS::Value(ColumnOrientHelper::OrientBeamsToSurface());
		}));
	jsACAPI->AddItem(new JS::Function("RotateSelectedOrientation", [](GS::Ref<JS::Base> param) {
		// Диагностика параметра (для отладки)
		double angleDeg = 0.0;
		// if (BrowserRepl::HasInstance()) {
		//	if (param == nullptr) {
		//		BrowserRepl::GetInstance().LogToBrowser("[JS] RotateSelectedOrientation: param is nullptr!");
		//	} else if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
		//		const auto t = v->GetType();
		//		if (t == JS::Value::DOUBLE) {
		//			angleDeg = v->GetDouble();
		//			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelectedOrientation: got DOUBLE %.6f", angleDeg));
		//		} else if (t == JS::Value::INTEGER) {
		//			angleDeg = static_cast<double>(v->GetInteger());
		//			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelectedOrientation: got INTEGER %d -> %.6f", (int)v->GetInteger(), angleDeg));
		//		} else if (t == JS::Value::STRING) {
		//			BrowserRepl::GetInstance().LogToBrowser("[JS] RotateSelectedOrientation: got string '" + v->GetString() + "'");
		//			angleDeg = GetDoubleFromJs(param, 0.0);
		//		} else {
		//			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelectedOrientation: got unknown type %d", (int)t));
		//			angleDeg = GetDoubleFromJs(param, 0.0);
		//		}
		//	} else {
		//		BrowserRepl::GetInstance().LogToBrowser("[JS] RotateSelectedOrientation: param is not JS::Value, using GetDoubleFromJs");
		//		angleDeg = GetDoubleFromJs(param, 0.0);
		//	}
		// } else {
			angleDeg = GetDoubleFromJs(param, 0.0);
		// }
		
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelectedOrientation final: angleDeg=%.3f", angleDeg));
		return new JS::Value(ColumnOrientHelper::RotateSelected(angleDeg));
		}));

	// --- Help / Log ---
	jsACAPI->AddItem(new JS::Function("OpenHelp", [](GS::Ref<JS::Base> param) {
		GS::UniString url;
		if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
			if (v->GetType() == JS::Value::STRING) url = v->GetString();
		}
		if (url.IsEmpty()) url = "https://landscape.227.info/help/start";
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenHelp] url=%s", false, url.ToCStr().Get());
#endif
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser("[C++] OpenHelp to " + url);
		HelpPalette::ShowWithURL(url);
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenDistributionPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenDistributionPalette] request", false);
#endif
		DistributionPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenOrientationPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenOrientationPalette] request", false);
#endif
		OrientationPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenGroundPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenGroundPalette] request", false);
#endif
		GroundPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenMarkupPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenMarkupPalette] request", false);
#endif
		MarkupPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenIdLayersPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenIdLayersPalette] request", false);
#endif
		IdLayersPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenAnglePalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenAnglePalette] request", false);
#endif
		AnglePalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenMeshPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenMeshPalette] request", false);
#endif
		MeshPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenContourPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenContourPalette] request", false);
#endif
		ContourPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenSendXlsPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenSendXlsPalette] request", false);
#endif
		SendXlsPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("OpenSelectionDetailsPalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[OpenSelectionDetailsPalette] request", false);
#endif
		SelectionDetailsPalette::ShowPalette();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("ClosePalette", [](GS::Ref<JS::Base>) {
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[ClosePalette] request", false);
#endif
		if (BrowserRepl::HasInstance() && BrowserRepl::GetInstance().IsVisible())
			BrowserRepl::GetInstance().Hide();
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("LogMessage", [](GS::Ref<JS::Base> param) {
		if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
			if (v->GetType() == JS::Value::STRING) {
				GS::UniString s = v->GetString();
				// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] " + s);
			}
		}
		return new JS::Value(true);
		}));

	// --- Markup API (разметка размерами) ---
	jsACAPI->AddItem(new JS::Function("SetMarkupStep", [](GS::Ref<JS::Base> param) {
		// Диагностика: проверяем что пришло
		// if (BrowserRepl::HasInstance()) {
		//	if (param == nullptr) {
		//		BrowserRepl::GetInstance().LogToBrowser("[JS] SetMarkupStep: param is nullptr!");
		//	} else if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
		//		const auto t = v->GetType();
		//		if (t == JS::Value::DOUBLE || t == JS::Value::INTEGER) {
		//			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetMarkupStep: got number %.6f", v->GetDouble()));
		//		} else if (t == JS::Value::STRING) {
		//			BrowserRepl::GetInstance().LogToBrowser("[JS] SetMarkupStep: got string '" + v->GetString() + "'");
		//		} else {
		//			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetMarkupStep: got type %d", (int)t));
		//		}
		//	} else {
		//		BrowserRepl::GetInstance().LogToBrowser("[JS] SetMarkupStep: param is not JS::Value");
		//	}
		// }
		
		const double stepMM = GetDoubleFromJs(param, 0.0);
		// if (BrowserRepl::HasInstance())
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetMarkupStep parsed=%.1f mm", stepMM));
		return new JS::Value(MarkupHelper::SetMarkupStep(stepMM));
		}));

	jsACAPI->AddItem(new JS::Function("CreateMarkupDimensions", [](GS::Ref<JS::Base>) {
		return new JS::Value(MarkupHelper::CreateMarkupDimensions());
		}));

	jsACAPI->AddItem(new JS::Function("CreateDimensionsToLine", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateDimensionsToLine()");
		return new JS::Value(MarkupHelper::CreateDimensionsToLine());
		}));

	jsACAPI->AddItem(new JS::Function("CreateDimensionsBetweenObjects", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateDimensionsBetweenObjects()");
		return new JS::Value(MarkupHelper::CreateDimensionsBetweenObjects());
		}));

	jsACAPI->AddItem(new JS::Function("CreateDimensionsToPoint", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateDimensionsToPoint()");
		return new JS::Value(MarkupHelper::CreateDimensionsToPoint());
		}));

	// --- Road API (выбор осевой линии) ---
	jsACAPI->AddItem(new JS::Function("SetBaseLineForShell", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetBaseLineForShell() [Shell]");
		
		// Используем ShellHelper для установки базовой линии (работает для контуров)
		const bool success = ShellHelper::SetBaseLineForShell();
		
		// Также синхронизируем с RoadHelper для совместимости
		if (success) {
			RoadHelper::SetCenterLine(); // Это также обновит RoadHelper, если нужно
		}
		
		return new JS::Value(success);
		}));

	jsACAPI->AddItem(new JS::Function("SetMeshSurfaceForShell", [](GS::Ref<JS::Base>) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetMeshSurfaceForShell() [Shell]");
		
		// Вызываем ShellHelper::SetMeshSurfaceForShell(), который установит mesh в ShellHelper и GroundHelper
		const bool success = ShellHelper::SetMeshSurfaceForShell();
		
		// Также синхронизируем с RoadHelper для совместимости
		if (success) {
			RoadHelper::SetTerrainMesh(); // Это также обновит RoadHelper, если нужно
		}
		
		return new JS::Value(success);
		}));

	// --- Create Mesh from Contour ---
	jsACAPI->AddItem(new JS::Function("CreateMeshFromContour", [](GS::Ref<JS::Base> param) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateMeshFromContour() - создание Mesh из контура");
		
		// Парсим параметры: принимаем строку "leftWidth:..,rightWidth:..,step:..,offset:.."
		double leftWidth = 600.0;  // мм по умолчанию
		double rightWidth = 600.0; // мм по умолчанию
		double step = 500.0;   // мм по умолчанию
		double offset = 0.0;   // мм по умолчанию
		
		if (param != nullptr) {
			if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
				if (v->GetType() == JS::Value::STRING) {
					GS::UniString s = v->GetString();
					const char* c = s.ToCStr().Get();
					
					const char* leftPtr = std::strstr(c, "leftWidth:");
					if (leftPtr != nullptr) {
						std::sscanf(leftPtr + 10, "%lf", &leftWidth);
					}
					const char* rightPtr = std::strstr(c, "rightWidth:");
					if (rightPtr != nullptr) {
						std::sscanf(rightPtr + 11, "%lf", &rightWidth);
					}
					const char* stepPtr = std::strstr(c, "step:");
					if (stepPtr != nullptr) {
						std::sscanf(stepPtr + 5, "%lf", &step);
					}
					const char* offsetPtr = std::strstr(c, "offset:");
					if (offsetPtr != nullptr) {
						std::sscanf(offsetPtr + 7, "%lf", &offset);
					}
				}
			}
		}
		
		// if (BrowserRepl::HasInstance()) {
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf(
		//		"[JS] CreateMeshFromContour parsed: leftWidth=%.1fmm, rightWidth=%.1fmm, step=%.1fmm, offset=%.1fmm",
		//		leftWidth, rightWidth, step, offset
		//	));
		// }
		
		bool success = ShellHelper::CreateMeshFromContour(leftWidth, rightWidth, step, offset);
		return new JS::Value(success);
	}));
	
	// --- Get Surface Finishes List (покрытия) ---
	jsACAPI->AddItem(new JS::Function("GetSurfaceFinishesList", [](GS::Ref<JS::Base> param) {
		// if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] GetSurfaceFinishesList()");
		
		GS::Array<RoadHelper::SurfaceFinishInfo> finishes = RoadHelper::GetSurfaceFinishesList();
		
		JS::Object* result = new JS::Object();
		JS::Array* finishArray = new JS::Array();
		
		for (UIndex i = 0; i < finishes.GetSize(); ++i) {
			JS::Object* finishObj = new JS::Object();
			finishObj->AddItem("index", new JS::Value((double)finishes[i].index));
			finishObj->AddItem("name", new JS::Value(finishes[i].name));
			finishArray->AddItem(finishObj);
		}
		
		result->AddItem("finishes", finishArray);
		result->AddItem("count", new JS::Value((double)finishes.GetSize()));
		
		return result;
	}));

	// --- Road API (создание дорожки по линии) ---
	jsACAPI->AddItem(new JS::Function("CreateShellFromLine", [](GS::Ref<JS::Base> param) {
		// Парсим параметры: принимаем строку "width:..,step:.." или просто число
		double width = 1000.0; // мм по умолчанию
		double step = 500.0;   // мм по умолчанию
		
		if (param != nullptr) {
			if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
				switch (v->GetType()) {
				case JS::Value::DOUBLE:
				case JS::Value::INTEGER: 
					width = v->GetDouble(); 
					break;
				case JS::Value::STRING: {
					GS::UniString s = v->GetString();
					for (UIndex i = 0; i < s.GetLength(); ++i) if (s[i] == ',') s[i] = '.';
					const char* c = s.ToCStr().Get();
					if (std::strncmp(c, "width:", 6) == 0) { 
						std::sscanf(c + 6, "%lf", &width); 
					}
					if (std::strstr(c, "step:") != nullptr) {
						const char* stepStart = std::strstr(c, "step:");
						if (stepStart != nullptr) {
							std::sscanf(stepStart + 5, "%lf", &step);
						}
					}
					break;
				}
				default: break;
				}
			}
		}
		
		// if (BrowserRepl::HasInstance()) {
		//	BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] CreateShellFromLine parsed: width=%.1fmm, step=%.1fmm", width, step));
		// }
		
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[BrowserRepl] Вызов RoadHelper::BuildRoad", false);
#endif
		RoadHelper::RoadParams params;
		params.widthMM = width;
		params.sampleStepMM = step;
		const bool success = RoadHelper::BuildRoad(params);
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[BrowserRepl] RoadHelper::BuildRoad вернул: %s", false, success ? "true" : "false");
#endif
		return new JS::Value(success);
		}));

	// --- Register object in the browser ---
	targetBrowser.RegisterAsynchJSObject(jsACAPI);
	// if (BrowserRepl::HasInstance())
	//	BrowserRepl::GetInstance().LogToBrowser("[C++] JS bridge registered");
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

	switch (buttonId) {
		case ToolbarButtonCloseId:
			Hide();
			break;
		case ToolbarButtonTableId:
			SelectionDetailsPalette::ShowPalette();
			break;
		case ToolbarButtonSplineId:
			DistributionPalette::ShowPalette();
			break;
		case ToolbarButtonRotateId:
			OrientationPalette::ShowPalette();
			break;
		case ToolbarButtonRotSurfId:
			AnglePalette::ShowPalette();
			break;
		case ToolbarButtonLandId:
			GroundPalette::ShowPalette();
			break;
		case ToolbarButtonDimsId:
			MarkupPalette::ShowPalette();
			break;
		case ToolbarButtonLayersId:
			IdLayersPalette::ShowPalette();
			break;
		case ToolbarButtonContourId:
			ContourPalette::ShowPalette();
			break;
		case ToolbarButtonMeshId:
			MeshPalette::ShowPalette();
			break;
		case ToolbarButtonCSVId:
			SendXlsPalette::ShowPalette();
			break;
		case ToolbarButtonSupportId:
			HelpPalette::ShowWithURL("https://landscape.227.info/help");
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

GSErrCode __ACENV_CALL BrowserRepl::PaletteControlCallBack(Int32, API_PaletteMessageID messageID, GS::IntPtr param)
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
