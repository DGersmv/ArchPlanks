#include "ACAPinc.h"
#include "APICommon.h"



#include "BrowserRepl.hpp"
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
#include "ColumnOrientHelper.hpp"



#include <cmath>
#include <cstdio>
#include <cstring>

// --------------------- Palette GUID / Instance ---------------------
static const GS::Guid paletteGuid("{11bd981d-f772-4a57-8709-42e18733a0cc}");
GS::Ref<BrowserRepl> BrowserRepl::instance;

// --------------------- Helpers (resource, js parsing, logging) ---------------------
static GS::UniString LoadHtmlFromResource()
{
	GS::UniString resourceData;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), 100);
	if (data != nullptr) {
		const GSSize handleSize = BMhGetSize(data);
		resourceData.Append(*data, handleSize);
		BMhKill(&data);
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[UI] HTML resource loaded, size=%u bytes", (unsigned)handleSize));
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[BrowserRepl] HTML resource loaded, size=%u bytes", false, (unsigned)handleSize);
#endif
	}
	else {
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser("[UI] ERROR: HTML resource not found (DATA 100)");
#ifdef DEBUG_UI_LOGS
		ACAPI_WriteReport("[BrowserRepl] ERROR: HTML resource not found (DATA 100)", false);
#endif
	}
	return resourceData;
}

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

// Буфер последнего ΔZ (м) — используется, если ApplyZDelta вызвали без аргумента
static double g_lastZDeltaMeters = 0.0;

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
	browser(GetReference(), BrowserId)
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] ctor", false);
#endif
	ACAPI_ProjectOperation_CatchProjectEvent(APINotify_Quit, NotificationHandler);

	// Подпишемся на изменение выделения (чтобы UI таблица актуализировалась)
	const GSErr selErr = ACAPI_Notification_CatchSelectionChange(SelectionChangeHandler);
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] CatchSelectionChange then err=%d", false, (int)selErr);
#endif

	Attach(*this);
	BeginEventProcessing();
	InitBrowserControl();
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

void BrowserRepl::InitBrowserControl()
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] InitBrowserControl: loading HTML", false);
#endif
	browser.LoadHTML(LoadHtmlFromResource());
	RegisterACAPIJavaScriptObject();
	// Страница сама дернёт UpdateSelectedElements() через whenACAPIReadyDo
	LogToBrowser("[C++] BrowserRepl initialized");
	
#ifndef DEBUG_UI_LOGS
	// В Release прячем чёрное поле логов (в HTML оно уже скрыто по умолчанию, но на всякий случай)
	browser.ExecuteJS("var logBox = document.getElementById('log-box'); if (logBox) logBox.style.display = 'none';");
#else
	// В Debug показываем поле логов (в HTML оно скрыто по умолчанию)
	browser.ExecuteJS("var logBox = document.getElementById('log-box'); if (logBox) logBox.style.display = 'block';");
#endif
}

void BrowserRepl::LogToBrowser(const GS::UniString& msg)
{
#ifdef DEBUG_UI_LOGS
	// UniString → UTF-8
	std::string utf8(msg.ToCStr(CC_UTF8));
	GS::UniString jsSafe(utf8.c_str(), CC_UTF8);

	// Экранируем спецсимволы для JS-строки
	jsSafe.ReplaceAll("\\", "\\\\");
	jsSafe.ReplaceAll("\"", "\\\"");
	jsSafe.ReplaceAll("\r", "");
	jsSafe.ReplaceAll("\n", "\\n");

	browser.ExecuteJS("AddLog(\"" + jsSafe + "\");");
#endif
}

// ------------------ JS API registration ---------------------
void BrowserRepl::RegisterACAPIJavaScriptObject()
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
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] ChangeSelectedElementsID " + baseID);
		const bool success = SelectionHelper::ChangeSelectedElementsID(baseID);
		return ConvertToJavaScriptVariable(success);
		}));

	jsACAPI->AddItem(new JS::Function("CreateLayerAndMoveElements", [](GS::Ref<JS::Base> param) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateLayerAndMoveElements");
		
		// Парсим параметры из JavaScript строки (передаем как JSON строку)
		LayerHelper::LayerCreationParams params;
		
		GS::UniString jsonStr = GetStringFromJavaScriptVariable(param);
		if (BrowserRepl::HasInstance()) {
			BrowserRepl::GetInstance().LogToBrowser("[C++] JSON строка: " + jsonStr);
		}
		
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
		
		if (BrowserRepl::HasInstance()) {
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[C++] Создание слоя: папка='%s', слой='%s', ID='%s', hide=%s", 
				params.folderPath.ToCStr().Get(), 
				params.layerName.ToCStr().Get(), 
				params.baseID.ToCStr().Get(),
				params.hideLayer ? "true" : "false"));
		}
		
		const bool success = LayerHelper::CreateLayerAndMoveElements(params);
		return ConvertToJavaScriptVariable(success);
		}));

	// --- ΔZ API (двухшаговый буфер + совместимость со старым мостом) ---
	jsACAPI->AddItem(new JS::Function("SetZDelta", [](GS::Ref<JS::Base> param) {
		g_lastZDeltaMeters = GetDoubleFromJs(param, 0.0);
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetZDelta=%.3f m", g_lastZDeltaMeters));
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("ApplyZDelta", [](GS::Ref<JS::Base> param) {
		// Диагностика: проверяем что пришло
		if (BrowserRepl::HasInstance()) {
			if (param == nullptr) {
				BrowserRepl::GetInstance().LogToBrowser("[JS] ApplyZDelta: param is nullptr!");
			} else if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
				const auto t = v->GetType();
				if (t == JS::Value::DOUBLE || t == JS::Value::INTEGER) {
					BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyZDelta: got number %.6f", v->GetDouble()));
				} else if (t == JS::Value::STRING) {
					BrowserRepl::GetInstance().LogToBrowser("[JS] ApplyZDelta: got string '" + v->GetString() + "'");
				} else {
					BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyZDelta: got type %d", (int)t));
				}
			} else {
				BrowserRepl::GetInstance().LogToBrowser("[JS] ApplyZDelta: param is not JS::Value");
			}
		}
		
		const double val = (param != nullptr) ? GetDoubleFromJs(param, g_lastZDeltaMeters) : g_lastZDeltaMeters;
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyZDelta(%.3f m)", val));
		const bool ok = GroundHelper::ApplyZDelta(val);
		return new JS::Value(ok);
		}));

	// --- Ground API (посадка на Mesh) ---
	jsACAPI->AddItem(new JS::Function("SetGroundSurface", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetGroundSurface()");
		const bool ok = GroundHelper::SetGroundSurface();
		return new JS::Value(ok);
		}));

	jsACAPI->AddItem(new JS::Function("SetGroundObjects", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetGroundObjects()");
		const bool ok = GroundHelper::SetGroundObjects();
		return new JS::Value(ok);
		}));

	jsACAPI->AddItem(new JS::Function("ApplyGroundOffset", [](GS::Ref<JS::Base> param) {
		const double offset = GetDoubleFromJs(param, 0.0); // offset на C++ стороне сейчас игнорируется, но логируем
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] ApplyGroundOffset(%.3f m)", offset));
		const bool ok = GroundHelper::ApplyGroundOffset(offset);
		return new JS::Value(ok);
		}));

	// --- Rotate API ---
	jsACAPI->AddItem(new JS::Function("RotateSelected", [](GS::Ref<JS::Base> param) {
		const double angle = GetDoubleFromJs(param, 0.0);
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelected angle=%.2f", angle));
		return new JS::Value(RotateHelper::RotateSelected(angle));
		}));

	jsACAPI->AddItem(new JS::Function("AlignSelectedX", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] AlignSelectedX()");
		return new JS::Value(RotateHelper::AlignSelectedX());
		}));

	jsACAPI->AddItem(new JS::Function("RandomizeSelectedAngles", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] RandomizeSelectedAngles()");
		return new JS::Value(RotateHelper::RandomizeSelectedAngles());
		}));

	jsACAPI->AddItem(new JS::Function("OrientObjectsToPoint", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] OrientObjectsToPoint()");
		return new JS::Value(RotateHelper::OrientObjectsToPoint());
		}));

	// --- GDL Generator ---
	jsACAPI->AddItem(new JS::Function("GenerateGDLFromSelection", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] GenerateGDLFromSelection()");
		return new JS::Value(GDLHelper::GenerateGDLFromSelection());
		}));

	// --- Landscape API (заглушки с логом) ---
	jsACAPI->AddItem(new JS::Function("SetDistributionLine", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetDistributionLine()");
		return new JS::Value(LandscapeHelper::SetDistributionLine());
		}));
	jsACAPI->AddItem(new JS::Function("SetDistributionObject", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetDistributionObject()");
		return new JS::Value(LandscapeHelper::SetDistributionObject());
		}));
	jsACAPI->AddItem(new JS::Function("SetDistributionStep", [](GS::Ref<JS::Base> param) {
		const double step = GetDoubleFromJs(param, 0.0);
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetDistributionStep step=%.3f", step));
		return new JS::Value(LandscapeHelper::SetDistributionStep(step));
		}));
	jsACAPI->AddItem(new JS::Function("SetDistributionCount", [](GS::Ref<JS::Base> param) {
		const int count = (int)std::llround(GetDoubleFromJs(param, 0.0));
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetDistributionCount count=%d", count));
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
		if (BrowserRepl::HasInstance()) {
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] DistributeNow parsed: step=%.6f, count=%d", step, count));
		}
		return new JS::Value(LandscapeHelper::DistributeSelected(step, count));
		}));

	// --- Beam Orientation API ---
	jsACAPI->AddItem(new JS::Function("SetBeams", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetBeams()");
		return new JS::Value(ColumnOrientHelper::SetBeams());
		}));
	jsACAPI->AddItem(new JS::Function("SetMeshForColumns", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetMeshForColumns()");
		return new JS::Value(ColumnOrientHelper::SetMesh());
		}));
	jsACAPI->AddItem(new JS::Function("OrientBeamsToSurface", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] OrientBeamsToSurface()");
		return new JS::Value(ColumnOrientHelper::OrientBeamsToSurface());
		}));
	jsACAPI->AddItem(new JS::Function("RotateSelectedOrientation", [](GS::Ref<JS::Base> param) {
		// Диагностика параметра (для отладки)
		double angleDeg = 0.0;
		if (BrowserRepl::HasInstance()) {
			if (param == nullptr) {
				BrowserRepl::GetInstance().LogToBrowser("[JS] RotateSelectedOrientation: param is nullptr!");
			} else if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
				const auto t = v->GetType();
				if (t == JS::Value::DOUBLE) {
					angleDeg = v->GetDouble();
					BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelectedOrientation: got DOUBLE %.6f", angleDeg));
				} else if (t == JS::Value::INTEGER) {
					angleDeg = static_cast<double>(v->GetInteger());
					BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelectedOrientation: got INTEGER %d -> %.6f", (int)v->GetInteger(), angleDeg));
				} else if (t == JS::Value::STRING) {
					BrowserRepl::GetInstance().LogToBrowser("[JS] RotateSelectedOrientation: got string '" + v->GetString() + "'");
					angleDeg = GetDoubleFromJs(param, 0.0);
				} else {
					BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelectedOrientation: got unknown type %d", (int)t));
					angleDeg = GetDoubleFromJs(param, 0.0);
				}
			} else {
				BrowserRepl::GetInstance().LogToBrowser("[JS] RotateSelectedOrientation: param is not JS::Value, using GetDoubleFromJs");
				angleDeg = GetDoubleFromJs(param, 0.0);
			}
		} else {
			angleDeg = GetDoubleFromJs(param, 0.0);
		}
		
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] RotateSelectedOrientation final: angleDeg=%.3f", angleDeg));
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
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser("[C++] OpenHelp to " + url);
		HelpPalette::ShowWithURL(url);
		return new JS::Value(true);
		}));

	jsACAPI->AddItem(new JS::Function("LogMessage", [](GS::Ref<JS::Base> param) {
		if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
			if (v->GetType() == JS::Value::STRING) {
				GS::UniString s = v->GetString();
				if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] " + s);
			}
		}
		return new JS::Value(true);
		}));

	// --- Markup API (разметка размерами) ---
	jsACAPI->AddItem(new JS::Function("SetMarkupStep", [](GS::Ref<JS::Base> param) {
		// Диагностика: проверяем что пришло
		if (BrowserRepl::HasInstance()) {
			if (param == nullptr) {
				BrowserRepl::GetInstance().LogToBrowser("[JS] SetMarkupStep: param is nullptr!");
			} else if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
				const auto t = v->GetType();
				if (t == JS::Value::DOUBLE || t == JS::Value::INTEGER) {
					BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetMarkupStep: got number %.6f", v->GetDouble()));
				} else if (t == JS::Value::STRING) {
					BrowserRepl::GetInstance().LogToBrowser("[JS] SetMarkupStep: got string '" + v->GetString() + "'");
				} else {
					BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetMarkupStep: got type %d", (int)t));
				}
			} else {
				BrowserRepl::GetInstance().LogToBrowser("[JS] SetMarkupStep: param is not JS::Value");
			}
		}
		
		const double stepMM = GetDoubleFromJs(param, 0.0);
		if (BrowserRepl::HasInstance())
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] SetMarkupStep parsed=%.1f mm", stepMM));
		return new JS::Value(MarkupHelper::SetMarkupStep(stepMM));
		}));

	jsACAPI->AddItem(new JS::Function("CreateMarkupDimensions", [](GS::Ref<JS::Base>) {
		return new JS::Value(MarkupHelper::CreateMarkupDimensions());
		}));

	jsACAPI->AddItem(new JS::Function("CreateDimensionsToLine", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateDimensionsToLine()");
		return new JS::Value(MarkupHelper::CreateDimensionsToLine());
		}));

	jsACAPI->AddItem(new JS::Function("CreateDimensionsBetweenObjects", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateDimensionsBetweenObjects()");
		return new JS::Value(MarkupHelper::CreateDimensionsBetweenObjects());
		}));

	jsACAPI->AddItem(new JS::Function("CreateDimensionsToPoint", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateDimensionsToPoint()");
		return new JS::Value(MarkupHelper::CreateDimensionsToPoint());
		}));

	// --- Road API (выбор осевой линии) ---
	jsACAPI->AddItem(new JS::Function("SetBaseLineForShell", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetBaseLineForShell() [Shell]");
		
		// Используем ShellHelper для установки базовой линии (работает для контуров)
		const bool success = ShellHelper::SetBaseLineForShell();
		
		// Также синхронизируем с RoadHelper для совместимости
		if (success) {
			RoadHelper::SetCenterLine(); // Это также обновит RoadHelper, если нужно
		}
		
		return new JS::Value(success);
		}));

	jsACAPI->AddItem(new JS::Function("SetMeshSurfaceForShell", [](GS::Ref<JS::Base>) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] SetMeshSurfaceForShell() [Shell]");
		
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
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] CreateMeshFromContour() - создание Mesh из контура");
		
		// Парсим параметры: принимаем строку "width:..,step:.."
		double width = 1000.0; // мм по умолчанию
		double step = 500.0;   // мм по умолчанию
		
		if (param != nullptr) {
			if (GS::Ref<JS::Value> v = GS::DynamicCast<JS::Value>(param)) {
				if (v->GetType() == JS::Value::STRING) {
					GS::UniString s = v->GetString();
					const char* c = s.ToCStr().Get();
					
					if (std::strncmp(c, "width:", 6) == 0) { 
						std::sscanf(c + 6, "%lf", &width); 
					}
					const char* stepPtr = std::strstr(c, "step:");
					if (stepPtr != nullptr) {
						std::sscanf(stepPtr + 5, "%lf", &step);
					}
				}
			}
		}
		
		if (BrowserRepl::HasInstance()) {
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] CreateMeshFromContour parsed: width=%.1fmm, step=%.1fmm", width, step));
		}
		
		bool success = ShellHelper::CreateMeshFromContour(width, step);
		return new JS::Value(success);
	}));
	
	// --- Get Surface Finishes List (покрытия) ---
	jsACAPI->AddItem(new JS::Function("GetSurfaceFinishesList", [](GS::Ref<JS::Base> param) {
		if (BrowserRepl::HasInstance()) BrowserRepl::GetInstance().LogToBrowser("[JS] GetSurfaceFinishesList()");
		
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
		
		if (BrowserRepl::HasInstance()) {
			BrowserRepl::GetInstance().LogToBrowser(GS::UniString::Printf("[JS] CreateShellFromLine parsed: width=%.1fmm, step=%.1fmm", width, step));
		}
		
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
	browser.RegisterAsynchJSObject(jsACAPI);
	LogToBrowser("[C++] JS bridge registered");
}

// ------------------- Palette and Events ----------------------
void BrowserRepl::UpdateSelectedElementsOnHTML()
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] UpdateSelectedElementsOnHTML()", false);
#endif
	browser.ExecuteJS("UpdateSelectedElements()");
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
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] PanelResized dx=%d dy=%d", false, (int)ev.GetHorizontalChange(), (int)ev.GetVerticalChange());
#endif
	BeginMoveResizeItems();
	browser.Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void BrowserRepl::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] PanelCloseRequested will Hide", false);
#endif
	Hide();
	*accepted = true;
}

GSErrCode __ACENV_CALL BrowserRepl::SelectionChangeHandler(const API_Neig*)
{
#ifdef DEBUG_UI_LOGS
	ACAPI_WriteReport("[BrowserRepl] Selection changed then update UI", false);
#endif
	if (BrowserRepl::HasInstance())
		BrowserRepl::GetInstance().UpdateSelectedElementsOnHTML();
	return NoError;
}

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
