#include "CutPlanBoardHelper.hpp"
#include "APICommon.h"
#include "CH.hpp"
#include <Windows.h>
#include <commdlg.h>
#include <cwchar>
#include <cstdarg>
#include <cstdio>
#include <cmath>

namespace CutPlanBoardHelper {

namespace {

static double GetAddParReal(API_AddParType* const* params, Int32 count, const char* parName)
{
	if (params == nullptr || *params == nullptr) return 0.0;
	for (Int32 i = 0; i < count; ++i) {
		const API_AddParType& p = (*params)[i];
		if (p.name != nullptr && CHEqualASCII(p.name, parName, GS::CaseInsensitive)) {
			return p.value.real;
		}
	}
	return 0.0;
}

static Int32 GetAddParCount(API_AddParType* const* params)
{
	if (params == nullptr || *params == nullptr) return 0;
	return static_cast<Int32>(BMGetHandleSize(reinterpret_cast<GSHandle>(const_cast<API_AddParType**>(params))) / sizeof(API_AddParType));
}

static GS::UniString BuildMaterialLabel(double widthMM, double heightMM)
{
	// Простой текст вида "50x200 мм"
	return GS::UniString::Printf("%.0f x %.0f \u043C\u043C", widthMM, heightMM);
}

static const char* AddParTypeName(API_AddParID typeID)
{
	switch (typeID) {
		case API_ZombieParT:        return "Zombie";
		case APIParT_Integer:       return "Integer";
		case APIParT_Length:       return "Length";
		case APIParT_Angle:        return "Angle";
		case APIParT_RealNum:      return "RealNum";
		case APIParT_LightSw:      return "LightSw";
		case APIParT_ColRGB:       return "ColRGB";
		case APIParT_Intens:       return "Intens";
		case APIParT_LineTyp:     return "LineTyp";
		case APIParT_Mater:        return "Mater";
		case APIParT_FillPat:      return "FillPat";
		case APIParT_PenCol:       return "PenCol";
		case APIParT_CString:      return "CString";
		case APIParT_Boolean:      return "Boolean";
		case APIParT_Separator:    return "Separator";
		case APIParT_Title:        return "Title";
		case APIParT_BuildingMaterial: return "BuildingMaterial";
		case APIParT_Profile:      return "Profile";
		case APIParT_Dictionary:  return "Dictionary";
		default:                   return "?";
	}
}

static void FillDefaultsFromLibPart(Int32 libInd, double* outHeight, double* outWidth, double* outLen, double* outMaxLen)
{
	double a = 0.0, b = 0.0;
	Int32 addParNum = 0;
	API_AddParType** addPars = nullptr;
	if (ACAPI_LibraryPart_GetParams(libInd, &a, &b, &addParNum, &addPars) != NoError || addPars == nullptr || *addPars == nullptr)
		return;
	for (Int32 i = 0; i < addParNum; ++i) {
		const API_AddParType& p = (*addPars)[i];
		if (CHEqualASCII(p.name, "iHeight", GS::CaseInsensitive) && outHeight && *outHeight == 0.0)
			*outHeight = p.value.real;
		else if (CHEqualASCII(p.name, "iWidth", GS::CaseInsensitive) && outWidth && *outWidth == 0.0)
			*outWidth = p.value.real;
		else if (CHEqualASCII(p.name, "iLen", GS::CaseInsensitive) && outLen && *outLen == 0.0)
			*outLen = p.value.real;
		else if (CHEqualASCII(p.name, "iMaxLen", GS::CaseInsensitive) && outMaxLen && *outMaxLen == 0.0)
			*outMaxLen = p.value.real;
	}
	ACAPI_DisposeAddParHdl(&addPars);
}

} // anonymous

bool IsArchiFramePlank(const API_Guid& guid)
{
	API_Element element = {};
	element.header.guid = guid;
	if (ACAPI_Element_Get(&element) != NoError)
		return false;
	if (element.header.type.typeID != API_ObjectID)
		return false;
	API_LibPart lp = {};
	lp.typeID = APILib_ObjectID;
	lp.index = element.object.libInd;
	if (ACAPI_LibraryPart_Get(&lp) != NoError)
		return false;
	GS::UniString fname(lp.file_UName);
	if (lp.location != nullptr) {
		delete lp.location;
		lp.location = nullptr;
	}
	return (fname.Contains("ArchiFramePlank") || fname == "ArchiFramePlank.gsm");
}

bool GetArchiFramePlankParams(const API_Guid& guid, ArchiFramePlankParams& out)
{
	out = {};
	API_Element element = {};
	element.header.guid = guid;
	if (ACAPI_Element_Get(&element) != NoError)
		return false;
	API_ElementMemo memo = {};
	if (ACAPI_Element_GetMemo(guid, &memo, APIMemoMask_AddPars) != NoError)
		return false;
	Int32 count = GetAddParCount(memo.params);
	out.iHeight = GetAddParReal(memo.params, count, "iHeight");
	out.iLen = GetAddParReal(memo.params, count, "iLen");
	out.iMaxLen = GetAddParReal(memo.params, count, "iMaxLen");
	out.iWidth = GetAddParReal(memo.params, count, "iWidth");
	ACAPI_DisposeElemMemoHdls(&memo);
	// Если в экземпляре 0 — подставляем значения по умолчанию из библиотечной части
	if (out.iHeight == 0.0 || out.iWidth == 0.0 || out.iLen == 0.0 || out.iMaxLen == 0.0)
		FillDefaultsFromLibPart(element.object.libInd, &out.iHeight, &out.iWidth, &out.iLen, &out.iMaxLen);
	// GDL Length-параметры приходят в метрах; переводим в мм для отображения и расчёта распила
	const double mToMM = 1000.0;
	out.iHeight *= mToMM;
	out.iWidth  *= mToMM;
	out.iLen    *= mToMM;
	out.iMaxLen *= mToMM;
	return (out.iLen > 0 && out.iMaxLen > 0);
}

GS::Array<CuttingStock::Part> CollectPartsFromSelection(double& outMaxStockLength)
{
	GS::Array<CuttingStock::Part> parts;
	outMaxStockLength = 6000.0;

	API_SelectionInfo selInfo = {};
	GS::Array<API_Neig> selNeigs;
	ACAPI_Selection_Get(&selInfo, &selNeigs, false, false);
	BMKillHandle((GSHandle*)&selInfo.marquee.coords);

	for (const API_Neig& n : selNeigs) {
		if (!IsArchiFramePlank(n.guid))
			continue;
		ArchiFramePlankParams p;
		if (!GetArchiFramePlankParams(n.guid, p))
			continue;
		if (p.iMaxLen > 0)
			outMaxStockLength = p.iMaxLen;
		CuttingStock::Part part;
		part.length = p.iLen;
		part.boardW = p.iHeight > 0 ? p.iHeight : 100.0;
		part.material = p.material;
		if (part.material.IsEmpty())
			part.material = GS::UniString::Printf("%.0f", p.iWidth);
		parts.Push(part);
	}
	return parts;
}

GS::Array<ArchiFrameSummaryRow> CollectArchiFrameSummaryFromSelection()
{
	GS::Array<ArchiFrameSummaryRow> rows;

	API_SelectionInfo selInfo = {};
	GS::Array<API_Neig> selNeigs;
	ACAPI_Selection_Get(&selInfo, &selNeigs, false, false);
	BMKillHandle((GSHandle*)&selInfo.marquee.coords);

	for (const API_Neig& n : selNeigs) {
		if (!IsArchiFramePlank(n.guid))
			continue;

		ArchiFramePlankParams p;
		if (!GetArchiFramePlankParams(n.guid, p))
			continue;

		const double widthMM = p.iWidth;
		const double heightMM = p.iHeight;
		const GS::UniString label = BuildMaterialLabel(widthMM, heightMM);
		const GS::UniString guidStr = APIGuidToString(n.guid);

		ArchiFrameSummaryRow* targetRow = nullptr;
		for (UIndex i = 0; i < rows.GetSize(); ++i) {
			ArchiFrameSummaryRow& row = rows[i];
			if (std::fabs(row.widthMM - widthMM) < 0.001 &&
				std::fabs(row.heightMM - heightMM) < 0.001) {
				targetRow = &row;
				break;
			}
		}

		if (targetRow == nullptr) {
			ArchiFrameSummaryRow row;
			row.widthMM = widthMM;
			row.heightMM = heightMM;
			row.materialLabel = label;
			row.count = 0;
			row.totalLenMM = 0.0;
			row.maxLenMM = p.iMaxLen;
			rows.Push(row);
			targetRow = &rows[rows.GetSize() - 1];
		}

		targetRow->count += 1;
		targetRow->totalLenMM += p.iLen;
		targetRow->guidStrs.Push(guidStr);
	}

	return rows;
}

CuttingStock::SolverParams DefaultSolverParams()
{
	CuttingStock::SolverParams p = {};
	p.slit = 4.0;
	p.trimLoss = 0.0;
	p.usefulMin = 300.0;
	p.wasteMax = 50.0;
	p.strictAB = false;
	p.maxImproveIter = 2000;
	return p;
}

GS::UniString BuildCutPlanCsv(const CuttingStock::SolverResult& result, double slit, FastProduction::ScenarioData* outScenarioData)
{
	GS::UniString csv;
	FastProduction::ScenarioData scenarioData;
	if (outScenarioData) {
		scenarioData = FastProduction::BuildScenarioData(result, 1, true, 2);
		*outScenarioData = scenarioData;
	}

	// Определяем максимальное количество отрезков на доску,
	// чтобы сформировать заголовок Cut1..CutN и строки полной ширины.
	UIndex maxCuts = 0;
	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		const CuttingStock::ResultBoard& rb = result.boards[b];
		if (rb.cuts.GetSize() > maxCuts)
			maxCuts = rb.cuts.GetSize();
	}

	csv += "Board;BoardW;";
	for (UIndex c = 0; c < maxCuts; ++c) {
		csv += GS::UniString::Printf("Cut%u;", (unsigned)(c + 1));
	}
	csv += "Remainder;Kerf";
	if (outScenarioData) {
		csv += ";ScenarioId;ScenarioOps;ScenarioSetups;ScenarioGroup";
	}
	csv += "\r\n";

	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		const CuttingStock::ResultBoard& rb = result.boards[b];
		csv += GS::UniString::Printf("%d;%.0f;", (int)(b + 1), rb.boardW);
		for (UIndex c = 0; c < maxCuts; ++c) {
			if (c < rb.cuts.GetSize())
				csv += GS::UniString::Printf("%.0f;", rb.cuts[c]);
			else
				csv += ";";
		}
		csv += GS::UniString::Printf("%.0f;%.0f", rb.remainder, slit);
		if (outScenarioData && b < scenarioData.boardScenarioId.GetSize()) {
			csv += ";";
			csv += scenarioData.boardScenarioId[b];
			csv += ";";
			csv += scenarioData.boardScenarioOps[b];
			csv += ";";
			csv += GS::UniString::Printf("%d", scenarioData.boardScenarioSetups[b]);
			csv += ";";
			csv += scenarioData.boardScenarioGroup[b];
		}
		csv += "\r\n";
	}
	csv += "\r\n";
	csv += "Remaining parts (length;boardW;material)\r\n";
	for (UIndex i = 0; i < result.remaining.GetSize(); ++i) {
		const CuttingStock::Part& p = result.remaining[i];
		csv += GS::UniString::Printf("%.0f;%.0f;%s\r\n", p.length, p.boardW, p.material.ToCStr().Get());
	}

	// Сводка по пиломатериалам: толщина/ширина, количество и объём в м3
	csv += "\r\n";
	csv += "Material summary (boardT_mm;boardW_mm;count;volume_m3)\r\n";
	GS::Array<ArchiFrameSummaryRow> summaryRows = CollectArchiFrameSummaryFromSelection();
	for (UIndex i = 0; i < summaryRows.GetSize(); ++i) {
		const ArchiFrameSummaryRow& row = summaryRows[i];
		const double boardTmm = row.widthMM;   // толщина (iWidth)
		const double boardWmm = row.heightMM;  // ширина (iHeight)
		const double maxLenMM = (row.maxLenMM > 0.0 ? row.maxLenMM : 6000.0);
		double boardsCountD = 0.0;
		if (maxLenMM > 0.0 && row.totalLenMM > 0.0)
			boardsCountD = std::ceil(row.totalLenMM / maxLenMM);
		const unsigned boardsCount = boardsCountD > 0.0 ? (unsigned)boardsCountD : 0u;

		const double boardTm = boardTmm / 1000.0;
		const double boardWm = boardWmm / 1000.0;
		const double maxLenM = maxLenMM / 1000.0;
		const double volumeM3 = boardsCount * maxLenM * boardTm * boardWm;
		csv += GS::UniString::Printf("%.0f;%.0f;%u;%.3f\r\n", boardTmm, boardWmm, boardsCount, volumeM3);
	}

	// Сводка по отрезкам: толщина/ширина доски и длина отрезка
	csv += "\r\n";
	csv += "Cut summary (boardT_mm;boardW_mm;cutLen_mm;count)\r\n";

	struct CutSummaryRow {
		double boardTmm;
		double boardWmm;
		double cutLenMM;
		unsigned count;
	};

	GS::Array<CutSummaryRow> cutSummary;

	// Карта height (ширина) -> thickness (толщина) по исходной сводке
	auto FindThicknessForWidth = [&summaryRows](double boardWmm) -> double {
		for (UIndex i = 0; i < summaryRows.GetSize(); ++i) {
			const ArchiFrameSummaryRow& row = summaryRows[i];
			if (std::fabs(row.heightMM - boardWmm) < 0.001) {
				return row.widthMM;
			}
		}
		return 0.0;
	};

	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		const CuttingStock::ResultBoard& rb = result.boards[b];
		const double boardWmm = rb.boardW;                // ширина
		const double boardTmm = FindThicknessForWidth(boardWmm); // толщина

		for (UIndex c = 0; c < rb.cuts.GetSize(); ++c) {
			double cutLenMM = rb.cuts[c];
			if (cutLenMM <= 0.0)
				continue;

			// Нормализуем до целых мм
			cutLenMM = std::round(cutLenMM);

			// Ищем существующую запись
			CutSummaryRow* row = nullptr;
			for (UIndex k = 0; k < cutSummary.GetSize(); ++k) {
				CutSummaryRow& r = cutSummary[k];
				if (std::fabs(r.boardTmm - boardTmm) < 0.001 &&
					std::fabs(r.boardWmm - boardWmm) < 0.001 &&
					std::fabs(r.cutLenMM - cutLenMM) < 0.001) {
					row = &r;
					break;
				}
			}

			if (row == nullptr) {
				CutSummaryRow r;
				r.boardTmm = boardTmm;
				r.boardWmm = boardWmm;
				r.cutLenMM = cutLenMM;
				r.count = 0;
				cutSummary.Push(r);
				row = &cutSummary[cutSummary.GetSize() - 1];
			}

			row->count += 1;
		}
	}

	// Сортируем по толщине, ширине, длине
	for (UIndex i = 0; i < cutSummary.GetSize(); ++i) {
		for (UIndex j = i + 1; j < cutSummary.GetSize(); ++j) {
			const CutSummaryRow& a = cutSummary[i];
			const CutSummaryRow& bRow = cutSummary[j];

			bool swap = false;
			if (a.boardTmm > bRow.boardTmm + 0.001) {
				swap = true;
			} else if (std::fabs(a.boardTmm - bRow.boardTmm) < 0.001 && a.boardWmm > bRow.boardWmm + 0.001) {
				swap = true;
			} else if (std::fabs(a.boardTmm - bRow.boardTmm) < 0.001 &&
					   std::fabs(a.boardWmm - bRow.boardWmm) < 0.001 &&
					   a.cutLenMM > bRow.cutLenMM + 0.001) {
				swap = true;
			}

			if (swap) {
				CutSummaryRow tmp = cutSummary[i];
				cutSummary[i] = cutSummary[j];
				cutSummary[j] = tmp;
			}
		}
	}

	for (UIndex i = 0; i < cutSummary.GetSize(); ++i) {
		const CutSummaryRow& r = cutSummary[i];
		csv += GS::UniString::Printf("%.0f;%.0f;%.0f;%u\r\n", r.boardTmm, r.boardWmm, r.cutLenMM, r.count);
	}

	return csv;
}

namespace {

static bool WriteUtf8File(const wchar_t* path, const GS::UniString& content)
{
	FILE* fp = _wfopen(path, L"wb");
	if (!fp)
		return false;
	const unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
	if (fwrite(bom, 1, 3, fp) != 3) {
		fclose(fp);
		return false;
	}
	const int wideLen = (int)content.GetLength();
	if (wideLen > 0) {
		const wchar_t* wstr = content.ToUStr().Get();
		int utf8Size = WideCharToMultiByte(CP_UTF8, 0, wstr, wideLen, NULL, 0, NULL, NULL);
		if (utf8Size > 0) {
			GS::Array<char> buf(utf8Size);
			WideCharToMultiByte(CP_UTF8, 0, wstr, wideLen, buf.GetContent(), utf8Size, NULL, NULL);
			if (fwrite(buf.GetContent(), 1, (size_t)utf8Size, fp) != (size_t)utf8Size) {
				fclose(fp);
				return false;
			}
		}
	}
	fclose(fp);
	return true;
}

/** Write text file in Windows-1251 so Notepad opens it without ? (Cyrillic). */
static bool WriteAnsiFile(const wchar_t* path, const GS::UniString& content)
{
	FILE* fp = _wfopen(path, L"wb");
	if (!fp)
		return false;
	const int wideLen = (int)content.GetLength();
	if (wideLen > 0) {
		const wchar_t* wstr = content.ToUStr().Get();
		const UINT cp = 1251;  /* Windows-1251 Cyrillic */
		int ansiSize = WideCharToMultiByte(cp, 0, wstr, wideLen, NULL, 0, NULL, NULL);
		if (ansiSize > 0) {
			GS::Array<char> buf(ansiSize);
			WideCharToMultiByte(cp, 0, wstr, wideLen, buf.GetContent(), ansiSize, NULL, NULL);
			if (fwrite(buf.GetContent(), 1, (size_t)ansiSize, fp) != (size_t)ansiSize) {
				fclose(fp);
				return false;
			}
		}
	}
	fclose(fp);
	return true;
}

static double FindThicknessForBoardW(const GS::Array<ArchiFrameSummaryRow>& summaryRows, double boardWmm)
{
	for (UIndex i = 0; i < summaryRows.GetSize(); ++i) {
		if (std::fabs(summaryRows[i].heightMM - boardWmm) < 0.001)
			return summaryRows[i].widthMM;
	}
	return 0.0;
}

static void AppendWideLine(GS::UniString& txt, const wchar_t* fmt, ...)
{
	wchar_t buf[1024];
	va_list args;
	va_start(args, fmt);
	int n = vswprintf_s(buf, fmt, args);
	va_end(args);
	if (n > 0)
		txt += GS::UniString(buf);
}

static GS::UniString BuildOperatorInstructions(const FastProduction::ScenarioData& scenarioData,
	const CuttingStock::SolverResult& result,
	const GS::Array<ArchiFrameSummaryRow>& summaryRows)
{
	GS::UniString txt;
	/* Use wide string literals (L"...") so Cyrillic is correct; then WriteAnsiFile(1251) will convert properly. */
	if (!result.boards.IsEmpty()) {
		GS::Array<double> widths;
		for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
			double w = result.boards[b].boardW;
			bool found = false;
			for (UIndex d = 0; d < widths.GetSize(); ++d) {
				if (std::fabs(widths[d] - w) < 0.001) { found = true; break; }
			}
			if (!found)
				widths.Push(w);
		}
		AppendWideLine(txt, L"=== \u0412\u0441\u0435\u0433\u043E \u0434\u043E\u0441\u043E\u043A ===\r\n");
		for (UIndex d = 0; d < widths.GetSize(); ++d) {
			double boardW = widths[d];
			UInt32 count = 0;
			for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
				if (std::fabs(result.boards[b].boardW - boardW) < 0.001)
					count++;
			}
			double thickness = FindThicknessForBoardW(summaryRows, boardW);
			if (thickness < 0.001) thickness = 0.0;
			AppendWideLine(txt, L"\u0412\u0441\u0435\u0433\u043E %u \u0448\u0442 %.0fx%.0f \u043C\u043C\r\n", (unsigned)count, thickness, boardW);
		}
		txt += GS::UniString(L"\r\n\r\n");
	}

	AppendWideLine(txt, L"=== \u041F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u044B \u0440\u0430\u0441\u043F\u0438\u043B\u0430 ===\r\n\r\n");
	for (UIndex s = 0; s < scenarioData.scenarios.GetSize(); ++s) {
		const FastProduction::ScenarioInfo& info = scenarioData.scenarios[s];
		double thickness = FindThicknessForBoardW(summaryRows, info.boardW);
		if (thickness < 0.001) thickness = 0.0;
		/* ScenarioId is ASCII (e.g. W195_S00); convert to wide for format */
		GS::UniString sid = info.scenarioId;
		AppendWideLine(txt, L"--- %s (\u0411\u0435\u0440\u0451\u043C %d \u0434\u043E\u0441\u043E\u043A %.0fx%.0f \u043C\u043C) ---\r\n",
			sid.ToUStr().Get(), info.boardsCount, thickness, info.boardW);
		const bool oneSetup = (info.steps.GetSize() == 1);
		for (UIndex r = 0; r < info.steps.GetSize(); ++r) {
			const FastProduction::ScenarioStepInfo& step = info.steps[r];
			const bool isFirst = (r == 0);
			const bool isLast = (r == info.steps.GetSize() - 1);
			if (oneSetup) {
				AppendWideLine(txt, L"\u0423\u043F\u043E\u0440 %d \u043C\u043C \u2014 \u043F\u0438\u043B\u0438\u043C \u0431\u0435\u0437 \u043E\u0441\u0442\u0430\u0442\u043A\u0430 (%d \u0440\u0435\u0437\u043E\u0432 \u0441 \u043A\u0430\u0436\u0434\u043E\u0439 \u0434\u043E\u0441\u043A\u0438). ", step.stopLength, step.cutsPerBoard);
				AppendWideLine(txt, L"\u041E\u0441\u0442\u0430\u0442\u043E\u043A %.0f \u043C\u043C \u043D\u0430 \u0434\u043E\u0441\u043A\u0443. \u0412\u0441\u0435\u0433\u043E %d \u0434\u043E\u0441\u043E\u043A, %d \u0440\u0435\u0437\u043E\u0432.\r\n", info.remainderMm, info.boardsCount, step.totalCuts);
			} else {
				if (isFirst) {
					AppendWideLine(txt, L"\u0423\u043F\u043E\u0440 %d \u043C\u043C \u2014 \u043F\u0438\u043B\u0438\u043C (%d \u0440\u0435\u0437 \u0441 \u043A\u0430\u0436\u0434\u043E\u0439). \u041E\u0441\u0442\u0430\u0442\u043E\u043A \u043E\u0442\u043A\u043B\u0430\u0434\u044B\u0432\u0430\u0435\u043C \u0432 \u043F\u0430\u0447\u043A\u0443.\r\n", step.stopLength, step.cutsPerBoard);
				} else if (isLast) {
					AppendWideLine(txt, L"\u0423\u043F\u043E\u0440 %d \u043C\u043C \u2014 \u0431\u0435\u0440\u0451\u043C \u043E\u0441\u0442\u0430\u0442\u043E\u043A \u0438\u0437 \u043F\u0430\u0447\u043A\u0438, \u043F\u0438\u043B\u0438\u043C (%d \u0440\u0435\u0437 \u0441 \u043A\u0430\u0436\u0434\u043E\u0439). ", step.stopLength, step.cutsPerBoard);
					AppendWideLine(txt, L"\u041E\u0441\u0442\u0430\u0442\u043E\u043A %.0f \u043C\u043C \u043D\u0430 \u0434\u043E\u0441\u043A\u0443.\r\n", info.remainderMm);
				} else {
					AppendWideLine(txt, L"\u0423\u043F\u043E\u0440 %d \u043C\u043C \u2014 \u0431\u0435\u0440\u0451\u043C \u043E\u0441\u0442\u0430\u0442\u043E\u043A \u0438\u0437 \u043F\u0430\u0447\u043A\u0438, \u043F\u0438\u043B\u0438\u043C (%d \u0440\u0435\u0437 \u0441 \u043A\u0430\u0436\u0434\u043E\u0439). \u041E\u0441\u0442\u0430\u0442\u043E\u043A \u0432 \u043F\u0430\u0447\u043A\u0443.\r\n", step.stopLength, step.cutsPerBoard);
				}
			}
		}
		txt += GS::UniString(L"\r\n");
	}
	return txt;
}

} // anonymous

bool PlaceScenarioTextOnFloor(const GS::UniString& instructionsTxt, short floorIndex)
{
	if (instructionsTxt.IsEmpty())
		return true;
	API_Element element = {};
	element.header.type = API_TextID;
	GSErrCode err = ACAPI_Element_GetDefaults(&element, nullptr);
	if (err != NoError)
		return false;
	element.text.head.floorInd = floorIndex;
	element.text.loc.x = 1.0;
	element.text.loc.y = 1.0;
	if (element.text.width < 100.0)
		element.text.width = 400.0;
	if (element.text.height < 50.0)
		element.text.height = 300.0;
	API_ElementMemo memo = {};
	GS::UniString contentCopy = instructionsTxt;
	memo.textContent = &contentCopy;
	err = ACAPI_Element_Create(&element, &memo);
	/* Do not call ACAPI_DisposeElemMemoHdls: we did not allocate memo.textContent. */
	return (err == NoError);
}

bool ExportCutPlanToExcel(const CuttingStock::SolverResult& result, double slit, short floorIndex)
{
	FastProduction::ScenarioData scenarioData;
	GS::UniString csv = BuildCutPlanCsv(result, slit, &scenarioData);
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

	if (!WriteUtf8File(pathBuf, csv))
		return false;

	// Base path without extension for additional files
	wchar_t basePath[MAX_PATH] = L"";
	wcscpy_s(basePath, pathBuf);
	wchar_t* dot = wcsrchr(basePath, L'.');
	if (dot && dot > basePath)
		*dot = L'\0';

	wchar_t extraPath[MAX_PATH] = L"";

	// set_list.csv
	swprintf_s(extraPath, L"%s_set_list.csv", basePath);
	GS::UniString setListCsv;
	setListCsv += "BoardW;ScenarioId;StopLength;CutsCount;BoardsCount;OpOrder\r\n";
	for (UIndex i = 0; i < scenarioData.setListRows.GetSize(); ++i) {
		const FastProduction::SetListRow& row = scenarioData.setListRows[i];
		setListCsv += GS::UniString::Printf("%.0f;%s;%d;%d;%d;%d\r\n",
			row.boardW, row.scenarioId.ToCStr().Get(), row.stopLength, row.cutsCount, row.boardsCount, row.opOrder);
	}
	WriteUtf8File(extraPath, setListCsv);

	// set_list_summary.csv
	swprintf_s(extraPath, L"%s_set_list_summary.csv", basePath);
	GS::UniString summaryCsv;
	summaryCsv += "BoardW;StopLength;TotalCuts\r\n";
	for (UIndex i = 0; i < scenarioData.setListSummaryRows.GetSize(); ++i) {
		const FastProduction::SetListSummaryRow& row = scenarioData.setListSummaryRows[i];
		summaryCsv += GS::UniString::Printf("%.0f;%d;%d\r\n", row.boardW, row.stopLength, row.totalCuts);
	}
	WriteUtf8File(extraPath, summaryCsv);

	// operator_instructions.txt (Windows-1251 so Notepad shows Cyrillic)
	swprintf_s(extraPath, L"%s_operator_instructions.txt", basePath);
	GS::Array<ArchiFrameSummaryRow> summaryRows = CollectArchiFrameSummaryFromSelection();
	GS::UniString instructionsTxt = BuildOperatorInstructions(scenarioData, result, summaryRows);
	WriteAnsiFile(extraPath, instructionsTxt);

	// Place scenario text on selected floor (API_TextID) so it displays in Archicad
	if (floorIndex >= 0) {
		API_StoryInfo storyInfo = {};
		if (ACAPI_ProjectSetting_GetStorySettings(&storyInfo) == NoError) {
			if (floorIndex >= storyInfo.firstStory && floorIndex <= storyInfo.lastStory)
				PlaceScenarioTextOnFloor(instructionsTxt, floorIndex);
		}
	}

	return true;
}

void DumpArchiFramePlankParamsToReport()
{
	API_SelectionInfo selInfo = {};
	GS::Array<API_Neig> selNeigs;
	ACAPI_Selection_Get(&selInfo, &selNeigs, false, false);
	BMKillHandle((GSHandle*)&selInfo.marquee.coords);

	Int32 dumped = 0;
	for (const API_Neig& n : selNeigs) {
		if (!IsArchiFramePlank(n.guid))
			continue;
		API_ElementMemo memo = {};
		if (ACAPI_Element_GetMemo(n.guid, &memo, APIMemoMask_AddPars) != NoError)
			continue;
		Int32 count = GetAddParCount(memo.params);
		GS::UniString guidStr = APIGuidToString(n.guid);
		ACAPI_WriteReport("=== ArchiFramePlank parameters (GUID: %s) ===", false, guidStr.ToCStr().Get());
		for (Int32 i = 0; i < count; ++i) {
			const API_AddParType& p = (*memo.params)[i];
			const char* typeStr = AddParTypeName(p.typeID);
			if (p.typeID == APIParT_Separator || p.typeID == APIParT_Title) {
				ACAPI_WriteReport("  [%s] %s", false, typeStr, p.name);
			} else if (p.typeID == APIParT_CString) {
				ACAPI_WriteReport("  %s | %s | (string)", false, p.name, typeStr);
			} else {
				ACAPI_WriteReport("  %s | %s | %.6g", false, p.name, typeStr, p.value.real);
			}
		}
		ACAPI_DisposeElemMemoHdls(&memo);
		++dumped;
		break; // только первый выбранный, чтобы не засорять отчёт
	}
	if (dumped == 0)
		ACAPI_WriteReport("No ArchiFramePlank in selection. Select at least one ArchiFramePlank and run again.", true);
}

bool RunCuttingPlan(double slitMM, double extraLenMM, short floorIndex)
{
	double maxStockLength = 0.0;
	GS::Array<CuttingStock::Part> parts = CollectPartsFromSelection(maxStockLength);
	if (parts.IsEmpty()) {
		ACAPI_WriteReport("No ArchiFramePlank objects in selection. Select ArchiFramePlank elements first.", true);
		return false;
	}

	CuttingStock::SolverParams params = DefaultSolverParams();
	if (slitMM > 0.0)
		params.slit = slitMM;
	const double baseMax = (maxStockLength > 0.0 ? maxStockLength : 6000.0);
	if (extraLenMM > 0.0)
		params.maxStockLength = baseMax + extraLenMM;
	else
		params.maxStockLength = baseMax;

	CuttingStock::SolverResult result = CuttingStock::Solve(parts, params);
	return ExportCutPlanToExcel(result, params.slit, floorIndex);
}

} // namespace CutPlanBoardHelper
