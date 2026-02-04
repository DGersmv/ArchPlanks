#include "CutPlanBoardHelper.hpp"
#include "APICommon.h"
#include "CH.hpp"
#include <Windows.h>
#include <commdlg.h>
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

GS::UniString BuildCutPlanCsv(const CuttingStock::SolverResult& result, double slit)
{
	GS::UniString csv;

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
	csv += "Remainder;Kerf\r\n";

	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		const CuttingStock::ResultBoard& rb = result.boards[b];
		csv += GS::UniString::Printf("%d;%.0f;", (int)(b + 1), rb.boardW);
		for (UIndex c = 0; c < maxCuts; ++c) {
			if (c < rb.cuts.GetSize())
				csv += GS::UniString::Printf("%.0f;", rb.cuts[c]);
			else
				csv += ";";
		}
		csv += GS::UniString::Printf("%.0f;%.0f\r\n", rb.remainder, slit);
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

bool ExportCutPlanToExcel(const CuttingStock::SolverResult& result, double slit)
{
	GS::UniString csv = BuildCutPlanCsv(result, slit);
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

	const unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
	if (fwrite(bom, 1, 3, fp) != 3) {
		fclose(fp);
		return false;
	}
	const int wideLen = (int)csv.GetLength();
	if (wideLen > 0) {
		const wchar_t* wstr = csv.ToUStr().Get();
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

bool RunCuttingPlan(double slitMM, double extraLenMM, short /*floorIndex*/)
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
	return ExportCutPlanToExcel(result, params.slit);
}

} // namespace CutPlanBoardHelper
