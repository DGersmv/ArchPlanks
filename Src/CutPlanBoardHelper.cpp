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
			rows.Push(row);
			targetRow = &rows[rows.GetSize() - 1];
		}

		targetRow->count += 1;
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
	csv += "Board;BoardW;Cut1;Cut2;Cut3;Cut4;Cut5;Cut6;Cut7;Cut8;Cut9;Cut10;Remainder;Kerf\r\n";
	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		const CuttingStock::ResultBoard& rb = result.boards[b];
		csv += GS::UniString::Printf("%d;%.2f;", (int)(b + 1), rb.boardW);
		for (UIndex c = 0; c < 10; ++c) {
			if (c < rb.cuts.GetSize())
				csv += GS::UniString::Printf("%.2f;", rb.cuts[c]);
			else
				csv += ";";
		}
		csv += GS::UniString::Printf("%.2f;%.2f\r\n", rb.remainder, slit);
	}
	csv += "\r\n";
	csv += "Remaining parts (length;boardW;material)\r\n";
	for (UIndex i = 0; i < result.remaining.GetSize(); ++i) {
		const CuttingStock::Part& p = result.remaining[i];
		csv += GS::UniString::Printf("%.2f;%.2f;%s\r\n", p.length, p.boardW, p.material.ToCStr().Get());
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

bool RunCuttingPlan(double slitMM, short /*floorIndex*/)
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
	params.maxStockLength = (maxStockLength > 0.0 ? maxStockLength : 6000.0);

	CuttingStock::SolverResult result = CuttingStock::Solve(parts, params);
	return ExportCutPlanToExcel(result, params.slit);
}

} // namespace CutPlanBoardHelper
