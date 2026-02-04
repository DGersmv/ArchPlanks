#ifndef CUTPLANBOARDHELPER_HPP
#define CUTPLANBOARDHELPER_HPP

#include "APIEnvir.h"
#include "ACAPinc.h"
#include "GSRoot.hpp"
#include "CuttingStockSolver.hpp"

namespace CutPlanBoardHelper {

struct ArchiFramePlankParams {
	double iHeight;
	double iLen;
	double iMaxLen;
	double iWidth;
	GS::UniString material;
};

struct ArchiFrameSummaryRow {
	double widthMM;              // iWidth
	double heightMM;             // iHeight
	GS::UniString materialLabel; // человекочитаемый текст (например, "50×200 мм")
	UInt32 count = 0;            // количество досок в этой группе
	GS::Array<GS::UniString> guidStrs; // GUID-ы досок в строковом виде
};

bool IsArchiFramePlank(const API_Guid& guid);
bool GetArchiFramePlankParams(const API_Guid& guid, ArchiFramePlankParams& out);

GS::Array<CuttingStock::Part> CollectPartsFromSelection(double& outMaxStockLength);
GS::Array<ArchiFrameSummaryRow> CollectArchiFrameSummaryFromSelection();

CuttingStock::SolverParams DefaultSolverParams();

GS::UniString BuildCutPlanCsv(const CuttingStock::SolverResult& result, double slit);

bool ExportCutPlanToExcel(const CuttingStock::SolverResult& result, double slit);

// Высокоуровневая обёртка для запуска алгоритма Cutting Plan из UI
// slitMM      — толщина пилы, мм (если <= 0, используется значение по умолчанию)
// floorIndex  — пока заглушка, для будущего размещения объектов на этаже
bool RunCuttingPlan(double slitMM, short floorIndex);

} // namespace CutPlanBoardHelper

#endif
