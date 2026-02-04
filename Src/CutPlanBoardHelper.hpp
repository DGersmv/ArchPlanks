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
	double widthMM;                  // iWidth (толщина, мм)
	double heightMM;                 // iHeight (ширина, мм)
	GS::UniString materialLabel;     // человекочитаемый текст (например, "50×200 мм")
	UInt32 count = 0;                // количество досок в этой группе
	double totalLenMM = 0.0;         // суммарная длина всех досок этого типа, мм
	double maxLenMM = 0.0;           // исходная длина доски (iMaxLen), мм
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
// extraLenMM  — допуск по длине доски, мм (сколько можно «добавить» к iMaxLen при расчёте)
// floorIndex  — пока заглушка, для будущего размещения объектов на этаже
bool RunCuttingPlan(double slitMM, double extraLenMM, short floorIndex);

// Выводит в отчёт Archicad все AddPar (имя, тип, значение) для выбранных ArchiFramePlank.
// Полезно для определения реальных имён параметров в GDL (iHeight, iWidth и т.д.).
void DumpArchiFramePlankParamsToReport();

} // namespace CutPlanBoardHelper

#endif
