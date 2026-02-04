#ifndef FASTPRODUCTION_HPP
#define FASTPRODUCTION_HPP

#include "GSRoot.hpp"
#include "UniString.hpp"
#include "Array.hpp"
#include "CuttingStockSolver.hpp"

namespace FastProduction {

struct SetListRow {
	double boardW;
	GS::UniString scenarioId;
	int stopLength;
	int cutsCount;
	int boardsCount;
	int opOrder;
};

struct SetListSummaryRow {
	double boardW;
	int stopLength;
	int totalCuts;
};

struct ScenarioStepInfo {
	int stopLength;
	int cutsPerBoard;
	int totalCuts;
	int opOrder;
};

struct ScenarioInfo {
	GS::UniString scenarioId;
	double boardW;
	int boardsCount;
	double remainderMm;
	GS::Array<ScenarioStepInfo> steps;
};

struct ScenarioData {
	GS::Array<GS::UniString> boardScenarioId;
	GS::Array<GS::UniString> boardScenarioOps;
	GS::Array<Int32> boardScenarioSetups;
	GS::Array<GS::UniString> boardScenarioGroup;
	GS::Array<SetListRow> setListRows;
	GS::Array<SetListSummaryRow> setListSummaryRows;
	GS::Array<ScenarioInfo> scenarios;
};

/** Build ScenarioOps string from cuts: e.g. "3000x2" or "3090x1|590x4|549x1". */
GS::UniString ComputeBoardScenarioOps(const GS::Array<double>& cuts, int roundStepMm, bool sortDesc);

/** Build full scenario data from solver result. */
ScenarioData BuildScenarioData(const CuttingStock::SolverResult& result, int roundStepMm, bool sortDesc, UInt32 minGroupSize);

struct Run {
	int stopLength;
	int count;
};

/** Parse ScenarioOps string into runs (e.g. "3000x2|590x4" -> Run array). */
void ParseScenarioOps(const GS::UniString& scenarioOps, GS::Array<Run>& outRuns);

} // namespace FastProduction

#endif
