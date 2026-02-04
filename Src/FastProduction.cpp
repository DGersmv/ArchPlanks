#include "FastProduction.hpp"
#include <cmath>
#include <cstdio>

namespace FastProduction {

namespace {

static int RoundToStep(double value, int stepMm)
{
	if (stepMm <= 0)
		return static_cast<int>(std::round(value));
	double v = std::round(value / stepMm) * stepMm;
	return static_cast<int>(std::round(v));
}

static void NormalizeCuts(const GS::Array<double>& cuts, int roundStepMm, GS::Array<int>& out)
{
	out.Clear();
	for (UIndex i = 0; i < cuts.GetSize(); ++i) {
		double c = cuts[i];
		if (c != c || c <= 0.0)  // NaN or non-positive
			continue;
		int len = RoundToStep(c, roundStepMm);
		if (len > 0)
			out.Push(len);
	}
}

static void SortLengthsDesc(GS::Array<int>& lengths)
{
	for (UIndex i = 0; i < lengths.GetSize(); ++i) {
		for (UIndex j = i + 1; j < lengths.GetSize(); ++j) {
			if (lengths[j] > lengths[i]) {
				int t = lengths[i];
				lengths[i] = lengths[j];
				lengths[j] = t;
			}
		}
	}
}

static int CountScenarioSetups(const GS::UniString& scenarioOps)
{
	GS::Array<Run> runs;
	ParseScenarioOps(scenarioOps, runs);
	return static_cast<int>(runs.GetSize());
}

} // anonymous

GS::UniString ComputeBoardScenarioOps(const GS::Array<double>& cuts, int roundStepMm, bool sortDesc)
{
	GS::Array<int> lengths;
	NormalizeCuts(cuts, roundStepMm, lengths);
	if (lengths.IsEmpty())
		return GS::UniString();

	if (sortDesc)
		SortLengthsDesc(lengths);

	GS::UniString result;
	int runLen = lengths[0];
	int runCount = 1;
	for (UIndex i = 1; i <= lengths.GetSize(); ++i) {
		int cur = (i < lengths.GetSize()) ? lengths[i] : -1;
		if (i < lengths.GetSize() && cur == runLen) {
			++runCount;
		} else {
			if (!result.IsEmpty())
				result += "|";
			result += GS::UniString::Printf("%dx%d", runLen, runCount);
			if (i < lengths.GetSize()) {
				runLen = cur;
				runCount = 1;
			}
		}
	}
	return result;
}

void ParseScenarioOps(const GS::UniString& scenarioOps, GS::Array<Run>& outRuns)
{
	outRuns.Clear();
	GS::UniString s = scenarioOps;
	s.Trim();
	if (s.IsEmpty())
		return;

	const char* p = s.ToCStr().Get();
	if (!p)
		return;

	for (;;) {
		int len = 0;
		int count = 0;
		int n = 0;
		if (std::sscanf(p, "%d%*[xX]%d%n", &len, &count, &n) >= 2 && len > 0 && count > 0) {
			Run r;
			r.stopLength = len;
			r.count = count;
			outRuns.Push(r);
			p += n;
		} else {
			break;
		}
		while (*p == '|' || *p == ' ' || *p == '\t')
			++p;
		if (*p == '\0')
			break;
	}
}

ScenarioData BuildScenarioData(const CuttingStock::SolverResult& result, int roundStepMm, bool sortDesc, UInt32 minGroupSize)
{
#ifdef DEBUG
	static bool testsRun = false;
	if (!testsRun) {
		testsRun = true;
		(void)RunFastProductionTests();
	}
#endif
	ScenarioData data;
	if (result.boards.IsEmpty())
		return data;

	const int kRoundStep = (roundStepMm > 0) ? roundStepMm : 1;

	// Per-board ScenarioOps
	GS::Array<GS::UniString> opsList;
	opsList.SetSize(result.boards.GetSize());
	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		const CuttingStock::ResultBoard& rb = result.boards[b];
		opsList[b] = ComputeBoardScenarioOps(rb.cuts, kRoundStep, sortDesc);
	}

	// Unique (BoardW, ScenarioOps) -> scenario index, and board count per scenario
	struct ScenarioKey {
		double boardW;
		GS::UniString scenarioOps;
		bool operator==(const ScenarioKey& o) const {
			return std::fabs(boardW - o.boardW) < 0.001 && scenarioOps == o.scenarioOps;
		}
	};
	GS::Array<ScenarioKey> uniqueKeys;
	GS::Array<UIndex> boardToScenarioIndex;
	boardToScenarioIndex.SetSize(result.boards.GetSize());

	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		double boardW = result.boards[b].boardW;
		const GS::UniString& ops = opsList[b];
		ScenarioKey key;
		key.boardW = boardW;
		key.scenarioOps = ops;

		UIndex idx = uniqueKeys.GetSize();
		for (UIndex k = 0; k < uniqueKeys.GetSize(); ++k) {
			if (uniqueKeys[k] == key) {
				idx = k;
				break;
			}
		}
		if (idx == uniqueKeys.GetSize())
			uniqueKeys.Push(key);
		boardToScenarioIndex[b] = idx;
	}

	// Count boards per scenario
	GS::Array<UInt32> scenarioBoardCount;
	scenarioBoardCount.SetSize(uniqueKeys.GetSize());
	for (UIndex k = 0; k < scenarioBoardCount.GetSize(); ++k)
		scenarioBoardCount[k] = 0;
	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		UIndex k = boardToScenarioIndex[b];
		scenarioBoardCount[k] += 1;
	}

	// ScenarioId: W{BoardW}_S{index:02d} per (BoardW) group
	GS::Array<Int32> scenarioIdPerKey;
	scenarioIdPerKey.SetSize(uniqueKeys.GetSize());
	GS::Array<double> distinctBoardWs;
	for (UIndex k = 0; k < uniqueKeys.GetSize(); ++k) {
		double w = uniqueKeys[k].boardW;
		bool found = false;
		for (UIndex d = 0; d < distinctBoardWs.GetSize(); ++d) {
			if (std::fabs(distinctBoardWs[d] - w) < 0.001) { found = true; break; }
		}
		if (!found)
			distinctBoardWs.Push(w);
	}
	// Sort distinct widths
	for (UIndex i = 0; i < distinctBoardWs.GetSize(); ++i) {
		for (UIndex j = i + 1; j < distinctBoardWs.GetSize(); ++j) {
			if (distinctBoardWs[j] < distinctBoardWs[i]) {
				double t = distinctBoardWs[i];
				distinctBoardWs[i] = distinctBoardWs[j];
				distinctBoardWs[j] = t;
			}
		}
	}
	GS::Array<Int32> scenarioIndexPerBoardW;
	for (UIndex k = 0; k < uniqueKeys.GetSize(); ++k) {
		double w = uniqueKeys[k].boardW;
		Int32 localIdx = 0;
		for (UIndex d = 0; d < distinctBoardWs.GetSize(); ++d) {
			if (std::fabs(distinctBoardWs[d] - w) < 0.001) {
				for (UIndex j = 0; j < k; ++j) {
					if (std::fabs(uniqueKeys[j].boardW - w) < 0.001)
						++localIdx;
				}
				break;
			}
		}
		scenarioIdPerKey[k] = localIdx;
	}
	// Recompute local index per width
	GS::Array<GS::UniString> scenarioIds;
	scenarioIds.SetSize(uniqueKeys.GetSize());
	for (UIndex k = 0; k < uniqueKeys.GetSize(); ++k) {
		double w = uniqueKeys[k].boardW;
		Int32 localIdx = 0;
		for (UIndex j = 0; j < k; ++j) {
			if (std::fabs(uniqueKeys[j].boardW - w) < 0.001)
				++localIdx;
		}
		scenarioIds[k] = GS::UniString::Printf("W%.0f_S%02d", w, localIdx);
	}

	// Fill per-board data
	data.boardScenarioId.SetSize(result.boards.GetSize());
	data.boardScenarioOps.SetSize(result.boards.GetSize());
	data.boardScenarioSetups.SetSize(result.boards.GetSize());
	data.boardScenarioGroup.SetSize(result.boards.GetSize());

	for (UIndex b = 0; b < result.boards.GetSize(); ++b) {
		UIndex k = boardToScenarioIndex[b];
		data.boardScenarioId[b] = scenarioIds[k];
		data.boardScenarioOps[b] = opsList[b];
		data.boardScenarioSetups[b] = CountScenarioSetups(opsList[b]);
		data.boardScenarioGroup[b] = (scenarioBoardCount[k] >= minGroupSize) ? GS::UniString("FAST") : GS::UniString("TAIL");
	}

	// Set list rows and summary
	GS::Array<double> summaryKeysBoardW;
	GS::Array<int> summaryKeysStop;
	GS::Array<int> summaryTotals;

	for (UIndex k = 0; k < uniqueKeys.GetSize(); ++k) {
		GS::Array<Run> runs;
		ParseScenarioOps(uniqueKeys[k].scenarioOps, runs);
		const int boardsCount = static_cast<int>(scenarioBoardCount[k]);
		const double boardW = uniqueKeys[k].boardW;
		const GS::UniString& scenarioId = scenarioIds[k];

		for (UIndex r = 0; r < runs.GetSize(); ++r) {
			SetListRow row;
			row.boardW = boardW;
			row.scenarioId = scenarioId;
			row.stopLength = runs[r].stopLength;
			row.cutsCount = runs[r].count * boardsCount;
			row.boardsCount = boardsCount;
			row.opOrder = static_cast<int>(r + 1);
			data.setListRows.Push(row);

			// Summary aggregate
			bool found = false;
			for (UIndex s = 0; s < summaryKeysBoardW.GetSize(); ++s) {
				if (std::fabs(summaryKeysBoardW[s] - boardW) < 0.001 && summaryKeysStop[s] == runs[r].stopLength) {
					summaryTotals[s] += runs[r].count * boardsCount;
					found = true;
					break;
				}
			}
			if (!found) {
				summaryKeysBoardW.Push(boardW);
				summaryKeysStop.Push(runs[r].stopLength);
				summaryTotals.Push(runs[r].count * boardsCount);
			}
		}

		// ScenarioInfo for operator instructions
		ScenarioInfo info;
		info.scenarioId = scenarioId;
		info.boardW = boardW;
		info.boardsCount = boardsCount;
		info.remainderMm = 0.0;
		for (UIndex bb = 0; bb < result.boards.GetSize(); ++bb) {
			if (boardToScenarioIndex[bb] == k) {
				info.remainderMm = result.boards[bb].remainder;
				break;
			}
		}
		for (UIndex r = 0; r < runs.GetSize(); ++r) {
			ScenarioStepInfo si;
			si.stopLength = runs[r].stopLength;
			si.cutsPerBoard = runs[r].count;
			si.totalCuts = runs[r].count * boardsCount;
			si.opOrder = static_cast<int>(r + 1);
			info.steps.Push(si);
		}
		data.scenarios.Push(info);
	}

	for (UIndex s = 0; s < summaryKeysBoardW.GetSize(); ++s) {
		SetListSummaryRow sr;
		sr.boardW = summaryKeysBoardW[s];
		sr.stopLength = summaryKeysStop[s];
		sr.totalCuts = summaryTotals[s];
		data.setListSummaryRows.Push(sr);
	}

	return data;
}

#ifdef DEBUG
namespace {
static bool RunFastProductionTests()
{
	GS::Array<double> cuts1;
	cuts1.Push(3000.0);
	cuts1.Push(3000.0);
	GS::UniString ops1 = ComputeBoardScenarioOps(cuts1, 1, true);
	if (ops1 != "3000x2") return false;

	GS::Array<double> cuts2;
	cuts2.Push(3090.0);
	cuts2.Push(590.0);
	cuts2.Push(590.0);
	cuts2.Push(590.0);
	cuts2.Push(590.0);
	cuts2.Push(549.0);
	GS::UniString ops2 = ComputeBoardScenarioOps(cuts2, 1, true);
	if (ops2 != "3090x1|590x4|549x1") return false;

	GS::Array<Run> runs;
	ParseScenarioOps(GS::UniString("3000x2"), runs);
	if (runs.GetSize() != 1 || runs[0].stopLength != 3000 || runs[0].count != 2) return false;

	ParseScenarioOps(GS::UniString("3090x1|590x4|549x1"), runs);
	if (runs.GetSize() != 3) return false;
	if (runs[0].stopLength != 3090 || runs[0].count != 1) return false;
	if (runs[1].stopLength != 590 || runs[1].count != 4) return false;
	if (runs[2].stopLength != 549 || runs[2].count != 1) return false;

	return true;
}
} // anonymous
#endif

} // namespace FastProduction
