#include "CuttingStockSolver.hpp"
#include <cmath>

namespace CuttingStock {

namespace {

const double W1 = 1e9;   // штраф за новую заготовку
const double W2 = 1.0;
const double W3 = 5000.0; // штраф за отход

struct PlacedPart {
	double length;
	GS::UniString material;
	double boardW;
};

struct BoardState {
	double used;
	double boardW;
	GS::Array<PlacedPart> placed;
};

static bool IsInvalidRemainder(double remainder, const SolverParams& p) {
	if (remainder < 0) return true;
	if (p.strictAB && remainder > p.wasteMax && remainder < p.usefulMin) return true;
	return false;
}

static double ScoreRemainder(double remainder, bool willOpenNew, const SolverParams& p) {
	if (IsInvalidRemainder(remainder, p)) return 1e12;
	double score = 0;
	if (willOpenNew) score += W1;
	score += W2 * remainder;
	if (remainder <= p.wasteMax) score += W3;
	return score;
}

static SolverResult SolveGreedy(GS::Array<Part> parts, const SolverParams& params) {
	SolverResult result;
	if (parts.IsEmpty()) return result;

	// Сортировка по убыванию длины
	for (UIndex i = 0; i < parts.GetSize(); ++i) {
		for (UIndex j = i + 1; j < parts.GetSize(); ++j) {
			if (parts[j].length > parts[i].length) {
				Part tmp = parts[i];
				parts[i] = parts[j];
				parts[j] = tmp;
			}
		}
	}

	GS::Array<BoardState> boards;
	const double maxL = params.maxStockLength;

	for (UIndex i = 0; i < parts.GetSize(); ++i) {
		const Part& p = parts[i];
		if (p.length + params.trimLoss > maxL) {
			result.remaining.Push(p);
			continue;
		}

		bool placed = false;
		double bestScore = 1e12;
		UIndex bestBoardIdx = 0;
		bool bestIsNew = false;

		for (UIndex b = 0; b < boards.GetSize(); ++b) {
			BoardState& st = boards[b];
			if (st.boardW != p.boardW) continue; // пока группируем по boardW
			double need = st.placed.IsEmpty() ? (params.trimLoss + p.length) : (st.used + params.slit + p.length);
			if (need > maxL) continue;
			double rem = maxL - need;
			double sc = ScoreRemainder(rem, false, params);
			if (sc < bestScore) {
				bestScore = sc;
				bestBoardIdx = b;
				bestIsNew = false;
				placed = true;
			}
		}

		if (!placed || bestScore > W1 * 0.5) {
			double remNew = maxL - params.trimLoss - p.length;
			double scNew = ScoreRemainder(remNew, true, params);
			if (scNew <= bestScore) {
				bestIsNew = true;
				placed = true;
			}
		}

		if (!placed) {
			result.remaining.Push(p);
			continue;
		}

		if (bestIsNew) {
			BoardState st;
			st.used = params.trimLoss + p.length;
			st.boardW = p.boardW;
			st.placed.Push({ p.length, p.material, p.boardW });
			boards.Push(st);
		} else {
			BoardState& st = boards[bestBoardIdx];
			if (!st.placed.IsEmpty())
				st.used += params.slit;
			st.used += p.length;
			st.placed.Push({ p.length, p.material, p.boardW });
		}
	}

	for (UIndex bi = 0; bi < boards.GetSize(); ++bi) {
		const BoardState& st = boards[bi];
		ResultBoard rb;
		rb.boardW = st.boardW;
		rb.remainder = maxL - st.used;
		for (UIndex pi = 0; pi < st.placed.GetSize(); ++pi)
			rb.cuts.Push(st.placed[pi].length);
		result.boards.Push(rb);
	}

	return result;
}

} // anonymous

SolverResult Solve(const GS::Array<Part>& parts, const SolverParams& params) {
	if (parts.IsEmpty()) return SolverResult();

	GS::Array<Part> sorted = parts;
	return SolveGreedy(sorted, params);
}

} // namespace CuttingStock
