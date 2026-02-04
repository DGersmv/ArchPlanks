#ifndef CUTTINGSTOCKSOLVER_HPP
#define CUTTINGSTOCKSOLVER_HPP

#include "GSRoot.hpp"
#include "UniString.hpp"
#include "Array.hpp"

namespace CuttingStock {

struct Part {
	double length;
	GS::UniString material;
	double boardW;  // iHeight — ширина доски для отображения/группировки
};

struct SolverParams {
	double maxStockLength;
	double slit;
	double trimLoss;
	double usefulMin;   // A — остаток >= usefulMin считается полезным
	double wasteMax;    // B — остаток <= wasteMax считается отходом
	bool strictAB;
	int maxImproveIter;
};

struct ResultBoard {
	GS::Array<double> cuts;
	double remainder;
	double boardW;
};

struct SolverResult {
	GS::Array<ResultBoard> boards;
	GS::Array<Part> remaining;
};

SolverResult Solve(const GS::Array<Part>& parts, const SolverParams& params);

} // namespace CuttingStock

#endif
