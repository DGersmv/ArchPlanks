#pragma once

#include "GSRoot.hpp"
#include "UniString.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

class SelectionMetricsHelper
{
public:
	struct Metric {
		GS::UniString	key;
		GS::UniString	name;
		double			grossValue = 0.0;	// without SEO
		double			netValue = 0.0;		// current (with SEO)
		double			diffValue = 0.0;	// |gross - net|
	};

	static GS::Array<Metric> CollectForFirstSelected();
	static GS::Array<Metric> CollectForGuid(const API_Guid& guid);
};


