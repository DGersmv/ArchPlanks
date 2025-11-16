#pragma once

#include "GSRoot.hpp"
#include "UniString.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

class SelectionPropertyHelper
{
public:
	struct PropertyInfo {
		API_Guid		propertyGuid;
		GS::UniString	propertyName;
		GS::UniString	valueString;
	};

	static GS::Array<PropertyInfo> CollectForFirstSelected();
	static GS::Array<PropertyInfo> CollectForGuid(const API_Guid& guid);
};
