#include "SelectionPropertyHelper.hpp"

namespace {

API_Guid GetFirstSelectedGuid()
{
	API_SelectionInfo selectionInfo = {};
	GS::Array<API_Neig> selNeigs;
	if (ACAPI_Selection_Get(&selectionInfo, &selNeigs, false, false) == NoError && !selNeigs.IsEmpty()) {
		const API_Guid guid = selNeigs[0].guid;
		BMKillHandle(reinterpret_cast<GSHandle*>(&selectionInfo.marquee.coords));
		return guid;
	}
	BMKillHandle(reinterpret_cast<GSHandle*>(&selectionInfo.marquee.coords));
	return APINULLGuid;
}

GS::UniString ToString(const API_Property& property)
{
	if (property.status != API_Property_HasValue) {
		return GS::UniString();
	}
	if (property.value.variantStatus != API_VariantStatusNormal) {
		return GS::UniString();
	}

	GS::UniString valueString;
	if (ACAPI_Property_GetPropertyValueString(property, &valueString) == NoError) {
		return valueString;
	}
	return GS::UniString();
}

} // namespace

GS::Array<SelectionPropertyHelper::PropertyInfo> SelectionPropertyHelper::CollectForGuid(const API_Guid& guid)
{
	GS::Array<PropertyInfo> results;
	if (guid == APINULLGuid) {
		return results;
	}

	API_Element element = {};
	element.header.guid = guid;
	if (ACAPI_Element_Get(&element) != NoError) {
		return results;
	}

	GS::Array<API_PropertyDefinition> definitions;
	if (ACAPI_Element_GetPropertyDefinitions(guid, API_PropertyDefinitionFilter_All, definitions) != NoError) {
		return results;
	}

	GS::Array<API_Property> properties;
	if (ACAPI_Element_GetPropertyValues(guid, definitions, properties) != NoError) {
		return results;
	}

	for (const API_Property& property : properties) {
		if (property.status != API_Property_HasValue) {
			continue;
		}
		if (property.value.variantStatus != API_VariantStatusNormal) {
			continue;
		}

		PropertyInfo info;
		info.propertyGuid = property.definition.guid;
		info.propertyName = property.definition.name;
		info.valueString = ToString(property);
		results.Push(info);
	}
	return results;
}

GS::Array<SelectionPropertyHelper::PropertyInfo> SelectionPropertyHelper::CollectForFirstSelected()
{
	const API_Guid guid = GetFirstSelectedGuid();
	return CollectForGuid(guid);
}


