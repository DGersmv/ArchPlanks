#include "SelectionMetricsHelper.hpp"

#include <cmath>

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

struct QuantitySnapshot {
	double	topSurface = 0.0;
	double	totalSurface = 0.0;
	double	volume = 0.0;

	bool	hasTopSurface = false;
	bool	hasTotalSurface = false;
	bool	hasVolume = false;

	struct LayerComp {
		API_AttributeIndex	buildMatIndex = APIInvalidAttributeIndex;
		double				area = 0.0;
		double				volume = 0.0;
	};

	GS::Array<LayerComp>	layerComps;		// послойные данные (для многослойных конструкций)
	bool					hasLayerComps = false;
};

static GSErrCode GetQuantities(const API_Element& element, QuantitySnapshot& snapshot)
{
	API_ElementQuantity quantity = {};

	GS::Array<API_CompositeQuantity>			composites;
	GS::Array<API_ElemPartQuantity>			elemPartQuantities;
	GS::Array<API_ElemPartCompositeQuantity>	elemPartComposites;

	GS::Array<API_Quantities> quantities;
	GS::Array<API_Guid>		elemGuids;

	quantities.Push(API_Quantities());
	quantities[0].elements = &quantity;
	quantities[0].composites = &composites;
	quantities[0].elemPartQuantities = &elemPartQuantities;
	quantities[0].elemPartComposites = &elemPartComposites;
	elemGuids.Push(element.header.guid);

	API_QuantityPar params = {};
	params.minOpeningSize = 0.0;	// минимальный размер отверстий (0 = без отсечения)

	API_QuantitiesMask mask;
	ACAPI_ELEMENT_QUANTITIES_MASK_SETFULL(mask);

	GSErrCode err = ACAPI_Element_GetMoreQuantities(&elemGuids, &params, &quantities, &mask);
	if (err != NoError) {
		return err;
	}

	switch (element.header.type.typeID) {
	case API_MeshID:
		snapshot.topSurface = quantity.mesh.topSurface;
		snapshot.totalSurface = quantity.mesh.bottomSurface;
		snapshot.volume = quantity.mesh.volume;
		snapshot.hasTopSurface = true;
		snapshot.hasTotalSurface = true;
		snapshot.hasVolume = true;
		break;
	case API_SlabID:
		snapshot.topSurface = quantity.slab.topSurface;
		snapshot.totalSurface = quantity.slab.bottomSurface;
		snapshot.volume = quantity.slab.volume;
		snapshot.hasTopSurface = true;
		snapshot.hasTotalSurface = true;
		snapshot.hasVolume = true;
		break;
	case API_RoofID:
		snapshot.topSurface = quantity.roof.topSurface;
		snapshot.totalSurface = quantity.roof.bottomSurface;
		snapshot.volume = quantity.roof.volume;
		snapshot.hasTopSurface = true;
		snapshot.hasTotalSurface = true;
		snapshot.hasVolume = true;
		break;
	case API_ShellID:
		snapshot.topSurface = quantity.shell.referenceSurface;
		snapshot.totalSurface = quantity.shell.oppositeSurface;
		snapshot.volume = quantity.shell.volume;
		snapshot.hasTopSurface = true;
		snapshot.hasTotalSurface = true;
		snapshot.hasVolume = true;
		break;
	case API_MorphID:
		snapshot.topSurface = quantity.morph.surface;
		snapshot.totalSurface = quantity.morph.surface;
		snapshot.volume = quantity.morph.volume;
		snapshot.hasTopSurface = true;
		snapshot.hasTotalSurface = true;
		snapshot.hasVolume = true;
		break;
	default:
		break;
	}

	// послойные значения (для многослойных конструкций)
	for (const API_CompositeQuantity& comp : composites) {
		QuantitySnapshot::LayerComp layer;
		layer.buildMatIndex = comp.buildMatIndices;
		layer.area = comp.projectedArea;
		layer.volume = comp.volumes;
		snapshot.layerComps.Push(layer);
	}
	snapshot.hasLayerComps = !snapshot.layerComps.IsEmpty();

	return NoError;
}

static void AppendMetric(GS::Array<SelectionMetricsHelper::Metric>& dest, const GS::UniString& key,
	const GS::UniString& name, double grossValue, double netValue)
{
	// Порог отсечения шума: около 0.0005 м³ (третьего знака)
	const double eps = 0.0005;

	SelectionMetricsHelper::Metric metric;
	metric.key = key;
	metric.name = name;
	metric.grossValue = grossValue;
	metric.netValue = netValue;
	const double rawDiff = std::fabs(grossValue - netValue);
	metric.diffValue = (rawDiff < eps) ? 0.0 : rawDiff;
	dest.Push(metric);
}

// Добавить метрики по слоям (для многослойных конструкций)
static void AppendLayerMetrics(GS::Array<SelectionMetricsHelper::Metric>& dest,
	const QuantitySnapshot& grossSnapshot,
	const QuantitySnapshot& netSnapshot)
{
	if (!grossSnapshot.hasLayerComps && !netSnapshot.hasLayerComps) {
		return;
	}

	GS::Array<API_AttributeIndex> allIndices;
	auto addIndex = [&allIndices](API_AttributeIndex idx) {
		if (!idx.IsPositive()) {
			return;
		}
		for (UIndex i = 0; i < allIndices.GetSize(); ++i) {
			if (allIndices[i] == idx) {
				return;
			}
		}
		allIndices.Push(idx);
	};

	for (const auto& lc : grossSnapshot.layerComps) {
		addIndex(lc.buildMatIndex);
	}
	for (const auto& lc : netSnapshot.layerComps) {
		addIndex(lc.buildMatIndex);
	}

	// Сначала посчитаем суммарные "сырые" объёмы по всем слоям
	double grossRawTotal = 0.0;
	double netRawTotal = 0.0;
	for (const auto& lc : grossSnapshot.layerComps) {
		grossRawTotal += lc.volume;
	}
	for (const auto& lc : netSnapshot.layerComps) {
		netRawTotal += lc.volume;
	}

	for (API_AttributeIndex idx : allIndices) {
		double grossRaw = 0.0;
		double netRaw = 0.0;

		for (const auto& lc : grossSnapshot.layerComps) {
			if (lc.buildMatIndex == idx) {
				grossRaw += lc.volume;
			}
		}
		for (const auto& lc : netSnapshot.layerComps) {
			if (lc.buildMatIndex == idx) {
				netRaw += lc.volume;
			}
		}

		// Нормализуем послойные объёмы так, чтобы их сумма совпадала
		// с общим объёмом элемента (до/после SEO)
		double grossVolume = grossRaw;
		double netVolume = netRaw;

		if (grossRawTotal > 0.0 && grossSnapshot.hasVolume && grossSnapshot.volume > 0.0) {
			const double k = grossSnapshot.volume / grossRawTotal;
			grossVolume = grossRaw * k;
		}
		if (netRawTotal > 0.0 && netSnapshot.hasVolume && netSnapshot.volume > 0.0) {
			const double k = netSnapshot.volume / netRawTotal;
			netVolume = netRaw * k;
		}

		if (grossVolume == 0.0 && netVolume == 0.0) {
			continue;
		}

		API_Attribute attr = {};
		attr.header.typeID = API_BuildingMaterialID;
		attr.header.index = idx;
		GS::UniString matName("Материал ");
		matName.Append(GS::UniString::Printf("#%d", (int)idx.ToInt32_Deprecated()));
		if (ACAPI_Attribute_Get(&attr) == NoError) {
			matName = attr.header.name;
		}

		GS::UniString keyBase = GS::UniString::Printf("layer_%d_", (int)idx.ToInt32_Deprecated());

		// По слоям считаем только объёмы (площади оставляем на уровне всего элемента)
		if (grossVolume != 0.0 || netVolume != 0.0) {
			GS::UniString key = keyBase;
			key.Append("volume");
			GS::UniString name("Слой ");
			name.Append(matName);
			name.Append(" – Объем");
			AppendMetric(dest, key, name, grossVolume, netVolume);
		}
	}
}

static GSErrCode DetachSeoLinks(const API_Guid& guid)
{
	if (guid == APINULLGuid) {
		return APIERR_BADGUID;
	}

	GS::Array<API_Guid> operators;
	GSErrCode err = ACAPI_Element_SolidLink_GetOperators(guid, &operators);
	if (err == APIERR_NO3D) {
		return NoError;
	}
	if (err != NoError) {
		return err;
	}

	for (const API_Guid& oper : operators) {
		ACAPI_Element_SolidLink_Remove(guid, oper);
	}
	return NoError;
}

class TemporaryElementCopy {
public:
	explicit TemporaryElementCopy(const API_Element& sourceElement)
		: m_sourceElement(sourceElement)
	{
	}

	GSErrCode Create()
	{
		if (m_sourceElement.header.guid == APINULLGuid) {
			return APIERR_BADGUID;
		}

		API_ElementMemo memo = {};
		GSErrCode memoErr = ACAPI_Element_GetMemo(m_sourceElement.header.guid, &memo);
		if (memoErr != NoError && memoErr != APIERR_BADID) {
			return memoErr;
		}

		API_Element copyElement = m_sourceElement;
		copyElement.header.guid = APINULLGuid;

		GSErrCode createErr = ACAPI_Element_Create(&copyElement, (memoErr == NoError) ? &memo : nullptr);
		ACAPI_DisposeElemMemoHdls(&memo);
		if (createErr != NoError) {
			return createErr;
		}

		m_copyGuid = copyElement.header.guid;
		DetachSeoLinks(m_copyGuid);
		return NoError;
	}

	GSErrCode Destroy()
	{
		if (m_copyGuid == APINULLGuid) {
			return NoError;
		}
		GS::Array<API_Guid> guids;
		guids.Push(m_copyGuid);
		GSErrCode deleteErr = ACAPI_Element_Delete(guids);
		m_copyGuid = APINULLGuid;
		return deleteErr;
	}

	API_Guid GetGuid() const { return m_copyGuid; }

	~TemporaryElementCopy()
	{
		Destroy();
	}

private:
	API_Element	m_sourceElement = {};
	API_Guid	m_copyGuid = APINULLGuid;
};

static GSErrCode GetGrossQuantitiesViaCopy(const API_Element& sourceElement, QuantitySnapshot& snapshot)
{
	return ACAPI_CallUndoableCommand("SelectionMetrics_TemporaryCopy", [&]() -> GSErrCode {
		TemporaryElementCopy tempCopy(sourceElement);
		GSErrCode createErr = tempCopy.Create();
		if (createErr != NoError) {
			return createErr;
		}

		API_Element copyElement = sourceElement;
		copyElement.header.guid = tempCopy.GetGuid();

		GSErrCode qtyErr = GetQuantities(copyElement, snapshot);
		GSErrCode destroyErr = tempCopy.Destroy();
		if (destroyErr != NoError) {
			return destroyErr;
		}
		return qtyErr;
	});
}

} // namespace

GS::Array<SelectionMetricsHelper::Metric> SelectionMetricsHelper::CollectForGuid(const API_Guid& guid)
{
	GS::Array<Metric> metrics;

	if (guid == APINULLGuid) {
		return metrics;
	}

	API_Element element = {};
	element.header.guid = guid;
	if (ACAPI_Element_Get(&element) != NoError) {
		return metrics;
	}

	QuantitySnapshot netSnapshot;
	if (GetQuantities(element, netSnapshot) != NoError) {
		return metrics;
	}

	QuantitySnapshot grossSnapshot = netSnapshot;
	if (GetGrossQuantitiesViaCopy(element, grossSnapshot) != NoError) {
		grossSnapshot = netSnapshot;
	}

	if (netSnapshot.hasTotalSurface || grossSnapshot.hasTotalSurface) {
		AppendMetric(metrics, "totalArea", "Площадь", grossSnapshot.totalSurface, netSnapshot.totalSurface);
	}

	if (netSnapshot.hasTopSurface || grossSnapshot.hasTopSurface) {
		AppendMetric(metrics, "topSurface", "Площадь верхней поверхности", grossSnapshot.topSurface, netSnapshot.topSurface);
	}

	if (netSnapshot.hasVolume || grossSnapshot.hasVolume) {
		AppendMetric(metrics, "volume", "Объем", grossSnapshot.volume, netSnapshot.volume);
	}

	// Дополнительно: послойные метрики для многослойных конструкций
	AppendLayerMetrics(metrics, grossSnapshot, netSnapshot);

	return metrics;
}

GS::Array<SelectionMetricsHelper::Metric> SelectionMetricsHelper::CollectForFirstSelected()
{
	const API_Guid guid = GetFirstSelectedGuid();
	return CollectForGuid(guid);
}


