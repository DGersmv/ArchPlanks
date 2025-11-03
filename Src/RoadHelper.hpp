#pragma once

#include "ACAPinc.h"
#include "APIdefs_Elements.h"
#include "GSRoot.hpp"
#include "Array.hpp"

// RoadHelper
// Генерация дорожки/ленты по рельефу:
// - осевая линия (Spline / PolyLine / Arc / Line)
// - поверхность рельефа (Mesh)
// - ширина дорожки
// - точность профиля (шаг замера вдоль оси)
// Результат:
// - Morph по рельефу (3D тело)
// - Hatch (штриховка) в плане по границам
// - Текст "Расчётная площадь: NN.N м²"
// - Логирование через BrowserRepl

namespace RoadHelper {

	// Пользовательские параметры генерации
	struct RoadParams {
		double widthMM;         // ширина дорожки в мм
		double sampleStepMM;    // шаг съёма профиля вдоль осевой в мм (точность профиля)
	};

	// Одна осевая линия, которую выбрал пользователь
	bool SetCenterLine ();          // сохраняет GUID выбранной линии
	bool SetTerrainMesh ();         // сохраняет GUID выбранной Mesh

	// Основная команда
	bool BuildRoad (const RoadParams& params);

	// Вспомогательные функции для построения перпендикуляров
	bool BuildPerpendicularPoints(const GS::Array<API_Coord>& centerPts, double halfWidthM, 
	                             GS::Array<API_Coord>& leftPts, GS::Array<API_Coord>& rightPts);

	// Перенесено из MeshHelper: создание Morph из точек (НЕ создает mesh!)
	// thicknessMM: толщина Morph в мм (0 = плоский, только верхняя поверхность)
	// materialTop, materialBottom, materialSide: индексы материалов для граней
	bool CreateMorphFromPoints(const GS::Array<API_Coord3D>& points, double thicknessMM = 0.0,
	                           API_AttributeIndex materialTop = ACAPI_CreateAttributeIndex(1),
	                           API_AttributeIndex materialBottom = ACAPI_CreateAttributeIndex(2),
	                           API_AttributeIndex materialSide = ACAPI_CreateAttributeIndex(3));
	
	// Вычисление площади верхней поверхности Morph и создание текстовой выноски
	double CalculateMorphSurfaceArea(const GS::Array<API_Coord3D>& points);
	bool CreateAreaLabel(const API_Coord& position, double areaM2);
	
	// Получить список всех доступных покрытий (материалы с текстурами) (для UI)
	struct SurfaceFinishInfo {
		Int32 index;  // числовой индекс материала/покрытия (1, 2, 3, ...)
		GS::UniString name;
	};
	GS::Array<SurfaceFinishInfo> GetSurfaceFinishesList();
	void InvalidateSurfaceFinishesCache(); // Сбросить кэш покрытий

	// ------------------------------------------------------------------
	// Внутренние структуры (чтобы .cpp мог работать понятно)
	// ------------------------------------------------------------------

	struct SampledSectionPoint {
		API_Coord3D leftPt;
		API_Coord3D rightPt;
	};

	// Сырые точки с шагом по осевой линии (до сглаживания)
	struct RawProfile {
		GS::Array<API_Coord3D> leftSide;   // точки левого края (в 3D с z от рельефа)
		GS::Array<API_Coord3D> rightSide;  // точки правого края
	};

	// Сглаженные кривые после фитинга через сплайн и повторной дискретизации
	struct SmoothProfile {
		GS::Array<API_Coord3D> leftSide;   // сглаженный левый край (плотный)
		GS::Array<API_Coord3D> rightSide;  // сглаженный правый край (плотный)
	};

} // namespace RoadHelper
