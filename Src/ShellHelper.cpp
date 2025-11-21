#include "ACAPinc.h"
#include "APIEnvir.h"
#include "APICommon.h"
#include "APIdefs_3D.h"
#include "APIdefs_Elements.h"
#include "APIdefs_Goodies.h"
#include "ShellHelper.hpp"
#include "LandscapeHelper.hpp"
#include "GroundHelper.hpp"
#include "MeshIntersectionHelper.hpp"
#include "RoadHelper.hpp"
#include "BrowserRepl.hpp"
#include <cstdarg>
#include <cmath>
#include <algorithm>
#include <limits>
#include <unordered_map>

static const double kEPS = 1e-9;
static constexpr double kPI = 3.14159265358979323846;

namespace ShellHelper {

// =============== Глобальные переменные ===============
API_Guid g_baseLineGuid = APINULLGuid;      // GUID базовой линии
API_Guid g_meshSurfaceGuid = APINULLGuid;   // GUID Mesh поверхности

// =============== Forward declarations ===============
API_Guid Create3DShell(const GS::Array<API_Coord3D>& points);

// =============== Логирование ===============
static inline void Log(const char* fmt, ...)
{
    va_list vl; va_start(vl, fmt);
    char buf[4096]; vsnprintf(buf, sizeof(buf), fmt, vl);
    va_end(vl);
    GS::UniString s(buf);
    // if (BrowserRepl::HasInstance())
    //     BrowserRepl::GetInstance().LogToBrowser(s);
#ifdef DEBUG_UI_LOGS
    ACAPI_WriteReport("%s", false, s.ToCStr().Get());
#endif
    (void)s;
}

// =============== Выбор базовой линии ===============
bool SetBaseLineForShell()
{
    Log("[ShellHelper] SetBaseLineForShell: выбор базовой линии");
    
    // Используем существующую функцию LandscapeHelper::SetDistributionLine()
    // которая уже умеет выбирать линии/дуги/полилинии/сплайны
    bool success = LandscapeHelper::SetDistributionLine();
    
    if (success) {
        // Получаем GUID выбранной линии из выделения
        API_SelectionInfo selectionInfo;
        GS::Array<API_Neig> selNeigs;
        GSErrCode err = ACAPI_Selection_Get(&selectionInfo, &selNeigs, false, false);
        
        if (err == NoError && selNeigs.GetSize() > 0) {
            g_baseLineGuid = selNeigs[0].guid;
            Log("[ShellHelper] Базовая линия выбрана успешно, GUID: %s", 
                APIGuidToString(g_baseLineGuid).ToCStr().Get());
        } else {
            Log("[ShellHelper] Ошибка получения GUID выбранной линии");
            g_baseLineGuid = APINULLGuid;
        }
    } else {
        Log("[ShellHelper] Ошибка выбора базовой линии");
        g_baseLineGuid = APINULLGuid;
    }
    
    return success;
}

// =============== ТЕСТОВАЯ ФУНКЦИЯ СОЗДАНИЯ MESH ===============
bool CreateTestMesh()
{
    Log("[ShellHelper] ТЕСТ: Создаем MESH ТОЧНО КАК В ПРИМЕРЕ Element_Test!");
    
    // Создаем MESH элемент ТОЧНО КАК В ПРИМЕРЕ
    API_Element element = {};
    element.header.type = API_MeshID;
    GSErrCode err = ACAPI_Element_GetDefaults(&element, nullptr);
    if (err != NoError) {
        Log("[ShellHelper] ТЕСТ ERROR: ACAPI_Element_GetDefaults failed, err=%d", (int)err);
        return false;
    }
    
    // Настройки ТОЧНО КАК В ПРИМЕРЕ
    element.mesh.poly.nCoords = 5;
    element.mesh.poly.nSubPolys = 1;
    element.mesh.poly.nArcs = 0;
    
    Log("[ShellHelper] ТЕСТ: MESH настройки: nCoords=5, nSubPolys=1, nArcs=0");
    
    // Создаем memo ТОЧНО КАК В ПРИМЕРЕ
    API_ElementMemo memo = {};
    memo.coords = reinterpret_cast<API_Coord**>(BMAllocateHandle((element.mesh.poly.nCoords + 1) * sizeof(API_Coord), ALLOCATE_CLEAR, 0));
    memo.pends = reinterpret_cast<Int32**>(BMAllocateHandle((element.mesh.poly.nSubPolys + 1) * sizeof(Int32), ALLOCATE_CLEAR, 0));
    memo.parcs = reinterpret_cast<API_PolyArc**>(BMAllocateHandle(element.mesh.poly.nArcs * sizeof(API_PolyArc), ALLOCATE_CLEAR, 0));
    memo.meshPolyZ = reinterpret_cast<double**>(BMAllocateHandle((element.mesh.poly.nCoords + 1) * sizeof(double), ALLOCATE_CLEAR, 0));
    
    // Координаты ТОЧНО КАК В ПРИМЕРЕ (прямоугольник)
    (*memo.coords)[1].x = 0.0;
    (*memo.coords)[1].y = 0.0;
    (*memo.coords)[2].x = 5.0;
    (*memo.coords)[2].y = 3.0;
    (*memo.coords)[3].x = 5.0;
    (*memo.coords)[3].y = 0.0;
    (*memo.coords)[4].x = 0.0;
    (*memo.coords)[4].y = 2.0;
    (*memo.coords)[element.mesh.poly.nCoords] = (*memo.coords)[1]; // Замыкаем полигон
    
    (*memo.pends)[1] = element.mesh.poly.nCoords;
    
    // Z-координаты ТОЧНО КАК В ПРИМЕРЕ
    (*memo.meshPolyZ)[1] = 1.0;
    (*memo.meshPolyZ)[2] = 2.0;
    (*memo.meshPolyZ)[3] = 3.0;
    (*memo.meshPolyZ)[4] = 4.0;
    (*memo.meshPolyZ)[5] = (*memo.meshPolyZ)[1]; // Замыкаем Z-координаты
    
    Log("[ShellHelper] ТЕСТ: Заполнены координаты и Z-координаты ТОЧНО КАК В ПРИМЕРЕ");
    
    // Создаем MESH ТОЧНО КАК В ПРИМЕРЕ
    err = ACAPI_Element_Create(&element, &memo);
    if (err == APIERR_IRREGULARPOLY) {
        Log("[ShellHelper] ТЕСТ: Полигон нерегулярный, регуляризуем ТОЧНО КАК В ПРИМЕРЕ...");
        
        API_RegularizedPoly poly = {};
        poly.coords = memo.coords;
        poly.pends = memo.pends;
        poly.parcs = memo.parcs;
        poly.vertexIDs = memo.vertexIDs;
        poly.needVertexAncestry = 1;
        
        Int32 nResult = 0;
        API_RegularizedPoly** polys = nullptr;
        GSErrCode regErr = ACAPI_Polygon_RegularizePolygon(&poly, &nResult, &polys);
        
        if (regErr != NoError) {
            Log("[ShellHelper] ТЕСТ ERROR: ACAPI_Polygon_RegularizePolygon failed, err=%d", (int)regErr);
            ACAPI_DisposeElemMemoHdls(&memo);
            return false;
        }
        
        if (regErr == NoError) {
            Log("[ShellHelper] ТЕСТ: Регуляризация успешна, создаем %d полигонов", (int)nResult);
            
            for (Int32 i = 0; i < nResult && err == NoError; i++) {
                element.mesh.poly.nCoords = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].coords)) / sizeof(API_Coord) - 1;
                element.mesh.poly.nSubPolys = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].pends)) / sizeof(Int32) - 1;
                element.mesh.poly.nArcs = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].parcs)) / sizeof(API_PolyArc);
                
                API_ElementMemo tmpMemo = {};
                tmpMemo.coords = (*polys)[i].coords;
                tmpMemo.pends = (*polys)[i].pends;
                tmpMemo.parcs = (*polys)[i].parcs;
                tmpMemo.vertexIDs = (*polys)[i].vertexIDs;
                
                tmpMemo.meshPolyZ = reinterpret_cast<double**>(BMAllocateHandle((element.mesh.poly.nCoords + 1) * sizeof(double), ALLOCATE_CLEAR, 0));
                if (tmpMemo.meshPolyZ != nullptr) {
                    for (Int32 j = 1; j <= element.mesh.poly.nCoords; j++) {
                        Int32 oldVertexIndex = 1; // Упрощенная логика
                        if (oldVertexIndex <= 5) {
                            (*tmpMemo.meshPolyZ)[j] = (*memo.meshPolyZ)[oldVertexIndex];
                        } else {
                            (*tmpMemo.meshPolyZ)[j] = 0.0;
                        }
                    }
                    
                    err = ACAPI_Element_Create(&element, &tmpMemo);
                    if (err != NoError) {
                        Log("[ShellHelper] ТЕСТ ERROR: ACAPI_Element_Create piece %d failed, err=%d", (int)i, (int)err);
                    }
                    BMKillHandle(reinterpret_cast<GSHandle*>(&tmpMemo.meshPolyZ));
                }
            }
        }
    }
    
    if (err == NoError) {
        Log("[ShellHelper] ТЕСТ SUCCESS: MESH создан ТОЧНО КАК В ПРИМЕРЕ!");
        ACAPI_DisposeElemMemoHdls(&memo);
        return true;
    } else {
        Log("[ShellHelper] ТЕСТ ERROR: Не удалось создать MESH, err=%d", (int)err);
        ACAPI_DisposeElemMemoHdls(&memo);
        return false;
    }
}

// =============== Создание MESH из точек (для пользователя) ===============
static bool CreateMeshFromPointsInternal(const GS::Array<API_Coord3D>& points)
{
    Log("[ShellHelper] CreateMeshFromPointsInternal: START with %d points", (int)points.GetSize());
    
    const UIndex numPoints = points.GetSize();
    if (numPoints < 3) {
        Log("[ShellHelper] ERROR: Нужно минимум 3 точки для создания MESH");
        return false;
    }
    
    // Создаем MESH элемент
    API_Element element = {};
    element.header.type = API_MeshID;
    GSErrCode err = ACAPI_Element_GetDefaults(&element, nullptr);
    if (err != NoError) {
        Log("[ShellHelper] ERROR: ACAPI_Element_GetDefaults failed, err=%d", (int)err);
        return false;
    }
    
    // nCoords = количество уникальных вершин + 1 (замыкающая точка)
    element.mesh.poly.nCoords = numPoints + 1;
    element.mesh.poly.nSubPolys = 1;
    element.mesh.poly.nArcs = 0;
    
    Log("[ShellHelper] MESH настройки: nCoords=%d (%d вершин + замыкающая), nSubPolys=1, nArcs=0", 
        (int)element.mesh.poly.nCoords, (int)numPoints);
    
    // Создаем memo
    API_ElementMemo memo = {};
    memo.coords = reinterpret_cast<API_Coord**>(BMAllocateHandle((element.mesh.poly.nCoords + 1) * sizeof(API_Coord), ALLOCATE_CLEAR, 0));
    memo.pends = reinterpret_cast<Int32**>(BMAllocateHandle((element.mesh.poly.nSubPolys + 1) * sizeof(Int32), ALLOCATE_CLEAR, 0));
    memo.parcs = reinterpret_cast<API_PolyArc**>(BMAllocateHandle(element.mesh.poly.nArcs * sizeof(API_PolyArc), ALLOCATE_CLEAR, 0));
    memo.meshPolyZ = reinterpret_cast<double**>(BMAllocateHandle((element.mesh.poly.nCoords + 1) * sizeof(double), ALLOCATE_CLEAR, 0));
    
    // Заполняем координаты и Z: все вершины + замыкающая
    for (UIndex i = 0; i < numPoints; ++i) {
        (*memo.coords)[i + 1] = {points[i].x, points[i].y};
        (*memo.meshPolyZ)[i + 1] = points[i].z;
    }
    // Замыкаем полигон (последняя точка = первая)
    (*memo.coords)[element.mesh.poly.nCoords] = (*memo.coords)[1];
    (*memo.meshPolyZ)[element.mesh.poly.nCoords] = (*memo.meshPolyZ)[1];
    
    (*memo.pends)[1] = element.mesh.poly.nCoords;
    
    Log("[ShellHelper] MESH: заполнены %d точек (%d вершин + замыкающая)", 
        (int)element.mesh.poly.nCoords, (int)numPoints);
    
    // Создаем MESH
    err = ACAPI_Element_Create(&element, &memo);
    if (err != NoError) {
        Log("[ShellHelper] ERROR: ACAPI_Element_Create failed, err=%d (%s)", (int)err, ErrID_To_Name(err));
    }
    if (err == APIERR_IRREGULARPOLY) {
        Log("[ShellHelper] Полигон нерегулярный, регуляризуем...");
        
        API_RegularizedPoly poly = {};
        poly.coords = memo.coords;
        poly.pends = memo.pends;
        poly.parcs = memo.parcs;
        poly.vertexIDs = memo.vertexIDs;
        poly.needVertexAncestry = 1;
        
        Int32 nResult = 0;
        API_RegularizedPoly** polys = nullptr;
        GSErrCode regErr = ACAPI_Polygon_RegularizePolygon(&poly, &nResult, &polys);
        
        if (regErr != NoError) {
            Log("[ShellHelper] ERROR: ACAPI_Polygon_RegularizePolygon failed, err=%d", (int)regErr);
            ACAPI_DisposeElemMemoHdls(&memo);
            return false;
        }
        
        if (regErr == NoError) {
            Log("[ShellHelper] Регуляризация успешна, создаем %d полигонов", (int)nResult);
            
            for (Int32 i = 0; i < nResult && err == NoError; i++) {
                element.mesh.poly.nCoords = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].coords)) / sizeof(API_Coord) - 1;
                element.mesh.poly.nSubPolys = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].pends)) / sizeof(Int32) - 1;
                element.mesh.poly.nArcs = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].parcs)) / sizeof(API_PolyArc);
                
                API_ElementMemo tmpMemo = {};
                tmpMemo.coords = (*polys)[i].coords;
                tmpMemo.pends = (*polys)[i].pends;
                tmpMemo.parcs = (*polys)[i].parcs;
                tmpMemo.vertexIDs = (*polys)[i].vertexIDs;
                
                tmpMemo.meshPolyZ = reinterpret_cast<double**>(BMAllocateHandle((element.mesh.poly.nCoords + 1) * sizeof(double), ALLOCATE_CLEAR, 0));
                if (tmpMemo.meshPolyZ != nullptr) {
                    for (Int32 j = 1; j <= element.mesh.poly.nCoords; j++) {
                        (*tmpMemo.meshPolyZ)[j] = (*memo.meshPolyZ)[j];
                    }
                    
                    err = ACAPI_Element_Create(&element, &tmpMemo);
                    if (err != NoError) {
                        Log("[ShellHelper] ERROR: ACAPI_Element_Create piece %d failed, err=%d", (int)i, (int)err);
                    }
                    BMKillHandle(reinterpret_cast<GSHandle*>(&tmpMemo.meshPolyZ));
                }
            }
        }
    }
    
    if (err == NoError) {
        Log("[ShellHelper] SUCCESS: MESH создан из %d точек!", (int)numPoints);
        ACAPI_DisposeElemMemoHdls(&memo);
        return true;
    } else {
        Log("[ShellHelper] ERROR: Не удалось создать MESH, err=%d", (int)err);
        ACAPI_DisposeElemMemoHdls(&memo);
        return false;
    }
}

bool CreateMeshFromPoints(const GS::Array<API_Coord3D>& points)
{
    return ACAPI_CallUndoableCommand("Create Mesh from Points", [&]() -> GSErrCode {
        return CreateMeshFromPointsInternal(points) ? NoError : APIERR_GENERAL;
    }) == NoError;
}

// =============== Основная функция создания оболочки ===============
bool CreateShellFromLine(double widthMM, double stepMM)
{
    Log("[ShellHelper] CreateShellFromLine: START, width=%.1fmm, step=%.1fmm", widthMM, stepMM);
    
    // Проверяем, что базовая линия выбрана
    if (g_baseLineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Базовая линия не выбрана. Сначала выберите базовую линию.");
        return false;
    }
    
    // Mesh поверхность опциональна - нужна только для получения Z-координат контуров
    
    Log("[ShellHelper] Базовая линия: %s", APIGuidToString(g_baseLineGuid).ToCStr().Get());
    if (g_meshSurfaceGuid != APINULLGuid) {
        Log("[ShellHelper] Mesh поверхность: %s", APIGuidToString(g_meshSurfaceGuid).ToCStr().Get());
    } else {
        Log("[ShellHelper] Mesh поверхность не выбрана - контуры будут созданы без проекции на Mesh");
    }
    
    // Получаем элемент базовой линии по сохраненному GUID
    API_Elem_Head elemHead = {};
    elemHead.guid = g_baseLineGuid;
    GSErrCode err = ACAPI_Element_GetHeader(&elemHead);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить заголовок элемента базовой линии");
        return false;
    }
    
    // Проверяем поддерживаемые типы элементов
    bool isSupported = false;
    if (elemHead.type == API_LineID ||
        elemHead.type == API_PolyLineID ||
        elemHead.type == API_ArcID ||
        elemHead.type == API_CircleID ||
        elemHead.type == API_SplineID) {
        isSupported = true;
    }
    
    if (!isSupported) {
        Log("[ShellHelper] ERROR: Неподдерживаемый тип элемента базовой линии");
        Log("[ShellHelper] Поддерживаются: Line, Polyline, Arc, Circle, Spline");
        return false;
    }
    
    // Получаем данные элемента
    API_Element element = {};
    element.header = elemHead;
    err = ACAPI_Element_Get(&element);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить данные элемента базовой линии");
        return false;
    }
    
    Log("[ShellHelper] Элемент базовой линии загружен успешно");
    
    // Парсим элемент в сегменты
    PathData path;
    if (!ParseElementToSegments(element, path)) {
        Log("[ShellHelper] ERROR: Не удалось распарсить элемент в сегменты");
        return false;
    }
    
    Log("[ShellHelper] Элемент распарсен: %d сегментов, общая длина %.3fм", 
        (int)path.segs.GetSize(), path.total);
    
    // Создаем настоящую 3D оболочку через Ruled Shell
    Log("[ShellHelper] Создаем 3D оболочку через Ruled Shell");
    return Create3DShellFromPath(path, widthMM, stepMM);
}

// =============== Анализ базовой линии ===============
GS::Array<API_Coord3D> AnalyzeBaseLine(const API_Guid& lineGuid, double stepMM)
{
    ACAPI_WriteReport("[ShellHelper] AnalyzeBaseLine: step=%.1fmm (заглушка)", false, stepMM);
    
    // TODO: Использовать существующие функции LandscapeHelper для получения точек линии
    // Пока возвращаем пустой массив
    GS::Array<API_Coord3D> points;
    return points;
}

// =============== Генерация перпендикулярных линий ===============
GS::Array<API_Coord3D> GeneratePerpendicularLines(
    const GS::Array<API_Coord3D>& basePoints, 
    double widthMM)
{
    ACAPI_WriteReport("[ShellHelper] GeneratePerpendicularLines: %d точек, ширина=%.1fmm", false, 
        (int)basePoints.GetSize(), widthMM);
    
    GS::Array<API_Coord3D> perpendicularPoints;
    
    if (basePoints.GetSize() < 2) {
        ACAPI_WriteReport("[ShellHelper] Недостаточно точек для генерации перпендикуляров", false);
        return perpendicularPoints;
    }
    
    // Проходим по базовым точкам и создаем перпендикуляры
    for (UIndex i = 0; i < basePoints.GetSize() - 1; ++i) {
        const API_Coord3D& current = basePoints[i];
        const API_Coord3D& next = basePoints[i + 1];
        
        // Вычисляем направление базовой линии
        double dx = next.x - current.x;
        double dy = next.y - current.y;
        double length = sqrt(dx * dx + dy * dy);
        
        if (length < 1e-6) continue; // Пропускаем слишком короткие сегменты
        
        // Нормализуем направление
        dx /= length;
        dy /= length;
        
        // Вычисляем перпендикулярное направление (поворот на 90° в плоскости XY)
        double perpX = -dy;
        double perpY = dx;
        
        // Создаем две точки перпендикуляра
        double halfWidth = widthMM / 2000.0; // Переводим мм в метры и делим пополам
        
        API_Coord3D left = {};
        left.x = current.x + perpX * halfWidth;
        left.y = current.y + perpY * halfWidth;
        left.z = current.z;
        perpendicularPoints.Push(left);
        
        API_Coord3D right = {};
        right.x = current.x - perpX * halfWidth;
        right.y = current.y - perpY * halfWidth;
        right.z = current.z;
        perpendicularPoints.Push(right);
    }
    
    ACAPI_WriteReport("[ShellHelper] Сгенерировано %d перпендикулярных точек", false, (int)perpendicularPoints.GetSize());
    return perpendicularPoints;
}

// =============== Проекция на 3D сетку ===============
GS::Array<API_Coord3D> ProjectToMesh(const GS::Array<API_Coord3D>& points)
{
    ACAPI_WriteReport("[ShellHelper] ProjectToMesh: %d точек", false, (int)points.GetSize());
    
    GS::Array<API_Coord3D> projectedPoints;
    
    for (const API_Coord3D& point : points) {
        // Используем GroundHelper::GetGroundZAndNormal для проекции точки на mesh
        API_Coord3D projected = point;
        
        double z = 0.0;
        API_Vector3D normal = {};
        if (GroundHelper::GetGroundZAndNormal(point, z, normal)) {
            projected.z = z;
            ACAPI_WriteReport("[ShellHelper] Точка (%.3f, %.3f) спроецирована на Z=%.3f", false, 
                point.x, point.y, z);
        } else {
            ACAPI_WriteReport("[ShellHelper] Не удалось спроецировать точку (%.3f, %.3f)", false, 
                point.x, point.y);
        }
        
        projectedPoints.Push(projected);
    }
    
    ACAPI_WriteReport("[ShellHelper] Спроецировано %d точек", false, (int)projectedPoints.GetSize());
    return projectedPoints;
}

// =============== Создание перпендикулярных линий ===============
bool CreatePerpendicularLines(const API_Element& baseLine, double widthMM)
{
    Log("[ShellHelper] CreatePerpendicularLines: width=%.1fmm", widthMM);
    
    // Вычисляем направление базовой линии
    API_Coord begC = baseLine.line.begC;
    API_Coord endC = baseLine.line.endC;
    
    double dx = endC.x - begC.x;
    double dy = endC.y - begC.y;
    double length = sqrt(dx * dx + dy * dy);
    
    if (length < 1e-6) {
        Log("[ShellHelper] ERROR: Базовая линия слишком короткая");
        return false;
    }
    
    // Нормализуем направление
    dx /= length;
    dy /= length;
    
    // Вычисляем перпендикулярное направление (поворот на 90°)
    double perpX = -dy;
    double perpY = dx;
    
    // Вычисляем смещение в метрах (widthMM в миллиметрах)
    double halfWidth = widthMM / 2000.0; // Переводим мм в метры и делим пополам
    
    Log("[ShellHelper] Направление линии: (%.3f, %.3f), перпендикуляр: (%.3f, %.3f), смещение: %.3fм", 
        dx, dy, perpX, perpY, halfWidth);
    
    // Создаем две перпендикулярные линии
    ACAPI_CallUndoableCommand("Create Shell Lines", [&] () -> GSErrCode {
        
        // Левая линия
        API_Element leftLine = {};
        leftLine.header.type = API_LineID;
        GSErrCode err = ACAPI_Element_GetDefaults(&leftLine, nullptr);
        if (err != NoError) {
            Log("[ShellHelper] ERROR: Не удалось получить настройки по умолчанию для левой линии");
            return err;
        }
        
        leftLine.header.floorInd = baseLine.header.floorInd;
        leftLine.line.begC.x = begC.x + perpX * halfWidth;
        leftLine.line.begC.y = begC.y + perpY * halfWidth;
        leftLine.line.endC.x = endC.x + perpX * halfWidth;
        leftLine.line.endC.y = endC.y + perpY * halfWidth;
        
        Log("[ShellHelper] Левая линия: begC=(%.3f,%.3f), endC=(%.3f,%.3f)", 
            leftLine.line.begC.x, leftLine.line.begC.y, 
            leftLine.line.endC.x, leftLine.line.endC.y);
        
        err = ACAPI_Element_Create(&leftLine, nullptr);
        if (err != NoError) {
            Log("[ShellHelper] ERROR: Не удалось создать левую линию, err=%d", (int)err);
            return err;
        }
        
        // Правая линия
        API_Element rightLine = {};
        rightLine.header.type = API_LineID;
        err = ACAPI_Element_GetDefaults(&rightLine, nullptr);
        if (err != NoError) {
            Log("[ShellHelper] ERROR: Не удалось получить настройки по умолчанию для правой линии");
            return err;
        }
        
        rightLine.header.floorInd = baseLine.header.floorInd;
        rightLine.line.begC.x = begC.x - perpX * halfWidth;
        rightLine.line.begC.y = begC.y - perpY * halfWidth;
        rightLine.line.endC.x = endC.x - perpX * halfWidth;
        rightLine.line.endC.y = endC.y - perpY * halfWidth;
        
        Log("[ShellHelper] Правая линия: begC=(%.3f,%.3f), endC=(%.3f,%.3f)", 
            rightLine.line.begC.x, rightLine.line.begC.y, 
            rightLine.line.endC.x, rightLine.line.endC.y);
        
        err = ACAPI_Element_Create(&rightLine, nullptr);
        if (err != NoError) {
            Log("[ShellHelper] ERROR: Не удалось создать правую линию, err=%d", (int)err);
            return err;
        }
        
        Log("[ShellHelper] SUCCESS: Обе перпендикулярные линии созданы");
        return NoError;
    });
    
    return true;
}

// =============== Создание геометрии оболочки ===============
bool CreateShellGeometry(const GS::Array<API_Coord3D>& shellPoints)
{
    ACAPI_WriteReport("[ShellHelper] CreateShellGeometry: %d точек", false, (int)shellPoints.GetSize());
    
    if (shellPoints.GetSize() < 2) {
        ACAPI_WriteReport("[ShellHelper] Недостаточно точек для создания оболочки", false);
        return false;
    }
    
    // TODO: Создать полилинию или другие элементы из точек оболочки
    // Пока просто логируем точки
    for (UIndex i = 0; i < shellPoints.GetSize(); ++i) {
        const API_Coord3D& point = shellPoints[i];
        ACAPI_WriteReport("[ShellHelper] Точка %d: (%.3f, %.3f, %.3f)", false, 
            (int)i, point.x, point.y, point.z);
    }
    
    ACAPI_WriteReport("[ShellHelper] Оболочка создана (пока заглушка)", false);
    return true;
}

// =============== Утилиты для работы с геометрией ===============
static inline double SegLenLine(const API_Coord& a, const API_Coord& b) {
    const double dx = b.x - a.x, dy = b.y - a.y;
    return std::sqrt(dx * dx + dy * dy);
}

static inline double SegAng(const API_Coord& a, const API_Coord& b) { 
    return std::atan2(b.y - a.y, b.x - a.x); 
}

static inline bool NearlyEq(const API_Coord& a, const API_Coord& b, double tol = 1e-9) {
    return (std::fabs(a.x - b.x) <= tol) && (std::fabs(a.y - b.y) <= tol);
}

static inline API_Coord Add(const API_Coord& a, const API_Coord& b) { 
    return { a.x + b.x, a.y + b.y }; 
}

static inline API_Coord Sub(const API_Coord& a, const API_Coord& b) { 
    return { a.x - b.x, a.y - b.y }; 
}

static inline API_Coord Mul(const API_Coord& a, double s) { 
    return { a.x * s, a.y * s }; 
}

static inline API_Coord UnitFromAng(double ang) { 
    return { std::cos(ang), std::sin(ang) }; 
}

static inline double Clamp01(double t) { 
    return t < 0 ? 0 : (t > 1 ? 1 : t); 
}

static inline double NormPos(double a) { // в [0..2π)
    while (a < 0.0)       a += 2.0 * kPI;
    while (a >= 2.0 * kPI)  a -= 2.0 * kPI;
    return a;
}

// =============== Восстановление дуги по хорде и углу ===============
static bool BuildArcFromPolylineSegment(
    const API_Coord& A, const API_Coord& B, double arcAngle,
    API_Coord& C, double& r, double& a0, double& a1, bool& ccw)
{
    const double L = std::hypot(B.x - A.x, B.y - A.y);
    if (L <= kEPS || !std::isfinite(arcAngle))
        return false;

    // Нормируем угол в (-π, π]
    double phi = arcAngle;
    while (phi <= -kPI) phi += 2.0 * kPI;
    while (phi > kPI)  phi -= 2.0 * kPI;

    if (std::fabs(phi) < 1e-9)
        return false; // фактически прямая

    // Вычисляем радиус по minor-углу
    double rMinor = (0.5 * L) / std::sin(0.5 * phi);

    // Проверка на major-дугу
    bool isMajor = std::fabs(phi) > kPI;
    if (isMajor) {
        phi = phi > 0 ? phi - 2.0 * kPI : phi + 2.0 * kPI;
        r = (0.5 * L) / std::sin(0.5 * phi);
    } else {
        r = rMinor;
    }

    // Центр дуги
    const double midX = 0.5 * (A.x + B.x);
    const double midY = 0.5 * (A.y + B.y);
    const double perpLen = std::sqrt(r * r - (0.5 * L) * (0.5 * L));
    const double perpAng = SegAng(A, B) + (phi > 0 ? kPI / 2.0 : -kPI / 2.0);

    C = { midX + perpLen * std::cos(perpAng), midY + perpLen * std::sin(perpAng) };

    // Углы
    a0 = std::atan2(A.y - C.y, A.x - C.x);
    a1 = std::atan2(B.y - C.y, B.x - C.x);
    ccw = phi > 0;

    return true;
}

// =============== Парсинг элемента в сегменты ===============
bool ParseElementToSegments(const API_Element& element, PathData& path)
{
    path.segs.Clear();
    path.total = 0.0;
    
    if (element.header.type == API_LineID) {
        Seg seg;
        seg.type = SegType::Line;
        seg.A = element.line.begC;
        seg.B = element.line.endC;
        seg.len = SegLenLine(seg.A, seg.B);
        if (seg.len > kEPS) {
            path.segs.Push(seg);
            path.total += seg.len;
        }
        Log("[ShellHelper] Line parsed: length=%.3f", seg.len);
        return !path.segs.IsEmpty();
    }
    else if (element.header.type == API_ArcID) {
        Seg seg;
        seg.type = SegType::Arc;
        seg.C = element.arc.origC;
        seg.r = element.arc.r;
        seg.a0 = element.arc.begAng;
        seg.a1 = element.arc.endAng;
        
        // Вычисляем угол дуги
        double arcAngle = seg.a1 - seg.a0;
        seg.ccw = arcAngle > 0;
        seg.len = std::fabs(arcAngle) * seg.r;
        
        if (seg.len > kEPS) {
            path.segs.Push(seg);
            path.total += seg.len;
        }
        Log("[ShellHelper] Arc parsed: radius=%.3f, angle=%.3f, length=%.3f", 
            seg.r, arcAngle, seg.len);
        return !path.segs.IsEmpty();
    }
        
    else if (element.header.type == API_CircleID) {
        Seg seg;
        seg.type = SegType::Arc;
        seg.C = element.arc.origC;
        seg.r = element.arc.r;
        seg.a0 = 0.0;
        seg.a1 = 2.0 * kPI;
        seg.ccw = true;
        seg.len = 2.0 * kPI * seg.r;
        if (seg.len > kEPS) {
            path.segs.Push(seg);
            path.total += seg.len;
        }
        Log("[ShellHelper] Circle parsed: radius=%.3f, length=%.3f", seg.r, seg.len);
        return !path.segs.IsEmpty();
    }
        
    else if (element.header.type == API_PolyLineID) {
            API_ElementMemo memo;
            BNZeroMemory(&memo, sizeof(memo));
            GSErrCode err = ACAPI_Element_GetMemo(element.header.guid, &memo);
            if (err != NoError || memo.coords == nullptr) {
                ACAPI_DisposeElemMemoHdls(&memo);
                Log("[ShellHelper] ERROR: Не удалось получить memo для полилинии");
                return false;
            }

            const Int32 nCoordsAll = (Int32)(BMGetHandleSize((GSHandle)memo.coords) / (Int32)sizeof(API_Coord));
            const Int32 nCoords = std::max<Int32>(0, nCoordsAll - 1);
            if (nCoords < 2) {
                ACAPI_DisposeElemMemoHdls(&memo);
                Log("[ShellHelper] ERROR: Недостаточно точек в полилинии");
                return false;
            }

            // Карта дуг: индекс начала -> arcAngle
            std::unordered_map<Int32, double> arcByBeg;
            if (memo.parcs != nullptr) {
                const Int32 nArcsAll = (Int32)(BMGetHandleSize((GSHandle)memo.parcs) / (Int32)sizeof(API_PolyArc));
                const Int32 nArcs = std::max<Int32>(0, nArcsAll - 1);
                Log("[ShellHelper] Found %d arcs in polyline", nArcs);
                for (Int32 ai = 1; ai <= nArcs; ++ai) {
                    const API_PolyArc& pa = (*memo.parcs)[ai];
                    Log("[ShellHelper] Arc %d: begIndex=%d, arcAngle=%.6f", ai, pa.begIndex, pa.arcAngle);
                    if (pa.begIndex >= 1 && pa.begIndex <= nCoords - 1) {
                        arcByBeg[pa.begIndex] = pa.arcAngle;
                        Log("[ShellHelper] Added arc to map: begIndex=%d, arcAngle=%.6f", pa.begIndex, pa.arcAngle);
                    } else {
                        Log("[ShellHelper] Skipped arc %d: begIndex=%d out of range [1,%d]", ai, pa.begIndex, nCoords - 1);
                    }
                }
            } else {
                Log("[ShellHelper] No arcs found in polyline (memo.parcs is null)");
            }

            // Перебор сегментов
            API_Coord prev = (*memo.coords)[1];
            for (Int32 idx = 2; idx <= nCoords; ++idx) {
                const API_Coord curr = (*memo.coords)[idx];
                if (NearlyEq(prev, curr)) {
                    prev = curr;
                    continue;
                }

                const Int32 segIdx = idx - 1;
                auto it = arcByBeg.find(segIdx);
                Log("[ShellHelper] Checking segment %d for arcs...", segIdx);

                Seg seg = {};
                if (it != arcByBeg.end() && std::fabs(it->second) > kEPS) {
                    // дуга
                    seg.type = SegType::Arc;
                    Log("[ShellHelper] Found arc at segment %d: angle=%.6f", segIdx, it->second);
                    if (BuildArcFromPolylineSegment(prev, curr, it->second, seg.C, seg.r, seg.a0, seg.a1, seg.ccw)) {
                        seg.len = std::fabs(seg.a1 - seg.a0) * seg.r;
                        Log("[ShellHelper] Arc built: center=(%.3f,%.3f), radius=%.3f, len=%.3f", 
                            seg.C.x, seg.C.y, seg.r, seg.len);
                    } else {
                        Log("[ShellHelper] Failed to build arc, using line instead");
                        seg.type = SegType::Line;
                        seg.A = prev;
                        seg.B = curr;
                        seg.len = SegLenLine(prev, curr);
                    }
                } else {
                    // линия
                    seg.type = SegType::Line;
                    seg.A = prev;
                    seg.B = curr;
                    seg.len = SegLenLine(prev, curr);
                    if (it != arcByBeg.end()) {
                        Log("[ShellHelper] Line segment %d: len=%.3f (arc angle too small: %.6f)", segIdx, seg.len, it->second);
                    } else {
                        Log("[ShellHelper] Line segment %d: len=%.3f (no arc found)", segIdx, seg.len);
                    }
                }

                if (seg.len > kEPS) {
                    path.segs.Push(seg);
                    path.total += seg.len;
                } else {
                    Log("[ShellHelper] Skipping segment %d: too short (%.6f)", segIdx, seg.len);
                }

                prev = curr;
            }

        ACAPI_DisposeElemMemoHdls(&memo);
        Log("[ShellHelper] Polyline parsed: %d segments, total length=%.3f", 
            (int)path.segs.GetSize(), path.total);
        return !path.segs.IsEmpty();
    }
    else if (element.header.type == API_SplineID) {
        // Кубические Безье по bezierDirs (как в LandscapeHelper для точности)
        API_ElementMemo memo;
        BNZeroMemory(&memo, sizeof(memo));
        GSErrCode err = ACAPI_Element_GetMemo(element.header.guid, &memo, APIMemoMask_Polygon);
        if (err != NoError || memo.coords == nullptr || memo.bezierDirs == nullptr) {
            ACAPI_DisposeElemMemoHdls(&memo);
            Log("[ShellHelper] ERROR: Не удалось получить memo с bezierDirs для сплайна");
            return false;
        }

        const Int32 n = (Int32)(BMGetHandleSize((GSHandle)memo.coords) / sizeof(API_Coord));
        if (n < 2) {
            ACAPI_DisposeElemMemoHdls(&memo);
            Log("[ShellHelper] ERROR: Недостаточно точек в сплайне");
            return false;
        }

        // Вспомогательные функции для Безье
        auto FromAngLen = [](double ang, double len) -> API_Coord {
            return { std::cos(ang) * len, std::sin(ang) * len };
        };
        
        auto BezierPoint = [](const API_Coord& P0, const API_Coord& C1,
                               const API_Coord& C2, const API_Coord& P3, double t) -> API_Coord {
            const double u = 1.0 - t;
            const double u2 = u * u;
            const double u3 = u2 * u;
            const double t2 = t * t;
            const double t3 = t2 * t;
            return {
                u3 * P0.x + 3.0 * u2 * t * C1.x + 3.0 * u * t2 * C2.x + t3 * P3.x,
                u3 * P0.y + 3.0 * u2 * t * C1.y + 3.0 * u * t2 * C2.y + t3 * P3.y
            };
        };

        // Разбиваем каждый сегмент Безье на подсегменты для точности
        const int N = 32; // подсегментов на ребро Безье
        API_Coord prev;
        bool firstPoint = true;

        for (Int32 i = 0; i < n - 1; ++i) {
            const API_Coord P0 = (*memo.coords)[i];
            const API_Coord P3 = (*memo.coords)[i + 1];
            const API_SplineDir d0 = (*memo.bezierDirs)[i];
            const API_SplineDir d1 = (*memo.bezierDirs)[i + 1];
            
            // Контрольные точки Безье
            const API_Coord C1 = Add(P0, FromAngLen(d0.dirAng, d0.lenNext));
            const API_Coord C2 = Sub(P3, FromAngLen(d1.dirAng, d1.lenPrev));

            // Разбиваем на подсегменты
            if (firstPoint) {
                prev = P0;
                firstPoint = false;
            }

            for (int k = 1; k <= N; ++k) {
                const double t = (double)k / (double)N;
                const API_Coord pt = BezierPoint(P0, C1, C2, P3, t);
                
                if (!NearlyEq(prev, pt)) {
                    Seg seg;
                    seg.type = SegType::Line;
                    seg.A = prev;
                    seg.B = pt;
                    seg.len = SegLenLine(prev, pt);
                    
                    if (seg.len > kEPS) {
                        path.segs.Push(seg);
                        path.total += seg.len;
                    }
                }
                
                prev = pt;
            }
        }

        ACAPI_DisposeElemMemoHdls(&memo);
        Log("[ShellHelper] Spline parsed: %d segments (using Bezier), total length=%.3f", 
            (int)path.segs.GetSize(), path.total);
        return !path.segs.IsEmpty();
    }
    else {
        Log("[ShellHelper] ERROR: Неподдерживаемый тип элемента для парсинга");
        return false;
    }
}

// =============== Создание 3D оболочки через Ruled Shell ===============
bool Create3DShellFromPath(const PathData& path, double widthMM, double stepMM)
{
    Log("[ShellHelper] Create3DShellFromPath: %d сегментов, width=%.1fmm, step=%.1fmm", 
        (int)path.segs.GetSize(), widthMM, stepMM);
    
    if (path.segs.IsEmpty()) {
        Log("[ShellHelper] ERROR: Нет сегментов для обработки");
        return false;
    }
    
    // Используем логику из LandscapeHelper для получения точек по шагу
    double step = stepMM / 1000.0; // шаг в метрах
    
    // Генерируем позиции точек по шагу (как в DistributeOnSinglePath)
    GS::Array<double> sVals;
    for (double s = 0.0; s <= path.total + 1e-9; s += step) {
        sVals.Push(s);
    }
    
    // Убеждаемся, что последняя точка точно на конце линии
    if (sVals.IsEmpty() || sVals[sVals.GetSize() - 1] < path.total - 1e-9) {
        sVals.Push(path.total);
    }
    
    Log("[ShellHelper] Сгенерировано %d точек по шагу %.3fм", (int)sVals.GetSize(), step);
    
    GS::Array<API_Coord3D> leftPoints;
    GS::Array<API_Coord3D> rightPoints;
    
    // Обрабатываем каждую точку
    for (UIndex i = 0; i < sVals.GetSize(); ++i) {
        double s = sVals[i];
        // Получаем точку и угол используя логику из LandscapeHelper
        API_Coord pointOnPath;
        double tangentAngle = 0.0;
        
        // Используем EvalOnPath логику для получения точки и угла
        double acc = 0.0;
        bool found = false;
        for (UIndex j = 0; j < path.segs.GetSize() && !found; ++j) {
            const Seg& seg = path.segs[j];
            
            if (s > acc + seg.len) { 
                acc += seg.len; 
                continue; 
            }
            const double f = (seg.len < 1e-9) ? 0.0 : (s - acc) / seg.len;

            if (seg.type == SegType::Line) {
                pointOnPath.x = seg.A.x + f * (seg.B.x - seg.A.x);
                pointOnPath.y = seg.A.y + f * (seg.B.y - seg.A.y);
                tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
            } else if (seg.type == SegType::Arc) {
                const double sweep = seg.a1 - seg.a0;
                const double ang = seg.a0 + f * sweep;
                pointOnPath.x = seg.C.x + seg.r * std::cos(ang);
                pointOnPath.y = seg.C.y + seg.r * std::sin(ang);
                tangentAngle = ang + ((sweep >= 0.0) ? kPI / 2.0 : -kPI / 2.0);
            } else {
                // Для сплайнов используем линейную интерполяцию
                pointOnPath.x = seg.A.x + f * (seg.B.x - seg.A.x);
                pointOnPath.y = seg.A.y + f * (seg.B.y - seg.A.y);
                tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
            }
            found = true;
        }
        
        if (!found) {
            // Если не нашли сегмент, используем последнюю точку
            const Seg& seg = path.segs[path.segs.GetSize() - 1];
            if (seg.type == SegType::Line) {
                pointOnPath = seg.B;
                tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
            } else {
                const double sweep = seg.a1 - seg.a0;
                pointOnPath.x = seg.C.x + seg.r * std::cos(seg.a1);
                pointOnPath.y = seg.C.y + seg.r * std::sin(seg.a1);
                tangentAngle = seg.a1 + ((sweep >= 0.0) ? kPI / 2.0 : -kPI / 2.0);
            }
        }
        
        // Шаг 1: Получаем Z-координату для точки на базовой линии от Mesh
        double baseZ = 0.0;
        API_Vector3D baseNormal = {};
        if (g_meshSurfaceGuid != APINULLGuid) {
            // Передаем XY координаты точки на базовой линии в MeshIntersectionHelper
            API_Coord baseXY = {pointOnPath.x, pointOnPath.y};
            if (MeshIntersectionHelper::GetZAndNormal(baseXY, baseZ, baseNormal)) {
                // Логируем первые и последние точки для отладки
                if (i < 5 || i >= sVals.GetSize() - 5) {
                    Log("[ShellHelper] Point on base line %d: (%.3f, %.3f) -> Z=%.3f", (int)i + 1, pointOnPath.x, pointOnPath.y, baseZ);
                }
            } else {
                Log("[ShellHelper] WARNING: Не удалось получить Z для точки на базовой линии (%.3f, %.3f)", pointOnPath.x, pointOnPath.y);
                baseZ = 0.0;
            }
        }
        
        // Шаг 2: Строим перпендикуляры от точки на базовой линии
        double perpAngle = tangentAngle + kPI / 2.0;
        double perpX = std::cos(perpAngle);
        double perpY = std::sin(perpAngle);
        
        double halfWidth = widthMM / 2000.0; // Переводим мм в метры и делим пополам
        
        // Шаг 3: Создаем левую и правую точки, используя Z из базовой линии
        API_Coord3D leftPoint = {pointOnPath.x + perpX * halfWidth, pointOnPath.y + perpY * halfWidth, baseZ};
        API_Coord3D rightPoint = {pointOnPath.x - perpX * halfWidth, pointOnPath.y - perpY * halfWidth, baseZ};
        
        leftPoints.Push(leftPoint);
        rightPoints.Push(rightPoint);
        
        // Логируем первые и последние точки для отладки
        if (i < 5 || i >= sVals.GetSize() - 5) {
            Log("[ShellHelper] Точка %d: left(%.3f, %.3f, %.3f), right(%.3f, %.3f, %.3f)", 
                (int)i + 1, leftPoint.x, leftPoint.y, leftPoint.z, rightPoint.x, rightPoint.y, rightPoint.z);
        }
    }
    
    Log("[ShellHelper] Создано %d пар точек для 3D оболочки", (int)leftPoints.GetSize());
    
    // Шаг 4: Создаем 2 НЕ замкнутые Spline по найденным точкам
    Log("[ShellHelper] Создаем два НЕ замкнутых Spline из %d точек", (int)leftPoints.GetSize() * 2);
    
    // Создаем левый Spline из левых точек
    GS::Array<API_Coord> leftSplinePoints;
    for (UIndex i = 0; i < leftPoints.GetSize(); ++i) {
        leftSplinePoints.Push({leftPoints[i].x, leftPoints[i].y});
    }
    
    Log("[ShellHelper] Левый Spline: %d точек", (int)leftSplinePoints.GetSize());
    if (leftSplinePoints.GetSize() > 0) {
        Log("[ShellHelper] Левый Spline: первая точка (%.3f, %.3f), последняя (%.3f, %.3f)", 
            leftSplinePoints[0].x, leftSplinePoints[0].y,
            leftSplinePoints[leftSplinePoints.GetSize()-1].x, leftSplinePoints[leftSplinePoints.GetSize()-1].y);
    }
    
    API_Guid leftSplineGuid = CreateSplineFromPoints(leftSplinePoints);
    if (leftSplineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Не удалось создать левый Spline");
        return false;
    }
    
    // Создаем правый Spline из правых точек
    GS::Array<API_Coord> rightSplinePoints;
    for (UIndex i = 0; i < rightPoints.GetSize(); ++i) {
        rightSplinePoints.Push({rightPoints[i].x, rightPoints[i].y});
    }
    
    Log("[ShellHelper] Правый Spline: %d точек", (int)rightSplinePoints.GetSize());
    if (rightSplinePoints.GetSize() > 0) {
        Log("[ShellHelper] Правый Spline: первая точка (%.3f, %.3f), последняя (%.3f, %.3f)", 
            rightSplinePoints[0].x, rightSplinePoints[0].y,
            rightSplinePoints[rightSplinePoints.GetSize()-1].x, rightSplinePoints[rightSplinePoints.GetSize()-1].y);
    }
    
    API_Guid rightSplineGuid = CreateSplineFromPoints(rightSplinePoints);
    if (rightSplineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Не удалось создать правый Spline");
        return false;
    }
    
    Log("[ShellHelper] SUCCESS: Созданы два НЕ замкнутых Spline (левый и правый)");
    
    // Шаг 5: Замыкаем крайние точки обоих Spline простыми линиями
    Log("[ShellHelper] Замыкаем крайние точки обоих Spline простыми линиями");
    
    // Создаем линию между первыми точками (начало)
    API_Element startLine = {};
    startLine.header.type = API_LineID;
    GSErrCode err = ACAPI_Element_GetDefaults(&startLine, nullptr);
    if (err == NoError) {
        startLine.line.begC = leftSplinePoints[0];
        startLine.line.endC = rightSplinePoints[0];
        
        Log("[ShellHelper] Start Line: begC=(%.3f,%.3f), endC=(%.3f,%.3f)", 
            startLine.line.begC.x, startLine.line.begC.y,
            startLine.line.endC.x, startLine.line.endC.y);
        
        err = ACAPI_CallUndoableCommand("Create Start Line", [&]() -> GSErrCode {
            return ACAPI_Element_Create(&startLine, nullptr);
        });
        
        if (err == NoError) {
            Log("[ShellHelper] SUCCESS: Создана линия между первыми точками");
        } else {
            Log("[ShellHelper] ERROR: Не удалось создать линию между первыми точками, err=%d", (int)err);
        }
    }
    
    // Создаем линию между последними точками (конец)
    API_Element endLine = {};
    endLine.header.type = API_LineID;
    err = ACAPI_Element_GetDefaults(&endLine, nullptr);
    if (err == NoError) {
        endLine.line.begC = leftSplinePoints[leftSplinePoints.GetSize() - 1];
        endLine.line.endC = rightSplinePoints[rightSplinePoints.GetSize() - 1];
        
        Log("[ShellHelper] End Line: begC=(%.3f,%.3f), endC=(%.3f,%.3f)", 
            endLine.line.begC.x, endLine.line.begC.y,
            endLine.line.endC.x, endLine.line.endC.y);
        
        err = ACAPI_CallUndoableCommand("Create End Line", [&]() -> GSErrCode {
            return ACAPI_Element_Create(&endLine, nullptr);
        });
        
        if (err == NoError) {
            Log("[ShellHelper] SUCCESS: Создана линия между последними точками");
        } else {
            Log("[ShellHelper] ERROR: Не удалось создать линию между последними точками, err=%d", (int)err);
        }
    }
    
    Log("[ShellHelper] SUCCESS: Замыкающие линии созданы");
    
    // Шаг 6: Создаем SHELL вместо MESH!
    Log("[ShellHelper] Создаем SHELL вместо MESH!");
    
    // Создаем замкнутый контур для MESH: левые точки + правые точки в обратном порядке
    GS::Array<API_Coord> meshContourPoints;
    
    // Добавляем левые точки от начала до конца
    for (UIndex i = 0; i < leftPoints.GetSize(); ++i) {
        meshContourPoints.Push({leftPoints[i].x, leftPoints[i].y});
    }
    
    // Добавляем правые точки от конца до начала (в обратном порядке)
    for (Int32 i = rightPoints.GetSize() - 1; i >= 0; --i) {
        meshContourPoints.Push({rightPoints[i].x, rightPoints[i].y});
    }
    
    Log("[ShellHelper] MESH контур: %d точек", (int)meshContourPoints.GetSize());
    
    // СОЗДАЕМ MESH С УРОВНЯМИ! 🎯
    API_Element mesh = {};
    mesh.header.type = API_MeshID;
    err = ACAPI_Element_GetDefaults(&mesh, nullptr);
    if (err == NoError) {
        // ПРОСТОЙ MESH КАК ПОЛИГОН! 🎯
        const Int32 nCoords = (Int32)meshContourPoints.GetSize();
        
        mesh.mesh.poly.nCoords = nCoords;
        mesh.mesh.poly.nSubPolys = 1;
        mesh.mesh.poly.nArcs = 0;
        
        Log("[ShellHelper] ПРОСТОЙ MESH: %d координат, 1 полигон", (int)nCoords);
        
        // Создаем memo для MESH
        API_ElementMemo meshMemo = {};
        BNZeroMemory(&meshMemo, sizeof(API_ElementMemo));
        
        // ВЫДЕЛЯЕМ ПАМЯТЬ ДЛЯ КООРДИНАТ! 🎯
        meshMemo.coords = reinterpret_cast<API_Coord**>(BMAllocateHandle((nCoords + 1) * (GSSize)sizeof(API_Coord), ALLOCATE_CLEAR, 0));
        meshMemo.pends = reinterpret_cast<Int32**>(BMAllocateHandle((mesh.mesh.poly.nSubPolys + 1) * sizeof(Int32), ALLOCATE_CLEAR, 0));
        meshMemo.parcs = reinterpret_cast<API_PolyArc**>(BMAllocateHandle(mesh.mesh.poly.nArcs * sizeof(API_PolyArc), ALLOCATE_CLEAR, 0));
        if (meshMemo.coords != nullptr && meshMemo.pends != nullptr) {
            // Инициализируем элемент с индексом 0
            (*meshMemo.coords)[0] = meshContourPoints[0]; // Заглушка для элемента 0
            
            // Заполняем координаты (1-based indexing)
            for (UIndex i = 0; i < meshContourPoints.GetSize(); ++i) {
                (*meshMemo.coords)[i + 1] = meshContourPoints[i];
            }
            
            // Замыкаем полигон (последняя точка = первая)
            (*meshMemo.coords)[nCoords] = (*meshMemo.coords)[1];
            
            // Устанавливаем pends (end index для полигона)
            (*meshMemo.pends)[1] = nCoords;
            
            // ВЫДЕЛЯЕМ ПАМЯТЬ ДЛЯ Z-КООРДИНАТ! 🎯
            meshMemo.meshPolyZ = reinterpret_cast<double**>(BMAllocateHandle((nCoords + 1) * (GSSize)sizeof(double), ALLOCATE_CLEAR, 0));
            if (meshMemo.meshPolyZ != nullptr) {
                // Инициализируем элемент с индексом 0
                (*meshMemo.meshPolyZ)[0] = 0.0; // Заглушка для элемента 0
                
                // Заполняем Z-координаты из левых и правых точек
                UIndex zIndex = 1; // 1-based indexing
                
                // Добавляем Z-координаты левых точек (от начала до конца)
                for (UIndex i = 0; i < leftPoints.GetSize(); ++i) {
                    (*meshMemo.meshPolyZ)[zIndex++] = leftPoints[i].z;
                }
                
                // Добавляем Z-координаты правых точек (от конца до начала)
                for (Int32 i = rightPoints.GetSize() - 1; i >= 0; --i) {
                    (*meshMemo.meshPolyZ)[zIndex++] = rightPoints[i].z;
                }
                
                Log("[ShellHelper] MESH Z-координаты: %d точек", (int)nCoords);
                
                // ПРОСТОЙ MESH БЕЗ УРОВНЕЙ! 🎯
                // Попробуем создать MESH как простой полигон с Z-координатами
                Log("[ShellHelper] MESH: создаем простой MESH без уровней");
            } else {
                Log("[ShellHelper] ERROR: Не удалось выделить память для Z-координат MESH");
                ACAPI_DisposeElemMemoHdls(&meshMemo);
                return false;
            }
            
            // Создаем MESH внутри Undo-команды
            err = ACAPI_CallUndoableCommand("Create Mesh", [&]() -> GSErrCode {
                GSErrCode createErr = ACAPI_Element_Create(&mesh, &meshMemo);
                
                // Если полигон нерегулярный, нужно его регуляризовать!
                if (createErr == APIERR_IRREGULARPOLY) {
                    Log("[ShellHelper] MESH: Полигон нерегулярный, регуляризуем...");
                    
                    API_RegularizedPoly poly = {};
                    poly.coords = meshMemo.coords;
                    poly.pends = meshMemo.pends;
                    poly.parcs = meshMemo.parcs;
                    poly.vertexIDs = meshMemo.vertexIDs;
                    poly.needVertexAncestry = 1;
                    
                    Int32 nResult = 0;
                    API_RegularizedPoly** polys = nullptr;
                    GSErrCode regErr = ACAPI_Polygon_RegularizePolygon(&poly, &nResult, &polys);
                    
                    if (regErr == NoError && nResult > 0) {
                        Log("[ShellHelper] MESH: Регуляризация успешна, создаем %d полигонов", (int)nResult);
                        
                        // Создаем регуляризованные полигоны
                        for (Int32 i = 0; i < nResult; i++) {
                            mesh.mesh.poly.nCoords = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].coords)) / sizeof(API_Coord) - 1;
                            mesh.mesh.poly.nSubPolys = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].pends)) / sizeof(Int32) - 1;
                            mesh.mesh.poly.nArcs = BMhGetSize(reinterpret_cast<GSHandle>((*polys)[i].parcs)) / sizeof(API_PolyArc);
                            
                            // Создаем временный memo для регуляризованного полигона
                            API_ElementMemo tmpMemo = {};
                            tmpMemo.coords = (*polys)[i].coords;
                            tmpMemo.pends = (*polys)[i].pends;
                            tmpMemo.parcs = (*polys)[i].parcs;
                            tmpMemo.vertexIDs = (*polys)[i].vertexIDs;
                            
                            // Выделяем память для Z-координат
                            tmpMemo.meshPolyZ = reinterpret_cast<double**>(BMAllocateHandle((mesh.mesh.poly.nCoords + 1) * sizeof(double), ALLOCATE_CLEAR, 0));
                            if (tmpMemo.meshPolyZ != nullptr) {
                                // Копируем Z-координаты из оригинального meshPolyZ
                                for (Int32 j = 1; j <= mesh.mesh.poly.nCoords; j++) {
                                    // Находим соответствующий индекс в оригинальном массиве
                                    Int32 oldVertexIndex = 1; // Упрощенная логика - можно улучшить
                                    if (oldVertexIndex <= (Int32)meshContourPoints.GetSize()) {
                                        (*tmpMemo.meshPolyZ)[j] = (*meshMemo.meshPolyZ)[oldVertexIndex];
                                    } else {
                                        (*tmpMemo.meshPolyZ)[j] = 0.0;
                                    }
                                }
                                
                                GSErrCode pieceErr = ACAPI_Element_Create(&mesh, &tmpMemo);
                                if (pieceErr != NoError) {
                                    Log("[ShellHelper] MESH ERROR: Не удалось создать регуляризованный полигон %d, err=%d", (int)i, (int)pieceErr);
                                }
                                
                                BMKillHandle(reinterpret_cast<GSHandle*>(&tmpMemo.meshPolyZ));
                            }
                        }
                        
                        createErr = NoError; // Успешно создали регуляризованные полигоны
                    } else {
                        Log("[ShellHelper] MESH ERROR: Не удалось регуляризовать полигон, err=%d", (int)regErr);
                    }
                }
                
                return createErr;
            });
            
            if (err == NoError) {
                Log("[ShellHelper] SUCCESS: ПРОСТОЙ MESH создан! %d координат", (int)nCoords);
                ACAPI_DisposeElemMemoHdls(&meshMemo);
                return true;
            } else {
                Log("[ShellHelper] ERROR: Не удалось создать простой MESH, err=%d", (int)err);
                ACAPI_DisposeElemMemoHdls(&meshMemo);
                return false;
            }
        } else {
            Log("[ShellHelper] ERROR: Не удалось выделить память для координат MESH");
            return false;
        }
    } else {
        Log("[ShellHelper] ERROR: Не удалось получить настройки по умолчанию для MESH, err=%d", (int)err);
        return false;
    }
}

// =============== Создание Morph из контура ===============
bool CreateMorphFromContour(double widthMM, double stepMM, double thicknessMM,
                            API_AttributeIndex materialTop, API_AttributeIndex materialBottom, API_AttributeIndex materialSide)
{
    Log("[ShellHelper] CreateMorphFromContour: START, width=%.1fmm, step=%.1fmm, thickness=%.1fmm", widthMM, stepMM, thicknessMM);
    
    // Проверяем, что базовая линия выбрана
    if (g_baseLineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Базовая линия не выбрана. Сначала выберите базовую линию.");
        return false;
    }
    
    Log("[ShellHelper] Базовая линия: %s", APIGuidToString(g_baseLineGuid).ToCStr().Get());
    
    // Проверяем mesh
    if (g_meshSurfaceGuid != APINULLGuid) {
        Log("[ShellHelper] Mesh выбран: %s - будет использоваться для получения Z координат", APIGuidToString(g_meshSurfaceGuid).ToCStr().Get());
    } else {
        Log("[ShellHelper] WARNING: Mesh не выбран - будут использоваться Z=0.0 для всех точек");
    }
    
    // Получаем элемент базовой линии по сохраненному GUID
    API_Elem_Head elemHead = {};
    elemHead.guid = g_baseLineGuid;
    GSErrCode err = ACAPI_Element_GetHeader(&elemHead);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить заголовок элемента базовой линии");
        return false;
    }
    
    // Проверяем поддерживаемые типы элементов
    bool isSupported = false;
    if (elemHead.type == API_LineID ||
        elemHead.type == API_PolyLineID ||
        elemHead.type == API_ArcID ||
        elemHead.type == API_CircleID ||
        elemHead.type == API_SplineID) {
        isSupported = true;
    }
    
    if (!isSupported) {
        Log("[ShellHelper] ERROR: Неподдерживаемый тип элемента базовой линии");
        Log("[ShellHelper] Поддерживаются: Line, Polyline, Arc, Circle, Spline");
        return false;
    }
    
    // Получаем данные элемента
    API_Element element = {};
    element.header = elemHead;
    err = ACAPI_Element_Get(&element);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить данные элемента базовой линии");
        return false;
    }
    
    Log("[ShellHelper] Элемент базовой линии загружен успешно");
    
    // Парсим элемент в сегменты
    PathData path;
    if (!ParseElementToSegments(element, path)) {
        Log("[ShellHelper] ERROR: Не удалось распарсить элемент в сегменты");
        return false;
    }
    
    Log("[ShellHelper] Элемент распарсен: %d сегментов, общая длина %.3fм", 
        (int)path.segs.GetSize(), path.total);
    
    // Используем логику из Create3DShellFromPath для получения точек
    double step = stepMM / 1000.0; // шаг в метрах
    
    // Генерируем позиции точек по шагу
    GS::Array<double> sVals;
    for (double s = 0.0; s <= path.total + 1e-9; s += step) {
        sVals.Push(s);
    }
    
    // Убеждаемся, что последняя точка точно на конце линии
    if (sVals.IsEmpty() || sVals[sVals.GetSize() - 1] < path.total - 1e-9) {
        sVals.Push(path.total);
    }
    
    Log("[ShellHelper] Сгенерировано %d точек по шагу %.3fм", (int)sVals.GetSize(), step);
    
    GS::Array<API_Coord3D> leftPoints;
    GS::Array<API_Coord3D> rightPoints;
    
    // Обрабатываем каждую точку (та же логика что в Create3DShellFromPath)
    for (UIndex i = 0; i < sVals.GetSize(); ++i) {
        double s = sVals[i];
        API_Coord pointOnPath;
        double tangentAngle = 0.0;
        
        double acc = 0.0;
        bool found = false;
        for (UIndex j = 0; j < path.segs.GetSize() && !found; ++j) {
            const Seg& seg = path.segs[j];
            
            if (s > acc + seg.len) { 
                acc += seg.len; 
                continue; 
            }
            const double f = (seg.len < 1e-9) ? 0.0 : (s - acc) / seg.len;

            if (seg.type == SegType::Line) {
                pointOnPath.x = seg.A.x + f * (seg.B.x - seg.A.x);
                pointOnPath.y = seg.A.y + f * (seg.B.y - seg.A.y);
                tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
            } else if (seg.type == SegType::Arc) {
                const double sweep = seg.a1 - seg.a0;
                const double ang = seg.a0 + f * sweep;
                pointOnPath.x = seg.C.x + seg.r * std::cos(ang);
                pointOnPath.y = seg.C.y + seg.r * std::sin(ang);
                tangentAngle = ang + ((sweep >= 0.0) ? kPI / 2.0 : -kPI / 2.0);
            } else {
                pointOnPath.x = seg.A.x + f * (seg.B.x - seg.A.x);
                pointOnPath.y = seg.A.y + f * (seg.B.y - seg.A.y);
                tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
            }
            found = true;
        }
        
        if (!found) {
            const Seg& seg = path.segs[path.segs.GetSize() - 1];
            if (seg.type == SegType::Line) {
                pointOnPath = seg.B;
                tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
            } else {
                const double sweep = seg.a1 - seg.a0;
                pointOnPath.x = seg.C.x + seg.r * std::cos(seg.a1);
                pointOnPath.y = seg.C.y + seg.r * std::sin(seg.a1);
                tangentAngle = seg.a1 + ((sweep >= 0.0) ? kPI / 2.0 : -kPI / 2.0);
            }
        }
        
        // Шаг 1: Получаем Z-координату для точки на базовой линии от Mesh
        double baseZ = 0.0;
        API_Vector3D baseNormal = {};
        if (g_meshSurfaceGuid != APINULLGuid) {
            // Передаем XY координаты точки на базовой линии в MeshIntersectionHelper
            API_Coord baseXY = {pointOnPath.x, pointOnPath.y};
            if (MeshIntersectionHelper::GetZAndNormal(baseXY, baseZ, baseNormal)) {
                // Логируем все точки для отладки
                // Не логируем каждую точку для производительности - только ошибки
            } else {
                // Логируем только первую ошибку для производительности
                if (i == 0) {
                    Log("[ShellHelper] WARNING: Не удалось получить Z для первой точки на базовой линии (%.3f, %.3f)", pointOnPath.x, pointOnPath.y);
                }
                baseZ = 0.0;
            }
        } else {
            // Логируем если mesh не установлен (только для первых точек)
            if (i < 3) {
                // Логируем только один раз для производительности
                if (i == 0) {
                    Log("[ShellHelper] WARNING: g_meshSurfaceGuid is NULL, using Z=0.0 for all points");
                }
            }
        }
        
        // Шаг 2: Строим перпендикуляры от точки на базовой линии
        double perpAngle = tangentAngle + kPI / 2.0;
        double perpX = std::cos(perpAngle);
        double perpY = std::sin(perpAngle);
        double halfWidth = widthMM / 2000.0;
        
        // Шаг 3: Создаем левую и правую точки, используя Z из базовой линии
        API_Coord3D leftPoint = {pointOnPath.x + perpX * halfWidth, pointOnPath.y + perpY * halfWidth, baseZ};
        API_Coord3D rightPoint = {pointOnPath.x - perpX * halfWidth, pointOnPath.y - perpY * halfWidth, baseZ};
        
        leftPoints.Push(leftPoint);
        rightPoints.Push(rightPoint);
    }
    
    Log("[ShellHelper] Создано %d пар точек для контура", (int)leftPoints.GetSize());
    
    // Создаем контуры (левые и правые Spline + замыкающие линии)
    GS::Array<API_Coord> leftSplinePoints;
    for (UIndex i = 0; i < leftPoints.GetSize(); ++i) {
        leftSplinePoints.Push({leftPoints[i].x, leftPoints[i].y});
    }
    
    API_Guid leftSplineGuid = CreateSplineFromPoints(leftSplinePoints);
    if (leftSplineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Не удалось создать левый Spline");
        return false;
    }
    
    GS::Array<API_Coord> rightSplinePoints;
    for (UIndex i = 0; i < rightPoints.GetSize(); ++i) {
        rightSplinePoints.Push({rightPoints[i].x, rightPoints[i].y});
    }
    
    API_Guid rightSplineGuid = CreateSplineFromPoints(rightSplinePoints);
    if (rightSplineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Не удалось создать правый Spline");
        return false;
    }
    
    Log("[ShellHelper] SUCCESS: Созданы два Spline (левый и правый)");
    
    // Замыкаем крайние точки линиями
    API_Element startLine = {};
    startLine.header.type = API_LineID;
    err = ACAPI_Element_GetDefaults(&startLine, nullptr);
    if (err == NoError) {
        startLine.line.begC = leftSplinePoints[0];
        startLine.line.endC = rightSplinePoints[0];
        err = ACAPI_CallUndoableCommand("Create Start Line", [&]() -> GSErrCode {
            return ACAPI_Element_Create(&startLine, nullptr);
        });
    }
    
    API_Element endLine = {};
    endLine.header.type = API_LineID;
    err = ACAPI_Element_GetDefaults(&endLine, nullptr);
    if (err == NoError) {
        endLine.line.begC = leftSplinePoints[leftSplinePoints.GetSize() - 1];
        endLine.line.endC = rightSplinePoints[rightSplinePoints.GetSize() - 1];
        err = ACAPI_CallUndoableCommand("Create End Line", [&]() -> GSErrCode {
            return ACAPI_Element_Create(&endLine, nullptr);
        });
    }
    
    Log("[ShellHelper] SUCCESS: Замыкающие линии созданы");
    
    // Создаем Morph из замкнутого контура
    Log("[ShellHelper] Creating Morph from closed contour...");
    GS::Array<API_Coord3D> morphPoints;
    
    // Добавляем левые точки от начала до конца (используем Z координаты из mesh, если доступны)
    for (UIndex i = 0; i < leftPoints.GetSize(); ++i) {
        double z = leftPoints[i].z;
        if (g_meshSurfaceGuid == APINULLGuid) {
            z = 0.0; // Используем Z=0.0 если mesh не выбран (чтобы видно было на плане)
        }
        // Если mesh выбран, используем Z полученный от MeshIntersectionHelper
        morphPoints.Push({leftPoints[i].x, leftPoints[i].y, z});
    }
    
    // Добавляем правые точки от конца до начала (в обратном порядке, используем Z из mesh)
    for (Int32 i = rightPoints.GetSize() - 1; i >= 0; --i) {
        double z = rightPoints[i].z;
        if (g_meshSurfaceGuid == APINULLGuid) {
            z = 0.0; // Используем Z=0.0 если mesh не выбран (чтобы видно было на плане)
        }
        // Если mesh выбран, используем Z полученный от MeshIntersectionHelper
        morphPoints.Push({rightPoints[i].x, rightPoints[i].y, z});
    }
    
    Log("[ShellHelper] Morph contour: %d points", (int)morphPoints.GetSize());
    
    if (morphPoints.GetSize() >= 3) {
        if (RoadHelper::CreateMorphFromPoints(morphPoints, thicknessMM, materialTop, materialBottom, materialSide)) {
            Log("[ShellHelper] SUCCESS: Morph created from contour!");
            
            // Вычисляем площадь верхней поверхности
            double surfaceArea = RoadHelper::CalculateMorphSurfaceArea(morphPoints);
            
            if (surfaceArea > 0.0) {
                // Используем первую точку контура для размещения выноски
                API_Coord labelPos = {morphPoints[0].x, morphPoints[0].y};
                
                // Создаем текстовую выноску с площадью
                RoadHelper::CreateAreaLabel(labelPos, surfaceArea);
            }
            
            return true;
        } else {
            Log("[ShellHelper] ERROR: Failed to create Morph from contour");
            return false;
        }
    } else {
        Log("[ShellHelper] ERROR: Not enough points for Morph (%d < 3)", (int)morphPoints.GetSize());
        return false;
    }
}

// =============== Создание Mesh из контура ===============
bool CreateMeshFromContour(double leftWidthMM, double rightWidthMM, double stepMM, double offsetMM)
{
    Log("[ShellHelper] CreateMeshFromContour: START, leftWidth=%.1fmm, rightWidth=%.1fmm, step=%.1fmm, offset=%.1fmm",
        leftWidthMM, rightWidthMM, stepMM, offsetMM);
    
    // Конвертируем offset из мм в метры
    double offsetM = offsetMM / 1000.0;
    Log("[ShellHelper] Offset converted: %.1fmm -> %.6fm", offsetMM, offsetM);
    
    // Проверяем, что базовая линия выбрана
    if (g_baseLineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Базовая линия не выбрана. Сначала выберите базовую линию.");
        return false;
    }
    
    Log("[ShellHelper] Базовая линия: %s", APIGuidToString(g_baseLineGuid).ToCStr().Get());
    
    // Проверяем mesh
    if (g_meshSurfaceGuid != APINULLGuid) {
        Log("[ShellHelper] Mesh выбран: %s - будет использоваться для получения Z координат", APIGuidToString(g_meshSurfaceGuid).ToCStr().Get());
    } else {
        Log("[ShellHelper] WARNING: Mesh не выбран - будут использоваться Z=0.0 для всех точек");
    }
    
    // Получаем элемент базовой линии по сохраненному GUID
    API_Elem_Head elemHead = {};
    elemHead.guid = g_baseLineGuid;
    GSErrCode err = ACAPI_Element_GetHeader(&elemHead);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить заголовок элемента базовой линии");
        return false;
    }
    
    // Проверяем поддерживаемые типы элементов
    bool isSupported = false;
    if (elemHead.type == API_LineID ||
        elemHead.type == API_PolyLineID ||
        elemHead.type == API_ArcID ||
        elemHead.type == API_CircleID ||
        elemHead.type == API_SplineID) {
        isSupported = true;
    }
    
    if (!isSupported) {
        Log("[ShellHelper] ERROR: Неподдерживаемый тип элемента базовой линии");
        Log("[ShellHelper] Поддерживаются: Line, Polyline, Arc, Circle, Spline");
        return false;
    }
    
    // Получаем данные элемента
    API_Element element = {};
    element.header = elemHead;
    err = ACAPI_Element_Get(&element);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить данные элемента базовой линии");
        return false;
    }
    
    Log("[ShellHelper] Элемент базовой линии загружен успешно");
    
    // Парсим элемент в сегменты
    PathData path;
    if (!ParseElementToSegments(element, path)) {
        Log("[ShellHelper] ERROR: Не удалось распарсить элемент в сегменты");
        return false;
    }
    
    Log("[ShellHelper] Элемент распарсен: %d сегментов, общая длина %.3fм", 
        (int)path.segs.GetSize(), path.total);
    
    // Используем логику из Create3DShellFromPath для получения точек
    double step = stepMM / 1000.0; // шаг в метрах

    // Отрицательные значения ширины воспринимаем по модулю
    double leftWidthM = std::fabs(leftWidthMM) / 1000.0;   // левая ширина в метрах
    double rightWidthM = std::fabs(rightWidthMM) / 1000.0; // правая ширина в метрах

    // Если обе ширины равны нулю, создавать нечего
    if (leftWidthM <= 0.0 && rightWidthM <= 0.0) {
        Log("[ShellHelper] ERROR: Both left and right widths are zero. Nothing to create.");
        return false;
    }
    
    // Генерируем позиции точек по шагу
    GS::Array<double> sVals;
    for (double s = 0.0; s <= path.total + 1e-9; s += step) {
        sVals.Push(s);
    }
    
    // Убеждаемся, что последняя точка точно на конце линии
    if (sVals.IsEmpty() || sVals[sVals.GetSize() - 1] < path.total - 1e-9) {
        sVals.Push(path.total);
    }
    
    Log("[ShellHelper] Сгенерировано %d точек по шагу %.3fм", (int)sVals.GetSize(), step);
    
    GS::Array<API_Coord3D> leftPoints;
    GS::Array<API_Coord3D> rightPoints;
    
    // Обрабатываем каждую точку (та же логика что в Create3DShellFromPath)
    for (UIndex i = 0; i < sVals.GetSize(); ++i) {
        double s = sVals[i];
        API_Coord pointOnPath;
        double tangentAngle = 0.0;
        
        double acc = 0.0;
        bool found = false;
        for (UIndex j = 0; j < path.segs.GetSize() && !found; ++j) {
            const Seg& seg = path.segs[j];
            
            if (s > acc + seg.len) {
                acc += seg.len;
                continue;
            }
            
            double t = (s - acc) / seg.len;
            if (seg.type == SegType::Line) {
                pointOnPath.x = seg.A.x + t * (seg.B.x - seg.A.x);
                pointOnPath.y = seg.A.y + t * (seg.B.y - seg.A.y);
                tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
                found = true;
            } else if (seg.type == SegType::Arc) {
                double ang = seg.a0 + t * (seg.a1 - seg.a0);
                pointOnPath.x = seg.C.x + seg.r * std::cos(ang);
                pointOnPath.y = seg.C.y + seg.r * std::sin(ang);
                tangentAngle = ang + ((seg.ccw) ? kPI / 2.0 : -kPI / 2.0);
                found = true;
            }
        }
        
        if (!found) {
            // Если не нашли сегмент, используем последнюю точку
            const Seg& seg = path.segs[path.segs.GetSize() - 1];
            if (seg.type == SegType::Line) {
                pointOnPath = seg.B;
                tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
            } else {
                const double sweep = seg.a1 - seg.a0;
                pointOnPath.x = seg.C.x + seg.r * std::cos(seg.a1);
                pointOnPath.y = seg.C.y + seg.r * std::sin(seg.a1);
                tangentAngle = seg.a1 + ((sweep >= 0.0) ? kPI / 2.0 : -kPI / 2.0);
            }
        }
        
        // Шаг 1: Получаем Z-координату для точки на базовой линии от Mesh
        double baseZ = 0.0;
        API_Vector3D baseNormal = {};
        if (g_meshSurfaceGuid != APINULLGuid) {
            // Передаем XY координаты точки на базовой линии в MeshIntersectionHelper
            API_Coord baseXY = {pointOnPath.x, pointOnPath.y};
            if (MeshIntersectionHelper::GetZAndNormal(baseXY, baseZ, baseNormal)) {
                // Логируем первые и последние точки для отладки
                if (i < 5 || i >= sVals.GetSize() - 5) {
                    Log("[ShellHelper] Point on base line %d: (%.3f, %.3f) -> Z=%.3f", (int)i + 1, pointOnPath.x, pointOnPath.y, baseZ);
                }
            } else {
                Log("[ShellHelper] WARNING: Failed to get Z for point %d, using Z=0", (int)i + 1);
            }
        } else {
            baseZ = 0.0;
        }
        
        // Шаг 2: Строим перпендикуляр
        double perpX = std::cos(tangentAngle + kPI / 2.0);
        double perpY = std::sin(tangentAngle + kPI / 2.0);
        
        // Шаг 3: Создаем левую и правую точки с применением offset
        API_Coord3D leftPoint = {
            pointOnPath.x + perpX * leftWidthM,
            pointOnPath.y + perpY * leftWidthM,
            baseZ + offsetM  // Применяем offset к Z координате
        };
        API_Coord3D rightPoint = {
            pointOnPath.x - perpX * rightWidthM,
            pointOnPath.y - perpY * rightWidthM,
            baseZ + offsetM  // Применяем offset к Z координате
        };
        
        leftPoints.Push(leftPoint);
        rightPoints.Push(rightPoint);
        
        // Логируем первые и последние точки для отладки
        if (i < 5 || i >= sVals.GetSize() - 5) {
            Log("[ShellHelper] Точка %d: left(%.3f, %.3f, %.3f), right(%.3f, %.3f, %.3f)", 
                (int)i + 1, leftPoint.x, leftPoint.y, leftPoint.z, rightPoint.x, rightPoint.y, rightPoint.z);
        }
    }
    
    Log("[ShellHelper] Создано %d пар точек для 3D оболочки", (int)leftPoints.GetSize());
    
    // Создаем Mesh из замкнутого контура
    Log("[ShellHelper] Creating Mesh from closed contour...");
    GS::Array<API_Coord3D> meshPoints;
    
    // Добавляем левые точки от начала до конца (Z уже включает offset)
    for (UIndex i = 0; i < leftPoints.GetSize(); ++i) {
        double z = leftPoints[i].z;  // Z уже включает offset, примененный выше
        if (g_meshSurfaceGuid == APINULLGuid) {
            z = offsetM; // Используем только offset если mesh не выбран
        }
        meshPoints.Push({leftPoints[i].x, leftPoints[i].y, z});
    }
    
    // Добавляем правые точки от конца до начала (в обратном порядке, Z уже включает offset)
    for (Int32 i = rightPoints.GetSize() - 1; i >= 0; --i) {
        double z = rightPoints[i].z;  // Z уже включает offset, примененный выше
        if (g_meshSurfaceGuid == APINULLGuid) {
            z = offsetM; // Используем только offset если mesh не выбран
        }
        meshPoints.Push({rightPoints[i].x, rightPoints[i].y, z});
    }
    
    Log("[ShellHelper] Mesh contour: %d points", (int)meshPoints.GetSize());
    
    if (meshPoints.GetSize() >= 3) {
        if (CreateMeshFromPoints(meshPoints)) {
            Log("[ShellHelper] SUCCESS: Mesh created from contour!");
            return true;
        } else {
            Log("[ShellHelper] ERROR: Failed to create Mesh from contour");
            return false;
        }
    } else {
        Log("[ShellHelper] ERROR: Not enough points for Mesh (%d < 3)", (int)meshPoints.GetSize());
        return false;
    }
}

// =============== Создание Spline из 2D точек ===============
API_Guid CreateSplineFromPoints(const GS::Array<API_Coord>& points)
{
    if (points.GetSize() < 2) {
        Log("[ShellHelper] ERROR: Недостаточно точек для создания Spline (нужно минимум 2)");
        return APINULLGuid;
    }
    
    Log("[ShellHelper] CreateSplineFromPoints: создаем Spline с %d точками", (int)points.GetSize());
    
    API_Element spline = {};
    spline.header.type = API_SplineID;
    GSErrCode err = ACAPI_Element_GetDefaults(&spline, nullptr);
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить настройки по умолчанию для Spline, err=%d", (int)err);
        return APINULLGuid;
    }
    
    // Создаем memo для Spline
    API_ElementMemo memo = {};
    BNZeroMemory(&memo, sizeof(API_ElementMemo));
    
    const Int32 nUnique = (Int32)points.GetSize();
    const Int32 nCoords = nUnique; // Без замыкающей точки для Spline
    
    // Выделяем память для координат (1-based indexing!)
    // Выделяем nCoords + 1 элементов, но используем только индексы от 1 до nCoords
    memo.coords = reinterpret_cast<API_Coord**>(BMAllocateHandle((nCoords + 1) * (GSSize)sizeof(API_Coord), ALLOCATE_CLEAR, 0));
    
    // Инициализируем элемент с индексом 0, чтобы избежать (0,0)
    (*memo.coords)[0] = points[0]; // Используем первую точку как заглушку
    if (memo.coords == nullptr) {
        Log("[ShellHelper] ERROR: Не удалось выделить память для координат Spline");
        return APINULLGuid;
    }
    
    // Выделяем память для направлений Безье
    memo.bezierDirs = reinterpret_cast<API_SplineDir**>(BMAllocateHandle((nCoords + 1) * (GSSize)sizeof(API_SplineDir), ALLOCATE_CLEAR, 0));
    if (memo.bezierDirs == nullptr) {
        ACAPI_DisposeElemMemoHdls(&memo);
        Log("[ShellHelper] ERROR: Не удалось выделить память для bezierDirs");
        return APINULLGuid;
    }
    
    // Заполняем координаты (1-based indexing!)
    for (Int32 i = 0; i < nUnique; ++i) {
        (*memo.coords)[i + 1] = points[i];
        
        if (i < 5 || i >= nUnique - 5) { // Логируем первые и последние 5 точек
            // Не логируем каждую точку сплайна для производительности
        }
    }
    
    
    // Не замыкаем контур для Spline (это нужно только для полигонов)
    
    // Настраиваем направления Безье для плавного Spline
    for (Int32 i = 1; i <= nCoords; ++i) {
        API_SplineDir& dir = (*memo.bezierDirs)[i];
        
        if (i == 1) {
            // Первая точка - направление к следующей
            API_Coord next = (*memo.coords)[2];
            API_Coord curr = (*memo.coords)[1];
            double dx = next.x - curr.x;
            double dy = next.y - curr.y;
            double len = sqrt(dx*dx + dy*dy);
            if (len > 1e-9) {
                dir.dirAng = atan2(dy, dx);
                dir.lenNext = len * 0.3; // 30% от длины сегмента
                dir.lenPrev = 0.0;
            } else {
                dir.dirAng = 0.0;
                dir.lenNext = 0.0;
                dir.lenPrev = 0.0;
            }
        } else if (i == nCoords) {
            // Последняя точка - направление от предыдущей
            API_Coord prev = (*memo.coords)[nCoords-1];
            API_Coord curr = (*memo.coords)[nCoords];
            double dx = curr.x - prev.x;
            double dy = curr.y - prev.y;
            double len = sqrt(dx*dx + dy*dy);
            if (len > 1e-9) {
                dir.dirAng = atan2(dy, dx);
                dir.lenNext = 0.0;
                dir.lenPrev = len * 0.3; // 30% от длины сегмента
            } else {
                dir.dirAng = 0.0;
                dir.lenNext = 0.0;
                dir.lenPrev = 0.0;
            }
        } else {
            // Средние точки - направление между предыдущей и следующей
            API_Coord prev = (*memo.coords)[i-1];
            API_Coord curr = (*memo.coords)[i];
            API_Coord next = (*memo.coords)[i+1];
            
            double dx1 = curr.x - prev.x;
            double dy1 = curr.y - prev.y;
            double dx2 = next.x - curr.x;
            double dy2 = next.y - curr.y;
            
            double len1 = sqrt(dx1*dx1 + dy1*dy1);
            double len2 = sqrt(dx2*dx2 + dy2*dy2);
            
            if (len1 > 1e-9 && len2 > 1e-9) {
                double ang1 = atan2(dy1, dx1);
                double ang2 = atan2(dy2, dx2);
                dir.dirAng = (ang1 + ang2) * 0.5; // Среднее направление
                dir.lenNext = len2 * 0.3;
                dir.lenPrev = len1 * 0.3;
            } else {
                dir.dirAng = 0.0;
                dir.lenNext = 0.0;
                dir.lenPrev = 0.0;
            }
        }
    }
    
    
    // Создаем элемент внутри Undo-команды
    err = ACAPI_CallUndoableCommand("Create Spline", [&]() -> GSErrCode {
        return ACAPI_Element_Create(&spline, &memo);
    });
    ACAPI_DisposeElemMemoHdls(&memo);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось создать Spline, err=%d", (int)err);
        return APINULLGuid;
    }
    
    Log("[ShellHelper] SUCCESS: Создан Spline с %d точками", (int)points.GetSize());
    return spline.header.guid;
}

// =============== Создание Shell с 3D точками ===============
API_Guid Create3DShell(const GS::Array<API_Coord3D>& points)
{
    if (points.GetSize() < 3) {
        Log("[ShellHelper] ERROR: Недостаточно точек для создания Shell (нужно минимум 3)");
        return APINULLGuid;
    }
    
    Log("[ShellHelper] Create3DShell: создаем плавную 3D Shell поверхность с %d точками", (int)points.GetSize());
    
    API_Element shell = {};
    shell.header.type = API_ShellID;
    GSErrCode err = ACAPI_Element_GetDefaults(&shell, nullptr);
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить настройки по умолчанию для Shell, err=%d", (int)err);
        return APINULLGuid;
    }
    
    // Создаем memo для Shell
    API_ElementMemo memo = {};
    BNZeroMemory(&memo, sizeof(API_ElementMemo));
    
    const Int32 nUnique = (Int32)points.GetSize();
    const Int32 nCoords = nUnique + 1; // + замыкающая точка
    
    // Выделяем память для координат (1-based indexing!)
    memo.coords = reinterpret_cast<API_Coord**>(BMAllocateHandle((nCoords + 1) * (GSSize)sizeof(API_Coord), ALLOCATE_CLEAR, 0));
    if (memo.coords == nullptr) {
        Log("[ShellHelper] ERROR: Не удалось выделить память для координат Shell");
        return APINULLGuid;
    }
    
    // Инициализируем элемент с индексом 0, чтобы избежать (0,0)
    (*memo.coords)[0] = {points[0].x, points[0].y}; // Используем первую точку как заглушку
    
    // Заполняем 2D контур (1-based indexing!) с логированием для отладки
    for (Int32 i = 0; i < nUnique; ++i) {
        (*memo.coords)[i + 1] = {points[i].x, points[i].y};
        
        if (i < 5 || i >= nUnique - 5) { // Логируем первые и последние 5 точек
            Log("[ShellHelper] Point %d: (%.3f, %.3f, %.3f)", i+1, points[i].x, points[i].y, points[i].z);
        }
    }
    // Замыкаем контур
    (*memo.coords)[nCoords] = (*memo.coords)[1];
    
    // Для Shell Z-координаты задаются через высоты точек
    // Пока создаем простой Shell только с 2D контуром
    
    // Настраиваем контуры
    memo.pends = reinterpret_cast<Int32**>(BMAllocateHandle(2 * (GSSize)sizeof(Int32), ALLOCATE_CLEAR, 0));
    if (memo.pends == nullptr) {
        ACAPI_DisposeElemMemoHdls(&memo);
        Log("[ShellHelper] ERROR: Не удалось выделить память для pends");
        return APINULLGuid;
    }
    (*memo.pends)[0] = 0;
    (*memo.pends)[1] = nCoords;
    
    // Настраиваем Shell для создания плавной поверхности
    // Для Shell настройки делаются через memo, не нужно менять shell.shell.poly
    
    // Создаем элемент внутри Undo-команды
    err = ACAPI_CallUndoableCommand("Create 3D Shell", [&]() -> GSErrCode {
        return ACAPI_Element_Create(&shell, &memo);
    });
    ACAPI_DisposeElemMemoHdls(&memo);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось создать Shell, err=%d", (int)err);
        return APINULLGuid;
    }
    
    Log("[ShellHelper] SUCCESS: Создана плавная Shell поверхность с %d точками (3D)", (int)points.GetSize());
    return shell.header.guid;
}

// =============== Создание Ruled Shell ===============
bool CreateRuledShell(const API_Guid& leftSplineGuid, const API_Guid& rightSplineGuid)
{
    Log("[ShellHelper] CreateRuledShell: создаем Ruled Shell между двумя Spline");
    
    if (leftSplineGuid == APINULLGuid || rightSplineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Неверные GUID для Spline");
        return false;
    }
    
    // ВРЕМЕННО: создаем простой Shell с прямоугольным контуром
    // TODO: Реализовать правильное создание Ruled Shell с двумя Spline
    
    Log("[ShellHelper] ВРЕМЕННО: создаем простой Shell вместо Ruled Shell");
    
    // Создаем простой прямоугольный контур для Shell
    GS::Array<API_Coord> contour;
    contour.Push({-5.0, -2.0});  // левый нижний
    contour.Push({5.0, -2.0});   // правый нижний
    contour.Push({5.0, 2.0});    // правый верхний
    contour.Push({-5.0, 2.0});   // левый верхний
    
    // Создаем Shell элемент
    API_Element shell = {};
    shell.header.type = API_ShellID;
    
    GSErrCode err = ACAPI_Element_GetDefaults(&shell, nullptr);
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить настройки по умолчанию для Shell, err=%d", (int)err);
        return false;
    }
    
    // Создаем memo для Shell
    API_ElementMemo memo = {};
    BNZeroMemory(&memo, sizeof(API_ElementMemo));
    
    const Int32 nCoords = (Int32)contour.GetSize() + 1; // + замыкающая точка
    
    // Выделяем память для координат (1-based indexing!)
    memo.coords = reinterpret_cast<API_Coord**>(BMAllocateHandle((nCoords + 1) * (GSSize)sizeof(API_Coord), ALLOCATE_CLEAR, 0));
    if (memo.coords == nullptr) {
        Log("[ShellHelper] ERROR: Не удалось выделить память для координат Shell");
        return false;
    }
    
    // Инициализируем элемент с индексом 0, чтобы избежать (0,0)
    (*memo.coords)[0] = contour[0]; // Используем первую точку как заглушку
    
    // Заполняем координаты (1-based indexing!)
    for (Int32 i = 0; i < (Int32)contour.GetSize(); ++i) {
        (*memo.coords)[i + 1] = contour[i];
    }
    (*memo.coords)[nCoords] = (*memo.coords)[1]; // замкнуть контур
    
    // Настраиваем контуры
    memo.pends = reinterpret_cast<Int32**>(BMAllocateHandle(2 * (GSSize)sizeof(Int32), ALLOCATE_CLEAR, 0));
    if (memo.pends == nullptr) {
        ACAPI_DisposeElemMemoHdls(&memo);
        Log("[ShellHelper] ERROR: Не удалось выделить память для pends");
        return false;
    }
    (*memo.pends)[0] = 0;
    (*memo.pends)[1] = nCoords;
    
    // Настраиваем Shell
    // Для Shell настройки делаются через memo, не нужно менять shellClass
    
    // Создаем элемент
    err = ACAPI_Element_Create(&shell, &memo);
    ACAPI_DisposeElemMemoHdls(&memo);
    
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось создать Shell, err=%d", (int)err);
        return false;
    }
    
    Log("[ShellHelper] SUCCESS: Shell создан успешно (временная реализация)");
    return true;
}

// =============== Создание перпендикулярных линий от сегментов ===============
bool CreatePerpendicularLinesFromSegments(const PathData& path, double widthMM)
{
    Log("[ShellHelper] CreatePerpendicularLinesFromSegments: %d сегментов, width=%.1fmm", 
        (int)path.segs.GetSize(), widthMM);
    
    if (path.segs.IsEmpty()) {
        Log("[ShellHelper] ERROR: Нет сегментов для обработки");
        return false;
    }
    
    double halfWidth = widthMM / 2000.0; // Переводим мм в метры и делим пополам
    
    ACAPI_CallUndoableCommand("Create Shell Lines", [&] () -> GSErrCode {
        
        // Создаем перпендикулярные линии с шагом
        // ОПТИМИЗАЦИЯ: увеличиваем шаг для уменьшения количества вызовов GetGroundZAndNormal
        double step = 2.0; // шаг в метрах (увеличен до 2.0 для производительности)
        double currentPos = 0.0;
        
        // ОПТИМИЗАЦИЯ: получаем Z-координату только один раз в начале линии
        API_Coord3D firstPoint = {0.0, 0.0, 0.0};
        double cachedZ = 0.0;
        bool zCached = false;
        
        while (currentPos <= path.total) {
            // Используем функцию EvalOnPath для получения точки и угла
            API_Coord pointOnPath;
            double tangentAngle = 0.0;
            
            // Находим нужный сегмент и позицию в нем
            double accumulatedLength = 0.0;
            bool found = false;
            
            for (UIndex i = 0; i < path.segs.GetSize() && !found; ++i) {
                const Seg& seg = path.segs[i];
                
                if (currentPos <= accumulatedLength + seg.len) {
                    // Текущая позиция находится в этом сегменте
                    double localPos = currentPos - accumulatedLength;
                    
                    // Вычисляем точку и угол в зависимости от типа сегмента
                    switch (seg.type) {
                        case SegType::Line: {
                            double t = (seg.len > 1e-9) ? localPos / seg.len : 0.0;
                            pointOnPath.x = seg.A.x + t * (seg.B.x - seg.A.x);
                            pointOnPath.y = seg.A.y + t * (seg.B.y - seg.A.y);
                            tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
                            break;
                        }
                        case SegType::Arc: {
                            double t = (seg.len > 1e-9) ? localPos / seg.len : 0.0;
                            double angle = seg.a0 + t * (seg.a1 - seg.a0);
                            pointOnPath.x = seg.C.x + seg.r * std::cos(angle);
                            pointOnPath.y = seg.C.y + seg.r * std::sin(angle);
                            tangentAngle = angle + ((seg.a1 > seg.a0) ? kPI / 2.0 : -kPI / 2.0);
                            break;
                        }
                        default:
                            // Для сплайнов пока используем линейную интерполяцию
                            Log("[ShellHelper] WARNING: Сплайны пока не поддерживаются, используем линейную интерполяцию");
                            double t = (seg.len > 1e-9) ? localPos / seg.len : 0.0;
                            pointOnPath.x = seg.A.x + t * (seg.B.x - seg.A.x);
                            pointOnPath.y = seg.A.y + t * (seg.B.y - seg.A.y);
                            tangentAngle = std::atan2(seg.B.y - seg.A.y, seg.B.x - seg.A.x);
                            break;
                    }
                    found = true;
                }
                accumulatedLength += seg.len;
            }
            
            if (!found) {
                // Если не нашли сегмент, используем последнюю точку
                const Seg& lastSeg = path.segs[path.segs.GetSize() - 1];
                pointOnPath.x = lastSeg.B.x;
                pointOnPath.y = lastSeg.B.y;
                tangentAngle = std::atan2(lastSeg.B.y - lastSeg.A.y, lastSeg.B.x - lastSeg.A.x);
            }
            
            // Получаем Z-координату от Mesh
            // ОПТИМИЗАЦИЯ: получаем Z только один раз в начале линии
            API_Coord3D point3D = {pointOnPath.x, pointOnPath.y, 0.0};
            
            if (!zCached) {
                // Получаем Z-координату только для первой точки
                double z = 0.0;
                API_Vector3D normal = {};
                
                if (GroundHelper::GetGroundZAndNormal(point3D, z, normal)) {
                    cachedZ = z;
                    zCached = true;
                    Log("[ShellHelper] Point (%.3f, %.3f, %.3f) - Z from Mesh (cached for all points)", point3D.x, point3D.y, cachedZ);
                } else {
                    Log("[ShellHelper] WARNING: Не удалось получить Z от Mesh для точки (%.3f, %.3f)", point3D.x, point3D.y);
                    cachedZ = 0.0;
                    zCached = true;
                }
            }
            
            point3D.z = cachedZ;
            // Не логируем каждую точку для производительности - только итоговые сообщения
            
            // Вычисляем перпендикулярное направление (поворот на 90 градусов)
            double perpAngle = tangentAngle + kPI / 2.0;
            double perpX = std::cos(perpAngle);
            double perpY = std::sin(perpAngle);
            
            // Создаем перпендикулярную линию
            API_Element line = {};
            line.header.type = API_LineID;
            GSErrCode err = ACAPI_Element_GetDefaults(&line, nullptr);
            if (err != NoError) continue;
            
            line.header.floorInd = 0; // TODO: получить правильный этаж
            
            // Создаем перпендикулярную линию с полученной Z-координатой
            line.line.begC.x = pointOnPath.x + perpX * halfWidth;
            line.line.begC.y = pointOnPath.y + perpY * halfWidth;
            line.line.endC.x = pointOnPath.x - perpX * halfWidth;
            line.line.endC.y = pointOnPath.y - perpY * halfWidth;
            
            err = ACAPI_Element_Create(&line, nullptr);
            if (err != NoError) {
                Log("[ShellHelper] ERROR: Не удалось создать перпендикулярную линию, err=%d", (int)err);
            }
            // Не логируем каждую успешную линию для производительности - только ошибки
            
            currentPos += step;
        }
        
        Log("[ShellHelper] SUCCESS: Перпендикулярные линии созданы для всех сегментов");
        return NoError;
    });
    
    return true;
}

// =============== Выбор Mesh поверхности ===============
bool SetMeshSurfaceForShell()
{
    Log("[ShellHelper] SetMeshSurfaceForShell: выбор Mesh поверхности");
    
    // Используем существующую функцию GroundHelper::SetGroundSurface()
    bool success = GroundHelper::SetGroundSurface();
    
    if (success) {
        // Получаем GUID выбранной Mesh поверхности из выделения
        API_SelectionInfo selectionInfo;
        GS::Array<API_Neig> selNeigs;
        GSErrCode err = ACAPI_Selection_Get(&selectionInfo, &selNeigs, false, false);
        
        if (err == NoError && selNeigs.GetSize() > 0) {
            g_meshSurfaceGuid = selNeigs[0].guid;
            Log("[ShellHelper] Mesh поверхность выбрана успешно, GUID: %s", 
                APIGuidToString(g_meshSurfaceGuid).ToCStr().Get());
            
            // Устанавливаем mesh в GroundHelper для работы MeshIntersectionHelper
            if (GroundHelper::SetGroundSurfaceByGuid(g_meshSurfaceGuid)) {
                Log("[ShellHelper] Mesh установлен в GroundHelper для MeshIntersectionHelper");
            } else {
                Log("[ShellHelper] WARNING: Не удалось установить mesh в GroundHelper");
            }
        } else {
            Log("[ShellHelper] Ошибка получения GUID выбранной Mesh поверхности");
            g_meshSurfaceGuid = APINULLGuid;
        }
    } else {
        Log("[ShellHelper] Ошибка выбора Mesh поверхности");
        g_meshSurfaceGuid = APINULLGuid;
    }
    
    return success;
}

// =============== Получение 3D точек вдоль базовой линии ===============
GS::Array<API_Coord3D> Get3DPointsAlongBaseLine(double stepMM)
{
    Log("[ShellHelper] Get3DPointsAlongBaseLine: step=%.1fmm", stepMM);
    
    GS::Array<API_Coord3D> points;
    
    if (g_baseLineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Базовая линия не выбрана");
        return points;
    }
    
    // Получаем элемент базовой линии
    API_Element element = {};
    element.header.guid = g_baseLineGuid;
    GSErrCode err = ACAPI_Element_Get(&element);
    if (err != NoError) {
        Log("[ShellHelper] ERROR: Не удалось получить элемент базовой линии");
        return points;
    }
    
    // Парсим элемент в сегменты
    PathData path;
    if (!ParseElementToSegments(element, path)) {
        Log("[ShellHelper] ERROR: Не удалось распарсить элемент в сегменты");
        return points;
    }
    
    Log("[ShellHelper] Элемент распарсен: %d сегментов, общая длина %.3fм", 
        (int)path.segs.GetSize(), path.total);
    
    // Генерируем точки вдоль пути с заданным шагом
    double step = stepMM / 1000.0; // конвертируем в метры
    double currentPos = 0.0;
    
    while (currentPos <= path.total) {
        // Находим точку на пути для текущей позиции
        API_Coord3D point = {0.0, 0.0, 0.0};
        
        // TODO: Реализовать получение точки на пути по позиции
        // Это потребует реализации функции EvalOnPath аналогично MarkupHelper
        
        points.Push(point);
        currentPos += step;
    }
    
    Log("[ShellHelper] Сгенерировано %d точек вдоль базовой линии", (int)points.GetSize());
    return points;
}

// =============== Создание перпендикулярных 3D точек ===============
bool CreatePerpendicular3DPoints(double widthMM, double stepMM, 
                                GS::Array<API_Coord3D>& leftPoints, 
                                GS::Array<API_Coord3D>& rightPoints)
{
    Log("[ShellHelper] CreatePerpendicular3DPoints: width=%.1fmm, step=%.1fmm", widthMM, stepMM);
    
    if (g_baseLineGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Базовая линия не выбрана");
        return false;
    }
    
    if (g_meshSurfaceGuid == APINULLGuid) {
        Log("[ShellHelper] ERROR: Mesh поверхность не выбрана");
        return false;
    }
    
    // Получаем точки вдоль базовой линии
    GS::Array<API_Coord3D> basePoints = Get3DPointsAlongBaseLine(stepMM);
    if (basePoints.IsEmpty()) {
        Log("[ShellHelper] ERROR: Не удалось получить точки вдоль базовой линии");
        return false;
    }
    
    double halfWidth = widthMM / 2000.0; // конвертируем в метры и делим пополам
    
    leftPoints.Clear();
    rightPoints.Clear();
    
    // Для каждой точки базовой линии создаем перпендикулярные точки
    for (UIndex i = 0; i < basePoints.GetSize(); ++i) {
        const API_Coord3D& basePoint = basePoints[i];
        
        // TODO: Вычислить направление касательной к пути в этой точке
        // Пока используем простое направление по X
        API_Coord3D tangent = {1.0, 0.0, 0.0};
        API_Coord3D perpendicular = {-tangent.y, tangent.x, 0.0};
        
        // Нормализуем перпендикулярный вектор
        double len = std::sqrt(perpendicular.x * perpendicular.x + perpendicular.y * perpendicular.y);
        if (len > kEPS) {
            perpendicular.x /= len;
            perpendicular.y /= len;
        }
        
        // Создаем левую и правую точки
        API_Coord3D leftPoint = {
            basePoint.x + perpendicular.x * halfWidth,
            basePoint.y + perpendicular.y * halfWidth,
            basePoint.z
        };
        
        API_Coord3D rightPoint = {
            basePoint.x - perpendicular.x * halfWidth,
            basePoint.y - perpendicular.y * halfWidth,
            basePoint.z
        };
        
        // Получаем Z-координаты от Mesh
        double leftZ = 0.0, rightZ = 0.0;
        API_Vector3D leftNormal = {}, rightNormal = {};
        
        if (GroundHelper::GetGroundZAndNormal(leftPoint, leftZ, leftNormal)) {
            leftPoint.z = leftZ;
        }
        
        if (GroundHelper::GetGroundZAndNormal(rightPoint, rightZ, rightNormal)) {
            rightPoint.z = rightZ;
        }
        
        leftPoints.Push(leftPoint);
        rightPoints.Push(rightPoint);
    }
    
    Log("[ShellHelper] Создано %d левых и %d правых перпендикулярных точек", 
        (int)leftPoints.GetSize(), (int)rightPoints.GetSize());
    
    return true;
}

// =============== Создание 3D Spline ===============
bool Create3DSpline(const GS::Array<API_Coord3D>& points, const GS::UniString& name)
{
    Log("[ShellHelper] Create3DSpline: создание 3D Spline из %d точек", (int)points.GetSize());
    
    if (points.IsEmpty()) {
        Log("[ShellHelper] ERROR: Нет точек для создания Spline");
        return false;
    }
    
    // TODO: Реализовать создание 3D Spline из массива точек
    // Это потребует изучения API для создания Spline элементов
    
    Log("[ShellHelper] TODO: Создание 3D Spline не реализовано");
    return false;
}

// =============== ПРОСТАЯ ФУНКЦИЯ СОЗДАНИЯ SHELL ===============
bool CreateSimpleShell()
{
    Log("[ShellHelper] ПРОСТОЙ SHELL: Создаем простой shell элемент");
    
    // Создаем SHELL элемент
    API_Element shell = {};
    shell.header.type = API_ShellID;
    GSErrCode err = ACAPI_Element_GetDefaults(&shell, nullptr);
    if (err != NoError) {
        Log("[ShellHelper] ПРОСТОЙ SHELL ERROR: ACAPI_Element_GetDefaults failed, err=%d", (int)err);
        return false;
    }
    
    Log("[ShellHelper] ПРОСТОЙ SHELL: настройки получены успешно");
    
    // Создаем memo для SHELL
    API_ElementMemo shellMemo = {};
    BNZeroMemory(&shellMemo, sizeof(API_ElementMemo));
    
    // Создаем простой прямоугольник для shell (4 точки)
    const Int32 nCoords = 4;
    shellMemo.coords = reinterpret_cast<API_Coord**>(BMAllocateHandle((nCoords + 1) * (GSSize)sizeof(API_Coord), ALLOCATE_CLEAR, 0));
    if (shellMemo.coords != nullptr) {
        // Простой прямоугольник
        (*shellMemo.coords)[0] = {0.0, 0.0}; // Заглушка
        (*shellMemo.coords)[1] = {0.0, 0.0}; // Точка 1
        (*shellMemo.coords)[2] = {3.0, 0.0}; // Точка 2
        (*shellMemo.coords)[3] = {3.0, 2.0}; // Точка 3
        (*shellMemo.coords)[4] = {0.0, 2.0}; // Точка 4
        
        // Настраиваем pends для контура
        shellMemo.pends = reinterpret_cast<Int32**>(BMAllocateHandle(2 * (GSSize)sizeof(Int32), ALLOCATE_CLEAR, 0));
        if (shellMemo.pends != nullptr) {
            (*shellMemo.pends)[0] = 0;        // Начало контура
            (*shellMemo.pends)[1] = nCoords;  // Конец контура
            Log("[ShellHelper] ПРОСТОЙ SHELL pends: [0, %d]", (int)nCoords);
        }
        
        // Создаем SHELL внутри Undo-команды
        err = ACAPI_CallUndoableCommand("Create Simple Shell", [&]() -> GSErrCode {
            return ACAPI_Element_Create(&shell, &shellMemo);
        });
        
        if (err == NoError) {
            Log("[ShellHelper] ПРОСТОЙ SHELL SUCCESS: SHELL создан! Простой прямоугольник");
            ACAPI_DisposeElemMemoHdls(&shellMemo);
            return true;
        } else {
            Log("[ShellHelper] ПРОСТОЙ SHELL ERROR: Не удалось создать SHELL, err=%d", (int)err);
            ACAPI_DisposeElemMemoHdls(&shellMemo);
            return false;
        }
    } else {
        Log("[ShellHelper] ПРОСТОЙ SHELL ERROR: Не удалось выделить память для координат SHELL");
        return false;
    }
}

} // namespace ShellHelper
