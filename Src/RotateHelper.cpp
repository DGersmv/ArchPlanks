#include "RotateHelper.hpp"
#include <random>
#include <cmath>

constexpr double PI = 3.14159265358979323846;
constexpr double DegToRad(double deg) { return deg * PI / 180.0; }

namespace RotateHelper {

// ---------------- Поворот ----------------
bool RotateSelected (double angleDeg)
{
    if (fabs(angleDeg) < 1e-6) return false;

    API_SelectionInfo selInfo = {};
    GS::Array<API_Neig> selNeigs;
    ACAPI_Selection_Get(&selInfo, &selNeigs, false, false);
    BMKillHandle((GSHandle*)&selInfo.marquee.coords);

    if (selNeigs.IsEmpty()) return false;

    const double addRad = DegToRad(angleDeg);

    GSErrCode err = ACAPI_CallUndoableCommand("Rotate Selected", [&]() -> GSErrCode {
        for (const API_Neig& n : selNeigs) {
            API_Element element = {};
            element.header.guid = n.guid;
            if (ACAPI_Element_Get(&element) != NoError) continue;

            API_Element mask; ACAPI_ELEMENT_MASK_CLEAR(mask);

            API_ElementMemo memo{};
            bool needsMemo = false;
            bool hasMemo = false;
            
            switch (element.header.type.typeID) {
            case API_ColumnID:
                element.column.axisRotationAngle += addRad;
                ACAPI_ELEMENT_MASK_SET(mask, API_ColumnType, axisRotationAngle);
                break;
            case API_ObjectID:
                element.object.angle += addRad;
                ACAPI_ELEMENT_MASK_SET(mask, API_ObjectType, angle);
                break;
            case API_LampID:
                element.lamp.angle += addRad;
                ACAPI_ELEMENT_MASK_SET(mask, API_LampType, angle);
                break;
            case API_BeamID: {
                // Поворот балки в плоскости XY вокруг центра балки (поворот вокруг оси Z)
                const double dx = element.beam.endC.x - element.beam.begC.x;
                const double dy = element.beam.endC.y - element.beam.begC.y;
                const double beamLength = std::hypot(dx, dy);
                const double halfLength = beamLength / 2.0;
                const double currentAngle = std::atan2(dy, dx);
                const double newAngle = currentAngle + addRad;
                
                // Центр балки
                const double centerX = (element.beam.begC.x + element.beam.endC.x) / 2.0;
                const double centerY = (element.beam.begC.y + element.beam.endC.y) / 2.0;
                
                // Поворот обеих точек вокруг центра
                element.beam.begC.x = centerX - halfLength * std::cos(newAngle);
                element.beam.begC.y = centerY - halfLength * std::sin(newAngle);
                element.beam.endC.x = centerX + halfLength * std::cos(newAngle);
                element.beam.endC.y = centerY + halfLength * std::sin(newAngle);
                
                ACAPI_ELEMENT_MASK_SET(mask, API_BeamType, begC);
                ACAPI_ELEMENT_MASK_SET(mask, API_BeamType, endC);
                needsMemo = true;
                hasMemo = (ACAPI_Element_GetMemo(n.guid, &memo, APIMemoMask_All) == NoError);
                break;
            }
            default:
                continue;
            }
            
            if (needsMemo && hasMemo) {
                (void)ACAPI_Element_Change(&element, &mask, &memo, 0, true);
                ACAPI_DisposeElemMemoHdls(&memo);
            } else {
                (void)ACAPI_Element_Change(&element, &mask, nullptr, 0, true);
            }
        }
        return NoError;
    });

    return err == NoError;
}

// ---------------- Выравнивание по X ----------------
bool AlignSelectedX ()
{
    API_SelectionInfo selInfo = {};
    GS::Array<API_Neig> selNeigs;
    ACAPI_Selection_Get(&selInfo, &selNeigs, false, false);
    BMKillHandle((GSHandle*)&selInfo.marquee.coords);

    if (selNeigs.IsEmpty()) return false;

    GSErrCode err = ACAPI_CallUndoableCommand("Align to X", [&]() -> GSErrCode {
        for (const API_Neig& n : selNeigs) {
            API_Element element = {};
            element.header.guid = n.guid;
            if (ACAPI_Element_Get(&element) != NoError) continue;

            API_Element mask; ACAPI_ELEMENT_MASK_CLEAR(mask);

            API_ElementMemo memo{};
            bool needsMemo = false;
            bool hasMemo = false;
            
            switch (element.header.type.typeID) {
            case API_ColumnID: 
                element.column.axisRotationAngle = 0.0; 
                ACAPI_ELEMENT_MASK_SET(mask, API_ColumnType, axisRotationAngle); 
                break;
            case API_ObjectID: 
                element.object.angle = 0.0; 
                ACAPI_ELEMENT_MASK_SET(mask, API_ObjectType, angle); 
                break;
            case API_LampID:   
                element.lamp.angle = 0.0;   
                ACAPI_ELEMENT_MASK_SET(mask, API_LampType, angle); 
                break;
            case API_BeamID: {
                // Выравнивание балки по X вокруг центра балки - поворачиваем в плоскости XY
                const double dx = element.beam.endC.x - element.beam.begC.x;
                const double dy = element.beam.endC.y - element.beam.begC.y;
                const double beamLength = std::hypot(dx, dy);
                const double halfLength = beamLength / 2.0;
                
                // Центр балки
                const double centerX = (element.beam.begC.x + element.beam.endC.x) / 2.0;
                const double centerY = (element.beam.begC.y + element.beam.endC.y) / 2.0;
                
                // Выравниваем балку по X вокруг центра
                element.beam.begC.x = centerX - halfLength; // Вдоль оси X
                element.beam.begC.y = centerY;              // Y не меняется
                element.beam.endC.x = centerX + halfLength;
                element.beam.endC.y = centerY;
                
                ACAPI_ELEMENT_MASK_SET(mask, API_BeamType, begC);
                ACAPI_ELEMENT_MASK_SET(mask, API_BeamType, endC);
                needsMemo = true;
                hasMemo = (ACAPI_Element_GetMemo(n.guid, &memo, APIMemoMask_All) == NoError);
                break;
            }
            default: 
                continue;
            }
            
            if (needsMemo && hasMemo) {
                (void)ACAPI_Element_Change(&element, &mask, &memo, 0, true);
                ACAPI_DisposeElemMemoHdls(&memo);
            } else {
                (void)ACAPI_Element_Change(&element, &mask, nullptr, 0, true);
            }
        }
        return NoError;
    });

    return err == NoError;
}

// ---------------- Случайные углы ----------------
bool RandomizeSelectedAngles ()
{
    API_SelectionInfo selInfo = {};
    GS::Array<API_Neig> selNeigs;
    ACAPI_Selection_Get(&selInfo, &selNeigs, false, false);
    BMKillHandle((GSHandle*)&selInfo.marquee.coords);

    if (selNeigs.IsEmpty()) return false;

    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_real_distribution<double> dist(0.0, 2.0 * PI);

    GSErrCode err = ACAPI_CallUndoableCommand("Randomize angles", [&]() -> GSErrCode {
        for (const API_Neig& n : selNeigs) {
            API_Element element = {};
            element.header.guid = n.guid;
            if (ACAPI_Element_Get(&element) != NoError) continue;

            const double rnd = dist(gen);

            API_Element mask; ACAPI_ELEMENT_MASK_CLEAR(mask);

            API_ElementMemo memo{};
            bool needsMemo = false;
            bool hasMemo = false;
            
            switch (element.header.type.typeID) {
            case API_ObjectID: 
                element.object.angle = rnd; 
                ACAPI_ELEMENT_MASK_SET(mask, API_ObjectType, angle); 
                break;
            case API_LampID:   
                element.lamp.angle = rnd;   
                ACAPI_ELEMENT_MASK_SET(mask, API_LampType, angle); 
                break;
            case API_ColumnID: 
                element.column.axisRotationAngle = rnd; 
                ACAPI_ELEMENT_MASK_SET(mask, API_ColumnType, axisRotationAngle); 
                break;
            case API_BeamID: {
                // Случайный поворот балки в плоскости XY вокруг центра (поворот вокруг оси Z)
                const double dx = element.beam.endC.x - element.beam.begC.x;
                const double dy = element.beam.endC.y - element.beam.begC.y;
                const double beamLength = std::hypot(dx, dy);
                const double halfLength = beamLength / 2.0;
                
                // Центр балки
                const double centerX = (element.beam.begC.x + element.beam.endC.x) / 2.0;
                const double centerY = (element.beam.begC.y + element.beam.endC.y) / 2.0;
                
                // Случайный поворот вокруг центра
                element.beam.begC.x = centerX - halfLength * std::cos(rnd);
                element.beam.begC.y = centerY - halfLength * std::sin(rnd);
                element.beam.endC.x = centerX + halfLength * std::cos(rnd);
                element.beam.endC.y = centerY + halfLength * std::sin(rnd);
                
                ACAPI_ELEMENT_MASK_SET(mask, API_BeamType, begC);
                ACAPI_ELEMENT_MASK_SET(mask, API_BeamType, endC);
                needsMemo = true;
                hasMemo = (ACAPI_Element_GetMemo(n.guid, &memo, APIMemoMask_All) == NoError);
                break;
            }
            default: 
                continue;
            }
            
            if (needsMemo && hasMemo) {
                (void)ACAPI_Element_Change(&element, &mask, &memo, 0, true);
                ACAPI_DisposeElemMemoHdls(&memo);
            } else {
                (void)ACAPI_Element_Change(&element, &mask, nullptr, 0, true);
            }
        }
        return NoError;
    });

    return err == NoError;
}

// ---------------- Ориентация на точку ----------------
bool OrientObjectsToPoint ()
{
    API_GetPointType pt = {};
    CHTruncate("Укажите точку для ориентации объектов", pt.prompt, sizeof(pt.prompt));
    if (ACAPI_UserInput_GetPoint(&pt) != NoError) return false;
    const API_Coord target = { pt.pos.x, pt.pos.y };

    API_SelectionInfo selInfo = {};
    GS::Array<API_Neig> selNeigs;
    ACAPI_Selection_Get(&selInfo, &selNeigs, false, false);
    BMKillHandle((GSHandle*)&selInfo.marquee.coords);

    if (selNeigs.IsEmpty()) return false;

    GSErrCode err = ACAPI_CallUndoableCommand("Orient Objects to Point", [&]() -> GSErrCode {
        for (const API_Neig& n : selNeigs) {
            API_Element element = {};
            element.header.guid = n.guid;
            if (ACAPI_Element_Get(&element) != NoError) continue;

            API_Coord objPos = {};
            switch (element.header.type.typeID) {
            case API_ObjectID: objPos = element.object.pos; break;
            case API_LampID:   objPos = element.lamp.pos;   break;
            case API_ColumnID: objPos.x = element.column.origoPos.x; objPos.y = element.column.origoPos.y; break;
            case API_BeamID:   objPos.x = (element.beam.begC.x + element.beam.endC.x) / 2.0; objPos.y = (element.beam.begC.y + element.beam.endC.y) / 2.0; break;  // Используем центр балки
            default: continue;
            }

            const double dx = target.x - objPos.x;
            const double dy = target.y - objPos.y;
            const double newAngle = std::atan2(dy, dx);

            API_Element mask; ACAPI_ELEMENT_MASK_CLEAR(mask);
            API_ElementMemo memo{};
            bool needsMemo = false;
            bool hasMemo = false;
            
            switch (element.header.type.typeID) {
            case API_ObjectID: 
                element.object.angle = newAngle; 
                ACAPI_ELEMENT_MASK_SET(mask, API_ObjectType, angle); 
                break;
            case API_LampID:   
                element.lamp.angle = newAngle;   
                ACAPI_ELEMENT_MASK_SET(mask, API_LampType, angle);   
                break;
            case API_ColumnID: 
                element.column.axisRotationAngle = newAngle; 
                ACAPI_ELEMENT_MASK_SET(mask, API_ColumnType, axisRotationAngle); 
                break;
            case API_BeamID: {
                // Ориентация балки в плоскости XY вокруг центра балки - направляем к целевой точке
                const double beamDx = element.beam.endC.x - element.beam.begC.x;
                const double beamDy = element.beam.endC.y - element.beam.begC.y;
                const double beamLength = std::hypot(beamDx, beamDy);
                const double halfLength = beamLength / 2.0;
                
                // Центр балки
                const double centerX = (element.beam.begC.x + element.beam.endC.x) / 2.0;
                const double centerY = (element.beam.begC.y + element.beam.endC.y) / 2.0;
                
                // Ориентируем балку вокруг центра к целевой точке
                element.beam.begC.x = centerX - halfLength * std::cos(newAngle);
                element.beam.begC.y = centerY - halfLength * std::sin(newAngle);
                element.beam.endC.x = centerX + halfLength * std::cos(newAngle);
                element.beam.endC.y = centerY + halfLength * std::sin(newAngle);
                
                ACAPI_ELEMENT_MASK_SET(mask, API_BeamType, begC);
                ACAPI_ELEMENT_MASK_SET(mask, API_BeamType, endC);
                needsMemo = true;
                hasMemo = (ACAPI_Element_GetMemo(n.guid, &memo, APIMemoMask_All) == NoError);
                break;
            }
            default: 
                continue;
            }
            
            if (needsMemo && hasMemo) {
                (void)ACAPI_Element_Change(&element, &mask, &memo, 0, true);
                ACAPI_DisposeElemMemoHdls(&memo);
            } else {
                (void)ACAPI_Element_Change(&element, &mask, nullptr, 0, true);
            }
        }
        return NoError;
    });

    return err == NoError;
}

} // namespace RotateHelper
