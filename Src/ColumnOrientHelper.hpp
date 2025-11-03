#ifndef COLUMNORIENTHELPER_HPP
#define COLUMNORIENTHELPER_HPP

#include "API_Guid.hpp"

// ============================================================================
// ColumnOrientHelper — ориентация балок по поверхности mesh
// ============================================================================
class ColumnOrientHelper {
public:
    // Установить балки для ориентации (из выделения)
    static bool SetBeams();
    
    // Установить mesh для ориентации (из выделения)
    static bool SetMesh();
    
    // Ориентировать балки по поверхности mesh
    static bool OrientBeamsToSurface();
    
    // Повернуть выделенные балки на заданный угол (в градусах)
    // Поворачивает profileAngle балок
    static bool RotateSelected(double angleDeg);
};

#endif // COLUMNORIENTHELPER_HPP


