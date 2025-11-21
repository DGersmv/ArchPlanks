#pragma once

#include "GSRoot.hpp"
#include "GSGuid.hpp"
#include "UniString.hpp"
#include "DGModule.hpp"
#include "ResourceIDs.hpp"

namespace DG { class Browser; }

class AnglePalette : public DG::Palette, public DG::PanelObserver
{
public:
    static bool           HasInstance();
    static void           CreateInstance();
    static AnglePalette&  GetInstance();
    static void           DestroyInstance();

    static void           ShowPalette();
    static void           HidePalette();
    static GSErrCode      RegisterPaletteControlCallBack();

    virtual ~AnglePalette();

private:
    AnglePalette();

    void                  Init();
    void                  LoadHtml();

    void PanelResized(const DG::PanelResizeEvent& ev) override;
    void PanelCloseRequested(const DG::PanelCloseRequestEvent& ev, bool* accepted) override;

private:
    static GS::Ref<AnglePalette> s_instance;
    static const GS::Guid        s_guid;

    DG::Browser* m_browserCtrl = nullptr;
};





