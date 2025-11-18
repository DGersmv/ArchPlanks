#ifndef BROWSERREPL_HPP
#define BROWSERREPL_HPP

#include "DGModule.hpp"
#include "APIEnvir.h"
#include "ACAPinc.h"
#include "ResourceIDs.hpp"
#include "DGDefs.h"

// Forward declaration for JavaScript API registration (used by other palettes)
namespace DG { class Browser; }

// Класс BrowserRepl управляет палитрой с нативными кнопками
class BrowserRepl : public DG::Palette, 
                    public DG::PanelObserver,
                    public DG::ButtonItemObserver,
                    public DG::CompoundItemObserver {
public:
    BrowserRepl();
    ~BrowserRepl();

    static bool HasInstance();
    static void CreateInstance();
    static BrowserRepl& GetInstance();
    static void DestroyInstance();

    void Show();
    void Hide();

    static GSErrCode __ACENV_CALL PaletteControlCallBack(Int32 referenceID, API_PaletteMessageID messageID, GS::IntPtr param);
    static GSErrCode RegisterPaletteControlCallBack();
    
    // JavaScript API registration - still needed for other palettes
    static void RegisterACAPIJavaScriptObject(DG::Browser& targetBrowser);

private:
    static GS::Ref<BrowserRepl> instance;
    
    // Native icon buttons
    DG::IconButton buttonClose;
    DG::IconButton buttonTable;
    DG::IconButton buttonSpline;
    DG::IconButton buttonRotate;
    DG::IconButton buttonRotSurf;
    DG::IconButton buttonLand;
    DG::IconButton buttonDims;
    DG::IconButton buttonLayers;
    DG::IconButton buttonContour;
    DG::IconButton buttonMesh;
    DG::IconButton buttonCSV;
    DG::IconButton buttonSupport;

    void SetMenuItemCheckedState(bool isChecked);
    
    // Button click handlers
    void ButtonClicked(const DG::ButtonClickEvent& ev) override;

    // DG overrides
    void PanelResized(const DG::PanelResizeEvent& ev) override;
    void PanelCloseRequested(const DG::PanelCloseRequestEvent& ev, bool* accepted) override;
};

#endif // BROWSERREPL_HPP
