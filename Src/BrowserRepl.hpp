#ifndef BROWSERREPL_HPP
#define BROWSERREPL_HPP

#include "DGModule.hpp"
#include "APIEnvir.h"
#include "ACAPinc.h"
#include "ResourceIDs.hpp"
#include "DGDefs.h"

namespace DG { class Browser; }

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

    static GSErrCode PaletteControlCallBack(Int32 referenceID, API_PaletteMessageID messageID, GS::IntPtr param);
    static GSErrCode RegisterPaletteControlCallBack();
    static void RegisterACAPIJavaScriptObject(DG::Browser& targetBrowser);

private:
    static GS::Ref<BrowserRepl> instance;
    
    DG::IconButton buttonClose;
    DG::IconButton buttonTable;
    DG::IconButton buttonSupport;

    void SetMenuItemCheckedState(bool isChecked);
    void ButtonClicked(const DG::ButtonClickEvent& ev) override;
    void PanelResized(const DG::PanelResizeEvent& ev) override;
    void PanelCloseRequested(const DG::PanelCloseRequestEvent& ev, bool* accepted) override;
};

#endif
