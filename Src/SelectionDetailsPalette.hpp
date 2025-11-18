#pragma once

#include "GSRoot.hpp"
#include "GSGuid.hpp"
#include "UniString.hpp"
#include "DGModule.hpp"
#include "ResourceIDs.hpp"

namespace DG { class Browser; }

// Forward declaration
struct API_Neig;

class SelectionDetailsPalette : public DG::Palette, public DG::PanelObserver
{
public:
	static bool         HasInstance();
	static void         CreateInstance();
	static SelectionDetailsPalette& GetInstance();
	static void         DestroyInstance();

	static void         ShowPalette();
	static void         HidePalette();
	static void         UpdateSelectedElementsOnHTML();
	static GSErrCode    RegisterPaletteControlCallBack();
	static GSErrCode    SelectionChangeHandler(const API_Neig* neig);

	virtual ~SelectionDetailsPalette();

private:
	SelectionDetailsPalette();

	void                Init();
	void                LoadHtml();

	void PanelResized(const DG::PanelResizeEvent& ev) override;
	void PanelCloseRequested(const DG::PanelCloseRequestEvent& ev, bool* accepted) override;

private:
	static GS::Ref<SelectionDetailsPalette> s_instance;
	static const GS::Guid                   s_guid;

	DG::Browser* m_browserCtrl = nullptr;
};

