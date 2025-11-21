#include "SelectionDetailsPalette.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

#include "DGBrowser.hpp"
#include "BrowserRepl.hpp"

// -------------------- local helpers --------------------
static GS::UniString LoadSelectionDetailsHtml()
{
	GS::UniString html;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), SelectionDetailsHtmlResId);
	if (data != nullptr) {
		const GSSize size = BMhGetSize(data);
		html.Append(*data, size);
		BMhKill(&data);
	} else {
		html = "<html><body style='font-family:Arial;color:#900;padding:12px;'>"
			"Не удалось загрузить HTML ресурса выбранных элементов.""</body></html>";
	}
	return html;
}

static GSErrCode __ACENV_CALL SelectionDetailsPaletteCallback(Int32 /*refCon*/, API_PaletteMessageID messageID, GS::IntPtr param)
{
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!SelectionDetailsPalette::HasInstance()) SelectionDetailsPalette::CreateInstance();
		SelectionDetailsPalette::GetInstance().Show();
		break;

	case APIPalMsg_ClosePalette:
		if (SelectionDetailsPalette::HasInstance()) SelectionDetailsPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_Begin:
		if (SelectionDetailsPalette::HasInstance() && SelectionDetailsPalette::GetInstance().IsVisible())
			SelectionDetailsPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_End:
		if (SelectionDetailsPalette::HasInstance() && !SelectionDetailsPalette::GetInstance().IsVisible())
			SelectionDetailsPalette::GetInstance().Show();
		break;

	case APIPalMsg_DisableItems_Begin:
		if (SelectionDetailsPalette::HasInstance() && SelectionDetailsPalette::GetInstance().IsVisible())
			SelectionDetailsPalette::GetInstance().DisableItems();
		break;

	case APIPalMsg_DisableItems_End:
		if (SelectionDetailsPalette::HasInstance() && SelectionDetailsPalette::GetInstance().IsVisible())
			SelectionDetailsPalette::GetInstance().EnableItems();
		break;

	case APIPalMsg_IsPaletteVisible:
		if (param != 0)
			*reinterpret_cast<bool*>(param) = SelectionDetailsPalette::HasInstance() && SelectionDetailsPalette::GetInstance().IsVisible();
		break;

	default:
		break;
	}
	return NoError;
}

// -------------------- static members --------------------
GS::Ref<SelectionDetailsPalette> SelectionDetailsPalette::s_instance(nullptr);
const GS::Guid SelectionDetailsPalette::s_guid("{a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d}");

// -------------------- lifecycle --------------------
SelectionDetailsPalette::SelectionDetailsPalette()
	: DG::Palette(ACAPI_GetOwnResModule(), SelectionDetailsPaletteResId, ACAPI_GetOwnResModule(), s_guid)
{
	m_browserCtrl = new DG::Browser(GetReference(), SelectionDetailsBrowserCtrlId);
	Attach(*this);
	BeginEventProcessing();
	
	// Подпишемся на изменение выделения
	ACAPI_Notification_CatchSelectionChange(SelectionChangeHandler);
	
	Init();
}

SelectionDetailsPalette::~SelectionDetailsPalette()
{
	EndEventProcessing();
	delete m_browserCtrl;
	m_browserCtrl = nullptr;
}

void SelectionDetailsPalette::Init()
{
	LoadHtml();
	if (m_browserCtrl != nullptr)
		BrowserRepl::RegisterACAPIJavaScriptObject(*m_browserCtrl);
}

void SelectionDetailsPalette::LoadHtml()
{
	if (m_browserCtrl == nullptr)
		return;

	const GS::UniString html = LoadSelectionDetailsHtml();
	m_browserCtrl->LoadHTML(html);
}

// -------------------- singleton helpers --------------------
bool SelectionDetailsPalette::HasInstance()
{
	return s_instance != nullptr;
}

void SelectionDetailsPalette::CreateInstance()
{
	DBASSERT(!HasInstance());
	s_instance = new SelectionDetailsPalette();
	ACAPI_KeepInMemory(true);
}

SelectionDetailsPalette& SelectionDetailsPalette::GetInstance()
{
	DBASSERT(HasInstance());
	return *s_instance;
}

void SelectionDetailsPalette::DestroyInstance()
{
	s_instance = nullptr;
}

// -------------------- public API --------------------
void SelectionDetailsPalette::ShowPalette()
{
	if (!HasInstance())
		CreateInstance();

	GetInstance().Show();
}

void SelectionDetailsPalette::HidePalette()
{
	if (!HasInstance())
		return;

	GetInstance().Hide();
}

void SelectionDetailsPalette::UpdateSelectedElementsOnHTML()
{
	if (!HasInstance() || !GetInstance().IsVisible())
		return;

	if (GetInstance().m_browserCtrl != nullptr)
		GetInstance().m_browserCtrl->ExecuteJS("UpdateSelectedElements()");
}

GSErrCode SelectionDetailsPalette::RegisterPaletteControlCallBack()
{
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(s_guid),
		SelectionDetailsPaletteCallback,
		API_PalEnabled_FloorPlan |
		API_PalEnabled_Section |
		API_PalEnabled_Elevation |
		API_PalEnabled_InteriorElevation |
		API_PalEnabled_3D |
		API_PalEnabled_Detail |
		API_PalEnabled_Worksheet |
		API_PalEnabled_Layout |
		API_PalEnabled_DocumentFrom3D,
		GSGuid2APIGuid(s_guid)
	);
}

// -------------------- DG::PanelObserver --------------------
void SelectionDetailsPalette::PanelResized(const DG::PanelResizeEvent& ev)
{
	BeginMoveResizeItems();
	if (m_browserCtrl != nullptr)
		m_browserCtrl->Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void SelectionDetailsPalette::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
	HidePalette();
	if (accepted != nullptr)
		*accepted = true;
}

// -------------------- Selection Change Handler --------------------
GSErrCode __ACENV_CALL SelectionDetailsPalette::SelectionChangeHandler(const API_Neig* neig)
{
	(void)neig; // unused parameter
	if (SelectionDetailsPalette::HasInstance())
		SelectionDetailsPalette::UpdateSelectedElementsOnHTML();
	return NoError;
}

