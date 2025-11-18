#include "MarkupPalette.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

#include "DGBrowser.hpp"
#include "BrowserRepl.hpp"

static GS::UniString LoadMarkupHtml()
{
	GS::UniString html;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), MarkupHtmlResId);
	if (data != nullptr) {
		const GSSize size = BMhGetSize(data);
		html.Append(*data, size);
		BMhKill(&data);
	} else {
		html = "<html><body style='font-family:Arial;color:#900;padding:12px;'>"
			"Не удалось загрузить HTML ресурса разметки.""</body></html>";
	}
	return html;
}

static GSErrCode __ACENV_CALL MarkupPaletteCallback(Int32 /*refCon*/, API_PaletteMessageID messageID, GS::IntPtr param)
{
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!MarkupPalette::HasInstance()) MarkupPalette::CreateInstance();
		MarkupPalette::GetInstance().Show();
		break;

	case APIPalMsg_ClosePalette:
		if (MarkupPalette::HasInstance()) MarkupPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_Begin:
		if (MarkupPalette::HasInstance() && MarkupPalette::GetInstance().IsVisible())
			MarkupPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_End:
		if (MarkupPalette::HasInstance() && !MarkupPalette::GetInstance().IsVisible())
			MarkupPalette::GetInstance().Show();
		break;

	case APIPalMsg_DisableItems_Begin:
		if (MarkupPalette::HasInstance() && MarkupPalette::GetInstance().IsVisible())
			MarkupPalette::GetInstance().DisableItems();
		break;

	case APIPalMsg_DisableItems_End:
		if (MarkupPalette::HasInstance() && MarkupPalette::GetInstance().IsVisible())
			MarkupPalette::GetInstance().EnableItems();
		break;

	case APIPalMsg_IsPaletteVisible:
		if (param != 0)
			*reinterpret_cast<bool*>(param) = MarkupPalette::HasInstance() && MarkupPalette::GetInstance().IsVisible();
		break;

	default:
		break;
	}
	return NoError;
}

GS::Ref<MarkupPalette> MarkupPalette::s_instance(nullptr);
const GS::Guid MarkupPalette::s_guid("{c1dc7b09-801d-4f30-b6f6-61ccd7e99833}");

MarkupPalette::MarkupPalette()
	: DG::Palette(ACAPI_GetOwnResModule(), MarkupPaletteResId, ACAPI_GetOwnResModule(), s_guid)
{
	m_browserCtrl = new DG::Browser(GetReference(), MarkupBrowserCtrlId);
	Attach(*this);
	BeginEventProcessing();
	Init();
}

MarkupPalette::~MarkupPalette()
{
	EndEventProcessing();
	delete m_browserCtrl;
	m_browserCtrl = nullptr;
}

void MarkupPalette::Init()
{
	LoadHtml();
	if (m_browserCtrl != nullptr)
		BrowserRepl::RegisterACAPIJavaScriptObject(*m_browserCtrl);
}

void MarkupPalette::LoadHtml()
{
	if (m_browserCtrl == nullptr)
		return;

	const GS::UniString html = LoadMarkupHtml();
	m_browserCtrl->LoadHTML(html);
}

bool MarkupPalette::HasInstance()
{
	return s_instance != nullptr;
}

void MarkupPalette::CreateInstance()
{
	DBASSERT(!HasInstance());
	s_instance = new MarkupPalette();
	ACAPI_KeepInMemory(true);
}

MarkupPalette& MarkupPalette::GetInstance()
{
	DBASSERT(HasInstance());
	return *s_instance;
}

void MarkupPalette::DestroyInstance()
{
	s_instance = nullptr;
}

void MarkupPalette::ShowPalette()
{
	if (!HasInstance())
		CreateInstance();

	GetInstance().Show();
}

void MarkupPalette::HidePalette()
{
	if (!HasInstance())
		return;

	GetInstance().Hide();
}

GSErrCode MarkupPalette::RegisterPaletteControlCallBack()
{
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(s_guid),
		MarkupPaletteCallback,
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

void MarkupPalette::PanelResized(const DG::PanelResizeEvent& ev)
{
	BeginMoveResizeItems();
	if (m_browserCtrl != nullptr)
		m_browserCtrl->Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void MarkupPalette::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
	HidePalette();
	if (accepted != nullptr)
		*accepted = true;
}





