#include "GroundPalette.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

#include "DGBrowser.hpp"
#include "BrowserRepl.hpp"

static GS::UniString LoadGroundHtml()
{
	GS::UniString html;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), GroundHtmlResId);
	if (data != nullptr) {
		const GSSize size = BMhGetSize(data);
		html.Append(*data, size);
		BMhKill(&data);
	} else {
		html = "<html><body style='font-family:Arial;color:#900;padding:12px;'>"
			"Не удалось загрузить HTML ресурса привязки.""</body></html>";
	}
	return html;
}

static GSErrCode __ACENV_CALL GroundPaletteCallback(Int32 /*refCon*/, API_PaletteMessageID messageID, GS::IntPtr param)
{
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!GroundPalette::HasInstance()) GroundPalette::CreateInstance();
		GroundPalette::GetInstance().Show();
		break;

	case APIPalMsg_ClosePalette:
		if (GroundPalette::HasInstance()) GroundPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_Begin:
		if (GroundPalette::HasInstance() && GroundPalette::GetInstance().IsVisible())
			GroundPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_End:
		if (GroundPalette::HasInstance() && !GroundPalette::GetInstance().IsVisible())
			GroundPalette::GetInstance().Show();
		break;

	case APIPalMsg_DisableItems_Begin:
		if (GroundPalette::HasInstance() && GroundPalette::GetInstance().IsVisible())
			GroundPalette::GetInstance().DisableItems();
		break;

	case APIPalMsg_DisableItems_End:
		if (GroundPalette::HasInstance() && GroundPalette::GetInstance().IsVisible())
			GroundPalette::GetInstance().EnableItems();
		break;

	case APIPalMsg_IsPaletteVisible:
		if (param != 0)
			*reinterpret_cast<bool*>(param) = GroundPalette::HasInstance() && GroundPalette::GetInstance().IsVisible();
		break;

	default:
		break;
	}
	return NoError;
}

GS::Ref<GroundPalette> GroundPalette::s_instance(nullptr);
const GS::Guid GroundPalette::s_guid("{d1e5af33-4b51-4d8f-bcf0-f3baf749f2c4}");

GroundPalette::GroundPalette()
	: DG::Palette(ACAPI_GetOwnResModule(), GroundPaletteResId, ACAPI_GetOwnResModule(), s_guid)
{
	m_browserCtrl = new DG::Browser(GetReference(), GroundBrowserCtrlId);
	Attach(*this);
	BeginEventProcessing();
	Init();
}

GroundPalette::~GroundPalette()
{
	EndEventProcessing();
	delete m_browserCtrl;
	m_browserCtrl = nullptr;
}

void GroundPalette::Init()
{
	LoadHtml();
	if (m_browserCtrl != nullptr)
		BrowserRepl::RegisterACAPIJavaScriptObject(*m_browserCtrl);
}

void GroundPalette::LoadHtml()
{
	if (m_browserCtrl == nullptr)
		return;

	const GS::UniString html = LoadGroundHtml();
	m_browserCtrl->LoadHTML(html);
}

bool GroundPalette::HasInstance()
{
	return s_instance != nullptr;
}

void GroundPalette::CreateInstance()
{
	DBASSERT(!HasInstance());
	s_instance = new GroundPalette();
	ACAPI_KeepInMemory(true);
}

GroundPalette& GroundPalette::GetInstance()
{
	DBASSERT(HasInstance());
	return *s_instance;
}

void GroundPalette::DestroyInstance()
{
	s_instance = nullptr;
}

void GroundPalette::ShowPalette()
{
	if (!HasInstance())
		CreateInstance();

	GetInstance().Show();
}

void GroundPalette::HidePalette()
{
	if (!HasInstance())
		return;

	GetInstance().Hide();
}

GSErrCode GroundPalette::RegisterPaletteControlCallBack()
{
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(s_guid),
		GroundPaletteCallback,
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

void GroundPalette::PanelResized(const DG::PanelResizeEvent& ev)
{
	BeginMoveResizeItems();
	if (m_browserCtrl != nullptr)
		m_browserCtrl->Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void GroundPalette::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
	HidePalette();
	if (accepted != nullptr)
		*accepted = true;
}





