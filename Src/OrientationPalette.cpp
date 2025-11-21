#include "OrientationPalette.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

#include "DGBrowser.hpp"
#include "BrowserRepl.hpp"

static GS::UniString LoadOrientationHtml()
{
	GS::UniString html;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), OrientationHtmlResId);
	if (data != nullptr) {
		const GSSize size = BMhGetSize(data);
		html.Append(*data, size);
		BMhKill(&data);
	} else {
		html = "<html><body style='font-family:Arial;color:#900;padding:12px;'>"
			"Не удалось загрузить HTML ресурса ориентации.""</body></html>";
	}
	return html;
}

static GSErrCode __ACENV_CALL OrientationPaletteCallback(Int32 /*refCon*/, API_PaletteMessageID messageID, GS::IntPtr param)
{
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!OrientationPalette::HasInstance()) OrientationPalette::CreateInstance();
		OrientationPalette::GetInstance().Show();
		break;

	case APIPalMsg_ClosePalette:
		if (OrientationPalette::HasInstance()) OrientationPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_Begin:
		if (OrientationPalette::HasInstance() && OrientationPalette::GetInstance().IsVisible())
			OrientationPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_End:
		if (OrientationPalette::HasInstance() && !OrientationPalette::GetInstance().IsVisible())
			OrientationPalette::GetInstance().Show();
		break;

	case APIPalMsg_DisableItems_Begin:
		if (OrientationPalette::HasInstance() && OrientationPalette::GetInstance().IsVisible())
			OrientationPalette::GetInstance().DisableItems();
		break;

	case APIPalMsg_DisableItems_End:
		if (OrientationPalette::HasInstance() && OrientationPalette::GetInstance().IsVisible())
			OrientationPalette::GetInstance().EnableItems();
		break;

	case APIPalMsg_IsPaletteVisible:
		if (param != 0)
			*reinterpret_cast<bool*>(param) = OrientationPalette::HasInstance() && OrientationPalette::GetInstance().IsVisible();
		break;

	default:
		break;
	}
	return NoError;
}

GS::Ref<OrientationPalette> OrientationPalette::s_instance(nullptr);
const GS::Guid OrientationPalette::s_guid("{4cf8d473-3a3d-4b64-9f7c-5e3434a7a7c1}");

OrientationPalette::OrientationPalette()
	: DG::Palette(ACAPI_GetOwnResModule(), OrientationPaletteResId, ACAPI_GetOwnResModule(), s_guid)
{
	m_browserCtrl = new DG::Browser(GetReference(), OrientationBrowserCtrlId);
	Attach(*this);
	BeginEventProcessing();
	Init();
}

OrientationPalette::~OrientationPalette()
{
	EndEventProcessing();
	delete m_browserCtrl;
	m_browserCtrl = nullptr;
}

void OrientationPalette::Init()
{
	LoadHtml();
	if (m_browserCtrl != nullptr)
		BrowserRepl::RegisterACAPIJavaScriptObject(*m_browserCtrl);
}

void OrientationPalette::LoadHtml()
{
	if (m_browserCtrl == nullptr)
		return;

	const GS::UniString html = LoadOrientationHtml();
	m_browserCtrl->LoadHTML(html);
}

bool OrientationPalette::HasInstance()
{
	return s_instance != nullptr;
}

void OrientationPalette::CreateInstance()
{
	DBASSERT(!HasInstance());
	s_instance = new OrientationPalette();
	ACAPI_KeepInMemory(true);
}

OrientationPalette& OrientationPalette::GetInstance()
{
	DBASSERT(HasInstance());
	return *s_instance;
}

void OrientationPalette::DestroyInstance()
{
	s_instance = nullptr;
}

void OrientationPalette::ShowPalette()
{
	if (!HasInstance())
		CreateInstance();

	GetInstance().Show();
}

void OrientationPalette::HidePalette()
{
	if (!HasInstance())
		return;

	GetInstance().Hide();
}

GSErrCode OrientationPalette::RegisterPaletteControlCallBack()
{
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(s_guid),
		OrientationPaletteCallback,
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

void OrientationPalette::PanelResized(const DG::PanelResizeEvent& ev)
{
	BeginMoveResizeItems();
	if (m_browserCtrl != nullptr)
		m_browserCtrl->Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void OrientationPalette::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
	HidePalette();
	if (accepted != nullptr)
		*accepted = true;
}





