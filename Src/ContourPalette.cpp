#include "ContourPalette.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

#include "DGBrowser.hpp"
#include "BrowserRepl.hpp"

static GS::UniString LoadContourHtml()
{
	GS::UniString html;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), ContourHtmlResId);
	if (data != nullptr) {
		const GSSize size = BMhGetSize(data);
		html.Append(*data, size);
		BMhKill(&data);
	} else {
		html = "<html><body style='font-family:Arial;color:#900;padding:12px;'>"
			"Не удалось загрузить HTML ресурса контуров.""</body></html>";
	}
	return html;
}

static GSErrCode __ACENV_CALL ContourPaletteCallback(Int32 /*refCon*/, API_PaletteMessageID messageID, GS::IntPtr param)
{
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!ContourPalette::HasInstance()) ContourPalette::CreateInstance();
		ContourPalette::GetInstance().Show();
		break;

	case APIPalMsg_ClosePalette:
		if (ContourPalette::HasInstance()) ContourPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_Begin:
		if (ContourPalette::HasInstance() && ContourPalette::GetInstance().IsVisible())
			ContourPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_End:
		if (ContourPalette::HasInstance() && !ContourPalette::GetInstance().IsVisible())
			ContourPalette::GetInstance().Show();
		break;

	case APIPalMsg_DisableItems_Begin:
		if (ContourPalette::HasInstance() && ContourPalette::GetInstance().IsVisible())
			ContourPalette::GetInstance().DisableItems();
		break;

	case APIPalMsg_DisableItems_End:
		if (ContourPalette::HasInstance() && ContourPalette::GetInstance().IsVisible())
			ContourPalette::GetInstance().EnableItems();
		break;

	case APIPalMsg_IsPaletteVisible:
		if (param != 0)
			*reinterpret_cast<bool*>(param) = ContourPalette::HasInstance() && ContourPalette::GetInstance().IsVisible();
		break;

	default:
		break;
	}
	return NoError;
}

GS::Ref<ContourPalette> ContourPalette::s_instance(nullptr);
const GS::Guid ContourPalette::s_guid("{ae9f8753-f1fe-4d6e-a0fb-7f0f64fa6a5c}");

ContourPalette::ContourPalette()
	: DG::Palette(ACAPI_GetOwnResModule(), ContourPaletteResId, ACAPI_GetOwnResModule(), s_guid)
{
	m_browserCtrl = new DG::Browser(GetReference(), ContourBrowserCtrlId);
	Attach(*this);
	BeginEventProcessing();
	Init();
}

ContourPalette::~ContourPalette()
{
	EndEventProcessing();
	delete m_browserCtrl;
	m_browserCtrl = nullptr;
}

void ContourPalette::Init()
{
	LoadHtml();
	if (m_browserCtrl != nullptr)
		BrowserRepl::RegisterACAPIJavaScriptObject(*m_browserCtrl);
}

void ContourPalette::LoadHtml()
{
	if (m_browserCtrl == nullptr)
		return;

	const GS::UniString html = LoadContourHtml();
	m_browserCtrl->LoadHTML(html);
}

bool ContourPalette::HasInstance()
{
	return s_instance != nullptr;
}

void ContourPalette::CreateInstance()
{
	DBASSERT(!HasInstance());
	s_instance = new ContourPalette();
	ACAPI_KeepInMemory(true);
}

ContourPalette& ContourPalette::GetInstance()
{
	DBASSERT(HasInstance());
	return *s_instance;
}

void ContourPalette::DestroyInstance()
{
	s_instance = nullptr;
}

void ContourPalette::ShowPalette()
{
	if (!HasInstance())
		CreateInstance();

	GetInstance().Show();
}

void ContourPalette::HidePalette()
{
	if (!HasInstance())
		return;

	GetInstance().Hide();
}

GSErrCode ContourPalette::RegisterPaletteControlCallBack()
{
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(s_guid),
		ContourPaletteCallback,
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

void ContourPalette::PanelResized(const DG::PanelResizeEvent& ev)
{
	BeginMoveResizeItems();
	if (m_browserCtrl != nullptr)
		m_browserCtrl->Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void ContourPalette::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
	HidePalette();
	if (accepted != nullptr)
		*accepted = true;
}





