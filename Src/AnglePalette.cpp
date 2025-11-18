#include "AnglePalette.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

#include "DGBrowser.hpp"
#include "BrowserRepl.hpp"

static GS::UniString LoadAngleHtml()
{
	GS::UniString html;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), AngleHtmlResId);
	if (data != nullptr) {
		const GSSize size = BMhGetSize(data);
		html.Append(*data, size);
		BMhKill(&data);
	} else {
		html = "<html><body style='font-family:Arial;color:#900;padding:12px;'>"
			"Не удалось загрузить HTML ресурса палитры угла."
			"</body></html>";
	}
	return html;
}

static GSErrCode __ACENV_CALL AnglePaletteCallback(Int32 /*refCon*/, API_PaletteMessageID messageID, GS::IntPtr param)
{
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!AnglePalette::HasInstance()) AnglePalette::CreateInstance();
		AnglePalette::GetInstance().Show();
		break;

	case APIPalMsg_ClosePalette:
		if (AnglePalette::HasInstance()) AnglePalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_Begin:
		if (AnglePalette::HasInstance() && AnglePalette::GetInstance().IsVisible())
			AnglePalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_End:
		if (AnglePalette::HasInstance() && !AnglePalette::GetInstance().IsVisible())
			AnglePalette::GetInstance().Show();
		break;

	case APIPalMsg_DisableItems_Begin:
		if (AnglePalette::HasInstance() && AnglePalette::GetInstance().IsVisible())
			AnglePalette::GetInstance().DisableItems();
		break;

	case APIPalMsg_DisableItems_End:
		if (AnglePalette::HasInstance() && AnglePalette::GetInstance().IsVisible())
			AnglePalette::GetInstance().EnableItems();
		break;

	case APIPalMsg_IsPaletteVisible:
		if (param != 0)
			*reinterpret_cast<bool*>(param) = AnglePalette::HasInstance() && AnglePalette::GetInstance().IsVisible();
		break;

	default:
		break;
	}
	return NoError;
}

GS::Ref<AnglePalette> AnglePalette::s_instance(nullptr);
const GS::Guid AnglePalette::s_guid("{1399b4e6-2ab5-4fbf-99eb-cc783e0fb9d9}");

AnglePalette::AnglePalette()
	: DG::Palette(ACAPI_GetOwnResModule(), AnglePaletteResId, ACAPI_GetOwnResModule(), s_guid)
{
	m_browserCtrl = new DG::Browser(GetReference(), AngleBrowserCtrlId);
	Attach(*this);
	BeginEventProcessing();
	Init();
}

AnglePalette::~AnglePalette()
{
	EndEventProcessing();
	delete m_browserCtrl;
	m_browserCtrl = nullptr;
}

void AnglePalette::Init()
{
	LoadHtml();
	if (m_browserCtrl != nullptr)
		BrowserRepl::RegisterACAPIJavaScriptObject(*m_browserCtrl);
}

void AnglePalette::LoadHtml()
{
	if (m_browserCtrl == nullptr)
		return;

	const GS::UniString html = LoadAngleHtml();
	m_browserCtrl->LoadHTML(html);
}

bool AnglePalette::HasInstance()
{
	return s_instance != nullptr;
}

void AnglePalette::CreateInstance()
{
	DBASSERT(!HasInstance());
	s_instance = new AnglePalette();
	ACAPI_KeepInMemory(true);
}

AnglePalette& AnglePalette::GetInstance()
{
	DBASSERT(HasInstance());
	return *s_instance;
}

void AnglePalette::DestroyInstance()
{
	s_instance = nullptr;
}

void AnglePalette::ShowPalette()
{
	if (!HasInstance())
		CreateInstance();

	GetInstance().Show();
}

void AnglePalette::HidePalette()
{
	if (!HasInstance())
		return;

	GetInstance().Hide();
}

GSErrCode AnglePalette::RegisterPaletteControlCallBack()
{
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(s_guid),
		AnglePaletteCallback,
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

void AnglePalette::PanelResized(const DG::PanelResizeEvent& ev)
{
	BeginMoveResizeItems();
	if (m_browserCtrl != nullptr)
		m_browserCtrl->Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void AnglePalette::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
	HidePalette();
	if (accepted != nullptr)
		*accepted = true;
}





