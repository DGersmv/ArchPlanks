#include "SendXlsPalette.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

#include "DGBrowser.hpp"
#include "BrowserRepl.hpp"

static GS::UniString LoadSendXlsHtml()
{
	GS::UniString html;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), SendXlsHtmlResId);
	if (data != nullptr) {
		const GSSize size = BMhGetSize(data);
		html.Append(*data, size);
		BMhKill(&data);
	} else {
		html = "<html><body style='font-family:Arial;color:#900;padding:12px;'>"
			"Не удалось загрузить HTML ресурса выгрузки в Excel."
			"</body></html>";
	}
	return html;
}

static GSErrCode __ACENV_CALL SendXlsPaletteCallback(Int32 /*refCon*/, API_PaletteMessageID messageID, GS::IntPtr param)
{
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!SendXlsPalette::HasInstance()) SendXlsPalette::CreateInstance();
		SendXlsPalette::GetInstance().Show();
		break;

	case APIPalMsg_ClosePalette:
		if (SendXlsPalette::HasInstance()) SendXlsPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_Begin:
		if (SendXlsPalette::HasInstance() && SendXlsPalette::GetInstance().IsVisible())
			SendXlsPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_End:
		if (SendXlsPalette::HasInstance() && !SendXlsPalette::GetInstance().IsVisible())
			SendXlsPalette::GetInstance().Show();
		break;

	case APIPalMsg_DisableItems_Begin:
		if (SendXlsPalette::HasInstance() && SendXlsPalette::GetInstance().IsVisible())
			SendXlsPalette::GetInstance().DisableItems();
		break;

	case APIPalMsg_DisableItems_End:
		if (SendXlsPalette::HasInstance() && SendXlsPalette::GetInstance().IsVisible())
			SendXlsPalette::GetInstance().EnableItems();
		break;

	case APIPalMsg_IsPaletteVisible:
		if (param != 0)
			*reinterpret_cast<bool*>(param) = SendXlsPalette::HasInstance() && SendXlsPalette::GetInstance().IsVisible();
		break;

	default:
		break;
	}
	return NoError;
}

GS::Ref<SendXlsPalette> SendXlsPalette::s_instance(nullptr);
const GS::Guid SendXlsPalette::s_guid("{4f40335e-7f3e-4891-90f2-4be2d146ba6e}");

SendXlsPalette::SendXlsPalette()
	: DG::Palette(ACAPI_GetOwnResModule(), SendXlsPaletteResId, ACAPI_GetOwnResModule(), s_guid)
{
	m_browserCtrl = new DG::Browser(GetReference(), SendXlsBrowserCtrlId);
	Attach(*this);
	BeginEventProcessing();
	Init();
}

SendXlsPalette::~SendXlsPalette()
{
	EndEventProcessing();
	delete m_browserCtrl;
	m_browserCtrl = nullptr;
}

void SendXlsPalette::Init()
{
	LoadHtml();
	if (m_browserCtrl != nullptr)
		BrowserRepl::RegisterACAPIJavaScriptObject(*m_browserCtrl);
}

void SendXlsPalette::LoadHtml()
{
	if (m_browserCtrl == nullptr)
		return;

	const GS::UniString html = LoadSendXlsHtml();
	m_browserCtrl->LoadHTML(html);
}

bool SendXlsPalette::HasInstance()
{
	return s_instance != nullptr;
}

void SendXlsPalette::CreateInstance()
{
	DBASSERT(!HasInstance());
	s_instance = new SendXlsPalette();
	ACAPI_KeepInMemory(true);
}

SendXlsPalette& SendXlsPalette::GetInstance()
{
	DBASSERT(HasInstance());
	return *s_instance;
}

void SendXlsPalette::DestroyInstance()
{
	s_instance = nullptr;
}

void SendXlsPalette::ShowPalette()
{
	if (!HasInstance())
		CreateInstance();

	GetInstance().Show();
}

void SendXlsPalette::HidePalette()
{
	if (!HasInstance())
		return;

	GetInstance().Hide();
}

GSErrCode SendXlsPalette::RegisterPaletteControlCallBack()
{
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(s_guid),
		SendXlsPaletteCallback,
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

void SendXlsPalette::PanelResized(const DG::PanelResizeEvent& ev)
{
	BeginMoveResizeItems();
	if (m_browserCtrl != nullptr)
		m_browserCtrl->Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void SendXlsPalette::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
	HidePalette();
	if (accepted != nullptr)
		*accepted = true;
}




