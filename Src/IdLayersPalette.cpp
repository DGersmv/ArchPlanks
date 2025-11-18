#include "IdLayersPalette.hpp"

#include "APIEnvir.h"
#include "ACAPinc.h"

#include "DGBrowser.hpp"
#include "BrowserRepl.hpp"

static GS::UniString LoadIdLayersHtml()
{
	GS::UniString html;
	GSHandle data = RSLoadResource('DATA', ACAPI_GetOwnResModule(), IdLayersHtmlResId);
	if (data != nullptr) {
		const GSSize size = BMhGetSize(data);
		html.Append(*data, size);
		BMhKill(&data);
	} else {
		html = "<html><body style='font-family:Arial;color:#900;padding:12px;'>"
			"Не удалось загрузить HTML ресурса палитры ID/Слои."
			"</body></html>";
	}
	return html;
}

static GSErrCode __ACENV_CALL IdLayersPaletteCallback(Int32 /*refCon*/, API_PaletteMessageID messageID, GS::IntPtr param)
{
	switch (messageID) {
	case APIPalMsg_OpenPalette:
		if (!IdLayersPalette::HasInstance()) IdLayersPalette::CreateInstance();
		IdLayersPalette::GetInstance().Show();
		break;

	case APIPalMsg_ClosePalette:
		if (IdLayersPalette::HasInstance()) IdLayersPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_Begin:
		if (IdLayersPalette::HasInstance() && IdLayersPalette::GetInstance().IsVisible())
			IdLayersPalette::GetInstance().Hide();
		break;

	case APIPalMsg_HidePalette_End:
		if (IdLayersPalette::HasInstance() && !IdLayersPalette::GetInstance().IsVisible())
			IdLayersPalette::GetInstance().Show();
		break;

	case APIPalMsg_DisableItems_Begin:
		if (IdLayersPalette::HasInstance() && IdLayersPalette::GetInstance().IsVisible())
			IdLayersPalette::GetInstance().DisableItems();
		break;

	case APIPalMsg_DisableItems_End:
		if (IdLayersPalette::HasInstance() && IdLayersPalette::GetInstance().IsVisible())
			IdLayersPalette::GetInstance().EnableItems();
		break;

	case APIPalMsg_IsPaletteVisible:
		if (param != 0)
			*reinterpret_cast<bool*>(param) = IdLayersPalette::HasInstance() && IdLayersPalette::GetInstance().IsVisible();
		break;

	default:
		break;
	}
	return NoError;
}

GS::Ref<IdLayersPalette> IdLayersPalette::s_instance(nullptr);
const GS::Guid IdLayersPalette::s_guid("{0a2e5c92-7f54-4cc7-a3db-1f5c1f1e4a8f}");

IdLayersPalette::IdLayersPalette()
	: DG::Palette(ACAPI_GetOwnResModule(), IdLayersPaletteResId, ACAPI_GetOwnResModule(), s_guid)
{
	m_browserCtrl = new DG::Browser(GetReference(), IdLayersBrowserCtrlId);
	Attach(*this);
	BeginEventProcessing();
	Init();
}

IdLayersPalette::~IdLayersPalette()
{
	EndEventProcessing();
	delete m_browserCtrl;
	m_browserCtrl = nullptr;
}

void IdLayersPalette::Init()
{
	LoadHtml();
	if (m_browserCtrl != nullptr)
		BrowserRepl::RegisterACAPIJavaScriptObject(*m_browserCtrl);
}

void IdLayersPalette::LoadHtml()
{
	if (m_browserCtrl == nullptr)
		return;

	const GS::UniString html = LoadIdLayersHtml();
	m_browserCtrl->LoadHTML(html);
}

bool IdLayersPalette::HasInstance()
{
	return s_instance != nullptr;
}

void IdLayersPalette::CreateInstance()
{
	DBASSERT(!HasInstance());
	s_instance = new IdLayersPalette();
	ACAPI_KeepInMemory(true);
}

IdLayersPalette& IdLayersPalette::GetInstance()
{
	DBASSERT(HasInstance());
	return *s_instance;
}

void IdLayersPalette::DestroyInstance()
{
	s_instance = nullptr;
}

void IdLayersPalette::ShowPalette()
{
	if (!HasInstance())
		CreateInstance();

	GetInstance().Show();
}

void IdLayersPalette::HidePalette()
{
	if (!HasInstance())
		return;

	GetInstance().Hide();
}

GSErrCode IdLayersPalette::RegisterPaletteControlCallBack()
{
	return ACAPI_RegisterModelessWindow(
		GS::CalculateHashValue(s_guid),
		IdLayersPaletteCallback,
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

void IdLayersPalette::PanelResized(const DG::PanelResizeEvent& ev)
{
	BeginMoveResizeItems();
	if (m_browserCtrl != nullptr)
		m_browserCtrl->Resize(ev.GetHorizontalChange(), ev.GetVerticalChange());
	EndMoveResizeItems();
}

void IdLayersPalette::PanelCloseRequested(const DG::PanelCloseRequestEvent&, bool* accepted)
{
	HidePalette();
	if (accepted != nullptr)
		*accepted = true;
}






