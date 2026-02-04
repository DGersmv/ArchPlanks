// *****************************************************************************
// Source code for ArchPlanks Add-On (Cutting Plan)
// *****************************************************************************

// Временно отключить проверку лицензии для тестов. Перед релизом поставить 0.
#define SKIP_LICENSE_FOR_TESTS 1

#include	"APIEnvir.h"
#include	"ACAPinc.h"
#include	"ResourceIDs.hpp"
#include	"BrowserRepl.hpp"
#include	"HelpPalette.hpp"
#include	"SelectionDetailsPalette.hpp"
#include	"SendXlsPalette.hpp"
#include	"LicenseManager.hpp"
#include	"APICommon.h"

// -----------------------------------------------------------------------------
// Show or Hide Browser Palette
// -----------------------------------------------------------------------------

static void	ShowOrHideBrowserRepl ()
{
	if (!BrowserRepl::HasInstance ()) {
		BrowserRepl::CreateInstance ();
	}
	
	if (BrowserRepl::GetInstance ().IsVisible ()) {
		BrowserRepl::GetInstance ().Hide ();
	} else {
		BrowserRepl::GetInstance ().Show ();
	}
}

// -----------------------------------------------------------------------------
// MenuCommandHandler
// -----------------------------------------------------------------------------

static bool g_isLicenseValid = false;
static bool g_isDemoExpired = false;

extern "C" {
	bool IsLicenseValid() { return g_isLicenseValid; }
	bool IsDemoExpired() { return g_isDemoExpired; }
}

GSErrCode MenuCommandHandler (const API_MenuParams *menuParams)
{
	if (!g_isLicenseValid && g_isDemoExpired) {
		short itemIndex = menuParams->menuItemRef.itemIndex;
		if (itemIndex != 2 && itemIndex != BrowserReplMenuItemIndex) {
			ACAPI_WriteReport("Demo period expired. Please purchase a license.", true);
			return NoError;
		}
	}
	
	switch (menuParams->menuItemRef.menuResID) {
		case BrowserReplMenuResId:
			switch (menuParams->menuItemRef.itemIndex) {
				case 1:
					SelectionDetailsPalette::ShowPalette();
					break;
				case 2:
					{
						GS::UniString url = LicenseManager::BuildLicenseUrl();
						HelpPalette::ShowWithURL(url);
					}
					break;
				case BrowserReplMenuItemIndex:
					ShowOrHideBrowserRepl();
					break;
				default:
					break;
			}
			break;
		default:
			break;
	}

	return NoError;
}

// -----------------------------------------------------------------------------
// Required functions
// -----------------------------------------------------------------------------

API_AddonType	CheckEnvironment (API_EnvirParams* envir)
{
	RSGetIndString (&envir->addOnInfo.name, 32000, 1, ACAPI_GetOwnResModule ());
	RSGetIndString (&envir->addOnInfo.description, 32000, 2, ACAPI_GetOwnResModule ());
	return APIAddon_Preload;
}

GSErrCode	RegisterInterface (void)
{
	GSErrCode err = ACAPI_MenuItem_RegisterMenu (BrowserReplMenuResId, 0, MenuCode_UserDef, MenuFlag_Default);
	if (DBERROR (err != NoError))
		return err;
	return err;
}

GSErrCode Initialize ()
{
#if SKIP_LICENSE_FOR_TESTS
	g_isLicenseValid = true;
	g_isDemoExpired = false;
#else
	LicenseManager::LicenseData licenseData;
	LicenseManager::LicenseStatus licenseStatus = LicenseManager::CheckLicense(licenseData);
	
	g_isLicenseValid = (licenseStatus == LicenseManager::LicenseStatus::Valid);
	bool isDemoActive = false;
	g_isDemoExpired = false;
	
	if (!g_isLicenseValid) {
		LicenseManager::DemoData demoData;
		isDemoActive = LicenseManager::CheckDemoPeriod(demoData);
		
		if (isDemoActive) {
			LicenseManager::UpdateDemoData();
		} else {
			g_isDemoExpired = true;
		}
	}
	
	if (!g_isLicenseValid && g_isDemoExpired) {
		DisableEnableMenuItem(BrowserReplMenuResId, 1, true);
		GS::UniString supportUrl = LicenseManager::BuildLicenseUrl();
		ACAPI_WriteReport("Demo period expired. Please purchase a license. Support: ", false);
		ACAPI_WriteReport(supportUrl.ToCStr().Get(), false);
	}
#endif

	GSErrCode err = ACAPI_MenuItem_InstallMenuHandler (BrowserReplMenuResId, MenuCommandHandler);
	if (DBERROR (err != NoError))
		return err;

	GSErrCode palErr = NoError;
	palErr |= BrowserRepl::RegisterPaletteControlCallBack ();
	palErr |= HelpPalette::RegisterPaletteControlCallBack ();
	palErr |= SelectionDetailsPalette::RegisterPaletteControlCallBack ();
	palErr |= SendXlsPalette::RegisterPaletteControlCallBack ();

	if (DBERROR (palErr != NoError))
		return palErr;

	return NoError;
}

GSErrCode	FreeData (void)
{
	return NoError;
}
