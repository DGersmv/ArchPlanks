#include "LicenseManager.hpp"
#include "APICommon.h"

#ifdef GS_WIN
#include <Windows.h>
#include <iphlpapi.h>
#pragma comment(lib, "iphlpapi.lib")
#include <shlobj.h>
#pragma comment(lib, "shell32.lib")
#include <cstdio>
#else
#include <sys/socket.h>
#include <sys/ioctl.h>
#include <net/if.h>
#include <unistd.h>
#include <cstdio>
#endif

#include <ctime>
#include <sstream>
#include <iomanip>
#include <cstring>

// =============================================================================
// Получить идентификатор компьютера (MAC адрес)
// =============================================================================

GS::UniString LicenseManager::GetComputerId()
{
	return GetMacAddress();
}

// =============================================================================
// Получить MAC адрес первой сетевой карты
// =============================================================================

GS::UniString LicenseManager::GetMacAddress()
{
#ifdef GS_WIN
	IP_ADAPTER_INFO adapterInfo[16];
	DWORD dwBufLen = sizeof(adapterInfo);
	DWORD dwStatus = GetAdaptersInfo(adapterInfo, &dwBufLen);

	if (dwStatus == ERROR_SUCCESS) {
		PIP_ADAPTER_INFO pAdapterInfo = adapterInfo;
		// Берем первый адаптер с MAC адресом
		do {
			if (pAdapterInfo->AddressLength == 6) {
				// Форматируем MAC адрес как XX-XX-XX-XX-XX-XX
				char macStr[32];
				sprintf_s(macStr, sizeof(macStr), "%02X-%02X-%02X-%02X-%02X-%02X",
					(unsigned char)pAdapterInfo->Address[0],
					(unsigned char)pAdapterInfo->Address[1],
					(unsigned char)pAdapterInfo->Address[2],
					(unsigned char)pAdapterInfo->Address[3],
					(unsigned char)pAdapterInfo->Address[4],
					(unsigned char)pAdapterInfo->Address[5]);
				return GS::UniString(macStr);
			}
			pAdapterInfo = pAdapterInfo->Next;
		} while (pAdapterInfo);
	}
	
	// Если не удалось получить MAC, возвращаем пустую строку
	return GS::UniString("UNKNOWN-MAC");
#else
	// Для Mac - упрощенная версия (можно расширить)
	return GS::UniString("MAC-UNKNOWN");
#endif
}

// =============================================================================
// Получить путь к .apx файлу плагина
// =============================================================================

GS::UniString LicenseManager::GetAddonFilePath()
{
#ifdef GS_WIN
	HMODULE hModule = reinterpret_cast<HMODULE>(ACAPI_GetOwnResModule());
	if (hModule == nullptr) {
		return GS::UniString();
	}

	wchar_t modulePath[MAX_PATH];
	DWORD length = GetModuleFileNameW(hModule, modulePath, MAX_PATH);
	if (length == 0 || length >= MAX_PATH) {
		return GS::UniString();
	}

	return GS::UniString(modulePath);
#else
	// Для Mac - нужно использовать другой подход
	return GS::UniString();
#endif
}

// =============================================================================
// Получить путь к директории пользовательских данных (AppData\Local\LandscapeHelper)
// =============================================================================

GS::UniString LicenseManager::GetUserDataDirectory()
{
#ifdef GS_WIN
	wchar_t appDataPath[MAX_PATH];
	if (SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA, nullptr, SHGFP_TYPE_CURRENT, appDataPath) == S_OK) {
		GS::UniString userDataDir(appDataPath);
		userDataDir += "\\LandscapeHelper";
		
		// Создаем директорию, если её нет
		CreateDirectoryW(userDataDir.ToUStr().Get(), nullptr);
		
		return userDataDir;
	}
#endif
	return GS::UniString();
}

// =============================================================================
// Получить путь к файлу лицензии (проверяет оба места: рядом с .apx и в AppData)
// Возвращает путь к найденному файлу или первый путь для проверки
// =============================================================================

GS::UniString LicenseManager::GetLicenseFilePath()
{
	// Сначала проверяем рядом с .apx файлом
	GS::UniString addonPath = GetAddonFilePath();
	if (!addonPath.IsEmpty()) {
		GS::UniString licensePath = addonPath;
		
		// Убираем расширение .apx и добавляем .lic
		UIndex lastDot = licensePath.FindLast('.');
		if (lastDot != MaxUIndex) {
			licensePath = licensePath.GetSubstring(0, lastDot);
		}
		licensePath += ".lic";
		
		// Проверяем существование файла
#ifdef GS_WIN
		if (GetFileAttributesW(licensePath.ToUStr().Get()) != INVALID_FILE_ATTRIBUTES) {
			return licensePath;
		}
#else
		// Для других платформ - просто возвращаем путь
		return licensePath;
#endif
	}
	
	// Если не нашли рядом с .apx, проверяем в AppData
	GS::UniString userDataDir = GetUserDataDirectory();
	if (!userDataDir.IsEmpty()) {
		GS::UniString licensePath = userDataDir;
		licensePath += "\\license.lic";
		return licensePath;
	}
	
	return GS::UniString();
}

// =============================================================================
// Парсить строку формата KEY=VALUE
// =============================================================================

bool LicenseManager::ParseLicenseLine(const GS::UniString& line, GS::UniString& key, GS::UniString& value)
{
	// Убираем пробелы в начале и конце
	GS::UniString trimmed = line;
	trimmed.Trim();
	
	// Пропускаем пустые строки и комментарии
	if (trimmed.IsEmpty() || trimmed.GetSubstring(0, 1) == GS::UniString("#")) {
		return false;
	}

	// Ищем знак равно
	Int32 equalPos = trimmed.FindFirst(GS::UniString("="));
	if (equalPos == -1) {
		return false;
	}

	key = trimmed.GetSubstring(0, equalPos);
	key.Trim();
	
	value = trimmed.GetSubstring(equalPos + 1, trimmed.GetLength() - equalPos - 1);
	value.Trim();

	return !key.IsEmpty();
}

// =============================================================================
// Прочитать файл лицензии
// =============================================================================

bool LicenseManager::ReadLicenseFile(const GS::UniString& filePath, LicenseData& licenseData)
{
#ifdef GS_WIN
	FILE* fp = _wfopen(filePath.ToUStr().Get(), L"r");
#else
	FILE* fp = fopen(filePath.ToCStr().Get(), "r");
#endif
	
	if (fp == nullptr) {
		return false;
	}

	char lineBuffer[512];
	while (fgets(lineBuffer, sizeof(lineBuffer), fp) != nullptr) {
		GS::UniString line(lineBuffer);
		GS::UniString key, value;

		if (ParseLicenseLine(line, key, value)) {
			if (key == GS::UniString("COMPUTER_ID")) {
				licenseData.computerId = value;
			} else if (key == GS::UniString("PLUGIN_NAME")) {
				licenseData.pluginName = value;
			} else if (key == GS::UniString("PLUGIN_VERSION")) {
				licenseData.pluginVersion = value;
			} else if (key == GS::UniString("VALID_UNTIL")) {
				licenseData.validUntil = value;
			} else if (key == GS::UniString("ISSUED_DATE")) {
				licenseData.issuedDate = value;
			} else if (key == GS::UniString("LICENSE_KEY")) {
				licenseData.licenseKey = value;
			}
		}
	}

	fclose(fp);
	return true;
}

// =============================================================================
// Проверить валидность даты (формат YYYY-MM-DD)
// =============================================================================

bool LicenseManager::IsValidDate(const GS::UniString& dateStr)
{
	if (dateStr.GetLength() != 10) {
		return false;
	}

	// Проверяем формат YYYY-MM-DD
	if (dateStr[4] != '-' || dateStr[7] != '-') {
		return false;
	}

	// Проверяем что остальные символы - цифры
	GS::UniString digits = GS::UniString("0123456789");
	for (UIndex i = 0; i < dateStr.GetLength(); ++i) {
		if (i != 4 && i != 7) {
			GS::UniString chStr = dateStr.GetSubstring(i, 1);
			if (!digits.Contains(chStr)) {
				return false;
			}
		}
	}

	return true;
}

// =============================================================================
// Сравнить даты (YYYY-MM-DD), возвращает true если date1 < date2
// =============================================================================

bool LicenseManager::CompareDates(const GS::UniString& date1, const GS::UniString& date2)
{
	// Простое сравнение строк работает для формата YYYY-MM-DD
	return date1 < date2;
}

// =============================================================================
// Получить текущую дату в формате YYYY-MM-DD
// =============================================================================

GS::UniString LicenseManager::GetCurrentDate()
{
	std::time_t now = std::time(nullptr);
	std::tm* timeinfo;
#ifdef GS_WIN
	std::tm timeinfoBuf;
	timeinfo = &timeinfoBuf;
	localtime_s(timeinfo, &now);
#else
	std::tm timeinfoBuf;
	timeinfo = localtime_r(&now, &timeinfoBuf);
#endif
	
	char dateStr[16];
#ifdef GS_WIN
	sprintf_s(dateStr, sizeof(dateStr), "%04d-%02d-%02d",
		timeinfo->tm_year + 1900,
		timeinfo->tm_mon + 1,
		timeinfo->tm_mday);
#else
	snprintf(dateStr, sizeof(dateStr), "%04d-%02d-%02d",
		timeinfo->tm_year + 1900,
		timeinfo->tm_mon + 1,
		timeinfo->tm_mday);
#endif
	
	return GS::UniString(dateStr);
}

// =============================================================================
// Проверить лицензию (главная функция)
// Проверяет в двух местах: рядом с .apx и в AppData\Local\LandscapeHelper
// =============================================================================

LicenseManager::LicenseStatus LicenseManager::CheckLicense(LicenseData& licenseData)
{
	bool licenseFound = false;
	
	// Сначала проверяем рядом с .apx файлом
	GS::UniString addonPath = GetAddonFilePath();
	if (!addonPath.IsEmpty()) {
		GS::UniString licensePath = addonPath;
		
		// Убираем расширение .apx и добавляем .lic
		UIndex lastDot = licensePath.FindLast('.');
		if (lastDot != MaxUIndex) {
			licensePath = licensePath.GetSubstring(0, lastDot);
		}
		licensePath += ".lic";
		
		// Проверяем существование и читаем файл
#ifdef GS_WIN
		if (GetFileAttributesW(licensePath.ToUStr().Get()) != INVALID_FILE_ATTRIBUTES) {
			if (ReadLicenseFile(licensePath, licenseData)) {
				licenseFound = true;
			}
		}
#else
		if (ReadLicenseFile(licensePath, licenseData)) {
			licenseFound = true;
		}
#endif
	}
	
	// Если не нашли рядом с .apx, проверяем в AppData
	if (!licenseFound) {
		GS::UniString userDataDir = GetUserDataDirectory();
		if (!userDataDir.IsEmpty()) {
			GS::UniString licensePath = userDataDir;
			licensePath += "\\license.lic";
			
			if (ReadLicenseFile(licensePath, licenseData)) {
				licenseFound = true;
			}
		}
	}
	
	// Если файл не найден
	if (!licenseFound) {
		WriteLicenseLog(LicenseStatus::NotFound, licenseData);
		return LicenseStatus::NotFound;
	}

	// Проверяем валидность лицензии

	// Проверяем что все поля заполнены
	if (licenseData.computerId.IsEmpty() ||
		licenseData.pluginName.IsEmpty() ||
		licenseData.validUntil.IsEmpty()) {
		WriteLicenseLog(LicenseStatus::ParseError, licenseData);
		return LicenseStatus::ParseError;
	}

	// Проверяем идентификатор компьютера
	GS::UniString currentComputerId = GetComputerId();
	if (currentComputerId != licenseData.computerId) {
		WriteLicenseLog(LicenseStatus::ComputerMismatch, licenseData);
		return LicenseStatus::ComputerMismatch;
	}

	// Проверяем название плагина (ожидаем ArchPlanks_AC29, LandscapeHelper_AC29 или Browser_Repl_Int)
	GS::UniString expectedPluginName1 = GS::UniString("ArchPlanks_AC29");
	GS::UniString expectedPluginName2 = GS::UniString("LandscapeHelper_AC29");
	GS::UniString expectedPluginName3 = GS::UniString("Browser_Repl_Int");
	if (licenseData.pluginName != expectedPluginName1 && 
		licenseData.pluginName != expectedPluginName2 &&
		licenseData.pluginName != expectedPluginName3) {
		WriteLicenseLog(LicenseStatus::PluginMismatch, licenseData);
		return LicenseStatus::PluginMismatch;
	}

	// Проверяем дату окончания действия
	if (!IsValidDate(licenseData.validUntil)) {
		WriteLicenseLog(LicenseStatus::ParseError, licenseData);
		return LicenseStatus::ParseError;
	}

	GS::UniString currentDate = GetCurrentDate();
	if (CompareDates(licenseData.validUntil, currentDate)) {
		WriteLicenseLog(LicenseStatus::Expired, licenseData);
		return LicenseStatus::Expired;
	}

	// Все проверки пройдены
	WriteLicenseLog(LicenseStatus::Valid, licenseData);
	return LicenseStatus::Valid;
}

// =============================================================================
// Записать лог проверки лицензии
// =============================================================================

void LicenseManager::WriteLicenseLog(LicenseStatus status, const LicenseData& licenseData)
{
	// Логи пишем в AppData\Local\LandscapeHelper
	GS::UniString userDataDir = GetUserDataDirectory();
	if (userDataDir.IsEmpty()) {
		return;
	}
	
	GS::UniString logPath = userDataDir;
	logPath += "\\license.log";

	// Открываем файл для добавления (append mode)
#ifdef GS_WIN
	FILE* fp = _wfopen(logPath.ToUStr().Get(), L"a");
#else
	FILE* fp = fopen(logPath.ToCStr().Get(), "a");
#endif
	
	if (fp == nullptr) {
		return;
	}

	// Формируем строку лога
	GS::UniString statusStr;
	GS::UniString currentDate = GetCurrentDate();

	switch (status) {
		case LicenseStatus::Valid:
			statusStr = "VALID";
			break;
		case LicenseStatus::Invalid:
			statusStr = "INVALID";
			break;
		case LicenseStatus::NotFound:
			statusStr = "NOT_FOUND";
			break;
		case LicenseStatus::Expired:
			statusStr = "EXPIRED";
			break;
		case LicenseStatus::ComputerMismatch:
			statusStr = "COMPUTER_MISMATCH";
			break;
		case LicenseStatus::PluginMismatch:
			statusStr = "PLUGIN_MISMATCH";
			break;
		case LicenseStatus::ParseError:
			statusStr = "PARSE_ERROR";
			break;
	}

	GS::UniString logLine = GS::UniString("[") + currentDate + GS::UniString("] ");
	logLine += GS::UniString("Status: ") + statusStr;
	logLine += GS::UniString(" | Computer: ") + GetComputerId();
	
	if (!licenseData.pluginName.IsEmpty()) {
		logLine += GS::UniString(" | Plugin: ") + licenseData.pluginName;
	}
	
	if (!licenseData.validUntil.IsEmpty()) {
		logLine += GS::UniString(" | Valid until: ") + licenseData.validUntil;
	}
	
	logLine += GS::UniString("\n");

	// Конвертируем UniString в UTF-8 для записи
#ifdef GS_WIN
	// На Windows используем UTF-16 напрямую
	fputws(logLine.ToUStr().Get(), fp);
#else
	// На других платформах конвертируем в UTF-8
	GS::UniString utf8Line = logLine;
	fputs(utf8Line.ToCStr().Get(), fp);
#endif

	fclose(fp);
}

// =============================================================================
// Записать общий лог в файл (для отладки)
// =============================================================================

void LicenseManager::WriteLog(const GS::UniString& message)
{
	// Логи пишем в AppData\Local\LandscapeHelper
	GS::UniString userDataDir = GetUserDataDirectory();
	if (userDataDir.IsEmpty()) {
		return;
	}
	
	GS::UniString logPath = userDataDir;
	logPath += "\\license.log";

	// Открываем файл для добавления (append mode)
#ifdef GS_WIN
	FILE* fp = _wfopen(logPath.ToUStr().Get(), L"a");
#else
	FILE* fp = fopen(logPath.ToCStr().Get(), "a");
#endif
	
	if (fp == nullptr) {
		return;
	}

	GS::UniString currentDate = GetCurrentDate();
	GS::UniString logLine = GS::UniString("[") + currentDate + GS::UniString("] ") + message + GS::UniString("\n");

	// Конвертируем UniString в UTF-8 для записи
#ifdef GS_WIN
	// На Windows используем UTF-16 напрямую
	fputws(logLine.ToUStr().Get(), fp);
#else
	// На других платформах конвертируем в UTF-8
	GS::UniString utf8Line = logLine;
	fputs(utf8Line.ToCStr().Get(), fp);
#endif

	fclose(fp);
}

// =============================================================================
// Демо-режим: получить путь к файлу с демо-данными
// =============================================================================

GS::UniString LicenseManager::GetDemoFilePath()
{
	// Демо-файл храним в AppData\Local\LandscapeHelper
	GS::UniString userDataDir = GetUserDataDirectory();
	if (!userDataDir.IsEmpty()) {
		GS::UniString demoPath = userDataDir;
		demoPath += "\\demo.dat";
		return demoPath;
	}
	return GS::UniString();
}

// =============================================================================
// Работа с реестром: получить дату первого запуска
// =============================================================================

GS::UniString LicenseManager::GetFirstLaunchDateFromRegistry()
{
#ifdef GS_WIN
	HKEY hKey;
	LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\LandscapeHelper", 0, KEY_READ, &hKey);
	
	if (result == ERROR_SUCCESS) {
		wchar_t valueBuffer[64];
		DWORD valueSize = sizeof(valueBuffer);
		DWORD valueType = REG_SZ;
		
		result = RegQueryValueExW(hKey, L"FirstLaunchDate", nullptr, &valueType, 
			reinterpret_cast<LPBYTE>(valueBuffer), &valueSize);
		
		RegCloseKey(hKey);
		
		if (result == ERROR_SUCCESS && valueType == REG_SZ) {
			return GS::UniString(valueBuffer);
		}
	}
#endif
	return GS::UniString();
}

// =============================================================================
// Работа с реестром: установить дату первого запуска
// =============================================================================

void LicenseManager::SetFirstLaunchDateToRegistry(const GS::UniString& date)
{
#ifdef GS_WIN
	HKEY hKey;
	LONG result = RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\LandscapeHelper", 
		0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr);
	
	if (result == ERROR_SUCCESS) {
		const wchar_t* dateStr = date.ToUStr().Get();
		RegSetValueExW(hKey, L"FirstLaunchDate", 0, REG_SZ, 
			reinterpret_cast<const BYTE*>(dateStr), 
			static_cast<DWORD>((wcslen(dateStr) + 1) * sizeof(wchar_t)));
		
		RegCloseKey(hKey);
	}
#endif
}

// =============================================================================
// Работа с реестром: получить счетчик запусков
// =============================================================================

int LicenseManager::GetLaunchCountFromRegistry()
{
#ifdef GS_WIN
	HKEY hKey;
	LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\LandscapeHelper", 0, KEY_READ, &hKey);
	
	if (result == ERROR_SUCCESS) {
		DWORD value = 0;
		DWORD valueSize = sizeof(DWORD);
		DWORD valueType = REG_DWORD;
		
		result = RegQueryValueExW(hKey, L"LaunchCount", nullptr, &valueType, 
			reinterpret_cast<LPBYTE>(&value), &valueSize);
		
		RegCloseKey(hKey);
		
		if (result == ERROR_SUCCESS && valueType == REG_DWORD) {
			// Валидация: только 0-100
			if (value <= 100) {
				return static_cast<int>(value);
			}
		}
	}
#endif
	return 0;
}

// =============================================================================
// Работа с реестром: установить счетчик запусков
// =============================================================================

void LicenseManager::SetLaunchCountToRegistry(int count)
{
#ifdef GS_WIN
	HKEY hKey;
	LONG result = RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\LandscapeHelper", 
		0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr);
	
	if (result == ERROR_SUCCESS) {
		DWORD value = static_cast<DWORD>(count);
		RegSetValueExW(hKey, L"LaunchCount", 0, REG_DWORD, 
			reinterpret_cast<const BYTE*>(&value), sizeof(DWORD));
		
		RegCloseKey(hKey);
	}
#endif
}

// =============================================================================
// Демо-режим: проверить, активен ли демо-период (22 дня или 22 запуска)
// Проверка срока идет по реестру (защита от удаления), счетчик запусков - из файла
// =============================================================================

bool LicenseManager::CheckDemoPeriod(DemoData& demoData)
{
	const int MAX_LAUNCHES = 22;
	const int MAX_DAYS = 22;

	GS::UniString currentDate = GetCurrentDate();
	
	// Получаем дату первого запуска из реестра (основной источник)
	GS::UniString firstLaunchDate = GetFirstLaunchDateFromRegistry();
	
	// Если в реестре нет даты - это первый запуск
	if (firstLaunchDate.IsEmpty()) {
		firstLaunchDate = currentDate;
		SetFirstLaunchDateToRegistry(firstLaunchDate);
		demoData.firstLaunchDate = firstLaunchDate;
		demoData.lastLaunchDate = currentDate;
		demoData.launchCount = 0; // Будет увеличен в UpdateDemoData
		return true; // Демо активен
	}
	
	demoData.firstLaunchDate = firstLaunchDate;

	// Получаем счетчик запусков из реестра (основной источник, защищен от изменений)
	demoData.launchCount = GetLaunchCountFromRegistry();
	demoData.lastLaunchDate = currentDate;

	// Проверяем лимит запусков (из реестра)
	if (demoData.launchCount >= MAX_LAUNCHES) {
		return false; // Демо истек по количеству запусков
	}

	// Проверяем лимит дней (ОСНОВНАЯ ПРОВЕРКА - по реестру)
	if (CompareDates(firstLaunchDate, currentDate)) {
		// Вычисляем разницу в днях
		int year1 = 0, month1 = 0, day1 = 0, year2 = 0, month2 = 0, day2 = 0;
		if (demoData.firstLaunchDate.GetLength() >= 10 && currentDate.GetLength() >= 10) {
			GS::UniString firstYear = demoData.firstLaunchDate.GetSubstring(0, 4);
			GS::UniString firstMonth = demoData.firstLaunchDate.GetSubstring(5, 2);
			GS::UniString firstDay = demoData.firstLaunchDate.GetSubstring(8, 2);
			GS::UniString currYear = currentDate.GetSubstring(0, 4);
			GS::UniString currMonth = currentDate.GetSubstring(5, 2);
			GS::UniString currDay = currentDate.GetSubstring(8, 2);

			// Конвертируем строки в числа
			for (UIndex i = 0; i < firstYear.GetLength(); ++i) {
				GS::UniString chStr = firstYear.GetSubstring(i, 1);
				char ch = chStr.ToCStr().Get()[0];
				if (ch >= '0' && ch <= '9') year1 = year1 * 10 + (ch - '0');
			}
			for (UIndex i = 0; i < firstMonth.GetLength(); ++i) {
				GS::UniString chStr = firstMonth.GetSubstring(i, 1);
				char ch = chStr.ToCStr().Get()[0];
				if (ch >= '0' && ch <= '9') month1 = month1 * 10 + (ch - '0');
			}
			for (UIndex i = 0; i < firstDay.GetLength(); ++i) {
				GS::UniString chStr = firstDay.GetSubstring(i, 1);
				char ch = chStr.ToCStr().Get()[0];
				if (ch >= '0' && ch <= '9') day1 = day1 * 10 + (ch - '0');
			}
			for (UIndex i = 0; i < currYear.GetLength(); ++i) {
				GS::UniString chStr = currYear.GetSubstring(i, 1);
				char ch = chStr.ToCStr().Get()[0];
				if (ch >= '0' && ch <= '9') year2 = year2 * 10 + (ch - '0');
			}
			for (UIndex i = 0; i < currMonth.GetLength(); ++i) {
				GS::UniString chStr = currMonth.GetSubstring(i, 1);
				char ch = chStr.ToCStr().Get()[0];
				if (ch >= '0' && ch <= '9') month2 = month2 * 10 + (ch - '0');
			}
			for (UIndex i = 0; i < currDay.GetLength(); ++i) {
				GS::UniString chStr = currDay.GetSubstring(i, 1);
				char ch = chStr.ToCStr().Get()[0];
				if (ch >= '0' && ch <= '9') day2 = day2 * 10 + (ch - '0');
			}

			// Вычисляем разницу в днях
			std::tm tm1 = {};
			tm1.tm_year = year1 - 1900;
			tm1.tm_mon = month1 - 1;
			tm1.tm_mday = day1;
			std::tm tm2 = {};
			tm2.tm_year = year2 - 1900;
			tm2.tm_mon = month2 - 1;
			tm2.tm_mday = day2;

			std::time_t time1 = std::mktime(&tm1);
			std::time_t time2 = std::mktime(&tm2);
			double diffSeconds = std::difftime(time2, time1);
			int diffDays = static_cast<int>(diffSeconds / (24 * 60 * 60));

			if (diffDays >= MAX_DAYS) {
				return false; // Демо истек по времени
			}
		}
	}

	return true; // Демо активен
}

// =============================================================================
// Демо-режим: обновить данные демо (увеличить счетчик запусков, обновить дату)
// =============================================================================

void LicenseManager::UpdateDemoData()
{
	DemoData demoData;
	GS::UniString currentDate = GetCurrentDate();

	// Получаем дату первого запуска из реестра (основной источник, защищен от изменений)
	demoData.firstLaunchDate = GetFirstLaunchDateFromRegistry();
	if (demoData.firstLaunchDate.IsEmpty()) {
		// Если в реестре нет - это первый запуск, устанавливаем текущую дату
		demoData.firstLaunchDate = currentDate;
		SetFirstLaunchDateToRegistry(currentDate);
	}

	// Получаем счетчик запусков из реестра (основной источник, защищен от изменений)
	int launchCount = GetLaunchCountFromRegistry();
	
	// Увеличиваем счетчик запусков
	launchCount++;
	
	// Ограничиваем максимум 22
	if (launchCount > 22) {
		launchCount = 22;
	}
	
	// Сохраняем в реестр (защита от подделки)
	SetLaunchCountToRegistry(launchCount);
	
	demoData.launchCount = launchCount;
	demoData.lastLaunchDate = currentDate; // Всегда текущая дата

	// Перезаписываем файл с актуальными данными из реестра (защита от модификации)
	GS::UniString demoPath = GetDemoFilePath();
	if (demoPath.IsEmpty()) {
		return;
	}
	
	FILE* fp;
#ifdef GS_WIN
	fp = _wfopen(demoPath.ToUStr().Get(), L"w");
#else
	fp = fopen(demoPath.ToCStr().Get(), "w");
#endif

	if (fp != nullptr) {
		// Переписываем файл с актуальными данными из реестра (защита от модификации)
		// Всегда записываем корректные данные, игнорируя что было в файле
		char line[256];
		
		const int MAX_LAUNCHES = 22;
		const int MAX_DAYS = 22;
		
		// Вычисляем дату окончания демо (первый запуск + 22 дня)
		GS::UniString expiresDate;
		int daysRemaining = MAX_DAYS;
		
		if (!demoData.firstLaunchDate.IsEmpty() && demoData.firstLaunchDate.GetLength() >= 10) {
			// Парсим дату первого запуска
			int year = 0, month = 0, day = 0;
			const char* dateStr = demoData.firstLaunchDate.ToCStr().Get();
			if (sscanf(dateStr, "%d-%d-%d", &year, &month, &day) == 3) {
				// Создаем структуру времени
				std::tm tm1 = {};
				tm1.tm_year = year - 1900;
				tm1.tm_mon = month - 1;
				tm1.tm_mday = day + MAX_DAYS; // Добавляем 22 дня
				tm1.tm_hour = 12;
				
				// Нормализуем дату (mktime исправит переполнение дней)
				std::time_t expiresTime = std::mktime(&tm1);
				std::tm* expiresTm = std::localtime(&expiresTime);
				
				if (expiresTm != nullptr) {
					char expDateStr[32];
#ifdef GS_WIN
					sprintf_s(expDateStr, sizeof(expDateStr), "%04d-%02d-%02d",
						expiresTm->tm_year + 1900, expiresTm->tm_mon + 1, expiresTm->tm_mday);
#else
					snprintf(expDateStr, sizeof(expDateStr), "%04d-%02d-%02d",
						expiresTm->tm_year + 1900, expiresTm->tm_mon + 1, expiresTm->tm_mday);
#endif
					expiresDate = GS::UniString(expDateStr);
				}
				
				// Вычисляем сколько дней осталось
				std::time_t now = std::time(nullptr);
				double diffSeconds = std::difftime(expiresTime, now);
				daysRemaining = static_cast<int>(diffSeconds / (24 * 60 * 60));
				if (daysRemaining < 0) daysRemaining = 0;
			}
		}
		
		int launchesRemaining = MAX_LAUNCHES - demoData.launchCount;
		if (launchesRemaining < 0) launchesRemaining = 0;

#ifdef GS_WIN
		// Записываем дату первого запуска из реестра (источник правды)
		if (!demoData.firstLaunchDate.IsEmpty()) {
			sprintf_s(line, sizeof(line), "FIRST_LAUNCH_DATE=%s\n", demoData.firstLaunchDate.ToCStr().Get());
			fputs(line, fp);
		}
		// Записываем дату окончания демо
		if (!expiresDate.IsEmpty()) {
			sprintf_s(line, sizeof(line), "DEMO_EXPIRES_DATE=%s\n", expiresDate.ToCStr().Get());
			fputs(line, fp);
		}
		// Записываем счетчик запусков
		sprintf_s(line, sizeof(line), "LAUNCH_COUNT=%d\n", demoData.launchCount);
		fputs(line, fp);
		// Записываем сколько осталось запусков
		sprintf_s(line, sizeof(line), "LAUNCHES_REMAINING=%d\n", launchesRemaining);
		fputs(line, fp);
		// Записываем сколько осталось дней
		sprintf_s(line, sizeof(line), "DAYS_REMAINING=%d\n", daysRemaining);
		fputs(line, fp);
#else
		if (!demoData.firstLaunchDate.IsEmpty()) {
			snprintf(line, sizeof(line), "FIRST_LAUNCH_DATE=%s\n", demoData.firstLaunchDate.ToCStr().Get());
			fputs(line, fp);
		}
		if (!expiresDate.IsEmpty()) {
			snprintf(line, sizeof(line), "DEMO_EXPIRES_DATE=%s\n", expiresDate.ToCStr().Get());
			fputs(line, fp);
		}
		snprintf(line, sizeof(line), "LAUNCH_COUNT=%d\n", demoData.launchCount);
		fputs(line, fp);
		snprintf(line, sizeof(line), "LAUNCHES_REMAINING=%d\n", launchesRemaining);
		fputs(line, fp);
		snprintf(line, sizeof(line), "DAYS_REMAINING=%d\n", daysRemaining);
		fputs(line, fp);
#endif
		fclose(fp);
	}
}

// =============================================================================
// Сформировать URL для страницы лицензии с данными о плагине
// =============================================================================

GS::UniString LicenseManager::BuildLicenseUrl()
{
	GS::UniString url = GS::UniString("https://landscape.227.info/license");
	
	// Получаем MAC-адрес
	GS::UniString macAddress = GetComputerId();
	if (!macAddress.IsEmpty()) {
		url += GS::UniString("?mac=") + macAddress;
	}
	
	// Сначала проверяем, есть ли валидная лицензия
	LicenseData licenseData;
	LicenseStatus licenseStatus = CheckLicense(licenseData);
	
	if (licenseStatus == LicenseStatus::Valid) {
		// Лицензия валидна - передаем информацию о лицензии
		url += GS::UniString("&mode=license");
		
		if (!licenseData.validUntil.IsEmpty()) {
			url += GS::UniString("&validUntil=") + licenseData.validUntil;
		}
		if (!licenseData.issuedDate.IsEmpty()) {
			url += GS::UniString("&issuedDate=") + licenseData.issuedDate;
		}
		if (!licenseData.licenseKey.IsEmpty()) {
			url += GS::UniString("&licenseKey=") + licenseData.licenseKey;
		}
		
		return url;
	}
	
	// Лицензии нет - передаем информацию о демо
	url += GS::UniString("&mode=demo");
	
	// Читаем демо-файл
	GS::UniString demoPath = GetDemoFilePath();
	if (demoPath.IsEmpty()) {
		return url; // Возвращаем URL только с MAC и mode=demo, если файл не найден
	}
	
	FILE* fp;
#ifdef GS_WIN
	fp = _wfopen(demoPath.ToUStr().Get(), L"r");
#else
	fp = fopen(demoPath.ToCStr().Get(), "r");
#endif
	
	if (fp == nullptr) {
		return url; // Возвращаем URL только с MAC и mode=demo, если файл не открылся
	}
	
	// Переменные для хранения данных
	GS::UniString firstLaunchDate;
	GS::UniString demoExpiresDate;
	GS::UniString launchCount;
	GS::UniString launchesRemaining;
	GS::UniString daysRemaining;
	
	// Читаем файл построчно
	char lineBuffer[256];
	while (fgets(lineBuffer, sizeof(lineBuffer), fp) != nullptr) {
		GS::UniString line(lineBuffer);
		line.Trim();
		
		if (line.IsEmpty()) {
			continue;
		}
		
		GS::UniString key, value;
		if (ParseLicenseLine(line, key, value)) {
			if (key == GS::UniString("FIRST_LAUNCH_DATE")) {
				firstLaunchDate = value;
			} else if (key == GS::UniString("DEMO_EXPIRES_DATE")) {
				demoExpiresDate = value;
			} else if (key == GS::UniString("LAUNCH_COUNT")) {
				launchCount = value;
			} else if (key == GS::UniString("LAUNCHES_REMAINING")) {
				launchesRemaining = value;
			} else if (key == GS::UniString("DAYS_REMAINING")) {
				daysRemaining = value;
			}
		}
	}
	
	fclose(fp);
	
	// Добавляем параметры к URL
	if (!firstLaunchDate.IsEmpty()) {
		url += GS::UniString("&firstLaunch=") + firstLaunchDate;
	}
	if (!demoExpiresDate.IsEmpty()) {
		url += GS::UniString("&demoExpires=") + demoExpiresDate;
	}
	if (!launchCount.IsEmpty()) {
		url += GS::UniString("&launchCount=") + launchCount;
	}
	if (!launchesRemaining.IsEmpty()) {
		url += GS::UniString("&launchesRemaining=") + launchesRemaining;
	}
	if (!daysRemaining.IsEmpty()) {
		url += GS::UniString("&daysRemaining=") + daysRemaining;
	}
	
	return url;
}

