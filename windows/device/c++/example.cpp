

#include "drive_controller.h"
#include <iostream>


// Link to library
#pragma comment(lib, "Wbemuuid.lib")
#pragma comment(lib, "Cfgmgr32.lib")
#pragma comment(lib, "Setupapi.lib")


int main()
{
	// Initialize COM API
	HRESULT hResult = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hResult))
		return -1;

	// Setup process-wide security context
	hResult = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
	if (FAILED(hResult))
	{
		CoUninitialize();
		return -1;
	}


	// Init controller
	CDriveController::Instance().Init();

	// List all available lockable drives
	std::list<char> lstLockableDrives = CDriveController::Instance().GetAvailableDriveLetters();
	for (char chLetter : lstLockableDrives)
		std::cout << chLetter << std::endl;

	// Clear
	CDriveController::Instance().Reset();


	// Pause program
	system("pause");

	return 0;
}