

#include "drive_watcher.h"
#include <iostream>


// Link to library
#pragma comment(lib, "Wbemuuid.lib")
#pragma comment(lib, "Cfgmgr32.lib")
#pragma comment(lib, "Setupapi.lib")


// Notification listener
class CNotificationListener : public CDriveWatcher::INotificationListener
{
public:
	void Notify(const CDriveWatcher::SNotification& oNotification) override
	{
		if (oNotification.eType == CDriveWatcher::ENotificationType::DriveArrival)
			std::cout << "Drive arrived: " << oNotification.chDriveLetter << std::endl;
	}
};


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


	// Init watcher
	CDriveWatcher oWatcher;
	oWatcher.Init();

	// Set listener
	CDriveWatcher::INotificationListenerPtr pListener = CDriveWatcher::INotificationListenerPtr(new CNotificationListener);
	oWatcher.SetNotificationListener(pListener);
	oWatcher.Start(); // Starts asynchronously 

	Sleep(10000);

	// Clear
	oWatcher.Stop();
	oWatcher.Reset();

	return 0;
}