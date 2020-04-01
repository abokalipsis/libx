/*
Implementation file for the drive watcher
*/


//! Includes
#include "drive_watcher.h"
#include <string>
#include <comutil.h>


CDriveWatcher::CDriveWatcher()
	: m_pListener()
{
}

CDriveWatcher::~CDriveWatcher()
{
	Reset();
}

bool CDriveWatcher::Init()
{
	return Base::Init();
}

void CDriveWatcher::Reset()
{
	m_pListener.reset();
	Base::Reset();
}

bool CDriveWatcher::Start()
{
	std::string sQueryString("SELECT * FROM __InstanceCreationEvent WITHIN 0.1 WHERE TargetInstance ISA 'Win32_LogicalDisk'");
	return Base::StartListening(sQueryString);
}

void CDriveWatcher::Stop()
{
	Base::StopListening();
}

bool CDriveWatcher::IsActive() const
{
	return Base::IsListening();
}

void CDriveWatcher::SetNotificationListener(INotificationListenerPtr pListener)
{
	m_pListener = pListener;
}

CDriveWatcher::INotificationListenerPtr CDriveWatcher::GetNotificationListener()
{
	return m_pListener;
}

void CDriveWatcher::NotificationQueryCallbackFunc(int nObjectCount, IWbemClassObject** arrObject)
{
	if (m_pListener == nullptr || nObjectCount < 1 || arrObject == nullptr)
		return;

	for (int i = 0; i < nObjectCount; i++)
	{
		IWbemClassObject* pObject = arrObject[i];
		// Get instance
		_variant_t instance;
		HRESULT hResult = pObject->Get(TEXT("TargetInstance"), 0, &instance, NULL, NULL);
		if (SUCCEEDED(hResult) && instance.vt != VT_NULL && instance.vt != VT_EMPTY)
		{
			// Query instance interface
			IUnknown* pInstance = instance;
			IWbemClassObject* pDrive = nullptr;
			hResult = pInstance->QueryInterface(IID_IWbemClassObject, reinterpret_cast<void**>(&pDrive));
			if (SUCCEEDED(hResult) && pDrive != nullptr)
			{
				// Get drive letter
				_variant_t driveLetter;
				hResult = pDrive->Get(TEXT("Name"), 0, &driveLetter, NULL, NULL);
				if (SUCCEEDED(hResult) && driveLetter.vt != VT_NULL && driveLetter.vt != VT_EMPTY)
				{
					std::string sDriveLetter = driveLetter.operator _bstr_t().operator const char* ();
					if (!sDriveLetter.empty())
					{
						// Notify
						SNotification oNotification;
						oNotification.eType = ENotificationType::DriveArrival;
						oNotification.chDriveLetter = sDriveLetter[0];
						m_pListener->Notify(oNotification);
					}
				}
				pDrive->Release();
			}
		}
	}
}