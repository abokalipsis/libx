/*
Declaration file for the drive locker
*/


//! Includes
#include "drive_controller.h"
#include <string>
#include <Setupapi.h>
#include <comdef.h>


CDriveController::CDriveController()
	: m_pWbemLocator(nullptr),
	m_pWbemServices(nullptr)
{
}

CDriveController::~CDriveController()
{
	Reset();
}

bool CDriveController::Init()
{
	// Check is already initialized
	if (m_pWbemLocator != nullptr)
		return false;

	HRESULT hResult = CoCreateInstance(CLSID_WbemAdministrativeLocator, NULL, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&m_pWbemLocator);
	if (FAILED(hResult) || m_pWbemLocator == nullptr)
	{
		Reset();
		return false;
	}

	// Connect to local Wbem service with current credentials
	hResult = m_pWbemLocator->ConnectServer(_bstr_t("\\\\.\\root\\cimv2"), NULL, NULL, NULL, WBEM_FLAG_CONNECT_USE_MAX_WAIT, NULL, NULL, &m_pWbemServices);
	if (FAILED(hResult) || m_pWbemServices == nullptr)
	{
		Reset();
		return false;
	}

	return true;
}

void CDriveController::Reset()
{
	if (m_pWbemServices != nullptr)
	{
		m_pWbemServices->Release();
		m_pWbemServices = nullptr;
	}

	if (m_pWbemLocator != nullptr)
	{
		m_pWbemLocator->Release();
		m_pWbemLocator = nullptr;
	}
}

std::list<char> CDriveController::GetAvailableDriveLetters() const
{
	std::list<char> lstDriveLetters;
	if (m_pWbemServices == nullptr)
		return lstDriveLetters;

	// Execute a query
	IEnumWbemClassObject* pEnumerator = nullptr;
	HRESULT hResult = m_pWbemServices->ExecQuery(_bstr_t("WQL"), _bstr_t(std::string("SELECT * FROM Win32_LogicalDisk").c_str()),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);
	if (FAILED(hResult) || pEnumerator == nullptr)
		return lstDriveLetters;

	ULONG count = 0;
	IWbemClassObject* arrDrives[1] = { nullptr };
	while (SUCCEEDED(pEnumerator->Next(WBEM_INFINITE, 1, arrDrives, &count)) && count > 0)
	{
		IWbemClassObject* pDrive = arrDrives[0];
		if (pDrive != nullptr)
		{
			// Get drive letter
			_variant_t driveLetter;
			hResult = pDrive->Get(TEXT("Name"), 0, &driveLetter, NULL, NULL);
			if (SUCCEEDED(hResult) && driveLetter.vt != VT_NULL && driveLetter.vt != VT_EMPTY)
			{
				std::string sDriveLetter = driveLetter.operator _bstr_t().operator const char* ();
				if (!sDriveLetter.empty())
					lstDriveLetters.push_back(sDriveLetter[0]);
			}
			pDrive->Release();
		}
	}
	pEnumerator->Release();

	return lstDriveLetters;
}

std::list<char> CDriveController::GetRemovableDriveLetters() const
{
	std::list<char> lstDriveLetters;
	if (m_pWbemServices == nullptr)
		return lstDriveLetters;

	// Execute a query
	IEnumWbemClassObject* pEnumerator = nullptr;
	HRESULT hResult = m_pWbemServices->ExecQuery(_bstr_t("WQL"), _bstr_t(std::string("SELECT * FROM Win32_LogicalDisk").c_str()),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);
	if (FAILED(hResult) || pEnumerator == nullptr)
		return lstDriveLetters;

	ULONG count = 0;
	IWbemClassObject* arrDrives[1] = { nullptr };
	while (SUCCEEDED(pEnumerator->Next(WBEM_INFINITE, 1, arrDrives, &count)) && count > 0)
	{
		IWbemClassObject* pDrive = arrDrives[0];
		if (pDrive != nullptr)
		{
			// Get drive type
			_variant_t driveType;
			hResult = pDrive->Get(TEXT("DriveType"), 0, &driveType, NULL, NULL);
			if (SUCCEEDED(hResult) && driveType.vt != VT_NULL && driveType.vt != VT_EMPTY)
			{
				int nDriveType = driveType.operator int();
				if (nDriveType == DRIVE_REMOVABLE)
				{
					// Get drive letter
					_variant_t driveLetter;
					hResult = pDrive->Get(TEXT("Name"), 0, &driveLetter, NULL, NULL);
					if (SUCCEEDED(hResult) && driveType.vt != VT_NULL && driveType.vt != VT_EMPTY)
					{
						std::string sDriveLetter = driveLetter.operator _bstr_t().operator const char* ();
						if (!sDriveLetter.empty())
							lstDriveLetters.push_back(sDriveLetter[0]);
					}
				}
			}
			pDrive->Release();
		}
	}
	pEnumerator->Release();

	return lstDriveLetters;
}

bool CDriveController::RemoveDrive(char driveLetter)
{
	// Make uppercase
	driveLetter &= ~0x20;
	// Check letter
	if (driveLetter < 'A' || driveLetter > 'Z') {
		return false;
	}

	// Root path
	WCHAR szRootPath[] = TEXT("X:\\");
	szRootPath[0] = driveLetter;
	// Device path
	WCHAR szDevicePath[] = TEXT("X:");
	szDevicePath[0] = driveLetter;
	// Volume access path
	WCHAR szVolumeAccessPath[] = TEXT("\\\\.\\X:");
	szVolumeAccessPath[4] = driveLetter;

	// Open the storage volume
	HANDLE hVolume = CreateFile(szVolumeAccessPath, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, NULL, NULL);
	if (hVolume == INVALID_HANDLE_VALUE) {
		return false;
	}
	// Get the volume's device number
	long nDeviceNumber = -1;
	STORAGE_DEVICE_NUMBER sdn;
	DWORD dwBytesReturned = 0;
	long nResult = DeviceIoControl(hVolume, IOCTL_STORAGE_GET_DEVICE_NUMBER, NULL, 0, &sdn, sizeof(sdn), &dwBytesReturned, NULL);
	if (nResult)
		nDeviceNumber = sdn.DeviceNumber;
	CloseHandle(hVolume);
	if (nDeviceNumber == -1)
		return false;

	// Get the dos device name (like \device\floppy0) to decide if it's a floppy or not
	WCHAR szDosDeviceName[MAX_PATH];
	nResult = QueryDosDevice(szDevicePath, szDosDeviceName, MAX_PATH);
	if (!nResult)
		return false;

	// Get the drive type which is required to match the device numbers correctly
	UINT DriveType = GetDriveType(szRootPath);
	// Get the device instance handle of the storage volume by means of a SetupDi enum and matching the device number
	DEVINST DevInst = GetDrivesDeviceInstanceByDeviceNumber(nDeviceNumber, DriveType, szDosDeviceName);
	if (DevInst == 0)
		return false;

	// Get drives's parent, e.g. the USB bridge, the SATA port, an IDE channel with two drives!
	DEVINST DevInstParent = 0;
	nResult = CM_Get_Parent(&DevInstParent, DevInst, 0);
	if (nResult != CR_SUCCESS)
		return false;

	// Tries some times
	bool bSuccess = false;
	for (long tries = 1; tries <= 5; tries++)
	{
		PNP_VETO_TYPE VetoType = PNP_VETO_TYPE::PNP_VetoTypeUnknown;
		WCHAR VetoName[MAX_PATH];
		VetoName[0] = 0;

		// Eject device
		nResult = CM_Request_Device_Eject(DevInstParent, &VetoType, VetoName, MAX_PATH, 0);
		bSuccess = (nResult == CR_SUCCESS && VetoType == PNP_VetoTypeUnknown);
		if (bSuccess)
			break;

		// Required to give the next tries a chance!
		Sleep(1500);
	}

	return bSuccess;
}

CDriveController& CDriveController::Instance()
{
	static CDriveController oInstance;
	return oInstance;
}

DEVINST CDriveController::GetDrivesDeviceInstanceByDeviceNumber(long DeviceNumber, UINT DriveType, WCHAR* szDosDeviceName)
{
	bool IsFloppy = (wcscmp(szDosDeviceName, TEXT("\\Floppy")) == 0);
	GUID* guid = nullptr;

	switch (DriveType) {
	case DRIVE_REMOVABLE:
		if (IsFloppy)
			guid = (GUID*)&GUID_DEVINTERFACE_FLOPPY;
		else
			guid = (GUID*)&GUID_DEVINTERFACE_DISK;
		break;
	case DRIVE_FIXED:
		guid = (GUID*)&GUID_DEVINTERFACE_DISK;
		break;
	case DRIVE_CDROM:
		guid = (GUID*)&GUID_DEVINTERFACE_CDROM;
		break;
	default:
		return 0;
	}

	// Get device interface info set handle for all devices attached to system
	HDEVINFO hDevInfo = SetupDiGetClassDevs(guid, NULL, NULL, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
	if (hDevInfo == INVALID_HANDLE_VALUE)
		return 0;

	// Retrieve a context structure for a device interface of a device information set
	BYTE Buf[1024];
	PSP_DEVICE_INTERFACE_DETAIL_DATA pspdidd = (PSP_DEVICE_INTERFACE_DETAIL_DATA)Buf;
	SP_DEVICE_INTERFACE_DATA spdid;
	spdid.cbSize = sizeof(spdid);
	SP_DEVINFO_DATA spdd;
	DWORD dwSize;

	DWORD dwIndex = 0;
	long nResult;
	while (true)
	{
		nResult = SetupDiEnumDeviceInterfaces(hDevInfo, NULL, guid, dwIndex, &spdid);
		if (!nResult)
			break;

		dwSize = 0;
		// Check the buffer size
		SetupDiGetDeviceInterfaceDetail(hDevInfo, &spdid, NULL, 0, &dwSize, NULL);

		if (dwSize != 0 && dwSize <= sizeof(Buf))
		{
			// 5 Bytes!
			pspdidd->cbSize = sizeof(*pspdidd);
			ZeroMemory(&spdd, sizeof(spdd));
			spdd.cbSize = sizeof(spdd);

			nResult = SetupDiGetDeviceInterfaceDetail(hDevInfo, &spdid, pspdidd, dwSize, &dwSize, &spdd);
			if (nResult)
			{
				// Open the disk or CD-ROM or floppy
				HANDLE hDevice = CreateFile(pspdidd->DevicePath, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
				if (hDevice != INVALID_HANDLE_VALUE)
				{
					// Get its device number
					STORAGE_DEVICE_NUMBER sdn;
					DWORD dwBytesReturned = 0;
					nResult = DeviceIoControl(hDevice, IOCTL_STORAGE_GET_DEVICE_NUMBER, NULL, 0, &sdn, sizeof(sdn), &dwBytesReturned, NULL);
					if (nResult)
					{
						// Match the given device number with the one of the current device
						if (DeviceNumber == (long)sdn.DeviceNumber)
						{
							CloseHandle(hDevice);
							SetupDiDestroyDeviceInfoList(hDevInfo);
							return spdd.DevInst;
						}
					}
					CloseHandle(hDevice);
				}
			}
		}
		dwIndex++;
	}

	SetupDiDestroyDeviceInfoList(hDevInfo);

	return 0;
}