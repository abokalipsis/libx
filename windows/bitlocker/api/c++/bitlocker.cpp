/*
Definition file for the Windows build-in BitLocker wrapper
*/


//! Includes
#include "bitlocker.h"
#include <algorithm>
#include <cctype>
#include <time.h>
#include <comdef.h>


CBitLocker::CBitLocker()
	: m_pWbemLocator(nullptr),
	  m_pWbemServices(nullptr)
{
}

CBitLocker::~CBitLocker()
{
	Reset();
}

bool CBitLocker::Init()
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
	hResult = m_pWbemLocator->ConnectServer(_bstr_t("\\\\.\\root\\cimv2\\Security\\MicrosoftVolumeEncryption"), NULL, NULL, NULL, WBEM_FLAG_CONNECT_USE_MAX_WAIT, NULL, NULL, &m_pWbemServices);
	if (FAILED(hResult) || m_pWbemServices == nullptr)
	{
		Reset();
		return false;
	}

	return true;
}

void CBitLocker::Reset()
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

std::list<char> CBitLocker::GetLockableDriveLetters() const
{
	std::list<char> lstDriveLetters;
	if (m_pWbemServices == nullptr)
		return lstDriveLetters;

	// Execute a query
	IEnumWbemClassObject* pEnumerator = nullptr;
	HRESULT hResult = m_pWbemServices->ExecQuery(_bstr_t("WQL"), _bstr_t(std::string("SELECT * FROM Win32_EncryptableVolume").c_str()),
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
			hResult = pDrive->Get(L"DriveLetter", 0, &driveLetter, NULL, NULL);
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

CBitLocker::EDriveLockStatus CBitLocker::GetDriveLockStatus(char chLetter) const
{
	EDriveLockStatus eStatus = EDriveLockStatus::Invalid;
	if (m_pWbemServices == nullptr)
		return eStatus;

	// Execute a query
	IEnumWbemClassObject* pEnumerator = nullptr;
	HRESULT hResult = m_pWbemServices->ExecQuery(_bstr_t("WQL"), _bstr_t(std::string("SELECT * FROM Win32_EncryptableVolume WHERE DriveLetter='" + std::string(1, chLetter) + ":'").c_str()),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);
	if (FAILED(hResult) || pEnumerator == nullptr)
		return eStatus;

	// Get drive
	ULONG count = 0;
	IWbemClassObject* arrDrives[1] = { nullptr };
	hResult = pEnumerator->Next(WBEM_INFINITE, 1, arrDrives, &count);
	if (SUCCEEDED(hResult) && count > 0)
	{
		IWbemClassObject* pDrive = arrDrives[0];
		if (pDrive != nullptr)
		{
			// Get protection status
			_variant_t prop;
			hResult = pDrive->Get(L"ProtectionStatus", 0, &prop, NULL, NULL);
			if (SUCCEEDED(hResult) && prop.vt != VT_NULL && prop.vt != VT_EMPTY)
			{
				switch (prop.uintVal)
				{
				case 0:
					eStatus = EDriveLockStatus::Unprotected;
					break;
				case 1:
					eStatus = EDriveLockStatus::Unlocked;
					break;
				case 2:
					eStatus = EDriveLockStatus::Locked;
					break;
				default:
					break;
				}
			}
			pDrive->Release();
		}
	}
	pEnumerator->Release();

	return eStatus;
}

bool CBitLocker::EnableDriveLocker(char chLetter, const std::string& sPassword, IProgressNotifier* pProgressNotifier)
{
	bool bSuccess = false;

	// Add key protector
	tArgs inArgs, outArgs;
	inArgs.insert({ "Passphrase", _variant_t(sPassword.c_str()) });
	const std::string sVolumeKeyProtectorID = "VolumeKeyProtectorID";
	outArgs.insert({ sVolumeKeyProtectorID, _variant_t() });
	if (CallDriveMethod(chLetter, "ProtectKeyWithPassphrase", inArgs, outArgs))
	{
		_variant_t value = outArgs["ReturnValue"];
		if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		{
			if (value.uintVal == S_OK)
			{
				value = outArgs[sVolumeKeyProtectorID];
				if (value.vt != VT_NULL && value.vt != VT_EMPTY)
					bSuccess = (value.operator _bstr_t().length() > 0);
			}
		}
	}

	// Encrypt
	if (bSuccess)
	{
		inArgs.clear();
		outArgs.clear();
		inArgs.insert({ "EncryptionMethod", _variant_t("3") }); // AES 128
		inArgs.insert({ "EncryptionFlags", _variant_t("1") }); // Used space only
		bSuccess = CallDriveMethod(chLetter, "Encrypt", inArgs, outArgs);
		if (bSuccess)
		{
			// Wait for finish
			IProgressNotifier::EStatus eStatus = IProgressNotifier::EStatus::Invalid;
			double fPercentage = 0.0;
			do
			{
				Sleep(500);

				// Get status
				bool bOk = GetConversionStatus(chLetter, eStatus, fPercentage);
				if (!bOk)
					break;

				// Notify
				if (pProgressNotifier != nullptr)
					pProgressNotifier->NotifyStatus(eStatus, fPercentage);
			} while (eStatus == IProgressNotifier::EStatus::EncryptionInProgress);

			// Check output
			_variant_t value = outArgs["ReturnValue"];
			if (value.vt != VT_NULL && value.vt != VT_EMPTY)
				bSuccess = (value.uintVal == S_OK);
		}

		// Delete added protector if not success
		if (!bSuccess)
			DeleteDriveAllKeyProtectors(chLetter);
	}
 
	return bSuccess;
}

bool CBitLocker::DisableDriveLocker(char chLetter, IProgressNotifier* pProgressNotifier)
{
	// Decrypt
	tArgs inArgs, outArgs;
	bool bSuccess = CallDriveMethod(chLetter, "Decrypt", inArgs, outArgs);
	if (bSuccess)
	{
		// Wait for finish
		IProgressNotifier::EStatus eStatus = IProgressNotifier::EStatus::Invalid;
		double fPercentage = 0.0;
		do
		{
			Sleep(200);

			// Get status
			bool bOk = GetConversionStatus(chLetter, eStatus, fPercentage);
			if (!bOk)
				break;

			// Notify
			if (pProgressNotifier != nullptr)
				pProgressNotifier->NotifyStatus(eStatus, 100 - fPercentage);
		} while (eStatus == IProgressNotifier::EStatus::DecryptionInProgress);

		// Check output
		_variant_t value = outArgs["ReturnValue"];
		if (value.vt != VT_NULL && value.vt != VT_EMPTY)
			bSuccess = (value.uintVal == S_OK);
	}

	return bSuccess;
}

bool CBitLocker::LockDrive(char chLetter)
{
	bool bLocked = false;
	tArgs inArgs, outArgs;
	if (!CallDriveMethod(chLetter, "Lock", inArgs, outArgs))
		return bLocked;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		bLocked = (value.uintVal == S_OK);

	return bLocked;
}

bool CBitLocker::HasDrivePassword(char chLetter) const
{
	std::vector<std::string> vecIDs = GetDriveKeyProtectors(chLetter, 8);
	return (!vecIDs.empty() && !vecIDs[0].empty());
}

bool CBitLocker::UnlockDriveByPassword(char chLetter, const std::string& sPassword)
{
	bool bUnlocked = false;
	tArgs inArgs, outArgs;
	inArgs.insert({ "Passphrase", _variant_t(sPassword.c_str()) });
	if (!CallDriveMethod(chLetter, "UnlockWithPassphrase", inArgs, outArgs))
		return bUnlocked;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		bUnlocked = (value.uintVal == S_OK);

	return bUnlocked;
}

bool CBitLocker::ChangeDrivePassword(char chLetter, const std::string& sNewPassword)
{
	bool bChanged = false;
	// First disables and deletes all key protectors
	bool bDisabled = EnableDriveAllKeyProtectors(chLetter, false);
	if (bDisabled)
	{
		bool bDeleted = DeleteDriveAllKeyProtectors(chLetter);
		if (bDeleted)
		{
			tArgs inArgs, outArgs;
			inArgs.insert({ "Passphrase", _variant_t(sNewPassword.c_str()) });
			const std::string sVolumeKeyProtectorID = "VolumeKeyProtectorID";
			outArgs.insert({ sVolumeKeyProtectorID, _variant_t() });
			if (CallDriveMethod(chLetter, "ProtectKeyWithPassphrase", inArgs, outArgs))
			{
				_variant_t value = outArgs["ReturnValue"];
				if (value.vt != VT_NULL && value.vt != VT_EMPTY)
				{
					if (value.uintVal == S_OK)
					{
						value = outArgs[sVolumeKeyProtectorID];
						if (value.vt != VT_NULL && value.vt != VT_EMPTY)
							bChanged = (value.operator _bstr_t().length() > 0);
					}
				}
			}
		}
	}

	// Re-enable
	EnableDriveAllKeyProtectors(chLetter, true);

	return bChanged;
}

bool CBitLocker::HasDriveNumericalPassword(char chLetter) const
{
	std::vector<std::string> vecIDs = GetDriveKeyProtectors(chLetter, 3);
	return (!vecIDs.empty() && !vecIDs[0].empty());
}

bool CBitLocker::UnlockDriveByNumericalPassword(char chLetter, const std::string& sNumericalPassword)
{
	bool bUnlocked = false;
	tArgs inArgs, outArgs;
	inArgs.insert({ "NumericalPassword", _variant_t(sNumericalPassword.c_str()) });
	if (!CallDriveMethod(chLetter, "UnlockWithNumericalPassword", inArgs, outArgs))
		return bUnlocked;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		bUnlocked = (value.uintVal == S_OK);

	return bUnlocked;
}

bool CBitLocker::SetDriveNumericalPassword(char chLetter, const std::string& sNumericalPassword)
{
	bool bSuccess = false;
	if (HasDriveNumericalPassword(chLetter) && RemoveDriveNumericalPassword(chLetter))
		return bSuccess;

	tArgs inArgs, outArgs;
	inArgs.insert({ "NumericalPassword", _variant_t(sNumericalPassword.c_str()) });
	const std::string sVolumeKeyProtectorID("VolumeKeyProtectorID");
	outArgs.insert({ sVolumeKeyProtectorID, _variant_t() });
	std::string sMethodName;
	if (!CallDriveMethod(chLetter, "ProtectKeyWithNumericalPassword", inArgs, outArgs))
		return bSuccess;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
	{
		if (value.uintVal == S_OK)
		{
			value = outArgs[sVolumeKeyProtectorID];
			if (value.vt != VT_NULL && value.vt != VT_EMPTY)
				bSuccess = !std::string(value.operator _bstr_t().operator const char* ()).empty();
		}
	}

	return bSuccess;
}

bool CBitLocker::RemoveDriveNumericalPassword(char chLetter)
{
	bool bDeleted = false;
	std::vector<std::string> vecIDs = GetDriveKeyProtectors(chLetter, 3);
	for (const std::string& sID : vecIDs)
		bDeleted = DeleteDriveKeyProtector(chLetter, sID);

	return bDeleted;
}

bool CBitLocker::IsDriveAutoUnlock(char chLetter) const
{
	bool bEnabled = false;
	tArgs inArgs, outArgs;
	const std::string sIsAutoUnlockEnabled("IsAutoUnlockEnabled");
	outArgs.insert({ sIsAutoUnlockEnabled, _variant_t() });
	if (!CallDriveMethod(chLetter, "IsAutoUnlockEnabled", inArgs, outArgs))
		return bEnabled;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
	{
		if (value.uintVal == S_OK)
		{
			value = outArgs[sIsAutoUnlockEnabled];
			if (value.vt != VT_NULL && value.vt != VT_EMPTY)
				bEnabled = value.operator bool();
		}
	}

	return bEnabled;
}

bool CBitLocker::SetDriveAutoUnlock(char chLetter, bool bEnable)
{
	if ((bEnable && IsDriveAutoUnlock(chLetter)) || (!bEnable && !IsDriveAutoUnlock(chLetter)))
		return true;

	bool bSuccess = false;
	tArgs inArgs, outArgs;
	if (bEnable)
	{
		std::vector<std::string> vecIDs = GetDriveKeyProtectors(chLetter, 2);
		std::string sID;
		if (!vecIDs.empty())
			sID = vecIDs[0];
		else
			sID = ProtectDriveByExternalKey(chLetter);
		if (!sID.empty())
		{
			inArgs.insert({ "VolumeKeyProtectorID", _variant_t(sID.c_str()) });
			if (!CallDriveMethod(chLetter, "EnableAutoUnlock", inArgs, outArgs))
				return bSuccess;
		}
		else
			return bSuccess;
	}
	else
	{
		if (!CallDriveMethod(chLetter, "DisableAutoUnlock", inArgs, outArgs))
			return bSuccess;

		std::vector<std::string> vecIDs = GetDriveKeyProtectors(chLetter, 2);
		for (const std::string& sID : vecIDs)
			DeleteDriveKeyProtector(chLetter, sID);
	}

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		bSuccess = (value.uintVal == S_OK);

	return bSuccess;
}

bool CBitLocker::SetDriveIdentifier(char chLetter, const std::string& sIdentifier)
{
	bool bSuccess = false;
	tArgs inArgs, outArgs;
	inArgs.insert({ "IdentificationField", _variant_t(sIdentifier.c_str()) });
	if (!CallDriveMethod(chLetter, "SetIdentificationField", inArgs, outArgs))
		return bSuccess;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		bSuccess = (value.uintVal == S_OK);

	return bSuccess;
}

std::string CBitLocker::GetDriveIdentifier(char chLetter) const
{
	std::string sIdentifier;
	tArgs inArgs, outArgs;
	const std::string sIdentificationField("IdentificationField");
	outArgs.insert({ sIdentificationField, _variant_t() });
	if (!CallDriveMethod(chLetter, "GetIdentificationField", inArgs, outArgs))
		return sIdentifier;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
	{
		if (value.uintVal == S_OK)
		{
			value = outArgs[sIdentificationField];
			if (value.vt != VT_NULL && value.vt != VT_EMPTY)
				sIdentifier = value.operator _bstr_t().operator const char* ();
		}
	}

	return sIdentifier;
}

CBitLocker& CBitLocker::Instance()
{
	static CBitLocker oInstance;
	return oInstance;
}

std::string CBitLocker::GenerateNumericalPassword()
{
	std::string sNumericalPassword;
	if (s_nNumPassGroupValueMax - s_nNumPassGroupValueMin < s_NumPassGroupValueMultiply || s_NumPassGroupValueMultiply == 0)
		return sNumericalPassword;

	// Lambda expression to make group according to group policies
	auto GenerateGroup = []()
	{
		// Generate number
		size_t nGroupBaseMin = s_nNumPassGroupValueMin / s_NumPassGroupValueMultiply;
		size_t nGroupBaseMax = s_nNumPassGroupValueMax / s_NumPassGroupValueMultiply;
		size_t nGroupNumber = nGroupBaseMin + size_t((double(nGroupBaseMax - nGroupBaseMin) / double(RAND_MAX)) * std::rand());
		nGroupNumber *= s_NumPassGroupValueMultiply;
		// Create group
		std::string sGroup = std::to_string(nGroupNumber);
		if (size_t(sGroup.size()) < 6)
			sGroup = std::string(6 - sGroup.size(), '0') + sGroup;
		return sGroup;
	};

	// Make numerical password
	for (size_t i = 0; i < s_nNumPassGroupCount; ++i)
	{
		std::string sGroup = GenerateGroup();
		if (i == 0)
			sNumericalPassword += sGroup;
		else
			sNumericalPassword += s_chNumPassGroupSeparator + sGroup;
	}

	return sNumericalPassword;
}

bool CBitLocker::IsValidNumericalPassword(const std::string& sNumericalPassword)
{
	if (s_nNumPassGroupValueMax - s_nNumPassGroupValueMin < s_NumPassGroupValueMultiply || s_NumPassGroupValueMultiply == 0)
		return false;

	size_t nPos = 0;
	size_t nOffset = 0;
	std::vector<std::string> vecGroups;
	// Split into groups
	while ((nPos = sNumericalPassword.find(s_chNumPassGroupSeparator, nOffset)) != std::string::npos)
	{
		std::string sGroup = sNumericalPassword.substr(nOffset, nPos - nOffset);
		vecGroups.push_back(sGroup);
		nOffset = nPos + 1;
	}
	if (nOffset != std::string::npos && nOffset < sNumericalPassword.size())
	{
		std::string sGroup = sNumericalPassword.substr(nOffset);
		vecGroups.push_back(sGroup);
	}
	// Check group count
	if (vecGroups.size() != s_nNumPassGroupCount)
		return false;

	// Check group value
	for (const std::string& sGroup : vecGroups)
	{
		std::string::const_iterator it = sGroup.begin();
		bool bIsIntegerNumber = (!sGroup.empty() && std::find_if(sGroup.begin(), sGroup.end(), [](unsigned char c) { return !std::isdigit(c); }) == sGroup.end());
		if (!bIsIntegerNumber)
			return false;
		size_t nNumber = std::stoi(sGroup);
		if (nNumber % s_NumPassGroupValueMultiply != 0 || nNumber < s_nNumPassGroupValueMin || nNumber > s_nNumPassGroupValueMax)
			return false;
	}

	return true;
}

std::vector<std::string> CBitLocker::GetDriveKeyProtectors(char chLetter, int nProtectorType) const
{
	std::vector<std::string> sKeyProtectorsIDs;
	tArgs inArgs, outArgs;
	inArgs.insert({ "KeyProtectorType", _variant_t(std::to_string(nProtectorType).c_str()) });
	const std::string sVolumeKeyProtectorID("VolumeKeyProtectorID");
	outArgs.insert({ sVolumeKeyProtectorID, _variant_t() });
	if (!CallDriveMethod(chLetter, "GetKeyProtectors", inArgs, outArgs))
		return sKeyProtectorsIDs;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
	{
		if (value.uintVal == S_OK)
		{
			value = outArgs[sVolumeKeyProtectorID];
			if (value.vt != VT_NULL && value.vt != VT_EMPTY)
			{
				SAFEARRAY* saValues = value.parray;
				BSTR* pVals = nullptr;
				HRESULT hr = SafeArrayAccessData(saValues, (void**)&pVals);
				if (SUCCEEDED(hr))
				{
					long lowerBound = 0, upperBound = 0;
					SafeArrayGetLBound(saValues, 1, &lowerBound);
					SafeArrayGetUBound(saValues, 1, &upperBound);
					long elemetsCount = upperBound - lowerBound + 1;
					for (int i = 0; i < elemetsCount; ++i)
					{
						_bstr_t bstr(pVals[i]);
						std::string sID = bstr.operator const char* ();
						sKeyProtectorsIDs.push_back(sID);
					}
					SafeArrayUnaccessData(saValues);
				}
			}
		}
	}

	return sKeyProtectorsIDs;
}

std::string CBitLocker::ProtectDriveByExternalKey(char chLetter)
{
	std::string sKeyProtectorID;
	tArgs inArgs, outArgs;
	const std::string sVolumeKeyProtectorID("VolumeKeyProtectorID");
	outArgs.insert({ sVolumeKeyProtectorID, _variant_t() });
	std::string sMethodName;
	if (!CallDriveMethod(chLetter, "ProtectKeyWithExternalKey", inArgs, outArgs))
		return sKeyProtectorID;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
	{
		if (value.uintVal == S_OK)
		{
			value = outArgs[sVolumeKeyProtectorID];
			if (value.vt != VT_NULL && value.vt != VT_EMPTY)
				sKeyProtectorID = value.operator _bstr_t().operator const char* ();
		}
	}

	return sKeyProtectorID;
}

bool CBitLocker::EnableDriveAllKeyProtectors(char chLetter, bool bEnable)
{
	bool bDisabled = false;
	tArgs inArgs, outArgs;
	std::string sMethodName;
	if (bEnable)
		sMethodName = "EnableKeyProtectors";
	else
		sMethodName = "DisableKeyProtectors";
	if (!CallDriveMethod(chLetter, sMethodName, inArgs, outArgs))
		return bDisabled;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		bDisabled = (value.uintVal == S_OK);

	return bDisabled;
}

bool CBitLocker::DeleteDriveKeyProtector(char chLetter, const std::string& sID)
{
	bool bDeleted = false;
	tArgs inArgs, outArgs;
	inArgs.insert({ "VolumeKeyProtectorID", _variant_t(sID.c_str()) });
	std::string sMethodName;
	if (!CallDriveMethod(chLetter, "DeleteKeyProtector", inArgs, outArgs))
		return bDeleted;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		bDeleted = (value.uintVal == S_OK);

	return bDeleted;
}

bool CBitLocker::DeleteDriveAllKeyProtectors(char chLetter)
{
	bool bDisabled = false;
	tArgs inArgs, outArgs;
	std::string sMethodName;
	if (!CallDriveMethod(chLetter, "DeleteKeyProtectors", inArgs, outArgs))
		return bDisabled;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
		bDisabled = (value.uintVal == S_OK);

	return bDisabled;
}

bool CBitLocker::GetConversionStatus(char chLetter, IProgressNotifier::EStatus& eStatus, double& fEncryptionPersentage)
{
	bool bSuccess = false;
	tArgs inArgs, outArgs;
	inArgs.insert({ "PrecisionFactor", _variant_t("1") });
	const std::string sConversionStatus("ConversionStatus");
	const std::string sEncryptionPercentage("EncryptionPercentage");
	outArgs.insert({ sConversionStatus, _variant_t() });
	outArgs.insert({ sEncryptionPercentage, _variant_t() });
	if (!CallDriveMethod(chLetter, "GetConversionStatus", inArgs, outArgs))
		return bSuccess;

	// Check output
	_variant_t value = outArgs["ReturnValue"];
	if (value.vt != VT_NULL && value.vt != VT_EMPTY)
	{
		if (value.uintVal == S_OK)
		{
			bSuccess = true;

			// Get status
			value = outArgs[sConversionStatus];
			if (value.vt != VT_NULL && value.vt != VT_EMPTY)
			{
				int nStatus = value.operator int();
				switch (nStatus)
				{
				case 0:
					eStatus = IProgressNotifier::EStatus::Decrypted;
					break;
				case 1:
					eStatus = IProgressNotifier::EStatus::Encrypted;
					break;
				case 2:
					eStatus = IProgressNotifier::EStatus::EncryptionInProgress;
					break;
				case 3:
					eStatus = IProgressNotifier::EStatus::DecryptionInProgress;
					break;
				case 4:
					eStatus = IProgressNotifier::EStatus::EncryptionPaused;
					break;
				case 5:
					eStatus = IProgressNotifier::EStatus::DecryptionPaused;
					break;
				default:
					break;
				}
			}

			// Get percentage
			value = outArgs[sEncryptionPercentage];
			if (value.vt != VT_NULL && value.vt != VT_EMPTY)
				fEncryptionPersentage = (value.operator int() / 10.0);
		}
	}

	return bSuccess;
}

bool CBitLocker::CallDriveMethod(char chLetter, const std::string& sMethodName, const tArgs& inArgs, tArgs& outArgs) const
{
	if (m_pWbemServices == nullptr)
		return false;

	// Execute a query
	IEnumWbemClassObject* pEnumerator = nullptr;
	HRESULT hResult = m_pWbemServices->ExecQuery(_bstr_t("WQL"), _bstr_t(std::string("SELECT * FROM Win32_EncryptableVolume WHERE DriveLetter='" + std::string(1, chLetter) + ":'").c_str()),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);
	if (FAILED(hResult) || pEnumerator == nullptr)
		return false;

	// Get drive object
	ULONG count = 0;
	IWbemClassObject* arrDrives[1] = { nullptr };
	hResult = pEnumerator->Next(WBEM_INFINITE, 1, arrDrives, &count);
	if (FAILED(hResult) || count <= 0)
	{
		pEnumerator->Release();
		return false;
	}
	IWbemClassObject* pDrive = arrDrives[0];
	if (pDrive == nullptr)
	{
		pEnumerator->Release();
		return false;
	}

	// Get object class name
	_variant_t className;
	hResult = pDrive->Get(TEXT("__CLASS"), 0, &className, NULL, NULL);
	if (FAILED(hResult) || className.vt == VT_NULL || className.vt == VT_EMPTY)
	{
		pDrive->Release();
		pEnumerator->Release();
		return false;
	}

	// Get class
	IWbemClassObject* pClass = NULL;
	hResult = m_pWbemServices->GetObject(className.bstrVal, 0, NULL, &pClass, NULL);
	if (FAILED(hResult) || pClass == nullptr)
	{
		pDrive->Release();
		pEnumerator->Release();
		return false;
	}

	// Prepare method input parameters
	IWbemClassObject* pInArgsDefinition = NULL;
	IWbemClassObject* pInArgs = NULL;
	hResult = pClass->GetMethod(_bstr_t(sMethodName.c_str()), 0, &pInArgsDefinition, NULL);
	if (FAILED(hResult))
	{
		pClass->Release();
		pDrive->Release();
		pEnumerator->Release();
		return false;
	}
	std::list<std::pair<std::string, _variant_t>> lstInArgs;
	for (auto arg : inArgs)
		lstInArgs.push_back(arg);
	if (pInArgsDefinition != nullptr)
	{
		hResult = pInArgsDefinition->SpawnInstance(0, &pInArgs);
		if (SUCCEEDED(hResult) && pInArgs != nullptr)
		{
			for (auto& arg : lstInArgs)
				pInArgs->Put(_bstr_t(arg.first.c_str()), 0, &arg.second, 0);
		}
	}

	// Get object path
	_variant_t path;
	hResult = pDrive->Get(TEXT("__PATH"), 0, &path, NULL, NULL);
	if (FAILED(hResult) || path.vt == VT_NULL || path.vt == VT_EMPTY)
	{
		if (pInArgs != nullptr)
			pInArgs->Release();
		if (pInArgsDefinition != nullptr)
			pInArgsDefinition->Release();
		pClass->Release();
		pDrive->Release();
		pEnumerator->Release();
		return false;
	}

	// Call method
	IWbemClassObject* pOutArgs = NULL;
	hResult = m_pWbemServices->ExecMethod(path.bstrVal, _bstr_t(sMethodName.c_str()), 0, NULL, pInArgs, &pOutArgs, NULL);
	if (FAILED(hResult) || pOutArgs == nullptr)
	{
		if (pInArgs != nullptr)
			pInArgs->Release();
		if (pInArgsDefinition != nullptr)
			pInArgsDefinition->Release();
		pClass->Release();
		pDrive->Release();
		pEnumerator->Release();
		return false;
	}

	// Resolve output argument names
	std::vector<std::string> argNames;
	argNames.reserve(outArgs.size() + 1);
	for (const auto& arg : outArgs)
		argNames.push_back(arg.first);
	argNames.push_back(std::string("ReturnValue"));
	outArgs.clear();

	// Collect output arguments
	for (const std::string& sName : argNames)
	{
		_variant_t value;
		hResult = pOutArgs->Get(_bstr_t(sName.c_str()), 0, &value, NULL, 0);
		if (SUCCEEDED(hResult))
			outArgs.insert(std::make_pair(sName, value));
	}

	// Release object
	if (pInArgs != nullptr)
		pInArgs->Release();
	if (pInArgsDefinition != nullptr)
		pInArgsDefinition->Release();
	pClass->Release();
	pDrive->Release();
	pEnumerator->Release();

	return true;
}
