/*
Declaration file for the Windows build-in BitLocker wrapper (OS version >= Vista)
Requirements: COM library initialization
*/


//! Include guard
#pragma once


//! Macro definitions
#define _WIN32_DCOM


//! Includes
#include <string>
#include <vector>
#include <unordered_map>
#include <comutil.h>
#include <Wbemidl.h>


//! Class CBitLocker (Singleton)
class CBitLocker
{
public://! Type definitions
	// Drive lock status
	enum class EDriveLockStatus
	{
		Invalid,
		Unprotected,
		Locked,
		Unlocked
	};

	// Conversion status notifier interface
	class IProgressNotifier;

private: //! Constructors and destructor
	CBitLocker();
	~CBitLocker();

public: //! Interface
	// Initializes the locker to start work
	bool Init();
	// Resets the locker
	void Reset();

	// Returns currently available lockable drive letters
	std::list<char> GetLockableDriveLetters() const;

	// Returns the specified drive lock status
	EDriveLockStatus GetDriveLockStatus(char chLetter) const;
	// Enables locker on the specified drive
	bool EnableDriveLocker(char chLetter, const std::string& sPassword, IProgressNotifier* pProgressNotifier);
	// Disables locker on the specified drive
	bool DisableDriveLocker(char chLetter, IProgressNotifier* pProgressNotifier);

	// Locks unlocked drive
	bool LockDrive(char chLetter);

	// Returns true if the drive has password
	bool HasDrivePassword(char chLetter) const;
	// Unlocks the specified drive by password
	bool UnlockDriveByPassword(char chLetter, const std::string& sPassword);
	// Changes drive password (Note: it will delete all another passwords)
	bool ChangeDrivePassword(char chLetter, const std::string& sNewPassword);

	// Returns true if the drive has numerical password
	bool HasDriveNumericalPassword(char chLetter) const;
	// Unlocks the specified drive by numerical password
	bool UnlockDriveByNumericalPassword(char chLetter, const std::string& sNumericalPassword);
	// Sets numerical password on drive (Note: currently drive may have only one numerical password)
	bool SetDriveNumericalPassword(char chLetter, const std::string& sNumericalPassword);
	// Removes numerical password from drive
	bool RemoveDriveNumericalPassword(char chLetter);
	
	// Returns true if auto-unlock enabled of the specified drive
	bool IsDriveAutoUnlock(char chLetter) const;
	// Enable/Disable drive auto unlock flag for the specified drive
	bool SetDriveAutoUnlock(char chLetter, bool bEnable);

	// Sets\Gets drive identifier
	bool SetDriveIdentifier(char chLetter, const std::string& sIdentifier);
	std::string GetDriveIdentifier(char chLetter) const;

public: //! Static members
	// Format policies of numerical password
	static const size_t s_nNumPassGroupCount		= 8;
	static const char	s_chNumPassGroupSeparator	= '-';
	static const size_t s_nNumPassGroupValueMin		= 1;
	static const size_t s_nNumPassGroupValueMax		= 720895;
	static const size_t s_NumPassGroupValueMultiply = 11;

public: //! Static methods
	// Returns the instance
	static CBitLocker& Instance();

	// Generates numerical password
	static std::string GenerateNumericalPassword();
	// Returns true is the specified numerical password is valid
	static bool IsValidNumericalPassword(const std::string& sNumericalPassword);

public:
	// Conversion status notifier interface
	class IProgressNotifier
	{
	public: //! Type definitions
		enum class EStatus
		{
			Invalid,

			Encrypted,
			Decrypted,
			EncryptionInProgress,
			DecryptionInProgress,
			EncryptionPaused,
			DecryptionPaused
		};

	public: //! Interface
		virtual void NotifyStatus(EStatus eStatus, double fPersentage) = 0;
	};

protected: //! Implementation
	// Returns key protectors IDs of the specified protector type
	std::vector<std::string> GetDriveKeyProtectors(char chLetter, int nProtectorType) const;
	// Adds external key on the specified drive returns ID
	std::string ProtectDriveByExternalKey(char chLetter);
	// Enable/Disable all key protectors of the specified drives
	bool EnableDriveAllKeyProtectors(char chLetter, bool bEnable);
	// Deletes key protector with the specified ID
	bool DeleteDriveKeyProtector(char chLetter, const std::string& sID);
	// Deletes all key protectors of the specified drive
	bool DeleteDriveAllKeyProtectors(char chLetter);

	// Returns conversion status
	bool GetConversionStatus(char chLetter, IProgressNotifier::EStatus& eStatus, double& fEncryptionPersentage);

	typedef std::unordered_map<std::string, _variant_t> tArgs;
	// Calls <sMethodName> method on <chLetter> drive with <inArgs> input parameters, returns true if the call success
	// to get output arguments need to pass them names, note that the return value is always added by the function with "ReturnValue" name
	bool CallDriveMethod(char chLetter, const std::string& sMethodName, const tArgs& inArgs, tArgs& outArgs) const;

private: //! Members
	IWbemLocator*		m_pWbemLocator;		// Wbem locator
	IWbemServices*		m_pWbemServices;	// Wbem services
};