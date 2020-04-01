/*
Declaration file for the drive controller
Requirements: COM library initialization
*/


//! Include guard
#pragma once


//! Macro definitions
#define _WIN32_DCOM


//! Includes
#include <list>
#include <Wbemidl.h>
#include <cfgmgr32.h>


//! Class CDriveController (Singleton)
class CDriveController
{
private: //! Constructors and destructor
	// Disable object creation
	CDriveController();
	~CDriveController();

public: //! Interface
	// Initializes the locker
	bool Init();
	// Resets the locker
	void Reset();

	// Returns currently available drive letters
	std::list<char> GetAvailableDriveLetters() const;
	// Returns currently available removable drive letters
	std::list<char> GetRemovableDriveLetters() const;
	// Removes the specified drive
	bool RemoveDrive(char driveLetter);

public: //! Static methods
	// Returns the instance
	static CDriveController& Instance();

private: //! Implementation
	DEVINST GetDrivesDeviceInstanceByDeviceNumber(long DeviceNumber, UINT DriveType, WCHAR* szDosDeviceName);

protected: //! Members
	IWbemLocator*		m_pWbemLocator;		// Wbem locator
	IWbemServices*		m_pWbemServices;	// Wbem services
};