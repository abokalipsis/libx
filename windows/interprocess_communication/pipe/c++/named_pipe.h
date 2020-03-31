/*
Declaration file for named pipe
*/


//! Include guard
#pragma once


//! Includes
#include <string>
#include <vector>
#include <windows.h>


//! Class CNamedPipe
class CNamedPipe
{
public: //! Types
	enum class EMode
	{
		Read,
		Write
	};

public: //! Constructors and destructor
	CNamedPipe(const std::string& sName, EMode eMode);
	~CNamedPipe();

public: //! Interface
	// Returns true if the pipe is opened and ready to use
	bool IsOpen();
	// Opens pipe and get ready to use
	bool Open();
	// Closes pipe
	bool Close();

	// Sends data
	bool SendData(const std::vector<byte>& vecData);
	// Waits until data will be received
	bool ReceiveData(std::vector<byte>& vecData);

public: //! Static helpers
	static std::string ToString(const std::vector<byte>& vecData);
	static std::vector<byte> FromString(const std::string& sStr);

protected: //! Static helpers
	static HANDLE CreatePipe(const std::string& sName, EMode eMode);

private: // Members
	std::string			m_sName;
	EMode				m_eMode;
	HANDLE				m_hHandler;
};