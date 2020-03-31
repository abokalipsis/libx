/*
Declaration file for named pipe
*/


//! Includes
#include "named_pipe.h"


CNamedPipe::CNamedPipe(const std::string& sName, EMode eMode)
	: m_eMode(eMode),
	  m_sName(sName),
	  m_hHandler(INVALID_HANDLE_VALUE)
{
}

CNamedPipe::~CNamedPipe()
{
	if (IsOpen())
		Close();
}

bool CNamedPipe::IsOpen()
{
	return (m_hHandler != INVALID_HANDLE_VALUE);
}

bool CNamedPipe::Open()
{
	if (IsOpen())
		return false;

	if (!m_sName.empty())
		m_hHandler = CreatePipe(m_sName, m_eMode);

	return (m_hHandler != INVALID_HANDLE_VALUE);
}

bool CNamedPipe::Close()
{
	if (!IsOpen())
		return false;

	if (CloseHandle(m_hHandler))
		m_hHandler = INVALID_HANDLE_VALUE;

	return (m_hHandler == INVALID_HANDLE_VALUE);
}

bool CNamedPipe::SendData(const std::vector<byte>& vecData)
{
	bool bOk = false;
	if (m_hHandler != INVALID_HANDLE_VALUE && m_eMode == EMode::Write)
	{
		void const* pBuffer = vecData.data();
		const DWORD nBufferSize = DWORD(vecData.size());
		DWORD nBufferWritten = 0;
		while (nBufferWritten != nBufferSize)
		{
			DWORD nWritten = 0;
			WriteFile(m_hHandler, pBuffer, nBufferSize - nBufferWritten, &nWritten, NULL);
			nBufferWritten += nWritten;
			pBuffer = ((byte*)pBuffer) + nWritten;
			nWritten = 0;
		}
		bOk = (nBufferWritten == nBufferSize);
	}

	return bOk;
}

bool CNamedPipe::ReceiveData(std::vector<byte>& vecData)
{
	bool bOk = false;
	if (m_hHandler != INVALID_HANDLE_VALUE && m_eMode == EMode::Read)
	{
		// Wait for someone to connect to the pipe
		if (ConnectNamedPipe(m_hHandler, NULL) != FALSE)
		{
			void* pBuffer = new byte[1024];
			DWORD nBufferRead = 0;
			bool bRead = true;
			while (bRead)
			{
				DWORD nRead = 0;
				BOOL bResult = ReadFile(m_hHandler, pBuffer, DWORD(1024), &nRead, NULL);
				if (!bResult && GetLastError() == ERROR_MORE_DATA)
					continue;
				if (bResult)
				{
					nBufferRead += nRead;
					nRead = 0;
				}
				bRead = false;
			}
			bOk = (nBufferRead > 0);
			if (bOk)
			{
				byte* pData = (byte*)(pBuffer);
				vecData.clear();
				vecData.resize(nBufferRead);
				for (size_t i = 0; i < nBufferRead; ++i)
					vecData[i] = pData[i];
			}
			delete[] pBuffer;
		}
		DisconnectNamedPipe(m_hHandler);
	}

	return bOk;
}

std::string CNamedPipe::ToString(const std::vector<byte>& vecData)
{
	std::string sStr;
	sStr.resize(vecData.size());
	for (size_t i = 0; i < vecData.size(); ++i)
		sStr[i] = char(vecData[i]);

	return sStr;
}

std::vector<byte> CNamedPipe::FromString(const std::string& sStr)
{
	std::vector<byte> vecData;
	vecData.resize(sStr.size());
	for (size_t i = 0; i < sStr.size(); ++i)
		vecData[i] = byte(sStr[i]);

	return vecData;
}

HANDLE CNamedPipe::CreatePipe(const std::string& sName, EMode eMode)
{
	// Set security
	PSECURITY_DESCRIPTOR psd = NULL;
	BYTE sd[SECURITY_DESCRIPTOR_MIN_LENGTH];
	psd = (PSECURITY_DESCRIPTOR)sd;
	InitializeSecurityDescriptor(psd, SECURITY_DESCRIPTOR_REVISION);
	SetSecurityDescriptorDacl(psd, TRUE, (PACL)NULL, FALSE);
	SECURITY_ATTRIBUTES sa = { sizeof(sa), psd, FALSE };

	std::string sPath = R"(\\.\pipe\)" + sName;
	HANDLE hHandler = INVALID_HANDLE_VALUE;
	if (eMode == EMode::Read)
		hHandler = CreateNamedPipeA(sPath.c_str(), PIPE_ACCESS_INBOUND, PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1, 1024 * 16, 1024 * 16, NMPWAIT_USE_DEFAULT_WAIT, &sa);
	else
		hHandler = CreateFileA(sPath.c_str(), GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, &sa);

	return hHandler;
}