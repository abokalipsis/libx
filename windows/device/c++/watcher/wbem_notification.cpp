/*
Implementation file for the Wbem notifications
*/


//! Includes
#include "wbem_notification.h"
#include <comutil.h>


CWbemNotificationAsyncQueryAbstract::CWbemNotificationAsyncQueryAbstract()
	: m_bListening(false),
	m_pWbemLocator(nullptr),
	m_pWbemServices(nullptr),
	m_pUnsecApartment(nullptr),
	m_pSink(nullptr),
	m_pStabObject(nullptr),
	m_pStabSink(nullptr)
{
}

CWbemNotificationAsyncQueryAbstract::~CWbemNotificationAsyncQueryAbstract()
{
	Reset();
}

bool CWbemNotificationAsyncQueryAbstract::Init()
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

	// Create unsecured apartment
	hResult = CoCreateInstance(CLSID_UnsecuredApartment, NULL, CLSCTX_LOCAL_SERVER, IID_IUnsecuredApartment, (void**)&m_pUnsecApartment);
	if (FAILED(hResult) || m_pUnsecApartment == nullptr)
	{
		Reset();
		return false;
	}

	// Create sink
	m_pSink = new CWbemQuerySink(this);
	m_pSink->AddRef();

	// Create stab-situate object
	m_pUnsecApartment->CreateObjectStub(m_pSink, &m_pStabObject);
	if (m_pStabObject == nullptr)
	{
		Reset();
		return false;
	}

	// Get stab-situate object interface
	m_pStabObject->QueryInterface(IID_IWbemObjectSink, (void**)&m_pStabSink);
	if (m_pStabSink == nullptr)
	{
		Reset();
		return false;
	}

	return true;
}

bool CWbemNotificationAsyncQueryAbstract::Reset()
{
	if (m_bListening)
		return false;

	if (m_pStabSink != nullptr)
	{
		m_pStabSink->Release();
		m_pStabSink = nullptr;
	}

	if (m_pStabObject != nullptr)
	{
		m_pStabObject->Release();
		m_pStabObject = nullptr;
	}

	if (m_pSink != nullptr)
	{
		m_pSink->Release();
		m_pSink = nullptr;
	}

	if (m_pUnsecApartment != nullptr)
	{
		m_pUnsecApartment->Release();
		m_pUnsecApartment = nullptr;
	}

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

	return true;
}

bool CWbemNotificationAsyncQueryAbstract::IsListening() const
{
	return m_bListening;
}

bool CWbemNotificationAsyncQueryAbstract::StartListening(const std::string& sWQL)
{
	if (m_bListening)
		return false;

	if (m_pWbemServices == nullptr || m_pStabSink == nullptr)
		return false;

	// Execute a query asynchronously
	_bstr_t queryStirng(sWQL.c_str());
	HRESULT hResult = m_pWbemServices->ExecNotificationQueryAsync(_bstr_t("WQL"), queryStirng, WBEM_FLAG_SEND_STATUS, NULL, m_pStabSink);
	m_bListening = SUCCEEDED(hResult);

	return m_bListening;
}

bool CWbemNotificationAsyncQueryAbstract::StopListening()
{
	if (!m_bListening)
		return false;

	if (m_pWbemServices == nullptr || m_pStabSink == nullptr)
		return false;

	HRESULT hResult = m_pWbemServices->CancelAsyncCall(m_pStabSink);
	if (SUCCEEDED(hResult))
	{
		m_bListening = false;
		return true;
	}

	return false;
}


CWbemQuerySink::CWbemQuerySink(CWbemNotificationAsyncQueryAbstract* pMaster)
	: m_lRefCount(0),
	  m_pMaster(pMaster)
{
}

ULONG CWbemQuerySink::AddRef()
{
	return InterlockedIncrement(&m_lRefCount);
}

ULONG CWbemQuerySink::Release()
{
	LONG lRef = InterlockedDecrement(&m_lRefCount);
	if (lRef == 0)
		delete this;

	return lRef;
}

HRESULT CWbemQuerySink::QueryInterface(REFIID riid, void** ppv)
{
	if (riid == IID_IUnknown || riid == IID_IWbemObjectSink)
	{
		*ppv = (IWbemObjectSink*)this;
		AddRef();

		return WBEM_S_NO_ERROR;
	}
	else
		return E_NOINTERFACE;
}

HRESULT CWbemQuerySink::Indicate(long lObjectCount, IWbemClassObject __RPC_FAR* __RPC_FAR* apObjArray)
{
	if (m_pMaster == nullptr)
		return WBEM_E_CALL_CANCELLED;

	m_pMaster->NotificationQueryCallbackFunc(lObjectCount, apObjArray);

	return WBEM_S_NO_ERROR;
}

HRESULT CWbemQuerySink::SetStatus(LONG lFlags, HRESULT hResult, BSTR strParam, IWbemClassObject __RPC_FAR* pObjParam)
{
	return WBEM_S_NO_ERROR;
}