/*
Declaration file for the Wbem notifications
Requirements: COM API initialization
*/


//! Include guard
#pragma once


//! Macro definitions
#define _WIN32_DCOM


//! Includes
#include <string>
#include <Wbemidl.h>


//! Type declarations
class CWbemQuerySink;


//! Abstract class CWbemNotificationAsyncQueryAbstract
class CWbemNotificationAsyncQueryAbstract
{
public: //! Constructors and destructor
	CWbemNotificationAsyncQueryAbstract();
	~CWbemNotificationAsyncQueryAbstract();

protected: //! Implementation
	// Initializes the notifier
	bool Init();
	// Resets the notifier
	bool Reset();

	// Returns true is it's in listening
	bool IsListening() const;
	// Starts listening notifications corresponding to the specified WQL
	bool StartListening(const std::string& sWQL);
	// Stops listening
	bool StopListening();

	// Notification query callback function
	virtual void NotificationQueryCallbackFunc(int nObjectCount, IWbemClassObject** arrObject) = 0;

protected: //! Members
	bool						m_bListening;			// Indicates current state

	IWbemLocator*				m_pWbemLocator;			// Wbem locator
	IWbemServices*				m_pWbemServices;		// Wbem services
	IUnsecuredApartment*		m_pUnsecApartment;		// Async calls helper
	IWbemObjectSink*			m_pSink;				// Original sink
	IUnknown*					m_pStabObject;			// Stab-situate sink object
	IWbemObjectSink*			m_pStabSink;			// Stab-situate sink interface

private:
	friend CWbemQuerySink; // This is helper class, I didn't want to write into this class
};


//! Class CWbemQuerySink
class CWbemQuerySink : public IWbemObjectSink
{
public: //! Constructors and destructor
	CWbemQuerySink(CWbemNotificationAsyncQueryAbstract* pMaster);
	virtual ~CWbemQuerySink() = default;

public: //! Interface (overrides base interface)
	ULONG STDMETHODCALLTYPE AddRef() override;
	ULONG STDMETHODCALLTYPE Release() override;
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override;
	HRESULT STDMETHODCALLTYPE Indicate(LONG lObjectCount, IWbemClassObject __RPC_FAR* __RPC_FAR* apObjArray) override;
	HRESULT STDMETHODCALLTYPE SetStatus(LONG lFlags, HRESULT hResult, BSTR strParam, IWbemClassObject __RPC_FAR* pObjParam) override;

private: //! Members
	LONG									m_lRefCount;		// Reference count
	CWbemNotificationAsyncQueryAbstract*	m_pMaster;			// Master
};