/*
Declaration file for the drive watcher
Requirements: COM library initialization
*/


//! Include guard
#pragma once


//! Includes
#include "wbem_notification.h"
#include <memory>


//! Class CDriveWatcher
class CDriveWatcher : private CWbemNotificationAsyncQueryAbstract
{
	typedef CWbemNotificationAsyncQueryAbstract Base;

public: //! Type declarations
	class INotificationListener;
	typedef std::shared_ptr<INotificationListener> INotificationListenerPtr;

public: //! Constructors and destructor
	CDriveWatcher();
	~CDriveWatcher();

public: //! Interface
	// Initializes the watcher
	bool Init();
	// Resets the watcher
	void Reset();

	// Starts watching
	bool Start();
	// Stops watching
	void Stop();
	// Returns true is the watcher is active
	bool IsActive() const;

	// Registers notification listener (currently may notify only one listener)
	void SetNotificationListener(INotificationListenerPtr pListener);
	// Returns listener
	INotificationListenerPtr GetNotificationListener();

public: //! Type definitions
	// Notification type
	enum class ENotificationType
	{
		Invalid,
		DriveArrival
	};

	// Notification info
	struct SNotification
	{
		ENotificationType	eType = ENotificationType::Invalid;
		char				chDriveLetter = '\n';
	};

	// Notification listener interface
	class INotificationListener
	{
	public:
		virtual void Notify(const SNotification& oNotification) = 0;
	};


protected: //! Implementation
	// Notification query callback function
	void NotificationQueryCallbackFunc(int nObjectCount, IWbemClassObject** arrObject) override;

private: //! Members
	INotificationListenerPtr	m_pListener;	// Notification Listener
};