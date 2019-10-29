#pragma once
#include "stdafx.h"
#include <vector>
#include <process.h>
#include <EdkGuid.h>
#include <EdkMdb.h>

namespace Microsoft 
{
	namespace Samples 
	{
		class CMailboxObject
		{
		public:
			CMailboxObject(LPMAPISESSION lpMapiSession);
			~CMailboxObject();
			STDMETHOD(Logon)();
			STDMETHOD(OpenInbox)(LPMAPIFOLDER * lppFldr);
		private:
			STDMETHOD(GetStoreEntryIDFromSession)(ULONG * pCbEid, LPENTRYID * lppEntryID);
			STDMETHOD(LogonToStore)(ULONG ulFlags, ULONG Cb, LPENTRYID lpEntryID, LPMDB * lppMDB);
		private:
			CComPtr<IMsgStore>		m_lpMsgStore;
			CComPtr<IMAPISession>	m_lpMAPISession;
		};
	} //end namespace Samples
} // end namespace Microsoft
