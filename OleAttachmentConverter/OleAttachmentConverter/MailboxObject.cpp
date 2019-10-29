#include "stdafx.h"
#include "MailboxObject.h"

using namespace std;

/// Constructs a MailObject given a MAPI Session
/// Note : the object never takes ownership of the MAPI Session to it still requires the calling code to Logoff of the MAPI Session with it's done.
Microsoft::Samples::CMailboxObject::CMailboxObject(LPMAPISESSION lpMapiSession)
{
	OutputDebugString(_T("Enter CMailboxObject::CMailboxObject"));
	m_lpMAPISession = lpMapiSession;
	m_lpMsgStore = nullptr;
	OutputDebugString(_T("Enter CMailboxObject::CMailboxObject"));
}

/// Destructor for the MailboxObject.
Microsoft::Samples::CMailboxObject::~CMailboxObject()
{
	OutputDebugString(_T("Enter CMailboxObject::~CMailboxObject"));
	m_lpMsgStore.Release();
	m_lpMsgStore = nullptr;
	m_lpMAPISession.Release();
	m_lpMAPISession = nullptr;
	OutputDebugString(_T("Exit CMailboxObject::~CMailboxObject"));
}

STDMETHODIMP Microsoft::Samples::CMailboxObject::GetStoreEntryIDFromSession(ULONG * pCbEid, LPENTRYID * lppEntryID)
{
	HRESULT						hRes = E_FAIL;
	CComPtr<IMAPITable>			lpMsgStoreTable = nullptr;
	SRestriction				sres;
	SPropValue					sprop;
	SPropertyRestriction		spropres;
	SBinary						bin;
	LPSRowSet					lpRows = NULL;
	LPENTRYID					lpEntryID = nullptr;
	ULONG						cbEid = 0;

	SizedSPropTagArray(2, spta) = { 2, { PR_DISPLAY_NAME, PR_ENTRYID } };

	if (pCbEid == nullptr || lppEntryID == nullptr)
		return E_INVALIDARG;

	*pCbEid = cbEid;
	*lppEntryID = nullptr;

	hRes = m_lpMAPISession->GetMsgStoresTable(0, &lpMsgStoreTable);

	if (FAILED(hRes) || nullptr == lpMsgStoreTable)
	{
		goto Cleanup;
	}

	ZeroMemory(&bin, sizeof(bin));
	ZeroMemory(&spropres, sizeof(spropres));
	ZeroMemory(&sprop, sizeof(sprop));
	ZeroMemory(&sres, sizeof(sres));

	bin.cb = 0x10;
	bin.lpb = (BYTE*)pbExchangeProviderPrimaryUserGuid;

	sprop.Value.bin = bin;

	spropres.ulPropTag = sprop.ulPropTag = PR_MDB_PROVIDER;
	spropres.lpProp = &sprop;
	spropres.relop = RELOP_EQ;

	sres.res.resProperty = spropres;
	sres.rt = RES_PROPERTY;

	hRes = HrQueryAllRows(lpMsgStoreTable, (LPSPropTagArray)&spta, &sres, 0, 0, &lpRows);

	if (FAILED(hRes) || lpRows == nullptr || lpRows->cRows != 1)
	{
		goto Cleanup;
	}

	if (lpRows->aRow[0].lpProps[1].ulPropTag != PR_ENTRYID)
	{
		goto Cleanup;
	}

	cbEid = lpRows->aRow[0].lpProps[1].Value.bin.cb;

	MAPIAllocateBuffer(cbEid, (LPVOID*)&lpEntryID);

	if (lpEntryID == nullptr)
	{
		hRes = E_OUTOFMEMORY;
		goto Cleanup;
	}

	ZeroMemory(lpEntryID, cbEid);
	CopyMemory(lpEntryID, lpRows->aRow[0].lpProps[1].Value.bin.lpb, cbEid);

	*pCbEid = cbEid;
	*lppEntryID = lpEntryID;
	hRes = S_OK;
Cleanup:
	if (lpRows)
		FreeProws(lpRows);
	return hRes;
}

STDMETHODIMP Microsoft::Samples::CMailboxObject::LogonToStore(ULONG ulFlags, ULONG Cb, LPENTRYID lpEntryID, LPMDB * lppMDB)
{
	HRESULT hRes = E_FAIL;
	LPMDB	pMsgStore = nullptr;

	if (m_lpMAPISession == nullptr || lppMDB == nullptr || Cb == 0 || lpEntryID == nullptr)
		return E_INVALIDARG;

	*lppMDB = nullptr;

	DWORD start = GetTickCount();

	hRes = m_lpMAPISession->OpenMsgStore(0,
										Cb,
										lpEntryID,
										NULL,
										ulFlags,
										&pMsgStore);

	DWORD end = GetTickCount();

	wcout << "COMPLETED! " << (end - start) << " msec" << endl;

	if (FAILED(hRes) || pMsgStore == nullptr)
	{
		wcout << "An error occurred whilst trying to get to the store. Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	*lppMDB = pMsgStore; // Transfer ownership
	hRes = S_OK;
Cleanup:
	return hRes;
}

STDMETHODIMP Microsoft::Samples::CMailboxObject::Logon()
{
	HRESULT		hRes = E_FAIL;
	LPENTRYID	lpEntryID = nullptr;
	ULONG		cbEID = 0;
	ULONG		ulFlags = MAPI_BEST_ACCESS | MDB_NO_DIALOG;

	hRes = GetStoreEntryIDFromSession(&cbEID, &lpEntryID);

	if (FAILED(hRes) || lpEntryID == nullptr || cbEID == 0)
	{
		goto Cleanup;
	}

	hRes = LogonToStore(ulFlags, cbEID, lpEntryID, &m_lpMsgStore);

	if (FAILED(hRes))
	{
		goto Cleanup;
	}

	wcout << "Successfully logged on to the primary store " << endl;
Cleanup:
	if (lpEntryID)
	{
		MAPIFreeBuffer(lpEntryID);
	}
	return hRes;
}

STDMETHODIMP Microsoft::Samples::CMailboxObject::OpenInbox(LPMAPIFOLDER * lppFldr)
{
	HRESULT										hRes = E_FAIL;
	LPENTRYID									lpEntryID = nullptr;
	ULONG										CbEid = 0;
	ULONG										ulObjType = 0;
	CComPtr<IUnknown>							lpUnk = nullptr;
	ULONG										ulFlags = MAPI_BEST_ACCESS;
	CComPtr<IUnknown>							pUnk = nullptr;
	CComQIPtr<IMAPIFolder, &IID_IMAPIFolder>	lpFldr;

	if (lppFldr == nullptr)
		return E_INVALIDARG;

	*lppFldr = nullptr;

	hRes = m_lpMsgStore->GetReceiveFolder(_T("IPM.Note"), fMapiUnicode, &CbEid, &lpEntryID, nullptr);
	
	if (FAILED(hRes) || lpEntryID == nullptr || CbEid == 0)
	{
		goto Cleanup;
	}

	hRes = m_lpMsgStore->OpenEntry(CbEid, lpEntryID, &IID_IMAPIFolder, ulFlags, &ulObjType, &pUnk);

	if (FAILED(hRes) || pUnk == nullptr || ulObjType == 0)
	{
		goto Cleanup;
	}

	lpFldr = pUnk;
	
	if (lpFldr != nullptr)
	{
		*lppFldr = lpFldr.Detach(); // Transfer ownership
	}

Cleanup:
	if (lpEntryID)
		MAPIFreeBuffer(lpEntryID);
	return hRes;
}


//STDMETHODIMP Microsoft::Samples::CMailboxObject::CreateStoreEntryID(LPSTR TargetStoreDN, LPSTR TargetMailboxDN, ULONG ulFlags, ULONG * pCbEID, LPENTRYID * lppEntryID)
//{
//	HRESULT hRes = E_FAIL;
//	ULONG cbEID;
//	LPENTRYID lpEntryID;
//	SPEXCHMANAGESTORE pManageStore;
//
//	if (pCbEID == nullptr || lppEntryID == nullptr || m_lpMsgStore == nullptr)
//		return E_INVALIDARG;
//
//	*lppEntryID = lpEntryID = nullptr;
//	*pCbEID = cbEID = 0;
//
//	pManageStore = m_lpMsgStore; // Use ATL to do the work of QI
//
//	if (pManageStore == nullptr)
//	{
//		goto Cleanup;
//	}
//	else {
//		hRes = S_OK;
//	}
//
//	hRes = pManageStore->CreateStoreEntryID(TargetStoreDN, TargetMailboxDN, ulFlags, &cbEID, &lpEntryID);
//
//	if (FAILED(hRes))
//	{
//		goto Cleanup;
//	}
//
//	*pCbEID = cbEID;
//	*lppEntryID = lpEntryID;
//Cleanup:
//	return hRes;
//}