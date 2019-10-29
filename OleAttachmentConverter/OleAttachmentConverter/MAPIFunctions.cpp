#include "stdafx.h"
#include "MailboxObject.h"
#include "main.h"
#include "MAPIFunctions.h"
#include "OleFunctions.h"

using namespace std;



STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::GetOleStreamFromMAPI(CMDLINEARGS CmdLineArgs, LPSTREAM * lppOleStream)
{
	HRESULT		hRes = E_FAIL;

	if (lppOleStream == nullptr)
		return E_INVALIDARG;

	*lppOleStream = nullptr;

	{
		CComPtr<IMAPISession>				lpMapiSession = nullptr;
		CComPtr<IMAPIFolder>				lpInbox = nullptr;
		CComPtr<IMessage>					lpMsg = nullptr;
		CComPtr<IStream>					lpOleStream = nullptr;
		unique_ptr<CMailboxObject>			pMailObject = unique_ptr<CMailboxObject>(nullptr);

		ULONG ulFlags = MAPI_NEW_SESSION | MAPI_EXPLICIT_PROFILE | MAPI_EXTENDED | fMapiUnicode;

		hRes = MAPILogonEx(0, (LPTSTR)CmdLineArgs.lpszProfile.c_str(), nullptr, ulFlags, &lpMapiSession);
		if (FAILED(hRes) || lpMapiSession == nullptr)
		{
			wcout << L"An error has occurred while trying to retrieve logon to MAPI! Error: " << hex << hRes << endl;
			goto Cleanup;
		}
		
		pMailObject.reset(new CMailboxObject(lpMapiSession));
		
		hRes = pMailObject->Logon();

		if (FAILED(hRes))
		{
			wcout << L"An error has occurred while trying to logon to the mailbox! Error: " << hex << hRes << endl;
			goto Cleanup;
		}

		hRes = pMailObject->OpenInbox(&lpInbox);

		if (FAILED(hRes) || lpInbox == nullptr)
		{
			wcout << L"An error has occurred while trying to retrieve the Inbox folder! Error: " << hex << hRes << endl;
			goto Cleanup;
		}

		wcout << L"Opened Inbox folder!" << endl;

		hRes = FindMessageInQuestion(lpInbox, CmdLineArgs.lpszSubject, &lpMsg);
		if (FAILED(hRes) || lpMsg == nullptr)
		{
			wcout << L"An error has occurred while trying to retrieve message from the Inbox folder! Error: " << hex << hRes << endl;
			goto Cleanup;
		}

		wcout << L"Found message in Inbox folder!" << endl;

		//hRes = GetOleStreamFromMessage(lpMsg, &lpOleStream);
		hRes = GetObjectFromMessage<IStream, OBJT_STREAM>(lpMsg, &lpOleStream);

		if (FAILED(hRes) || lpOleStream == nullptr)
		{
			wcout << L"An error has occurred while trying to retrieve the Ole Attachment stream from the message! Error: " << hex << hRes << endl;
			goto Cleanup;
		}

		wcout << L"Found Ole Stream within message!" << endl;
		*lppOleStream = lpOleStream.Detach();
		hRes = S_OK;
Cleanup:
		if (lpMapiSession)
			lpMapiSession->Logoff(0,0,0);
	}
	return hRes;
}

STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::FindMessageInQuestion(LPMAPIFOLDER lpFldr, wstring Subject, LPMESSAGE * lppMsg)
{
	HRESULT									hRes = E_FAIL;
	CComPtr<IMAPITable>						spTblFldr = nullptr;
	LPSRowSet								lpRows = nullptr;
	SizedSPropTagArray(2, spta) = {2, {PR_DISPLAY_NAME, PR_ENTRYID}};
	SRestriction							sres;
	SContentRestriction						srescontentres;
	LPSPropValue							lpspv = nullptr;
	CComPtr<IUnknown>						lpUnk = nullptr;
	CComQIPtr<IMessage, &IID_IMessage>		lpMsg;
	ULONG					ulObjType = 0;

		
	if (lpFldr == nullptr || lppMsg == nullptr)
		return E_INVALIDARG;

	*lppMsg = nullptr;

	ZeroMemory(&sres, sizeof(sres));
	ZeroMemory(&srescontentres, sizeof(srescontentres));

	MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)&lpspv);
	ZeroMemory(lpspv, sizeof(SPropValue));
	srescontentres.lpProp = lpspv;
	srescontentres.lpProp->ulPropTag = PR_SUBJECT_W;
	srescontentres.lpProp->Value.lpszW = (LPWSTR)Subject.c_str();
	srescontentres.ulFuzzyLevel = FL_SUBSTRING;
	srescontentres.ulPropTag = srescontentres.lpProp->ulPropTag;

	sres.rt = RES_CONTENT;
	sres.res.resContent = srescontentres;

	hRes = lpFldr->GetContentsTable(0, &spTblFldr);

	if (FAILED(hRes) || spTblFldr == nullptr)
	{
	wcout << L"An error has occurred while trying to retrieve the Inbox contents! Error: " << hex << hRes << endl;
	goto Cleanup;
	}

	hRes = HrQueryAllRows(spTblFldr, (LPSPropTagArray)&spta, &sres, nullptr, 1, &lpRows);

	if (FAILED(hRes) || lpRows == nullptr || lpRows->cRows == 0)
	{
		goto Cleanup;
	}

	// Since I did a max row of 1 I am only going to get 1 returned at all times regardless of the number of matches 
	// the query returns
	if (lpRows->aRow[0].lpProps[1].ulPropTag != PR_ENTRYID)
	{
		if (PT_ERROR == PROP_TYPE(lpRows->aRow[0].lpProps[1].ulPropTag != PR_ENTRYID))
		{
			hRes = lpRows->aRow[0].lpProps[1].Value.err;
		}
		goto Cleanup;
	}

	LPENTRYID	lpEntryID = (LPENTRYID)lpRows->aRow[0].lpProps[1].Value.bin.lpb;
	ULONG		cbEid = lpRows->aRow[0].lpProps[1].Value.bin.cb;
	
	hRes = lpFldr->OpenEntry(cbEid, lpEntryID, &IID_IMessage, 0, &ulObjType, &lpUnk);

	if (FAILED(hRes) || lpUnk == nullptr || ulObjType == 0)
	{
		goto Cleanup;
	}

	lpMsg = lpUnk;

	if (lpMsg != nullptr)
	{
		*lppMsg = lpMsg.Detach(); // Transfer ownership
		hRes = S_OK;
	}
Cleanup:
	if (lpRows != nullptr)
		FreeProws(lpRows);
	if (lpspv)
		MAPIFreeBuffer(lpspv);
	return hRes;
}

/// Given an attachment object this will pull either the OleStream or an IStorage out of it.  It will create a temporary file on disk as a 
// side effect but that will be deleted when the stream is released.
STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::GetOleObjectFromAttachment(LPATTACH lpAttach, ULONG * pulObjectType, LPUNKNOWN * lppUnk)
{
	HRESULT					hRes = E_FAIL;
	CComPtr<IStream>		lpStream = nullptr;
	CComPtr<IStorage>		lpStg = nullptr;
	ULONG					PropTagSought = PR_ATTACH_DATA_OBJ;
	IID						InterfaceSought = IID_IStreamDocfile;

	if (lpAttach == nullptr || lppUnk == nullptr || pulObjectType == nullptr)
		return E_INVALIDARG;

	*lppUnk = nullptr;
	*pulObjectType = 0;

	while (FAILED(hRes) && (lpStream == nullptr && lpStg == nullptr))
	{
		hRes = lpAttach->OpenProperty(PropTagSought,
										&InterfaceSought,
										0,
										0,
										(LPUNKNOWN*)&lpStream);

		if (FAILED(hRes) && lpStream == nullptr)
		{
			// h ttps://msdn.microsoft.com/en-us/library/cc842411(v=office.14).aspx
			if (PropTagSought == PR_ATTACH_DATA_OBJ && InterfaceSought == IID_IStreamDocfile)
			{
				// This is the first time through with a failure
				// Change the prop sought
				PropTagSought = PR_ATTACH_DATA_BIN;
			}
			else if (PropTagSought == PR_ATTACH_DATA_BIN && InterfaceSought == IID_IStreamDocfile)
			{
				// This is the second time through with a failure
				// change the prop sought back to _DATA_OBJ
				// but this time change the itf sought
				PropTagSought = PR_ATTACH_DATA_OBJ;
				InterfaceSought = IID_IStorage;

				hRes = lpAttach->OpenProperty(PropTagSought,
												&InterfaceSought,
												0,
												0,
												(LPUNKNOWN*)&lpStg);
			}
			else
			{
				// This is the third time though with a failure
				// I am out of ideas and so is MSDN
				// bail
				break;
			}
		}
	}

	if (SUCCEEDED(hRes) && lpStream != nullptr)
	{
		// Ok, we have found the stream!
		// Now we need to copy it out and return an uncumbered stream
		// so we can close out the IAttach and the lpStream but leave the other stream open
		// It will create a temporary file on disk as a 
		// side effect but that will be deleted when the stream is released.
		LPSTREAM lpDestStream = nullptr;
		hRes = PersistOleStream(lpStream, L"", &lpDestStream);
		if (SUCCEEDED(hRes) && lpDestStream != nullptr)
		{
			*lppUnk = lpDestStream;
			*pulObjectType = OBJT_STREAM;
		}
	}

	if (SUCCEEDED(hRes) && lpStg != nullptr)
	{
		*lppUnk = lpStg.Detach();
		*pulObjectType = OBJT_STORAGE;
		hRes = S_OK;
	}

	return hRes;
}

STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::MyOpenStreamOnFile(std::wstring lpszFileName, LPSTREAM * lppDestStream, ULONG ulFlags = STGM_CREATE | STGM_READWRITE)
{
	HRESULT					hRes = E_FAIL;
	int						cch = -1;
	unique_ptr<char[]>		szFileName;

	// Convert the file name to ANSI
	// This is a requirement for the OpenStreamOnFile API
	// even though the header lists it as taking a LPTSTR for the file name.
	if (!lpszFileName.empty())
	{
		cch = wcslen(lpszFileName.c_str());
		szFileName.reset(new char[(cch + 1)]);
		ZeroMemory(szFileName.get(), cch + 1);
		if (0 == WideCharToMultiByte(CP_ACP, 0, lpszFileName.c_str(), cch, szFileName.get(), (cch + 1), nullptr, nullptr))
		{
			hRes = HRESULT_FROM_WIN32(GetLastError());
			goto Cleanup;
		}
	}

	*lppDestStream = nullptr;

	hRes = OpenStreamOnFile(MAPIAllocateBuffer,
							MAPIFreeBuffer,
							ulFlags,
							lpszFileName.empty() ? nullptr : (LPTSTR)szFileName.get(), // The header says it takes a LPTSTR but it really takes a LPSTR
							nullptr,
							lppDestStream);
Cleanup:
	return hRes;
}
