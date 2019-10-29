// main.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <string>
#include <memory>
#include <atlcomcli.h>
#include <initguid.h>
#define USES_IID_IMessage
#define USES_IID_IStreamDocfile
#include <MAPIX.h>
#include "MailboxObject.h"
#include "main.h"
#include "MAPIFunctions.h"
#include "OleFunctions.h"
#include "IMessage.h"

using namespace ATL;
using namespace Microsoft::Samples;
using namespace std;

int wmain(int argc, wchar_t * argv[])
{
	CMDLINEARGS CmdLineArgs = {};
	if (!OleAttachmentConverter::ParseCmdLineArgs(argc, argv, &CmdLineArgs))
	{
		OleAttachmentConverter::DisplayUsage();
		return 0;
	}
	return OleAttachmentConverter::Main(CmdLineArgs);
}

int Microsoft::Samples::OleAttachmentConverter::Main(CMDLINEARGS CmdLineArgs)
{
	MAPIINIT_0	mapiInit = {MAPI_INIT_VERSION, 0};
	MAPIInitialize(&mapiInit);
	{
		if (!CmdLineArgs.bUseMSG)
		{
			PullFromInbox(CmdLineArgs);
		}
		else
		{
			PullFromMSG(CmdLineArgs);
		}
	}
	MAPIUninitialize();
	return 0;
}

void Microsoft::Samples::OleAttachmentConverter::DisplayUsage()
{
	wcout << L"OleAttachmentConverter.exe -? -b <Path> -e -f <Subject> -h -i <Path> -p <Profile Name>" << endl;
	wcout << L"-? - Shows this help. Optional" << endl;
	wcout << L"-b - The name of the BMP to create - Required" << endl;
	wcout << L"-e - Enumerate the OleCache - Optional" << endl;
	wcout << L"-f <Subject> - The subject of the email to find.  It will choose the first one if there are duplicate matches - Required when loading from a mailbox." << endl;
	wcout << L"-h - Shows this help. Optional" << endl;
	wcout << L"-i <Path to the MSG> - Opens a MSG and starts pulling out the data from there - Required when loading from a MSG file." << endl;
	wcout << L"-p <Profile Name> - The profile to be used - Required when loading from a mailbox" << endl;
	wcout << L"Examples:" << endl;
	wcout << L"\tOleAttachmentConverter.exe -b \"test.bmp\" -i \"D:\\MSGs\\Test RTF.msg\"" << endl;
	wcout << L"\tOleAttachmentConverter.exe -b \"test3.bmp\" -p \"Outlook\" -f \"TEST RTF\"" << endl;
	wcout << L"\tOleAttachmentConverter.exe -b \"test3.bmp\" -p \"Outlook\" -f \"TEST RTF\" -e" << endl;
}

bool Microsoft::Samples::OleAttachmentConverter::ParseCmdLineArgs(int argc, wchar_t * argv[], LPCMDLINEARGS lpCmdLineArgs)
{
	for(int i = 1;i<argc;i++)
	{
		switch (argv[i][1])
		{
			case 'b':
			case 'B':
				lpCmdLineArgs->lpszBmpOutputFileName = wstring(argv[++i]);
				break;
			case 'e':
			case 'E':
				lpCmdLineArgs->bEnumCache = true;
				break;
			case 'f':
			case 'F':
				lpCmdLineArgs->lpszSubject = wstring(argv[++i]);
				break;
			case '?':
			case 'h':
			case 'H':
				return false;
				break;
			case 'i':
			case 'I':
				lpCmdLineArgs->bUseMSG = true;
				lpCmdLineArgs->lpszFileName = wstring(argv[++i]);
				break;
			case 'p':
			case 'P':
				lpCmdLineArgs->lpszProfile = wstring(argv[++i]);
				break;
		}
	}
	if (lpCmdLineArgs->lpszBmpOutputFileName.empty())
	{
		wcout << L"You must supply a BMP file path" << endl;
		return false;
	}
	return true;
}
void Microsoft::Samples::OleAttachmentConverter::PullFromInbox(CMDLINEARGS CmdLineArgs)
{
	HRESULT							hRes = E_FAIL;
	CComPtr<IStream>				lpOleStream = nullptr;
	CComQIPtr<IStorage>				lpStg;
	CComQIPtr<IOleCache>			pOleCache;
	CComQIPtr<IPersistStorage>		pPersistStg;
	STGMEDIUM						StgMed;

	ZeroMemory(&StgMed, sizeof(StgMed));
	
	hRes = GetOleStreamFromMAPI(CmdLineArgs, &lpOleStream);
	
	if (FAILED(hRes))
	{
		wcout << L"Couldn't get the OleStream From MAPI! Error: " << hex << hRes << endl;
		goto Cleanup;
	}
	if (!CmdLineArgs.lpszFileName.empty())
	{
		// The user has chosen to persist the stream
		hRes = PersistOleStream(lpOleStream, CmdLineArgs.lpszFileName, nullptr);
		if (FAILED(hRes))
		{
			wcout << L"An error has occurred while trying to save the Ole Attachment stream to disk! Error: " << hex << hRes << endl;
			goto Cleanup;
		}
	}
	hRes = PersistStructuredStorage(lpOleStream, &lpStg);
	if (FAILED(hRes))
	{
		wcout << L"An error has occurred while trying to save the Ole Attachment stream to disk! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	hRes = CreateDataCache(nullptr, CLSID_NULL, __uuidof(IPersistStorage), (LPVOID*)&pPersistStg);
	if (FAILED(hRes) || pPersistStg == nullptr)
	{
		wcout << L"Couldn't get the IPersistStorage! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	pOleCache = pPersistStg;

	hRes = pPersistStg->Load(lpStg);

	if (FAILED(hRes) || pPersistStg == nullptr)
	{
		wcout << L"Couldn't load the storage! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	if (CmdLineArgs.bEnumCache)
	{
		hRes = EnumCache(pOleCache);

		if (FAILED(hRes) )
		{
			wcout << L"Couldn't enumm the Ole Cache! Error: " << hex << hRes << endl;
			goto Cleanup;
		}
	}

	hRes = GetSTGMEDIUMFromCache(pOleCache, &StgMed);

	if (FAILED(hRes) || StgMed.tymed == TYMED_NULL)
	{
		wcout << L"The call to GetSTGMEDIUMFromCache() failed! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	hRes = SaveBitmapToFileSystem(StgMed, CmdLineArgs.lpszBmpOutputFileName);
	if (FAILED(hRes))
	{
		wcout << L"An error has occurred while trying to convert the Ole Stream to a SS Pointer! Error: " << hex << hRes << endl;
		goto Cleanup;
	}
	
Cleanup:
	return;
}

void Microsoft::Samples::OleAttachmentConverter::PullFromMSG(CMDLINEARGS CmdLineArgs)
{
	LPMSGSESS						lpMsgSess = nullptr;
	LPMALLOC						lpMalloc = MAPIGetDefaultMalloc();
	CComPtr<IStorage>				lpMsgStg = nullptr;
	CComPtr<IMessage>				lpMsg = nullptr;
	CComPtr<IStorage>				lpStg = nullptr;
	CComQIPtr<IOleCache>			pOleCache;
	CComQIPtr<IPersistStorage>		pPersistStg;
	STGMEDIUM						StgMed;

	ZeroMemory(&StgMed, sizeof(StgMed));

	HRESULT hRes  = OpenIMsgSession(lpMalloc, 0, &lpMsgSess);

	if (FAILED(hRes))
	{
		wcout << L"Couldn't create sess ptr! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	hRes = GetStructuredStorage(CmdLineArgs.lpszFileName, &lpMsgStg);
			
	if (FAILED(hRes) || !lpMsgStg)
	{
		wcout << L"Couldn't get the storage! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	hRes = OpenIMsgOnIStg(lpMsgSess,
							MAPIAllocateBuffer,
							MAPIAllocateMore,
							MAPIFreeBuffer,
							lpMalloc,
							nullptr,
							lpMsgStg,
							nullptr,
							0,
							0,
							&lpMsg);
	if (FAILED(hRes) || lpMsg == nullptr)
	{
		wcout << L"Couldn't get the IMessage ptr! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	hRes = GetObjectFromMessage<IStorage, OBJT_STORAGE>(lpMsg, &lpStg);

	if (FAILED(hRes) || lpStg == nullptr)
	{
		wcout << L"Couldn't get the Ole Stream ptr! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	wcout << L"Retried the Ole Object from the file " << endl;

	hRes = CreateDataCache(nullptr, CLSID_NULL, __uuidof(IPersistStorage), (LPVOID*)&pPersistStg);
	if (FAILED(hRes) || pPersistStg == nullptr)
	{
		wcout << L"Couldn't get the IPersistStorage! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	pOleCache = pPersistStg;

	hRes = pPersistStg->Load(lpStg);

	if (FAILED(hRes) || pPersistStg == nullptr)
	{
		wcout << L"Couldn't load the storage! Error: " << hex << hRes << endl;
		goto Cleanup;
	}


	wcout << L"Created the Cached and loaded the Persisted Storage" << endl;

	if (CmdLineArgs.bEnumCache)
	{
		hRes = EnumCache(pOleCache);

		if (FAILED(hRes) )
		{
			wcout << L"Couldn't enumm the Ole Cache! Error: " << hex << hRes << endl;
			goto Cleanup;
		}
	}

	hRes = GetSTGMEDIUMFromCache(pOleCache, &StgMed);

	if (FAILED(hRes) || StgMed.tymed == TYMED_NULL)
	{
		wcout << L"The call to GetSTGMEDIUMFromCache() failed! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	wcout << L"Retried the StgMed from the cache " << endl;

	hRes = SaveBitmapToFileSystem(StgMed, CmdLineArgs.lpszBmpOutputFileName);
	if (FAILED(hRes))
	{
		wcout << L"An error has occurred while trying to convert the Ole Stream to a SS Pointer! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	wcout << L"Save the BMP to the file system" << endl;

	// Now copy the file from one point to another.

Cleanup:
	if (StgMed.pUnkForRelease)
	{
		StgMed.pUnkForRelease->Release();
	}
	else
	{
		ReleaseStgMedium(&StgMed);
	}

	pPersistStg.Release();

	pOleCache.Release();

	lpStg.Release();

	lpMsg.Release();

	lpMsgStg.Release();

	if (lpMalloc)
	{
		lpMalloc->Release();
	}

	if (lpMsgSess)
	{
		CloseIMsgSession(lpMsgSess);
	}
}


