#include "stdafx.h"
#include "main.h"
#include "OleFunctions.h"
#include "MAPIFunctions.h"
#include "SimpleHGlobalWrapper.h"
#include <fstream>

using namespace std;
using namespace ATL;
using namespace Microsoft::Samples;

#define TEMP_FILE_PATH _T("C:\\Users\\dvespa\\AppData\\Local\\Temp")
#define APP_SETTINGS _T("Software\\Microsoft\\OleAttachmentConverter")

// This will persist a given stream to disk.  If the file path and file name are not passed
// it will create a temporary file that is deleted on release of the returned stream.
// If a return stream is not passed (i.e. lppDestStream == nullptr), then it will just persist the stream
// where the caller has asked.
STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::PersistOleStream(LPSTREAM lpOleStream, wstring lpszFileName, LPSTREAM * lppDestStream)
{
	HRESULT					hRes = E_FAIL;
	LARGE_INTEGER			li;
	STATSTG					StgInfo;
	CComPtr<IStream>		lpDest = nullptr;
	ULONG					ulFlags = STGM_CREATE | STGM_READWRITE;
	ULARGE_INTEGER			cBytesRead;
	ULARGE_INTEGER			cBytesWritten;
	bool					bUseTempFile = lpszFileName.empty();
	
	if (lpOleStream == nullptr)
		return E_INVALIDARG;

	if (lppDestStream != nullptr)
	{
		*lppDestStream = nullptr;
	}

	ZeroMemory(&li, sizeof(li));

	// Seek to the beginning of the stream
	// just in case it has been manipulated before getting to this method.
	hRes = lpOleStream->Seek(li, BOOKMARK_BEGINNING, nullptr);

	if (FAILED(hRes))
	{
		// Not sure why we can't Seek but that could be a problem
		hRes = E_FAIL;
		goto Cleanup;
	}

	if (bUseTempFile)
	{
		if (lppDestStream == nullptr)
		{
			// This doesn't make sense
			return hRes;
		}
		ulFlags |= SOF_UNIQUEFILENAME | STGM_DELETEONRELEASE;
	}

	hRes = MyOpenStreamOnFile(lpszFileName, &lpDest, ulFlags);

	if (FAILED(hRes) || lpDest == nullptr)
	{
		goto Cleanup;
	}

	ZeroMemory(&StgInfo, sizeof(StgInfo));

	hRes = lpOleStream->Stat(&StgInfo, STATFLAG_NONAME);

	if (FAILED(hRes) || lpDest == nullptr)
	{
		goto Cleanup;
	}

	// Confirm that there is data to read
	if (StgInfo.cbSize.HighPart == 0 && StgInfo.cbSize.LowPart == 0)
	{
		hRes = E_FAIL;
		goto Cleanup;
	}

	ZeroMemory(&cBytesRead, sizeof(cBytesRead));
	ZeroMemory(&cBytesWritten, sizeof(cBytesWritten));

	hRes = lpOleStream->CopyTo(lpDest, StgInfo.cbSize, &cBytesRead, &cBytesWritten);

	if (FAILED(hRes))
	{
		goto Cleanup;
	}

	// Confirm that we succesfully read from the source stream and successfully copied to the
	// destination stream
	if ((cBytesWritten.QuadPart == 0) || (cBytesRead.QuadPart == 0))
	{
		hRes = E_FAIL;
		goto Cleanup;
	}

	// Confirm that we succesfully wrote and read the same amount of bytes
	if ((cBytesWritten.QuadPart != cBytesRead.QuadPart))
	{
		hRes = E_FAIL;
		goto Cleanup;
	}

	hRes = lpDest->Commit(STGC_DEFAULT);

	if (SUCCEEDED(hRes) && lppDestStream != nullptr)
	{
		*lppDestStream = lpDest.Detach();
	}

Cleanup:
	return hRes;
}

STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::PersistStructuredStorage(LPSTREAM lpOleStream, LPSTORAGE * lppStg)
{
	if (lpOleStream == nullptr || lppStg == nullptr)
		return E_INVALIDARG;

	HRESULT					hRes = E_FAIL;
	unique_ptr<wchar_t[]>	szTempFilePath(new wchar_t[MAX_PATH]);
	unique_ptr<wchar_t[]>	szTempFileName(new wchar_t[MAX_PATH]);
	DWORD					nChars = MAX_PATH;
	UINT					err;

	*lppStg = nullptr;

	ZeroMemory(szTempFilePath.get(), MAX_PATH);
	ZeroMemory(szTempFileName.get(), MAX_PATH);

	RegGetValue(HKEY_CURRENT_USER, APP_SETTINGS, _T("TempPath"), RRF_RT_REG_SZ, nullptr, szTempFilePath.get(), &nChars); 
	
	if (nChars > 0)
	{
		err = GetTempFileName(szTempFilePath.get(), _T(""), 0, szTempFileName.get());
	}
	else
	{
		err = GetTempFileName(TEMP_FILE_PATH, _T(""), 0, szTempFileName.get());
	}
	if (err == 0)
	{
		return hRes;
	}

	// Save the stream to disk
	hRes = PersistOleStream(lpOleStream, szTempFileName.get(), nullptr);

	return GetStructuredStorage(szTempFileName.get(), lppStg);
}

STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::GetStructuredStorage(wstring lpszFileName, LPSTORAGE * lppStg)
{

	if (lppStg == nullptr)
		return E_INVALIDARG;

	HRESULT					hRes = E_FAIL;
	CComPtr<IStorage>		lpDestStg = nullptr;

	*lppStg = nullptr;

	// Now open that same stream but do so using the Structured Storage API
	hRes = StgOpenStorage(lpszFileName.c_str(), 
							nullptr,
							STGM_TRANSACTED | STGM_READWRITE | STGM_SHARE_DENY_WRITE,
							0, 
							0,
							&lpDestStg);

	if (SUCCEEDED(hRes) && lpDestStg != nullptr)
	{
		*lppStg = lpDestStg.Detach();
	}
	return hRes;
}

void Microsoft::Samples::OleAttachmentConverter::EnumStorage(LPSTORAGE lpStorage)
{
	HRESULT								hRes = E_FAIL;
	CComPtr<IEnumSTATSTG>				lpEnum = nullptr;
	unique_ptr<STATSTG[]>				rgElt = unique_ptr<STATSTG[]>(new STATSTG[3]);
	ULONG								cValues;

	// Now let's just loop through the streams and see what we have
	hRes = lpStorage->EnumElements(0, 
									0, 
									0, 
									&lpEnum);

	if (FAILED(hRes) || lpEnum == nullptr)
	{
		wcout << L"An error has occurred while trying enumerate the storage! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	do {
		cValues = -1;
		ZeroMemory(rgElt.get(), (sizeof(STATSTG) * 3));

		// The return value can be S_OK || S_FALSE so be careful
		hRes = lpEnum->Next(3, rgElt.get(), &cValues);

		if (FAILED(hRes))
		{
			wcout << L"An error has occurred while trying enumerate the storage! Error: " << hex << hRes << endl;
			goto Cleanup;
		}

		for (auto i = 0UL; i < cValues; i++)
		{
			wcout << "Stream " << rgElt[i].pwcsName << " found" << endl;
		}

	} while (cValues != 0 && SUCCEEDED(hRes));

Cleanup:
	return;
}

STDMETHODIMP LoadMetaFilePict(LPSTORAGE pStgSrc, HMETAFILEPICT * phmfp)
{
	HRESULT hRes = E_FAIL;
	return hRes;
}

STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::EnumCache(IOleCache * pOleCache)
{
	HRESULT						hRes = E_FAIL;
	CComPtr<IEnumSTATDATA>		pStatData = nullptr;
	hRes = pOleCache->EnumCache(&pStatData);

	if (FAILED(hRes) || pStatData == nullptr)
	{
		wcout << L"Couldn't get the Cache Enumerator! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	STATDATA rgData[2];
	ULONG cValues = -1;

	hRes = pStatData->Next(2, rgData, &cValues);

	if (FAILED(hRes))
	{
		wcout << L"Couldn't enumerate the cache! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	ULONG DIB_FORMAT_IDX = -1;

	for (ULONG i=0;i<cValues;i++)
	{
		FORMATETC currentFormat = rgData[i].formatetc;				
		switch (currentFormat.cfFormat)
		{
			case CF_BITMAP:
				wcout << "CF Format: BitMap" << endl;
				break;
			case CF_DIB:
				wcout << "CF Format: DIB" << endl;
				break;
			case CF_METAFILEPICT:
				wcout << "CF Format: MetaFilePict" << endl;
				break;
		}
		switch (currentFormat.tymed)
		{
			case TYMED_HGLOBAL:
				wcout << "Medium: Global" << endl;
				break;
			case TYMED_MFPICT:
				wcout << "Medium: MetaFilePict" << endl;
				break;
			case TYMED_GDI:
				wcout << "Medium: GDI" << endl;
				break;
			default:
				wcout << currentFormat.tymed << endl;
		}

		switch (currentFormat.dwAspect)
		{
			case DVASPECT_CONTENT:
				wcout << "Aspect: Content" << endl;
				break;
			case DVASPECT_ICON:
				wcout << "Aspect: Icon" << endl;
				break;
			case DVASPECT_THUMBNAIL:
				wcout << "Aspect: Thumbnail" << endl;
				break;
		}
	}
Cleanup:
	return S_OK;
}

STDMETHODIMP  Microsoft::Samples::OleAttachmentConverter::GetSTGMEDIUMFromCache(IOleCache * pOleCache, LPSTGMEDIUM lpStgMedium)
{
	HRESULT					hRes = E_FAIL;
	CComQIPtr<IDataObject>	pDataObj;
	unique_ptr<FORMATETC>	pFormatEtc = unique_ptr<FORMATETC>(new FORMATETC);


	if (pOleCache == nullptr || lpStgMedium == nullptr)
		return E_INVALIDARG;

	ZeroMemory(lpStgMedium, sizeof(STGMEDIUM));
	ZeroMemory(pFormatEtc.get(), sizeof(FORMATETC));

	pDataObj = pOleCache;
	
	if (pDataObj == nullptr)
	{
		wcout << L"QI for IDataObject failed! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

	pFormatEtc->cfFormat = CF_DIB;
	pFormatEtc->dwAspect = DVASPECT_CONTENT;
	pFormatEtc->tymed = TYMED_HGLOBAL;
	pFormatEtc->lindex = -1;

	hRes = pDataObj->GetData(pFormatEtc.get(), lpStgMedium);

	if (FAILED(hRes))
	{
		wcout << L"The call to IDataObject::GetData Failed! Error: " << hex << hRes << endl;
		goto Cleanup;
	}

#ifdef DEBUG
	if (pFormatEtc->tymed != lpStgMedium->tymed)
	{
		wcout << L"The call to IDataObject::GetData returned a different format than requested." << endl;
		hRes = MAPI_W_ERRORS_RETURNED; // This isn't necessarily bad, just the developer needs to be aware of it.  ALso this is my personal logic
		goto Cleanup;
	}
#endif
Cleanup:
	return hRes;
}

STDMETHODIMP Microsoft::Samples::OleAttachmentConverter::SaveBitmapToFileSystem(STGMEDIUM StgMed, wstring lpszBmpOutputFileName)
{
	HRESULT					hRes = E_FAIL;
	HANDLE					hFile = INVALID_HANDLE_VALUE;
	SimpleHGlobalWrapper	HGlobal = SimpleHGlobalWrapper(StgMed.hGlobal);
	DWORD					cBytesWritten = 0;
	BITMAPFILEHEADER		bitMapFileHeader;
	//string					bfType("BM");
	// TODO: I need to confirm that this doesn't have a null terminator
	LPCSTR					lpszBfType = "\x42\x4D"; // BM

	if (StgMed.tymed != TYMED_HGLOBAL)
		return hRes;

	ZeroMemory(&bitMapFileHeader, sizeof(bitMapFileHeader));

	hFile = CreateFile(lpszBmpOutputFileName.c_str(),
							GENERIC_READ | GENERIC_WRITE, 
							0, // Prevent others from doing anything with this file
							nullptr, 
							CREATE_ALWAYS, 
							FILE_ATTRIBUTE_NORMAL, 
							nullptr);

	if (hFile == INVALID_HANDLE_VALUE)
	{
		wcout << L"The call to CreateFile() failed! Error: " << hex << GetLastError() << endl;
		goto Cleanup;
	}

	//CopyMemory(&bitMapFileHeader.bfType, bfType.c_str(), bfType.size());
	CopyMemory(&bitMapFileHeader.bfType, lpszBfType, 2);

	bitMapFileHeader.bfSize =  sizeof(bitMapFileHeader) + HGlobal.Size();
	bitMapFileHeader.bfOffBits = sizeof(bitMapFileHeader);

	if (FALSE == WriteFile(hFile,
			&bitMapFileHeader, 
			sizeof(bitMapFileHeader), 
			&cBytesWritten,
			nullptr) || cBytesWritten != sizeof(bitMapFileHeader))
	{
		wcout << L"The call to WriteFile() failed when writing the BITMAPFILEHEADER! Error: " << hex << GetLastError() << endl;
		goto Cleanup;
	}

	cBytesWritten = 0;

	if (FALSE == WriteFile(hFile,
							HGlobal.GetPtr(), 
							HGlobal.Size(), 
							&cBytesWritten, 
							nullptr) || cBytesWritten != HGlobal.Size())
	{
		wcout << L"The call to WriteFile() failed when writing the HGlobal bytes! Error: " << hex << GetLastError() << endl;
		goto Cleanup;
	}

	hRes = S_OK;
Cleanup:
		if (hFile)
		{
			CloseHandle(hFile);
		}
		// I am not entirely sure that you have to do this
		// but just in case.
		//if (StgMed.pUnkForRelease)
		//{
		//	StgMed.pUnkForRelease->Release();
		//}
		//else
		//{
		//	ReleaseStgMedium(&StgMed);
		//}
	return hRes;
}