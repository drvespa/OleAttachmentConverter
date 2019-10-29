namespace Microsoft
{
	namespace Samples 
	{
		namespace OleAttachmentConverter
		{
			STDMETHODIMP GetOleStreamFromMAPI(CMDLINEARGS CmdLineArgs, LPSTREAM * lppOleStream);
			STDMETHODIMP FindMessageInQuestion(LPMAPIFOLDER lpFldr, std::wstring Subject, LPMESSAGE * lppMsg);
			STDMETHODIMP GetOleObjectFromAttachment(LPATTACH lpAttach, ULONG * pulObjectType, LPUNKNOWN * lppUnk);
			STDMETHODIMP MyOpenStreamOnFile(std::wstring lpszFileName, LPSTREAM * lppDestStream, ULONG ulFlags);
			#define OBJT_STREAM 1
			#define OBJT_STORAGE 2
			template <typename T, ULONG K> STDMETHODIMP GetObjectFromMessage (LPMESSAGE lpMsg, T ** lppType)
			{
				HRESULT							hRes = E_FAIL;
				ATL::CComPtr<IMAPITable>		spTblAttach = nullptr;
				LPSRowSet						lpRows = nullptr;
				SizedSPropTagArray(2, spta) = {2, {PR_ATTACH_METHOD, PR_ATTACH_NUM}};
				SRestriction					sres;
				SPropertyRestriction			sresprop;
				ATL::CComPtr<IAttach>			lpAttach = nullptr;
				ATL::CComPtr<T>					lpDest = nullptr;
				ULONG							ulObjType = 0;

				if (lpMsg == nullptr || lppType == nullptr)
					return E_INVALIDARG;

				*lppType = nullptr;
	
				hRes = lpMsg->GetAttachmentTable(0, &spTblAttach);

				if (FAILED(hRes) || spTblAttach == nullptr)
				{
					goto Cleanup;
				}

				// Find the atttachments based on the attach method 
				ZeroMemory(&sres, sizeof(sres));
				ZeroMemory(&sresprop, sizeof(sresprop));

				MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)&sresprop.lpProp);
				ZeroMemory(sresprop.lpProp, sizeof(SPropValue));

				sresprop.lpProp->ulPropTag = PR_ATTACH_METHOD;
				sresprop.lpProp->Value.l = ATTACH_OLE;
				sresprop.ulPropTag = sresprop.lpProp->ulPropTag;
				sresprop.relop = RELOP_EQ;

				sres.rt = RES_PROPERTY;
				sres.res.resProperty = sresprop;

				// This will return all rows
				hRes = HrQueryAllRows(spTblAttach, (LPSPropTagArray)&spta, &sres, nullptr, 0, &lpRows);

				if (FAILED(hRes) || lpRows == nullptr)
				{
					goto Cleanup;
				}

				if (lpRows->cRows == 0)
				{
					hRes = MAPI_E_NOT_FOUND;
					goto Cleanup;
				}
	
				// TODO: I am hardcoding this to only get the first Ole Attachment
				// There may be others and they won't be picked up by this code.
				hRes = lpMsg->OpenAttach(lpRows->aRow[0].lpProps[1].Value.l, nullptr, 0, &lpAttach);

				if (FAILED(hRes) || lpAttach == nullptr)
				{
					goto Cleanup;
				}

				hRes = GetOleObjectFromAttachment(lpAttach, &ulObjType, (LPUNKNOWN*)&lpDest);

				if (FAILED(hRes) || (lpDest == nullptr))
				{
					goto Cleanup;
				}

				if (ulObjType == K)
				{
					*lppType = lpDest.Detach();
				}
				else
				{
					hRes = E_FAIL;
				}
	
			Cleanup:
				if (sresprop.lpProp)
					MAPIFreeBuffer(sresprop.lpProp);
				if (lpRows)
					FreeProws(lpRows);
				return hRes;
			}
		}
	}
}