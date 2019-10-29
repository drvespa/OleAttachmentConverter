#include "StdAfx.h"
#include "EWSMailboxItem.h"


using namespace std;

using namespace Microsoft::Exchange::WebServices::Data;

EWSMailboxItem::EWSMailboxItem() : m_cRef(1)
{
	
}

EWSMailboxItem::~EWSMailboxItem()
{

}

STDMETHODIMP EWSMailboxItem::QueryInterface(REFIID riid, LPVOID * lppvObj)
{
	return E_NOTIMPL;
}

STDMETHODIMP_(ULONG) EWSMailboxItem::AddRef()
{
	return ++m_cRef;
}

STDMETHODIMP_(ULONG) EWSMailboxItem::Release()
{
	ULONG cRef = --m_cRef;
	if (cRef == 0)
	{
		delete this;
	}
	return cRef;
}


STDMETHODIMP EWSMailboxItem::SaveChanges(ULONG ulFlags)
{
	for (auto& i : m_MessageProps)
	{
		//auto propDef = m_Prop2PropDef.find(i.first);
		//if (*propDef.)
	}
	return S_OK;
}

STDMETHODIMP EWSMailboxItem::ModifyRecipients(ULONG ulFlags, LPADRLIST lpAdrList)
{
	return S_OK;
}

STDMETHODIMP EWSMailboxItem::SetProps(ULONG cValues, LPSPropValue lpPropArray, LPSPropProblemArray FAR * lppProblems)
{
	for (ULONG i = 0; i<cValues;)
	{
		_PV pv;
		ZeroMemory(&pv, sizeof(pv));
		switch (PROP_TYPE(lpPropArray[i].ulPropTag))
		{
		case PT_STRING8:
			{
				int cch = strlen(lpPropArray[i].Value.lpszA) + 1;
				int cb = cch * sizeof(char);
				auto pChar = unique_ptr<char[]>(new char[cch]);
				ZeroMemory(pChar.get(), cb);
				CopyMemory(pChar.get(), lpPropArray[i].Value.lpszA, cb);
				m_MessageProps.emplace(lpPropArray[i].ulPropTag, pv);
			}
			break;
		case PT_UNICODE:
			{
				int cch = wcslen(lpPropArray[i].Value.lpszW) + 1;
				int cb = cch * sizeof(wchar_t);
				auto pwChar = unique_ptr<wchar_t[]>(new wchar_t[cch]);
				ZeroMemory(pwChar.get(), cb);
				CopyMemory(pwChar.get(), lpPropArray[i].Value.lpszW, cb);
				pv.lpszW = pwChar.get();
				m_MessageProps.emplace(lpPropArray[i].ulPropTag, pv);
			}
			break;
		case PT_I2:
		case PT_I4:
		case PT_I8:
			break;
		case PT_ERROR:
			// Don't copy errrors
			break;
		default:
			break;
		}
	}
	return S_OK;
}

//STDMETHODIMP EWSMailboxItem::CopyTo(ULONG ciidExclude, LPCIID rgiidExclude, LPSPropTagArray lpExcludedProps, ULONG_PTR ulUIParam, LPMAPIPROGRESS lpProgress, LPCIID lpInterface, LPVOID lpDestObj, ULONG ulFlags, LPSPropProblemArray * lppProblems)
STDMETHODIMP EWSMailboxItem::CopyTo()
{
	return E_NOTIMPL;
}

STDMETHODIMP EWSMailboxItem::GetIDsFromNames(ULONG cPropNames, LPMAPINAMEID FAR * lppPropNames, ULONG ulFlags, LPSPropTagArray FAR * lppPropTags)
{
	//LPSPropTagArray lpSpta = nullptr;

	//if (lppPropTags == nullptr)
	//{
	//	return E_INVALIDARG;
	//}

	//*lppPropTags = nullptr;

	//if (cPropNames == 0)
	//	return S_OK;

	//if (m_NamedPropIdx > 0xFFFE)
	//{
	//	return MAPI_E_CALL_FAILED;
	//}

	//MAPIAllocateBuffer(cPropNames * sizeof(SPropTagArray), (LPVOID*)&lpSpta);

	//ZeroMemory(lpSpta, cPropNames * sizeof(SPropTagArray));

	//for (ULONG i = 0; i<cPropNames; i++)
	//{
	//	lpSpta->aulPropTag[i] = PROP_TAG(PT_UNSPECIFIED, m_NamedPropIdx);
	//	m_NamedPropIdx++;
	//}

	//lpSpta->cValues = cPropNames;

	////*lppPropTags = lpSpta;

	return S_OK;
}

STDMETHODIMP EWSMailboxItem::CreateAttach(LPCIID lpInterface, ULONG ulFlags, ULONG FAR * lpulAttachmentNum, LPATTACH FAR * lppAttach)
{
	//if (lpulAttachmentNum == nullptr || lppAttach == nullptr)
	//	return E_INVALIDARG;
	//*lpulAttachmentNum = m_cAttach;
	//m_cAttach++;
	//// The vtable of a IMessage and IAttach are very similar
	//// They differ at the point after IMAPProp
	//// However there are no methods in IAttach after IMAPIProp.
	//// Therefore, whatever hangs off the vtable at that point will never be accessed
	//// so just force it and regret it later. :)
	//LPATTACH lpAttach = nullptr;
	//lpAttach = reinterpret_cast<LPATTACH>(this);
	//lpAttach->AddRef();
	//*lppAttach = lpAttach;
	return S_OK;
}
