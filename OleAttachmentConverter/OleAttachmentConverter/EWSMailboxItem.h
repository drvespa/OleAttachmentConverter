#include "stdafx.h"
#include <Unknwnbase.h>
#include <map>
#include <unordered_map>


#pragma once
class EWSMailboxItem : IUnknown
{
private:
	ULONG m_cRef;
	ULONG m_NamedPropIdx;
	ULONG m_cAttach;
	std::unordered_map<ULONG, _PV> m_MessageProps;
public:
	STDMETHOD(QueryInterface)(REFIID riid, LPVOID * lppvObj);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();
	// IMAPIProp methods
	STDMETHOD(GetLastError)() { return E_NOTIMPL; };
	STDMETHOD(SaveChanges)(ULONG ulFlags);
	STDMETHOD(GetProps)() { return E_NOTIMPL; };
	STDMETHOD(GetPropList)() { return E_NOTIMPL; };
	STDMETHOD(OpenProperty)() { return E_NOTIMPL; };
	STDMETHOD(SetProps)(ULONG cValues, LPSPropValue lpPropArray, LPSPropProblemArray FAR * lppProblems);
	STDMETHOD(DeleteProps)() { return E_NOTIMPL; };
	STDMETHOD(CopyTo)();
	//STDMETHOD(CopyTo)(ULONG ciidExclude, LPCIID rgiidExclude, LPSPropTagArray lpExcludedProps, ULONG_PTR ulUIParam, LPMAPIPROGRESS lpProgress, LPCIID lpInterface, LPVOID lpDestObj, ULONG ulFlags, LPSPropProblemArray FAR * lppProblems);
	STDMETHOD(CopyProps)() { return E_NOTIMPL; };
	STDMETHOD(GetNamesFromIDs)() { return E_NOTIMPL; };
	STDMETHOD(GetIDsFromNames)(ULONG cPropNames, LPMAPINAMEID FAR * lppPropNames, ULONG ulFlags, LPSPropTagArray FAR * lppPropTags);
	// IMessage
	STDMETHOD(GetAttachmentTable)() { return E_NOTIMPL; };
	STDMETHOD(OpenAttach)() { return E_NOTIMPL; };
	STDMETHOD(CreateAttach)(LPCIID lpInterface, ULONG ulFlags, ULONG FAR * lpulAttachmentNum, LPATTACH FAR * lppAttach);
	STDMETHOD(DeleteAttach)() { return E_NOTIMPL; };
	STDMETHOD(GetRecipientTable)() { return E_NOTIMPL; };
	STDMETHOD(ModifyRecipients)(ULONG ulFlags, LPADRLIST lpAdrList);
	STDMETHOD(SubmitMessage)() { return E_NOTIMPL; };
	STDMETHOD(SetReadFlag)() { return E_NOTIMPL; };
	EWSMailboxItem();
	~EWSMailboxItem();
};

