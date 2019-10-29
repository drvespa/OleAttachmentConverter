#include "StdAfx.h"
#include "SimpleHGlobalWrapper.h"

Microsoft::Samples::SimpleHGlobalWrapper::SimpleHGlobalWrapper(HGLOBAL hMem) : m_hMem(nullptr), m_lpv(nullptr) 
{
	m_hMem = hMem;
	m_lpv = GlobalLock(m_hMem);
}

Microsoft::Samples::SimpleHGlobalWrapper::~SimpleHGlobalWrapper(void)
{
	m_lpv = nullptr;
	if (TRUE != GlobalUnlock(m_hMem))
	{
		// Handle the error somehow
	}
	m_hMem = nullptr;
}

STDMETHODIMP_(LPVOID) Microsoft::Samples::SimpleHGlobalWrapper::GetPtr()
{
	return m_lpv;
}

STDMETHODIMP_(SIZE_T) Microsoft::Samples::SimpleHGlobalWrapper::Size()
{
	return GlobalSize(m_hMem);
}
