#include "stdafx.h"
#include "ManagedWrapper.h"


CManagedWrapper::CManagedWrapper()
{

}

CManagedWrapper::~CManagedWrapper()
{

}

CManagedWrapper ^ CManagedWrapper::CreateWrapper()
{
	if (m_context == nullptr)
	{
		m_context == gcnew CManagedWrapper();
	}
	return m_context;
}

void CManagedWrapper::CreateService()
{
	throw gcnew System::NotImplementedException();
}

void CManagedWrapper::CreateItem()
{
	throw gcnew System::NotImplementedException();
}

void CManagedWrapper::SetProperties()
{
	throw gcnew System::NotImplementedException();
}

void CManagedWrapper::SaveChanges()
{
	throw gcnew System::NotImplementedException();
}
