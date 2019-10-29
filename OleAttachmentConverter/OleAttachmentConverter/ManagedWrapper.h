#pragma once
ref class CManagedWrapper
{
private:
	CManagedWrapper ^ m_context;
	CManagedWrapper();
	~CManagedWrapper();
public:
	CManagedWrapper();
	static CManagedWrapper ^ CreateWrapper();
	virtual void CreateService();
	virtual void CreateItem();
	virtual void SetProperties();
	virtual void SaveChanges();
};

