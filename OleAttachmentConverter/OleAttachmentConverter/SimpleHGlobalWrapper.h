#pragma once
namespace Microsoft
{
	namespace Samples
	{
		class SimpleHGlobalWrapper
		{
		public:
			SimpleHGlobalWrapper(void);
			SimpleHGlobalWrapper(HGLOBAL);
			~SimpleHGlobalWrapper(void);
			STDMETHOD_(LPVOID, GetPtr)();
			STDMETHOD_(SIZE_T, Size)();
		private:
			HGLOBAL m_hMem;
			LPVOID	m_lpv;
		};
	}
}
