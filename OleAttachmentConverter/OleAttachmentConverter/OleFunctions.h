namespace Microsoft
{
	namespace Samples 
	{
		namespace OleAttachmentConverter
		{
			STDMETHODIMP PersistOleStream(LPSTREAM lpOleStream, std::wstring lpszFileName, LPSTREAM * lppDestStream);
			STDMETHODIMP GetStructuredStorage(std::wstring lpszFileName, LPSTORAGE * lppStg);
			STDMETHODIMP PersistStructuredStorage(LPSTREAM lpOleStream, LPSTORAGE * lppStg);
			void EnumStorage(LPSTORAGE lppStg);
			STDMETHODIMP LoadMetaFilePict(LPSTORAGE pStgSrc, HMETAFILEPICT * phmfp);
			STDMETHODIMP EnumCache(IOleCache * pCache);
			STDMETHODIMP GetSTGMEDIUMFromCache(IOleCache * pCache, LPSTGMEDIUM lpStgMedium);
			STDMETHODIMP SaveBitmapToFileSystem(STGMEDIUM StgMed, std::wstring lpszBmpOutputFileName);
		}
	}
}