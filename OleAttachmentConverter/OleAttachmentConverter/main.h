
typedef struct tagCmdLineArgs 
{
	std::wstring lpszProfile;
	std::wstring lpszSubject;
	std::wstring lpszFileName;
	std::wstring lpszBmpOutputFileName;
	bool bUseMSG;
	bool bEnumCache;
} CMDLINEARGS, *LPCMDLINEARGS;

namespace Microsoft
{
	namespace Samples 
	{
		namespace OleAttachmentConverter
		{
			bool ParseCmdLineArgs(int argc, wchar_t * argv[], LPCMDLINEARGS lpCmdLineArgs);
			void DisplayUsage();
			int Main(CMDLINEARGS CmdLineArgs);
			void PullFromInbox(CMDLINEARGS CmdLineArgs);
			void PullFromMSG(CMDLINEARGS CmdLineArgs);
		}
	}
}