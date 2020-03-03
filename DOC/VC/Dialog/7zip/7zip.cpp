#include <windows.h>

HINSTANCE hinst = NULL;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
		hinst = (HINSTANCE)hModule;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

enum TSPFileType {ftDelphiDll, ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE};

extern int WINAPI WinMain(HINSTANCE hInstance,HINSTANCE hPrevInst,LPSTR lpCmdLine,int nCmdShow);

extern "C" __declspec(dllexport) void db_ShowDllForm_Plugins(void** frm, int* spFileType, char** strParentName, char** strSubModuleName, char** strClassName, char** strWindowName, char** strIconFileName, const bool show = true)
{
    * spFileType       = ftVCDialogDll;  // TSPFileType
    * frm              = NULL;           // VC 置空
    * strParentName    = "文件管理";     // 父模块名称
    * strSubModuleName = "7-zip";        // 子模块名称
    * strClassName     = "FM";           // 窗体类名
    * strWindowName    = "7-zip";        // 窗体名
    * strIconFileName  = "";

    if (show) 
    {
      WinMain(hinst, 0, (LPSTR)"", (int)show);
    }
}
