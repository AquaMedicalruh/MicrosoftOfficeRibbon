#include <Windows.h>
#include <atlbase.h>
#include <atlcom.h>
#include <atlstr.h>
#include <oleacc.h>
#include <Shlwapi.h>
#include <msclr\marshal_cppstd.h>
#include <comdef.h>
#include <office.h>
#include <iostream>

int main()
{
    CoInitialize(NULL);
    _ApplicationPtr pApp("PowerPoint.Application");
    HWND hWnd = (HWND)pApp->hWnd;
    HWND hRibbon = FindWindowEx(hWnd, NULL, L"Ribbon", NULL);
    SetWindowPos(hRibbon, NULL, 0, 0, 500, 100, SWP_NOMOVE);
    CoUninitialize();
    return 0;
}
