

#include <windows.h>


LRESULT CALLBACK CheckKey (int nCode, WPARAM wParam, LPARAM lParam); 
#pragma data_seg(".shared")
  HHOOK g_hHook=NULL;
  HWND g_hWnd=NULL;
#pragma data_seg()
#pragma comment(linker, "/section:.shared,RWS") 

HINSTANCE g_hInstance=NULL; 

long (__stdcall SetHook(HWND hWnd))
{    
	if (g_hHook == NULL)
	{

		g_hWnd = hWnd;    
		g_hHook = SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)CheckKey, g_hInstance, NULL);
		if (!g_hHook)
		{
			return -1;
		}
		return 1;
	}
	return -1;
}

long (__stdcall RemoveHook())
{    
	long retval;
	retval = UnhookWindowsHookEx(g_hHook);
	g_hHook = NULL;
	return retval;
}

LRESULT (__stdcall CALLBACK CheckKey(int nCode, WPARAM wParam, LPARAM lParam))
{    
	if (nCode == HC_ACTION)
	{
	   SendMessage(g_hWnd, WM_USER, wParam, lParam);
	   //PostMessage(g_hWnd, WM_KEYUP, wParam, lParam);
	}
    return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}

BOOL WINAPI DllMain(HINSTANCE hInst, ULONG uReason, LPVOID)
{
  	if(uReason == DLL_THREAD_DETACH || uReason == DLL_PROCESS_DETACH)
	{
		RemoveHook();
		return TRUE;
	}
	g_hInstance = hInst;        
	//DisableThreadLibraryCalls(hInst);    
    return TRUE;
}

 
