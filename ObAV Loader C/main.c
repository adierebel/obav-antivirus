#define WIN32_LEAN_AND_MEAN
#define NOCRYPT
#define NOSERVICE
#define NOMCX
#define NOIME

#include <windows.h>
#include <shellapi.h>

/****************************************************************************
 * Function: WinMain
 ****************************************************************************/

int APIENTRY wWinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPWSTR lpwszCmdline, int nCmdShow)
{
	ShellExecuteW( 0, L"open", lpwszCmdline, 0, 0, 1 );
	return 0;
}
