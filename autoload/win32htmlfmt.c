/* win32htmlfmt.c
 * Written by wz520 [wingzero1040@gmail.com]
 * Last Update: 2013-03-22
 *
 * This is a win32-DLL module. It is designed to be called by Vim script's
 * libcall() function.

 -------
 COMPILE
 -------
mingw32:
 On the Vim command-line:
 :!gcc % -shared -mwindows -s -O2 -o %<.dll

 */

#include <windows.h>

#define DLLEXPORT				_declspec(dllexport)

#define _FILE_WIN32HTMLFMT  // change behavior of GetClipData()

// *NOTE*
// FreeClipDataPointer() must be called before setting new value to pClipData
static void* pClipData = NULL;
void FreeClipDataPointer(void) {
	free(pClipData);
	pClipData = NULL;
}

// EXPORTS
//
// return "HTML Format" data, do not free() the returned pointer.
char* DLLEXPORT GetHTMLFormat(void);

// Get clipboard data with specified format
// The pointer to the clipboard data will be stored in *ppClipData
// Return the size of the data.
// If there was any error(e.g. not found), *ppClipData will be NULL, and the
// return value is undefined.
//
// !!*NOTE*!!
// For portability, use *ppClipData to return the pointer to the new data,
// but in this module, ppClipData argument does not exist.(takes only 1 arg)
size_t GetClipData(
#ifndef _FILE_WIN32HTMLFMT
		void** ppClipData,
#endif
		UINT uFormat)
{
	size_t size = 0;

#ifndef _FILE_WIN32HTMLFMT
	*ppClipData = NULL;
#else
	void** ppClipData = &pClipData;
	FreeClipDataPointer();
#endif

	if (IsClipboardFormatAvailable(uFormat) && OpenClipboard(NULL))
	{
		void* pBuf = NULL;
		HGLOBAL hMem = GetClipboardData(uFormat);
		if (hMem)
		{
			size = GlobalSize(hMem);
			pBuf = (void*)GlobalLock(hMem);
			if (pBuf)
			{
				*ppClipData = (void*)malloc( size );
				if (*ppClipData)
				{
					memcpy(*ppClipData, pBuf, size);
				}
			}
			GlobalUnlock(hMem);
		}
		CloseClipboard();
	}

	return size;
}

char* DLLEXPORT GetHTMLFormat(void)
{
	UINT cf_html = RegisterClipboardFormat("HTML Format");
	(void)GetClipData(cf_html);
	return pClipData;
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_DETACH:
		FreeClipDataPointer();
		break;
	}
	return TRUE;
}

