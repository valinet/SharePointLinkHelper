//#undef UNICODE
#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#include <Windows.h>
#include <Shlobj.h>
#include <shlwapi.h>
//#pragma comment(linker, "/NODEFAULTLIB")
//#pragma comment(linker, "/ENTRY:wWinMain")

#ifdef UNICODE
#define len wcslen
#define cat wcscat_s
#define str wcsstr
#define chr wcschr
#else
#define len strlen
#define cat strcat_s
#define str strstr
#define chr strchr
static CHAR* widestr2str(const WCHAR* wstr)
{
    int wstr_len = (int)wcslen(wstr);
    int num_chars = WideCharToMultiByte(CP_UTF8, 0, wstr, wstr_len, NULL, 0, NULL, NULL);
    CHAR* strTo = (CHAR*)malloc((num_chars + 1) * sizeof(CHAR));
    if (strTo)
    {
        WideCharToMultiByte(CP_UTF8, 0, wstr, wstr_len, strTo, num_chars, NULL, NULL);
        strTo[num_chars] = '\0';
    }
    return strTo;
}
#endif

#define APP_NAME                                                                          "SharePointLinkHelper"
#define APP_NAME_LENGTH                                                                      len(TEXT(APP_NAME))
#define PROTOCOL_NAME                                                                             "splinkhelper"
#define EXE_NAME                                                                     TEXT(APP_NAME) TEXT(".exe")
#define SP_SITE                                                                              "http://sharepoint"
#define WELCOME_PAGE                            TEXT(SP_SITE) TEXT("/SitePages/SharePoint%20Link%20Helper.aspx")

const char reg[] =
"Windows Registry Editor Version 5.00\r\n"
"[HKEY_CURRENT_USER\\SOFTWARE\\Policies\\Microsoft\\Edge]\r\n"
"\"AutoLaunchProtocolsFromOrigins\" = \"[{\\\"allowed_origins\\\": [\\\"" SP_SITE "\\\"], \\\"protocol\\\": \\\"" PROTOCOL_NAME "\\\"}]\"\r\n"
"[HKEY_CURRENT_USER\\SOFTWARE\\Policies\\Google\\Chrome]\r\n"
"\"AutoLaunchProtocolsFromOrigins\" = \"[{\\\"allowed_origins\\\": [\\\"" SP_SITE "\\\"], \\\"protocol\\\": \\\"" PROTOCOL_NAME "\\\"}]\"\r\n"
;

HRESULT GetWorkingDirectory(TCHAR* directory)
{
    TCHAR* szAppName = EXE_NAME;
    BOOL bRet = FALSE;
    HRESULT hr = SHGetFolderPath(
        NULL,
        CSIDL_APPDATA,
        NULL,
        SHGFP_TYPE_CURRENT,
        directory
    );
    if (SUCCEEDED(hr))
    {
        memcpy(
            directory + len(directory),
            TEXT("\\") TEXT(APP_NAME),
            (1 + APP_NAME_LENGTH + 1) * sizeof(TCHAR)
        );
        bRet = CreateDirectory(
            directory,
            NULL
        );
        if (bRet || (bRet == 0 && GetLastError() == ERROR_ALREADY_EXISTS))
        {
            memcpy(
                directory + len(directory),
                TEXT("\\"),
                (1 + 1) * sizeof(TCHAR)
            );
            memcpy(
                directory + len(directory),
                szAppName,
                (len(szAppName) + 1) * sizeof(TCHAR)
            );
        }
        else
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }
    else
    {
        return hr;
    }
    return S_OK;
}

HKEY GetRootKey()
{
    /*
    if (IsUserAnAdmin())
    {
        return HKEY_CLASSES_ROOT;
    }
    */
    return HKEY_CURRENT_USER;
}

HRESULT InstallOrUninstall(HRESULT uninstall)
{
    HRESULT hr = S_OK;
    LSTATUS ret = 0;
    DWORD dwLen;
    TCHAR szCommandName[MAX_PATH + 7];
    TCHAR szAppName[MAX_PATH];
    HKEY root = GetRootKey();
    DWORD dwDisposition;
    DWORD dwVal = 2162688;
    HKEY key;
    HMODULE hModule;
    HINSTANCE hErr;

    if ((hModule = GetModuleHandle(NULL)) == NULL)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    if (!GetModuleFileName(
        hModule,
        szAppName,
        MAX_PATH
    ))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    if (FAILED(hr = GetWorkingDirectory(szCommandName)))
    {
        return hr;
    }
    if (uninstall == S_FALSE)
    {
        if (!CopyFile(
            szAppName,
            szCommandName,
            FALSE
        ))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }
    else
    {
        if (!DeleteFile(
            szCommandName
        ))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        cat(
            szCommandName,
            MAX_PATH,
            TEXT(".reg")
        );
        if (!DeleteFile(
            szCommandName
        ))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        for (int i = len(szCommandName) - 1; i >= 0; i--)
        {
            if (szCommandName[i] == '\\')
            {
                szCommandName[i] = 0;
                break;
            }
        }
        if (!RemoveDirectory(szCommandName))
        {
        }
    }
    if (FAILED(hr = GetWorkingDirectory(szCommandName + 1)))
    {
        return hr;
    }
    dwLen = len(szCommandName);
    szCommandName[0] = '"';
    szCommandName[dwLen] = '"';
    szCommandName[dwLen + 1] = ' ';
    szCommandName[dwLen + 2] = '"';
    szCommandName[dwLen + 3] = '%';
    szCommandName[dwLen + 4] = '1';
    szCommandName[dwLen + 5] = '"';
    szCommandName[dwLen + 6] = 0;
    if (uninstall == S_FALSE)
    {
        if ((ret = RegCreateKeyEx(
            root,
            TEXT("SOFTWARE\\Classes\\")
            TEXT(PROTOCOL_NAME),
            0,
            NULL,
            REG_OPTION_NON_VOLATILE,
            KEY_ALL_ACCESS,
            NULL,
            &key,
            &dwDisposition
        )) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }
        if ((ret = RegSetValueEx(
            key,
            TEXT(""),
            0,
            REG_SZ,
            TEXT("URL:")
            TEXT(PROTOCOL_NAME)
            TEXT(" protocol"),
            (4 + len(TEXT(PROTOCOL_NAME)) + 9) * sizeof(TCHAR)
        )) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }
        if ((ret = RegSetValueEx(
            key,
            TEXT("EditFlags"),
            0,
            REG_DWORD,
            &dwVal,
            sizeof(DWORD)
        )) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }
        if ((ret = RegSetValueEx(
            key,
            TEXT("URL Protocol"),
            0,
            REG_SZ,
            L"",
            sizeof(TCHAR)
        )) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }
        if ((ret = RegCloseKey(key)) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }

        if ((ret = RegCreateKeyEx(
            root,
            TEXT("SOFTWARE\\Classes\\")
            TEXT(PROTOCOL_NAME)
            TEXT("\\shell\\open\\command"),
            0,
            NULL,
            REG_OPTION_NON_VOLATILE,
            KEY_ALL_ACCESS,
            NULL,
            &key,
            &dwDisposition
        )) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }
        if ((ret = RegSetValueEx(
            key,
            L"",
            0,
            REG_SZ,
            szCommandName,
            (len(szCommandName) + 1) * sizeof(TCHAR)
        )) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }
        if ((ret = RegCloseKey(key)) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }

        if (FAILED(hr = GetWorkingDirectory(szCommandName)))
        {
            return hr;
        }
        cat(
            szCommandName,
            MAX_PATH,
            TEXT(".reg")
        );
        HANDLE hFile = CreateFile(
            szCommandName,
            GENERIC_WRITE,
            FILE_SHARE_READ,
            NULL,
            CREATE_ALWAYS,
            FILE_ATTRIBUTE_NORMAL,
            NULL
        );
        if (hFile == INVALID_HANDLE_VALUE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        DWORD bytesWritten;
        if (!WriteFile(
            hFile,
            reg,
            len(reg) * sizeof(TCHAR),
            &bytesWritten,
            NULL
        ))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (!CloseHandle(hFile))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if ((hErr = ShellExecute(
            NULL,
            TEXT("open"),
            WELCOME_PAGE,
            NULL,
            NULL,
            SW_SHOWDEFAULT
        )) <= 32)
        {
            return hErr;
        }
        if ((hErr = ShellExecute(
            NULL,
            TEXT("open"),
            szCommandName,
            NULL,
            NULL,
            SW_SHOWDEFAULT
        )) <= 32)
        {
            return hErr;
        }
    }
    else
    {
        if ((ret = RegDeleteTree(
            root,
            TEXT("SOFTWARE\\Classes\\")
            TEXT(PROTOCOL_NAME)
        )) != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(ret);
        }
        MessageBox(
            NULL,
            TEXT("The product has been uninstalled successfully."),
            TEXT(APP_NAME),
            MB_ICONINFORMATION
        );
    }
    return S_OK;
}

HRESULT CheckIfInstalled()
{
    LSTATUS ret = 0;
    HKEY root = GetRootKey();

    HKEY key;
    if ((ret = RegOpenKeyEx(
        root,
        TEXT("SOFTWARE\\Classes\\")
        TEXT(PROTOCOL_NAME),
        0,
        KEY_ALL_ACCESS,
        &key
    )) != ERROR_SUCCESS)
    {
        return S_FALSE;
    }
    if ((ret = RegCloseKey(key)) != ERROR_SUCCESS)
    {
        return HRESULT_FROM_WIN32(ret);
    }
    return S_OK;
}

HRESULT LaunchExecutableForUrl(TCHAR* url)
{
    HINSTANCE hErr;
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr))
    {
        return hr;
    }
    /*MessageBox(
        NULL,
        url,
        NULL,
        NULL
    );*/
    if (str(url, TEXT(PROTOCOL_NAME ":///#")))
    {
        TCHAR explorerPath[MAX_PATH];
        ZeroMemory(explorerPath, MAX_PATH * sizeof(TCHAR));
        hr = SHGetFolderPath(
            NULL,
            CSIDL_WINDOWS,
            NULL,
            SHGFP_TYPE_CURRENT,
            explorerPath
        );
        if (SUCCEEDED(hr))
        {
            cat(
                explorerPath,
                MAX_PATH,
                TEXT("\\explorer.exe")
            );
            // 5 = len(":///#")
            // 7 = len("http://")
            // 2 = len("\"") + len("\"")
            // 2 = len("\\\\")
            // 10 = len("davwwwroot")
            // 100 = some safety padding
            TCHAR* path = HeapAlloc(
                GetProcessHeap(),
                HEAP_ZERO_MEMORY,
                (len(url) - len(TEXT(PROTOCOL_NAME)) - 5 - 7 + 2 + 2 + 10 + 100) * sizeof(TCHAR)
            );
            if (!path)
            {
                return HRESULT_FROM_WIN32(STATUS_NO_MEMORY);
            }
            path[0] = TEXT('"');
            path[1] = TEXT('\\');
            path[2] = TEXT('\\');
            memcpy(
                path + 3,
                url + len(TEXT(PROTOCOL_NAME)) + 5 + 7,
                (len(url) - len(TEXT(PROTOCOL_NAME)) - 5 - 7) * sizeof(TCHAR)
            );
            TCHAR* test = chr(path, TEXT('/'));
            memcpy(
                test + 1,
                TEXT("davwwwroot"),
                10 * sizeof(TCHAR)
            );
            TCHAR* p = chr(url + len(TEXT(PROTOCOL_NAME)) + 5 + 7, '/');
            memcpy(
                test + 11,
                p,
                len(p) * sizeof(TCHAR)
            );
            test[11 + len(p)] = TEXT('\0');
            DWORD i;
            for (i = 0; i < len(path); ++i)
            {
                if (path[i] == TEXT('/'))
                {
                    path[i] = TEXT('\\');
                }
            }
            path[i] = TEXT('"');
            path[i + 1] = TEXT('\0');
            /*MessageBox(
                NULL,
                path,
                NULL,
                NULL
            );*/
            DWORD pcchUnescaped;
            hr = UrlUnescape(
                path,
                NULL,
                &pcchUnescaped,
                URL_UNESCAPE_INPLACE | URL_UNESCAPE_AS_UTF8
            );
            if (FAILED(hr))
            {
                return hr;
            }
            /*MessageBox(
                NULL,
                path,
                NULL,
                NULL
            );*/
            if ((hErr = ShellExecute(
                NULL,
                NULL,
                explorerPath,
                path,
                NULL,
                SW_SHOWDEFAULT
            )) <= 32)
            {
                return hErr;
            }
            if (!HeapFree(
                GetProcessHeap(),
                0,
                path
            ))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }
    }
    else
    {
        TCHAR* ext = PathFindExtension(url);
        TCHAR* path;
        DWORD len = 0;
        hr = AssocQueryString(
            ASSOCF_NONE,
            ASSOCSTR_EXECUTABLE,
            ext,
            NULL,
            NULL,
            &len
        );
        if (hr != S_FALSE)
        {
            return hr;
        }
        path = HeapAlloc(
            GetProcessHeap(),
            0,
            len * sizeof(TCHAR)
        );
        if (!path)
        {
            return HRESULT_FROM_WIN32(STATUS_NO_MEMORY);
        }
        hr = AssocQueryString(
            ASSOCF_NONE,
            ASSOCSTR_EXECUTABLE,
            ext,
            NULL,
            path,
            &len
        );
        if (FAILED(hr))
        {
            return hr;
        }
        // 1 = string null terminator
        // 3 = len("://")
        if ((hErr = ShellExecute(
            NULL,
            NULL,
            path,
            url + (sizeof(TEXT(PROTOCOL_NAME)) / sizeof(TCHAR) - 1) + 3,
            NULL,
            SW_SHOWDEFAULT
        )) <= 32)
        {
            return hErr;
        }
        if (!HeapFree(
            GetProcessHeap(),
            0,
            path
        ))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }
    CoUninitialize();
    return S_OK;
}


int WINAPI wWinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine,
    _In_ int nShowCmd
)
{
    HRESULT hr = S_OK;

    if (!SetProcessDpiAwarenessContext(
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    ))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    size_t argc = 0;
    LPCWSTR cmdLine = GetCommandLineW();
    LPWSTR* Wargv = CommandLineToArgvW(cmdLine, &argc);
    if (argc == 2)
    {
#ifdef UNICODE
        hr = LaunchExecutableForUrl(Wargv[1]);
#else
        char* argv1 = widestr2str(Wargv[1]);
        hr = LaunchExecutableForUrl(argv1);
        free(argv1);
#endif
    }
    else if (argc == 1)
    {
        hr = CheckIfInstalled();
        if (SUCCEEDED(hr) || hr == S_FALSE)
        {
            hr = InstallOrUninstall(hr);
        }
    }
    if (LocalFree(Wargv))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    /*TCHAR buf[10];
    _itow_s(hr, buf, 10, 10);
    MessageBox(
        NULL,
        buf,
        NULL,
        NULL
    );*/
    ExitProcess(hr);
}