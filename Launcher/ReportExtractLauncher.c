#include <windows.h>
#include <wchar.h>
#include <wctype.h>

#define PATH_CAPACITY 32768

static void ShowError(const wchar_t* message)
{
    MessageBoxW(NULL, message, L"ReportExtract", MB_OK | MB_ICONERROR);
}

static int RemoveFileName(wchar_t* path)
{
    wchar_t* separator = wcsrchr(path, L'\\');
    if (separator == NULL)
    {
        return 0;
    }

    *separator = L'\0';
    return 1;
}

static const wchar_t* GetForwardedArguments(const wchar_t* commandLine)
{
    const wchar_t* cursor = commandLine;
    if (*cursor == L'"')
    {
        cursor++;
        while (*cursor != L'\0' && *cursor != L'"')
        {
            cursor++;
        }
        if (*cursor == L'"')
        {
            cursor++;
        }
    }
    else
    {
        while (*cursor != L'\0' && !iswspace(*cursor))
        {
            cursor++;
        }
    }

    while (iswspace(*cursor))
    {
        cursor++;
    }

    return cursor;
}

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE previousInstance, PWSTR commandLine, int showCommand)
{
    wchar_t applicationRoot[PATH_CAPACITY];
    wchar_t backendDirectory[PATH_CAPACITY];
    wchar_t backendExecutable[PATH_CAPACITY];
    wchar_t backendCommandLine[PATH_CAPACITY];
    STARTUPINFOW startupInfo = { 0 };
    PROCESS_INFORMATION processInfo = { 0 };
    DWORD executablePathLength;
    DWORD exitCode = 1;
    const wchar_t* forwardedArguments;

    UNREFERENCED_PARAMETER(instance);
    UNREFERENCED_PARAMETER(previousInstance);
    UNREFERENCED_PARAMETER(commandLine);
    UNREFERENCED_PARAMETER(showCommand);

    executablePathLength = GetModuleFileNameW(NULL, applicationRoot, PATH_CAPACITY);
    if (executablePathLength == 0 || executablePathLength == PATH_CAPACITY)
    {
        ShowError(L"ReportExtract could not resolve its application folder.");
        return 1;
    }

    if (!RemoveFileName(applicationRoot)
        || swprintf_s(backendDirectory, PATH_CAPACITY, L"%s\\Backend", applicationRoot) < 0
        || swprintf_s(backendExecutable, PATH_CAPACITY, L"%s\\ReportExtract.exe", backendDirectory) < 0)
    {
        ShowError(L"ReportExtract could not resolve Backend\\ReportExtract.exe.");
        return 1;
    }

    if (GetFileAttributesW(backendExecutable) == INVALID_FILE_ATTRIBUTES)
    {
        ShowError(L"ReportExtract could not find Backend\\ReportExtract.exe. Re-extract the complete portable folder.");
        return 1;
    }

    forwardedArguments = GetForwardedArguments(GetCommandLineW());
    if (swprintf_s(
            backendCommandLine,
            PATH_CAPACITY,
            L"\"%s\"%s%s",
            backendExecutable,
            *forwardedArguments == L'\0' ? L"" : L" ",
            forwardedArguments) < 0)
    {
        ShowError(L"ReportExtract command-line arguments are too long to forward.");
        return 1;
    }

    startupInfo.cb = sizeof(startupInfo);
    if (!CreateProcessW(
            backendExecutable,
            backendCommandLine,
            NULL,
            NULL,
            FALSE,
            0,
            NULL,
            backendDirectory,
            &startupInfo,
            &processInfo))
    {
        ShowError(L"ReportExtract could not start Backend\\ReportExtract.exe.");
        return 1;
    }

    WaitForSingleObject(processInfo.hProcess, INFINITE);
    GetExitCodeProcess(processInfo.hProcess, &exitCode);
    CloseHandle(processInfo.hThread);
    CloseHandle(processInfo.hProcess);
    return (int)exitCode;
}
