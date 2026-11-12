// QuillCnv - Microsoft Word import converter for Sinclair QL Quill documents.
//
// Word loads registered text converters (.cnv DLLs) through the converter API described
// in the WinWord Converter SDK (GC1039, convapi.h). When Word cannot read a file natively it
// asks each converter's IsFormatCorrect32 whether it recognises the file, so a Quill ".doc"
// is detected by its contents even though Word's own .doc format shares the extension.
//
// The conversion itself is done by QuillToDocx.exe (installed next to this DLL), which
// writes RTF; ForeignToRtf32 streams that RTF to Word. Import only: Word saves the
// document in its own formats.

#include <windows.h>

#include <cstdarg>
#include <cstdio>
#include <cstring>
#include <string>

typedef short FCE;

#define fceNoErr 0
#define fceOpenInFileErr (-1)
#define fceReadErr (-2)
#define fceInvalidFile (-5)
#define fceNoMemory (-8)
#define fceUserCancel (-13)
#define fceWrongFileType (-14)
#define fceTrue 1
#define fceFalse 0

#define fRegAppSupportNonOem 0x00000008

#define RegAppOpcodeVer 0x01
#define RegAppOpcodeDocfile 0x02
#define RegAppOpcodeCharset 0x03

typedef long(PASCAL* PFN_RTF)(long cchBuff, long nPercent);

#pragma pack(push, 1)
struct RegAppRet
{
    short cbStruct;
    char cbSizefDocfile;
    char opcodefDocfile;
    short grfType;  // bit 0 docfile, bit 1 non-docfile
    char cbSizeVer;
    char opcodeVer;
    short verMajor;
    short verMinor;
    char cbSizeCharset;
    char opcodeCharset;
    char charset;
};
#pragma pack(pop)

namespace
{
    const char kClassName[] = "Quill";
    const char kDescription[] = "Sinclair QL Quill";
    const char kExtensions[] = "doc";
    const SIZE_T kChunk = 16 * 1024;
    const DWORD kConvertTimeoutMs = 60 * 1000;

    HMODULE g_module = nullptr;
    // Filenames arrive in the OEM code page unless RegisterApp agreed on ANSI.
    UINT g_filenameCodePage = CP_OEMCP;
    FCE g_lastError = fceNoErr;
    char g_lastMessage[512] = "";

    std::wstring ModuleDirectory()
    {
        wchar_t path[MAX_PATH];
        DWORD n = GetModuleFileNameW(g_module, path, MAX_PATH);
        if (n == 0 || n == MAX_PATH)
            return L"";
        std::wstring dir(path, n);
        size_t slash = dir.find_last_of(L"\\/");
        return slash == std::wstring::npos ? L"" : dir.substr(0, slash);
    }

    // Diagnostics go to %TEMP%\QuillCnv.log, but only if that file already exists.
    void Log(const char* format, ...)
    {
        wchar_t temp[MAX_PATH];
        DWORD tempLength = GetTempPathW(MAX_PATH, temp);
        if (tempLength == 0 || tempLength >= MAX_PATH)
            return;
        std::wstring logPath = std::wstring(temp) + L"QuillCnv.log";
        if (GetFileAttributesW(logPath.c_str()) == INVALID_FILE_ATTRIBUTES)
            return;

        char line[2048];
        va_list args;
        va_start(args, format);
        vsnprintf(line, sizeof(line), format, args);
        va_end(args);

        FILE* f = nullptr;
        if (_wfopen_s(&f, logPath.c_str(), L"a") == 0 && f)
        {
            SYSTEMTIME t;
            GetLocalTime(&t);
            fprintf(f, "%02d:%02d:%02d.%03d [%lu] %s\n", t.wHour, t.wMinute, t.wSecond, t.wMilliseconds,
                    GetCurrentProcessId(), line);
            fclose(f);
        }
    }

    std::string Utf8(const std::wstring& s)
    {
        int n = WideCharToMultiByte(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0, nullptr, nullptr);
        std::string out(n, '\0');
        WideCharToMultiByte(CP_UTF8, 0, s.c_str(), (int)s.size(), &out[0], n, nullptr, nullptr);
        return out;
    }

    FCE Fail(FCE fce, const char* message)
    {
        g_lastError = fce;
        strncpy_s(g_lastMessage, message, _TRUNCATE);
        Log("error %d: %s", fce, message);
        return fce;
    }

    bool SetGlobalString(HANDLE h, const char* s)
    {
        if (!h)
            return false;
        SIZE_T n = strlen(s) + 1;
        if (GlobalSize(h) < n)
        {
            HGLOBAL resized = GlobalReAlloc(h, n, GMEM_MOVEABLE);
            if (resized != h)
                return false;
        }
        char* p = (char*)GlobalLock(h);
        if (!p)
            return false;
        memcpy(p, s, n);
        GlobalUnlock(h);
        return true;
    }

    std::wstring Decode(const char* s, UINT codePage)
    {
        int n = MultiByteToWideChar(codePage, 0, s, -1, nullptr, 0);
        if (n <= 1)
            return L"";
        std::wstring w(n, L'\0');
        MultiByteToWideChar(codePage, 0, s, -1, &w[0], n);
        w.resize(n - 1);
        return w;
    }

    // The filename's code page depends on what RegisterApp negotiated; if the decoded name
    // doesn't exist, try the other one before giving up.
    std::wstring FileNameFromHandle(HANDLE gh)
    {
        if (!gh)
            return L"";
        const char* p = (const char*)GlobalLock(gh);
        if (!p)
            return L"";
        std::string raw(p);
        GlobalUnlock(gh);

        std::wstring name = Decode(raw.c_str(), g_filenameCodePage);
        if (GetFileAttributesW(name.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            std::wstring other = Decode(raw.c_str(), g_filenameCodePage == CP_ACP ? CP_OEMCP : CP_ACP);
            if (GetFileAttributesW(other.c_str()) != INVALID_FILE_ATTRIBUTES)
                return other;
        }
        return name;
    }

    // Quill files start with the header length (20) as a big-endian (QL) or little-endian (PC)
    // word, followed by the identifier "vrm1qdf0".
    bool IsQuillFile(const std::wstring& path)
    {
        HANDLE f = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                               nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (f == INVALID_HANDLE_VALUE)
            return false;
        unsigned char b[10];
        DWORD got = 0;
        BOOL ok = ReadFile(f, b, sizeof(b), &got, nullptr);
        CloseHandle(f);
        return ok && got == sizeof(b) && ((b[0] == 0x00 && b[1] == 0x14) || (b[0] == 0x14 && b[1] == 0x00)) &&
               memcmp(b + 2, "vrm1qdf0", 8) == 0;
    }

    std::wstring TempRtfPath()
    {
        wchar_t dir[MAX_PATH];
        DWORD n = GetTempPathW(MAX_PATH, dir);
        if (n == 0 || n >= MAX_PATH)
            return L"";
        wchar_t name[64];
        swprintf_s(name, L"QuillCnv-%lu-%llu.rtf", GetCurrentProcessId(), GetTickCount64());
        return std::wstring(dir) + name;
    }

    // Runs QuillToDocx.exe <input> <output.rtf> without a console window.
    FCE RunConverter(const std::wstring& input, const std::wstring& output)
    {
        std::wstring exe = ModuleDirectory() + L"\\QuillToDocx.exe";
        if (GetFileAttributesW(exe.c_str()) == INVALID_FILE_ATTRIBUTES)
            return Fail(fceOpenInFileErr, "QuillToDocx.exe is missing from the converter's folder.");

        std::wstring cmd = L"\"" + exe + L"\" \"" + input + L"\" \"" + output + L"\"";
        STARTUPINFOW si = {};
        si.cb = sizeof(si);
        si.dwFlags = STARTF_USESHOWWINDOW;
        si.wShowWindow = SW_HIDE;
        PROCESS_INFORMATION pi = {};

        if (!CreateProcessW(exe.c_str(), &cmd[0], nullptr, nullptr, FALSE, CREATE_NO_WINDOW, nullptr, nullptr, &si,
                            &pi))
            return Fail(fceOpenInFileErr, "Could not start QuillToDocx.exe.");

        DWORD wait = WaitForSingleObject(pi.hProcess, kConvertTimeoutMs);
        DWORD exitCode = 1;
        if (wait == WAIT_OBJECT_0)
            GetExitCodeProcess(pi.hProcess, &exitCode);
        else
            TerminateProcess(pi.hProcess, 1);
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);

        if (wait != WAIT_OBJECT_0)
            return Fail(fceReadErr, "QuillToDocx.exe timed out.");
        if (exitCode != 0)
            return Fail(fceInvalidFile, "QuillToDocx.exe could not convert this Quill document.");
        return fceNoErr;
    }

    bool ReadWholeFile(const std::wstring& path, std::string& out)
    {
        HANDLE f = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING,
                               FILE_ATTRIBUTE_NORMAL, nullptr);
        if (f == INVALID_HANDLE_VALUE)
            return false;
        LARGE_INTEGER size;
        bool ok = GetFileSizeEx(f, &size) && size.QuadPart < 0x7FFFFFFF;
        if (ok)
        {
            out.resize((size_t)size.QuadPart);
            DWORD got = 0;
            ok = out.empty() || (ReadFile(f, &out[0], (DWORD)out.size(), &got, nullptr) && got == out.size());
        }
        CloseHandle(f);
        return ok;
    }

    // Hands the RTF to Word through ghBuff, one chunk per lpfnOut call.
    FCE StreamRtf(const std::string& rtf, HANDLE ghBuff, PFN_RTF lpfnOut)
    {
        size_t offset = 0;
        while (offset < rtf.size())
        {
            if (GlobalSize(ghBuff) < kChunk)
            {
                HGLOBAL resized = GlobalReAlloc(ghBuff, kChunk, GMEM_MOVEABLE);
                if (resized != ghBuff)
                    return Fail(fceNoMemory, "Could not size Word's RTF buffer.");
            }
            SIZE_T capacity = GlobalSize(ghBuff);
            size_t n = rtf.size() - offset < capacity ? rtf.size() - offset : (size_t)capacity;

            char* p = (char*)GlobalLock(ghBuff);
            if (!p)
                return Fail(fceNoMemory, "Could not lock Word's RTF buffer.");
            memcpy(p, rtf.data() + offset, n);
            GlobalUnlock(ghBuff);

            offset += n;
            long percent = (long)(offset * 100 / rtf.size());
            long rc = lpfnOut((long)n, percent);
            if (rc < 0)
            {
                Log("Word stopped the import (callback returned %ld)", rc);
                return (FCE)rc;
            }
        }
        return fceNoErr;
    }
}

BOOL APIENTRY DllMain(HMODULE module, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_module = module;
        DisableThreadLibraryCalls(module);
    }
    return TRUE;
}

extern "C" long PASCAL InitConverter32(HANDLE /*hWnd*/, char* szModule)
{
    Log("InitConverter32 module=%s", szModule ? szModule : "(null)");
    return 1;
}

extern "C" void PASCAL UninitConverter(void)
{
    Log("UninitConverter");
}

extern "C" HGLOBAL PASCAL RegisterApp(unsigned long lFlags, void* /*lpFuture*/)
{
    bool ansiNames = (lFlags & fRegAppSupportNonOem) != 0;
    g_filenameCodePage = ansiNames ? CP_ACP : CP_OEMCP;
    Log("RegisterApp flags=0x%lx -> filenames in %s", lFlags, ansiNames ? "ANSI" : "OEM");

    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, sizeof(RegAppRet));
    if (!h)
        return nullptr;
    RegAppRet* r = (RegAppRet*)GlobalLock(h);
    r->cbStruct = sizeof(RegAppRet);
    r->cbSizefDocfile = 4;
    r->opcodefDocfile = RegAppOpcodeDocfile;
    r->grfType = 0x0002;  // non-docfile only
    r->cbSizeVer = 6;
    r->opcodeVer = RegAppOpcodeVer;
    r->verMajor = 8;  // RTF as understood by Word 97 and later
    r->verMinor = 0;
    r->cbSizeCharset = 3;
    r->opcodeCharset = RegAppOpcodeCharset;
    r->charset = (char)(ansiNames ? ANSI_CHARSET : OEM_CHARSET);
    GlobalUnlock(h);
    return h;
}

extern "C" void PASCAL GetReadNames(HANDLE haszClass, HANDLE haszDescrip, HANDLE haszExt)
{
    SetGlobalString(haszClass, kClassName);
    SetGlobalString(haszDescrip, kDescription);
    SetGlobalString(haszExt, kExtensions);
}

extern "C" FCE PASCAL IsFormatCorrect32(HANDLE ghszFile, HANDLE ghszClass)
{
    std::wstring path = FileNameFromHandle(ghszFile);
    bool quill = IsQuillFile(path);
    Log("IsFormatCorrect32 %s -> %s", Utf8(path).c_str(), quill ? "Quill" : "not Quill");
    if (quill && ghszClass)
        SetGlobalString(ghszClass, kClassName);
    return quill ? fceTrue : fceFalse;
}

extern "C" FCE PASCAL ForeignToRtf32(HANDLE ghszFile, void* /*pstgForeign*/, HANDLE ghBuff, HANDLE /*ghszClass*/,
                                     HANDLE /*ghszSubset*/, PFN_RTF lpfnOut)
{
    std::wstring path = FileNameFromHandle(ghszFile);
    Log("ForeignToRtf32 %s (buffer %zu bytes)", Utf8(path).c_str(), (size_t)(ghBuff ? GlobalSize(ghBuff) : 0));

    if (!ghBuff || !lpfnOut)
        return Fail(fceNoMemory, "Word passed no RTF buffer or callback.");
    if (!IsQuillFile(path))
        return Fail(fceWrongFileType, "Not a Quill document.");

    std::wstring rtfPath = TempRtfPath();
    if (rtfPath.empty())
        return Fail(fceOpenInFileErr, "No temporary folder.");

    FCE fce = RunConverter(path, rtfPath);
    std::string rtf;
    if (fce == fceNoErr && !ReadWholeFile(rtfPath, rtf))
        fce = Fail(fceReadErr, "Could not read the converted RTF.");
    DeleteFileW(rtfPath.c_str());
    if (fce != fceNoErr)
        return fce;

    fce = StreamRtf(rtf, ghBuff, lpfnOut);
    Log("ForeignToRtf32 streamed %zu bytes, result %d", rtf.size(), fce);
    return fce;
}

extern "C" long PASCAL CchFetchLpszError(long fce, char* lpszError, long cb)
{
    const char* message = (fce == g_lastError && g_lastMessage[0]) ? g_lastMessage : "The Quill converter failed.";
    if (!lpszError || cb <= 0)
        return 0;
    strncpy_s(lpszError, (size_t)cb, message, _TRUNCATE);
    return (long)strlen(lpszError);
}

// Registration is done by Install-QuillConverter.ps1, not by Word.
extern "C" long PASCAL FRegisterConverter(HANDLE /*hkeyRoot*/)
{
    return 0;
}
