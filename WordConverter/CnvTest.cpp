// CnvTest - drives QuillCnv.cnv the way Word does, without Word.
//
//   CnvTest <converter.cnv> <input file> [output.rtf]
//
// Prints the converter's names and RegisterApp reply, whether IsFormatCorrect32 accepts the
// input, and (given an output path) writes the RTF that ForeignToRtf32 streams.

#include <windows.h>

#include <cstdio>
#include <cstring>
#include <string>

typedef short FCE;
typedef long(PASCAL* PFN_RTF)(long, long);
typedef long(PASCAL* InitConverter32Fn)(HANDLE, char*);
typedef void(PASCAL* UninitConverterFn)(void);
typedef HGLOBAL(PASCAL* RegisterAppFn)(unsigned long, void*);
typedef void(PASCAL* GetReadNamesFn)(HANDLE, HANDLE, HANDLE);
typedef FCE(PASCAL* IsFormatCorrect32Fn)(HANDLE, HANDLE);
typedef FCE(PASCAL* ForeignToRtf32Fn)(HANDLE, void*, HANDLE, HANDLE, HANDLE, PFN_RTF);
typedef long(PASCAL* CchFetchLpszErrorFn)(long, char*, long);

static HANDLE g_buffer;
static FILE* g_out;
static long g_calls;
static long g_bytes;

static long PASCAL OnRtf(long cch, long percent)
{
    const char* p = (const char*)GlobalLock(g_buffer);
    if (g_out && p)
        fwrite(p, 1, (size_t)cch, g_out);
    GlobalUnlock(g_buffer);
    g_calls++;
    g_bytes += cch;
    (void)percent;
    return 0;
}

static HGLOBAL GlobalString(const char* s)
{
    SIZE_T n = strlen(s) + 1;
    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE, n);
    memcpy(GlobalLock(h), s, n);
    GlobalUnlock(h);
    return h;
}

static std::string Read(HGLOBAL h)
{
    const char* p = (const char*)GlobalLock(h);
    std::string s = p ? p : "";
    GlobalUnlock(h);
    return s;
}

int wmain(int argc, wchar_t** argv)
{
    if (argc < 3)
    {
        fprintf(stderr, "usage: CnvTest <converter.cnv> <input file> [output.rtf]\n");
        return 2;
    }

    HMODULE lib = LoadLibraryW(argv[1]);
    if (!lib)
    {
        fprintf(stderr, "cannot load converter (error %lu)\n", GetLastError());
        return 2;
    }

    auto init = (InitConverter32Fn)GetProcAddress(lib, "InitConverter32");
    auto uninit = (UninitConverterFn)GetProcAddress(lib, "UninitConverter");
    auto registerApp = (RegisterAppFn)GetProcAddress(lib, "RegisterApp");
    auto getReadNames = (GetReadNamesFn)GetProcAddress(lib, "GetReadNames");
    auto isFormatCorrect = (IsFormatCorrect32Fn)GetProcAddress(lib, "IsFormatCorrect32");
    auto foreignToRtf = (ForeignToRtf32Fn)GetProcAddress(lib, "ForeignToRtf32");
    auto fetchError = (CchFetchLpszErrorFn)GetProcAddress(lib, "CchFetchLpszError");
    if (!init || !uninit || !registerApp || !getReadNames || !isFormatCorrect || !foreignToRtf || !fetchError)
    {
        fprintf(stderr, "converter is missing an entry point\n");
        return 2;
    }

    char module[] = "CnvTest";
    printf("InitConverter32 -> %ld\n", init(nullptr, module));

    HGLOBAL reply = registerApp(0x00000008 /* fRegAppSupportNonOem */, nullptr);
    if (reply)
    {
        const unsigned char* r = (const unsigned char*)GlobalLock(reply);
        short cb = *(const short*)r;
        printf("RegisterApp -> %d bytes:", cb);
        for (short i = 0; i < cb; i++)
            printf(" %02x", r[i]);
        printf("\n");
        GlobalUnlock(reply);
        GlobalFree(reply);
    }

    HGLOBAL cls = GlobalAlloc(GMEM_MOVEABLE, 1), desc = GlobalAlloc(GMEM_MOVEABLE, 1), ext = GlobalAlloc(GMEM_MOVEABLE, 1);
    getReadNames(cls, desc, ext);
    printf("GetReadNames -> class='%s' description='%s' extensions='%s'\n", Read(cls).c_str(), Read(desc).c_str(),
           Read(ext).c_str());

    // RegisterApp agreed on ANSI filenames.
    char ansi[MAX_PATH * 2] = "";
    WideCharToMultiByte(CP_ACP, 0, argv[2], -1, ansi, sizeof(ansi), nullptr, nullptr);
    HGLOBAL file = GlobalString(ansi);

    HGLOBAL fileClass = GlobalAlloc(GMEM_MOVEABLE, 1);
    FCE correct = isFormatCorrect(file, fileClass);
    printf("IsFormatCorrect32 -> %d (class '%s')\n", correct, correct == 1 ? Read(fileClass).c_str() : "");

    int status = correct == 1 ? 0 : 1;
    if (argc >= 4 && correct == 1)
    {
        if (_wfopen_s(&g_out, argv[3], L"wb") != 0 || !g_out)
        {
            fprintf(stderr, "cannot open output\n");
            return 2;
        }
        g_buffer = GlobalAlloc(GMEM_MOVEABLE, 0);
        DWORD start = GetTickCount();
        FCE fce = foreignToRtf(file, nullptr, g_buffer, fileClass, nullptr, OnRtf);
        DWORD elapsed = GetTickCount() - start;
        fclose(g_out);
        printf("ForeignToRtf32 -> %d, %ld bytes in %ld callbacks, %lu ms\n", fce, g_bytes, g_calls, elapsed);
        if (fce != 0)
        {
            char message[256];
            fetchError(fce, message, sizeof(message));
            printf("CchFetchLpszError -> %s\n", message);
            status = 1;
        }
        GlobalFree(g_buffer);
    }

    uninit();
    FreeLibrary(lib);
    return status;
}
