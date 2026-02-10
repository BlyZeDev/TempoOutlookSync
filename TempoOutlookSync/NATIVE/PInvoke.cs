namespace TempoOutlookSync.NATIVE;

using System.Runtime.InteropServices;

public static partial class PInvoke
{
    private const string Kernel32 = "kernel32.dll";
    private const string Shell32 = "shell32.dll";
    private const string User32 = "user32.dll";

    public const int ICON_BIG = 1;
    public const int WM_SETICON = 0x80;
    public const int SW_HIDE = 0;

    [LibraryImport(Kernel32, SetLastError = true)]
    public static partial nint GetConsoleWindow();

    [LibraryImport(Shell32, SetLastError = true, EntryPoint = $"{nameof(ExtractIcon)}W", StringMarshalling = StringMarshalling.Utf16)]
    public static partial nint ExtractIcon(nint hInst, string lpszExeFileName, int nIconIndex);

    [LibraryImport(Shell32, SetLastError = true, EntryPoint = $"{nameof(ExtractIconEx)}W", StringMarshalling = StringMarshalling.Utf16)]
    public static unsafe partial uint ExtractIconEx(string lpszFile, int nIconIndex, nint* phiconLarge, nint* phiconSmall, uint nIcons);

    [LibraryImport(User32, SetLastError = true, EntryPoint = $"{nameof(PostMessage)}W")]
    public static partial nint PostMessage(nint hWnd, int msg, nint wParam, nint lParam);

    [LibraryImport(User32, SetLastError = true)]
    public static partial void ShowWindowAsync(nint hWnd, int nCmdShow);
}