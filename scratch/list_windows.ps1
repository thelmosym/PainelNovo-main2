Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class WinFind {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@

$proc = Get-Process excel -ErrorAction SilentlyContinue | Select-Object -First 1
if ($proc) {
    [WinFind]::EnumWindows({
        param($hwnd, $lparam)
        $pid = 0
        [WinFind]::GetWindowThreadProcessId($hwnd, [ref]$pid)
        if ($pid -eq $proc.Id) {
            $sb = New-Object System.Text.StringBuilder 512
            [WinFind]::GetWindowText($hwnd, $sb, 512) | Out-Null
            $txt = $sb.ToString()
            if ($txt) {
                Write-Host "Window: '$txt'"
            }
        }
        return $true
    }, 0)
}

