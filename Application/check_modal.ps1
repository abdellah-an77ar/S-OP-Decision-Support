Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public class WinFinder {
    [DllImport("user32.dll")]
    public static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);
    public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumThreadDelegate lpfn, IntPtr lParam);

    public static string GetDialogText(IntPtr hWnd) {
        var sbAll = new StringBuilder();
        EnumChildWindows(hWnd, (hChild, lParam) => {
            StringBuilder sb = new StringBuilder(512);
            GetWindowText(hChild, sb, 512);
            if (sb.Length > 0) {
                sbAll.AppendLine("      child: " + sb.ToString());
            }
            return true;
        }, IntPtr.Zero);
        return sbAll.ToString();
    }

    public static List<string> GetWindowsForProcess(int pid) {
        var list = new List<string>();
        var proc = System.Diagnostics.Process.GetProcessById(pid);
        foreach (System.Diagnostics.ProcessThread t in proc.Threads) {
            EnumThreadWindows(t.Id, (hWnd, lParam) => {
                StringBuilder sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                StringBuilder cb = new StringBuilder(256);
                GetClassName(hWnd, cb, 256);
                string diag = GetDialogText(hWnd);
                list.Add(string.Format("Handle: {0}, Class: {1}, Text: '{2}', Children:\n{3}", hWnd, cb, sb, diag));
                return true;
            }, IntPtr.Zero);
        }
        return list;
    }
}
"@

$p = Get-Process -Name excel -ErrorAction SilentlyContinue | Select-Object -First 1
if ($p) {
    Write-Host "Excel PID: $($p.Id)"
    $wins = [WinFinder]::GetWindowsForProcess($p.Id)
    foreach ($w in $wins) {
        Write-Host $w
    }
} else {
    Write-Host "No Excel process found."
}
