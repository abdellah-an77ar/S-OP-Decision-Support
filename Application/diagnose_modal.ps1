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
                if (cb.ToString().Contains("#32770")) {
                    string diag = GetDialogText(hWnd);
                    list.Add(string.Format("DIALOG Handle: {0}, Title: '{1}', Content:\n{2}", hWnd, sb, diag));
                } else if (sb.Length > 0 && !sb.ToString().Contains("Cicero") && !sb.ToString().Contains("IME")) {
                    list.Add(string.Format("Handle: {0}, Class: {1}, Text: '{2}'", hWnd, cb, sb));
                }
                return true;
            }, IntPtr.Zero);
        }
        return list;
    }
}
"@

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

$excelPid = (Get-Process -Name excel | Sort-Object StartTime -Descending | Select-Object -First 1).Id
Write-Host "Excel PID: $excelPid"

try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "Workbook opened."

    $job = Start-Job -ScriptBlock {
        param($targetPid)
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $excel.Run("AtlasFood_SOP.xlsm!LancerCapacityPlan")
    } -ArgumentList $excelPid

    $timeout = 8
    $elapsed = 0
    while ($job.State -eq 'Running' -and $elapsed -lt $timeout) {
        Start-Sleep -Seconds 1
        $elapsed++
        $wins = [WinFinder]::GetWindowsForProcess($excelPid)
        foreach ($w in $wins) {
            Write-Host "  $w" -ForegroundColor Magenta
        }
    }

    if ($job.State -eq 'Running') {
        Write-Host "Macro STILL RUNNING after $timeout seconds. Inspecting dialogs:" -ForegroundColor Red
        $wins = [WinFinder]::GetWindowsForProcess($excelPid)
        foreach ($w in $wins) { Write-Host "   $w" -ForegroundColor Yellow }
        Stop-Job $job
    } else {
        $res = Receive-Job $job
        Write-Host "Macro COMPLETED! Result: $res" -ForegroundColor Green
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}
