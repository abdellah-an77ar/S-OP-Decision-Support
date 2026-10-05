Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public class WinChecker {
    [DllImport("user32.dll")]
    public static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);
    public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumThreadDelegate lpfn, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    public static string CheckAndDismissDialog(int pid) {
        string found = null;
        var proc = System.Diagnostics.Process.GetProcessById(pid);
        foreach (System.Diagnostics.ProcessThread t in proc.Threads) {
            EnumThreadWindows(t.Id, (hWnd, lParam) => {
                StringBuilder cb = new StringBuilder(256);
                GetClassName(hWnd, cb, 256);
                if (cb.ToString().Contains("#32770")) {
                    StringBuilder allText = new StringBuilder();
                    IntPtr okBtn = IntPtr.Zero;
                    EnumChildWindows(hWnd, (hChild, l) => {
                        StringBuilder sb = new StringBuilder(256);
                        GetWindowText(hChild, sb, 256);
                        if (sb.Length > 0) allText.Append(sb.ToString() + " ");
                        if (sb.ToString() == "OK") okBtn = hChild;
                        return true;
                    }, IntPtr.Zero);
                    found = allText.ToString();
                    // Click OK (BM_CLICK = 0x00F5)
                    if (okBtn != IntPtr.Zero) PostMessage(okBtn, 0x00F5, IntPtr.Zero, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
        }
        return found;
    }
}
"@

$WorkspaceRoot = "c:\Users\RZR\Documents\EMINES\CI2A\app vba"
$basDir = Join-Path $WorkspaceRoot "VBA_Modules"
$modNames = @("mod01_CONFIG", "mod02_IMPORT", "mod03_FORECAST", "mod04_PLANS", "mod05_SCENARIOS", "mod06_KPI", "mod07_EXPORT")

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

$excelPid = (Get-Process -Name excel | Sort-Object StartTime -Descending | Select-Object -First 1).Id

try {
    $wb = $xl.Workbooks.Add()
    $vbaProj = $wb.VBProject

    foreach ($m in $modNames) {
        $basFile = Join-Path $basDir "$m.bas"
        Write-Host "Importing $m..." -NoNewline
        $vbc = $vbaProj.VBComponents.Import($basFile)
        Write-Host " OK ($($vbc.CodeModule.CountOfLines) lines)"
    }

    Write-Host "`nAttempting Compile..."
    $compileCmd = $xl.VBE.CommandBars.FindControl([System.Reflection.Missing]::Value, 578)
    if ($compileCmd -ne $null -and $compileCmd.Enabled) {
        # Run compile in a job so we can monitor dialogs
        $job = Start-Job -ScriptBlock {
            param($targetPid)
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            $cmd = $excel.VBE.CommandBars.FindControl([System.Reflection.Missing]::Value, 578)
            $cmd.Execute()
        } -ArgumentList $excelPid

        $diag = $null
        for ($i = 0; $i -lt 10; $i++) {
            Start-Sleep -Milliseconds 500
            $diag = [WinChecker]::CheckAndDismissDialog($excelPid)
            if ($diag) { break }
            if ($job.State -ne 'Running') { break }
        }

        if ($diag) {
            Write-Host "  COMPILE FAILED WITH DIALOG: $diag" -ForegroundColor Red
        } else {
            Write-Host "  COMPILE PASSED!" -ForegroundColor Green
        }
        Stop-Job $job -ErrorAction SilentlyContinue
    }

} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}
