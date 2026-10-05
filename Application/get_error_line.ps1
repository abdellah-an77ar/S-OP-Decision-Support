Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class Clicker {
    [DllImport("user32.dll")]
    public static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);
    public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumThreadDelegate lpfn, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    public static void DismissDialog(int pid) {
        var proc = System.Diagnostics.Process.GetProcessById(pid);
        foreach (System.Diagnostics.ProcessThread t in proc.Threads) {
            EnumThreadWindows(t.Id, (hWnd, lParam) => {
                StringBuilder cb = new StringBuilder(256);
                GetClassName(hWnd, cb, 256);
                if (cb.ToString().Contains("#32770")) {
                    StringBuilder sbTitle = new StringBuilder(256);
                    GetWindowText(hWnd, sbTitle, 256);
                    Console.WriteLine("FOUND DIALOG: " + sbTitle.ToString());
                    EnumChildWindows(hWnd, (hChild, l) => {
                        StringBuilder sb = new StringBuilder(256);
                        GetWindowText(hChild, sb, 256);
                        Console.WriteLine("   child text: " + sb.ToString());
                        if (sb.ToString() == "OK") {
                            PostMessage(hChild, 0x00F5, IntPtr.Zero, IntPtr.Zero);
                        }
                        return true;
                    }, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
        }
    }
}
"@

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

$excelPid = (Get-Process -Name excel | Sort-Object StartTime -Descending | Select-Object -First 1).Id

try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "Workbook opened. Excel PID: $excelPid"

    # Trigger CalculerKPI in a job
    $job = Start-Job -ScriptBlock {
        param($tPid)
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $excel.Run("AtlasFood_SOP.xlsm!CalculerKPI")
    } -ArgumentList $excelPid

    # Wait 2 seconds for dialog to appear
    Start-Sleep -Seconds 2

    # Dismiss dialog by clicking OK
    Write-Host "Dismissing dialog..."
    [Clicker]::DismissDialog($excelPid)
    Start-Sleep -Milliseconds 500

    # Now inspect VBE ActiveCodePane!
    try {
        $pane = $xl.VBE.ActiveCodePane
        if ($pane -ne $null) {
            $modName = $pane.CodeModule.Name
            $startLine = 0; $startCol = 0; $endLine = 0; $endCol = 0
            $pane.GetSelection([ref]$startLine, [ref]$startCol, [ref]$endLine, [ref]$endCol)
            $errLine = $pane.CodeModule.Lines($startLine, 1)
            Write-Host ">>> ERROR IN MODULE: $modName at LINE $startLine <<<" -ForegroundColor Red
            Write-Host ">>> LINE CONTENT: $errLine <<<" -ForegroundColor Red
        } else {
            Write-Host "No ActiveCodePane found."
        }
    } catch {
        Write-Host "VBE inspection error: $_"
    }

    Stop-Job $job -ErrorAction SilentlyContinue

} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}
