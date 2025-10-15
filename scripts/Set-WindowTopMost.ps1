[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Part or all of the window title you want to match.")]
    [string]$WindowTitle,

    [Parameter(HelpMessage = "Remove the always-on-top flag instead of enabling it.")]
    [switch]$Disable,

    [Parameter(HelpMessage = "When multiple windows match, pick the Nth one (1-based index).")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$Index = 1
)

begin {
    $signature = @'
using System;
using System.Runtime.InteropServices;

public static class WindowTopMost
{
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int X,
        int Y,
        int cx,
        int cy,
        uint uFlags);
}
'@;

    if (-not ([System.Management.Automation.PSTypeName]"WindowTopMost".Type)) {
        Add-Type -TypeDefinition $signature -ErrorAction Stop
    }

    $HWND_TOPMOST = [IntPtr]::new(-1)
    $HWND_NOTOPMOST = [IntPtr]::new(-2)
    $SWP_NOSIZE = 0x0001
    $SWP_NOMOVE = 0x0002
    $SWP_SHOWWINDOW = 0x0040
}

process {
    $matchingWindows = Get-Process | Where-Object {
        $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "*$WindowTitle*"
    } | Sort-Object MainWindowTitle

    if (-not $matchingWindows) {
        throw "No windows found whose title contains '$WindowTitle'."
    }

    if ($Index -gt $matchingWindows.Count) {
        throw "Only $($matchingWindows.Count) window(s) matched '$WindowTitle'. Index $Index is out of range."
    }

    $target = $matchingWindows[$Index - 1]
    $handle = $target.MainWindowHandle

    if ($handle -eq 0) {
        throw "Selected process '$($target.ProcessName)' does not have an interactive window."
    }

    $insertAfter = if ($Disable.IsPresent) { $HWND_NOTOPMOST } else { $HWND_TOPMOST }
    $flags = $SWP_NOMOVE -bor $SWP_NOSIZE -bor $SWP_SHOWWINDOW

    $result = [WindowTopMost]::SetWindowPos($handle, $insertAfter, 0, 0, 0, 0, $flags)

    if (-not $result) {
        $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "SetWindowPos failed with error code $errorCode."
    }

    $action = if ($Disable.IsPresent) { "removed" } else { "enabled" }
    Write-Output "Always-on-top $action for window '$($target.MainWindowTitle)' (process $($target.ProcessName))."
}
