<#
#########################################
Author: Ali Umar | www.aliumar.dev
Revision date: 2024-12-24
Description: 
Remarks: This is for windows only. Maybe someone can port it to powershell-core if you really want to... or an
equivilant bash script or something lol.
#########################################
#>

$ScriptName = "i_hate_work.ps1"
$Version = "1.0"
$Author = "Ali Umar"

# Some embedded C# code calling FFI and user32.dll windows shared library

# BEGIN
Add-Type -TypeDefinition @" 
using System;
using System.Runtime.InteropServices;

public class FFI {
    public struct Point {
        public int X;
        public int Y;
    }

    // Set mouse position
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetCursorPos(int X, int Y);

    // Get Mouse position
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool GetCursorPos(out Point lpPoint);
    
    // Get window size
    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetSystemMetrics(int nIndex);

    // Import SetForegroundWindow function (bring the window to the foreground)
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    // Import GetWindowThreadProcessId function (get the process ID associated with a window)
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

function Focus-WindowByPid {
    param ([int]$tpid = -1)
    $processes = Get-Process
    $process = $processes | Where-Object { $_.Id -eq $tpid }

    if ($process) {
        $hWnd = (Get-Process -Id $tpid).MainWindowHandle
        if ($hWnd -ne [IntPtr]::Zero) {
            # windows API is stupid, minimize first & and then restore
            [FFI]::ShowWindow($hWnd, 6)
            [FFI]::ShowWindow($hWnd, 9)
            [FFI]::SetForegroundWindow($hWnd)
        } else {
            Write-Host "An open instance of teams was not found."
        }
    } else {
            Write-Host "An open instance of teams was not found."
    }
}


# EDIT ME BEGIN
$delay = 30
# EDIT ME END

# EDIT ME END
$cursorPos = New-Object FFI+Point
$mouseStartPos = New-Object FFI+Point
$windowWidth = [FFI]::GetSystemMetrics(0)
$windowHeight = [FFI]::GetSystemMetrics(1)
$offset = 100
$mouseStartPos.X = $windowWidth / 2 - $offset
$mouseStartPos.Y = $windowHeight / 2 - $offset
$teams = Get-Process | Where-Object { $_.ProcessName -like "*ms-teams*" } 

Write-Host "Starting script: $ScriptName (Version $Version)" 
Write-Host "Starting script: $ScriptName (Version $Version)" 

# Main loop
while($true) {
  [FFI]::GetCursorPos([ref]$cursorPos) | Out-Null;
  [FFI]::SetCursorPos($mouseStartPos.X, $mouseStartPos.Y) | Out-Null
  echo $teams | ForEach-Object { Focus-WindowByPid -tpid $_.Id }
  Start-Sleep -Seconds $delay
}

