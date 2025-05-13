# MIT License
#
# Copyright (c) 2025 iappyx
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class KeyboardListener {
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
}

public class WindowHelper {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern int ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@ -Language CSharp

$global:hWnd = [WindowHelper]::GetForegroundWindow()

$keyMap = @{
    "Ctrl"  = 0x11
    "Shift" = 0x10
    "Alt"   = 0x12
    "Del"   = 0x2E
    "Up"    = 0x26
    "Down"  = 0x28
    "Left"  = 0x25
    "Right" = 0x27
}

$instrumentaKeysVersion = "0.12"

Write-Host "██╗███╗   ██╗███████╗████████╗██████╗ ██╗   ██╗███╗   ███╗███████╗███╗   ██╗████████╗ █████╗ "
Write-Host "██║████╗  ██║██╔════╝╚══██╔══╝██╔══██╗██║   ██║████╗ ████║██╔════╝████╗  ██║╚══██╔══╝██╔══██╗"
Write-Host "██║██╔██╗ ██║███████╗   ██║   ██████╔╝██║   ██║██╔████╔██║█████╗  ██╔██╗ ██║   ██║   ███████║"
Write-Host "██║██║╚██╗██║╚════██║   ██║   ██╔══██╗██║   ██║██║╚██╔╝██║██╔══╝  ██║╚██╗██║   ██║   ██╔══██║"
Write-Host "██║██║ ╚████║███████║   ██║   ██║  ██║╚██████╔╝██║ ╚═╝ ██║███████╗██║ ╚████║   ██║   ██║  ██║"
Write-Host "╚═╝╚═╝  ╚═══╝╚══════╝   ╚═╝   ╚═╝  ╚═╝ ╚═════╝ ╚═╝     ╚═╝╚══════╝╚═╝  ╚═══╝   ╚═╝   ╚═╝  ╚═╝"
Write-Host "██╗  ██╗███████╗██╗   ██╗███████╗                                                            "
Write-Host "██║ ██╔╝██╔════╝╚██╗ ██╔╝██╔════╝                                                            "
Write-Host "█████╔╝ █████╗   ╚████╔╝ ███████╗      Keyboard Shortcut Companion (v. $instrumentaKeysVersion)"
Write-Host "██╔═██╗ ██╔══╝    ╚██╔╝  ╚════██║                                                            "
Write-Host "██║  ██╗███████╗   ██║   ███████║                                                            "
Write-Host "╚═╝  ╚═╝╚══════╝   ╚═╝   ╚══════╝                                                            "
Write-Host ""

$shortcuts = @{}
$friendlyShortcuts = @{}
$shortcutList = @()
$csvPath = "shortcuts.csv"
Write-Host "Loading $csvPath..."


if (Test-Path $csvPath) {
    $csvData = Import-Csv -Path $csvPath -Header "Key", "Macro"

    foreach ($entry in $csvData) {
        if ($entry.Key -and $entry.Macro) {
            $keyCombo = $entry.Key -split '\+'
            $virtualKeys = @()
            
            $shortcutList += [PSCustomObject]@{
            Shortcut = $entry.Key
            Macro    = $entry.Macro
            }

            foreach ($key in $keyCombo) {
                if ($keyMap.ContainsKey($key)) {
                    $virtualKeys += $keyMap[$key]
                } else {
                    try {
                        $virtualKeys += [int][char]$key 
                    } catch {
                        Write-Host "Error: Invalid key '$key' in CSV."
                    }
                }
            }

            $macroName = $entry.Macro.Trim()
            if ($macroName -ne "") {

                $keyIdentifier = $virtualKeys -join " "
                $shortcuts[$keyIdentifier] = $macroName
                $friendlyShortcuts[$keyIdentifier] = $entry.Key
                # Write-Host "Enabled shortcut $($entry.Key) for macro $macroName"
            } else {
                Write-Host "Warning: Macro name is empty for shortcut $($entry.Key) ($keyIdentifier)"
            }
        } else {
            Write-Host "Error: Missing key or macro in CSV entry."
        }
    }
} else {
    Write-Host "Error: CSV file not found!"
    exit
}

Write-Host ""
$shortcutList | Format-Table -AutoSize

function GetActiveWindowHandle {
    $hWnd = [WindowHelper]::GetForegroundWindow()

    if ($hWnd -eq [IntPtr]::Zero) {
        return $null
    }

    return $hWnd
}

function IsPowerPointActive {
    $activeWindow = GetActiveWindowHandle
    if ($null -eq $activeWindow) { 
        return $false 
    }

    $activeProcessId = 0  
    [WindowHelper]::GetWindowThreadProcessId($activeWindow, [ref]$activeProcessId)

    $activeProcess = Get-Process -Id $activeProcessId -ErrorAction SilentlyContinue
    if ($activeProcess) {
        $exeName = $activeProcess.ProcessName
        $isActive = ($exeName -eq "POWERPNT")

        if ($isActive -eq "True") {
	return "active"
	} else {
	return "not-active"
	}
    } else {
        return "not-active"
    }
}

function ConnectToPowerpoint {
    try {
        return [System.Runtime.Interopservices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        return $null
    }
}

$ppt = ConnectToPowerpoint

Write-Host ""
Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Listening for shortcuts, and hiding this window to the systray in three seconds" -NoNewline
for ($i = 1; $i -le 3; $i++) {
    Start-Sleep -Seconds 1
    Write-Host "." -NoNewline
}
Write-Host ""

[void] [WindowHelper]::ShowWindow($global:hWnd, 0)

Add-Type -AssemblyName System.Windows.Forms

$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$trayIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
$trayIcon.Text = "Instrumenta Keys is running (click to view/hide)"
$trayIcon.Visible = $true

$trayIcon.Add_MouseClick({
    $windowState = [WindowHelper]::ShowWindow($global:hWnd, 0)
    
    if ($windowState -eq 0) {
        [WindowHelper]::ShowWindow($global:hWnd, 9)
    } else {
        [WindowHelper]::ShowWindow($global:hWnd, 0)
    }
})

$waiting = $true
$inPresentationMode = $false

Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Waiting for PowerPoint instance."

while ($true) {
    $ppt = ConnectToPowerpoint

    if ($null -eq $ppt) {
        if (-not $waiting) { 
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Waiting for PowerPoint instance."
            $waiting = $true
        }
        Start-Sleep -Milliseconds 5000 
        continue
    }

    if ($waiting) { 
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - PowerPoint instance found. Accepting shortcuts when PowerPoint window is active."
        $waiting = $false
    }

    $testIfActive = IsPowerPointActive
    if ($testIfActive -eq "not-active") { 
        Start-Sleep -Milliseconds 1000
        continue
    }

    if ($ppt.SlideShowWindows.Count -gt 0) {
        if (-not $inPresentationMode) {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Presentation Mode detected. Shortcuts are temporarily disabled."
            $inPresentationMode = $true  # Mark that we're in presentation mode
        }
        Start-Sleep -Milliseconds 1000
        continue
    } else {
        if ($inPresentationMode) {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Exited Presentation Mode. Shortcuts are active again."
            $inPresentationMode = $false  # Reset flag when exiting presentation mode
        }
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    foreach ($virtualKeyCombo in $shortcuts.Keys) {
        $pressed = $true

        foreach ($virtualKey in $virtualKeyCombo -split ' ') {  
            $keyInt = [int]$virtualKey  

            if ([KeyboardListener]::GetAsyncKeyState($keyInt) -eq 0) {
                $pressed = $false
                break
            }
        }

        if ($pressed) { 
            if ($shortcuts.ContainsKey($virtualKeyCombo)) {
                $macroName = $shortcuts[$virtualKeyCombo]
                Write-Host "$timestamp - Detected shortcut $($friendlyShortcuts[$virtualKeyCombo]), executing macro $macroName"
                try {
                    $ppt.Run($macroName)
                } catch {
                    Write-Host "$timestamp - Error: Failed execution of $macroName with message $_ "
                }

                Start-Sleep -Milliseconds 300
            } else {
                Write-Host "$timestamp - Error: Macro not found for shortcut $($friendlyShortcuts[$virtualKeyCombo])."
            }
        }
    }

    Start-Sleep -Milliseconds 100
}

Write-Output "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Execution complete. Press Enter to exit..."
Read-Host