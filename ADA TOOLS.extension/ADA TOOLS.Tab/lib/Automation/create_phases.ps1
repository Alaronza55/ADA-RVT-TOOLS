<#
ADA Tools - drives Revit's Phasing dialog (Manage > Phases) to add phases.

Launched by Planning.panel/ImportPlanning before it posts the Phases command.
Revit's API cannot create phases, so this helper waits for the Phasing dialog,
then for every name: clicks "Insert After" and types the name into the new
row. It NEVER clicks OK - the user checks the result and clicks OK / Cancel.

  -Mode create : add the phases listed in -NamesFile (one per line, UTF-8)
  -Mode probe  : only dump the dialog's UI Automation tree to the log

Everything is logged to -LogFile (the dialog tree is always dumped first).
#>
param(
    [int]$RevitPid,
    [string]$NamesFile,
    [string]$LogFile,
    [string]$Mode = 'create',
    [int]$TimeoutSec = 60,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms, System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class AdaNative {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
}
"@
[void][AdaNative]::SetProcessDPIAware()

$AE   = [System.Windows.Automation.AutomationElement]
$CT   = [System.Windows.Automation.ControlType]
$Tree = [System.Windows.Automation.TreeScope]
$True_ = [System.Windows.Automation.Condition]::TrueCondition

# Button captions per Revit UI language (first match wins)
$INSERT_AFTER = @('Insert After', 'Insérer après', 'Inserer apres', 'Danach einfügen',
                  'Nach einfügen', 'Inserisci dopo', 'Insertar después', 'Invoegen na')

function Log([string]$msg) {
    Add-Content -LiteralPath $LogFile -Value ('{0:HH:mm:ss.fff}  {1}' -f (Get-Date), $msg) -Encoding UTF8
}

function Popup([string]$msg, [int]$icon = 64) {
    # 4096 = system modal (stays on top of Revit)
    if ($Quiet) { Log "POPUP: $msg"; return }
    $shell = New-Object -ComObject WScript.Shell
    [void]$shell.Popup($msg, 0, 'ADA Tools - Import Planning', $icon + 4096)
}

function Describe($el) {
    try {
        $c = $el.Current
        $r = $c.BoundingRectangle
        $patterns = ($el.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName -replace 'Identifiers.Pattern', '' }) -join ','
        return ('{0} | name="{1}" | id="{2}" | class="{3}" | rect={4},{5},{6},{7} | {8}' -f
            $c.ControlType.ProgrammaticName, $c.Name, $c.AutomationId, $c.ClassName,
            [int]$r.X, [int]$r.Y, [int]$r.Width, [int]$r.Height, $patterns)
    } catch { return '<unreadable element>' }
}

function Dump($el, [int]$depth, [int]$maxDepth) {
    Log ((' ' * ($depth * 2)) + (Describe $el))
    if ($depth -ge $maxDepth) { return }
    try { $children = $el.FindAll($Tree::Children, $True_) } catch { return }
    foreach ($child in $children) { Dump $child ($depth + 1) $maxDepth }
}

function Find-Button($scope, [string[]]$names) {
    # matched by caption only: depending on the UI framework, buttons can be
    # reported as Button, Pane or Custom
    foreach ($b in $scope.FindAll($Tree::Descendants, $True_)) {
        $n = ($b.Current.Name -replace '&', '').Trim()
        foreach ($wanted in $names) { if ($n -ieq $wanted) { return $b } }
    }
    return $null
}

function Press($el) {
    $ip = $null
    if ($el.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$ip)) {
        $ip.Invoke()
        return 'invoke'
    }
    $r = $el.Current.BoundingRectangle
    [AdaNative]::Click([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
    return 'click'
}

function Find-PhasingDialog {
    $pidCond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $RevitPid)
    $windows = $AE::RootElement.FindAll($Tree::Children, $pidCond)
    foreach ($w in $windows) {
        $candidates = @($w)
        try {
            $candidates += @($w.FindAll($Tree::Children,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window))))
        } catch {}
        foreach ($c in $candidates) {
            $btn = Find-Button $c $INSERT_AFTER
            if ($btn) { return @{ Dialog = $c; InsertAfter = $btn } }
        }
    }
    return $null
}

function Find-Grid($dialog) {
    # biggest data-like control that is not the whole dialog
    $dr = $dialog.Current.BoundingRectangle
    $best = $null; $bestArea = 0
    $types = @($CT::DataGrid, $CT::Table, $CT::List, $CT::Custom, $CT::Pane)
    foreach ($el in $dialog.FindAll($Tree::Descendants, $True_)) {
        $c = $el.Current
        if ($types -notcontains $c.ControlType) { continue }
        $r = $c.BoundingRectangle
        $area = $r.Width * $r.Height
        if ($area -le 0 -or $area -ge ($dr.Width * $dr.Height * 0.8)) { continue }
        if ($area -gt $bestArea) { $best = $el; $bestArea = $area }
    }
    return $best
}

function Get-NameCells($grid) {
    # cells exposing an editable value (if Revit's grid exposes them at all)
    $result = @()
    foreach ($el in $grid.FindAll($Tree::Descendants, $True_)) {
        $vp = $null
        if ($el.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$vp)) {
            $result += ,@($el, $vp.Current.Value)
        }
    }
    return $result
}

function Find-SelectedRowY($grid) {
    # screen-capture the grid and find the highlighted (selected) row band
    $r = $grid.Current.BoundingRectangle
    $w = [int]$r.Width; $h = [int]$r.Height
    if ($w -lt 20 -or $h -lt 20) { return $null }
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen([int]$r.X, [int]$r.Y, 0, 0, $bmp.Size)
    $g.Dispose()

    $x0 = [int]($w * 0.15); $x1 = [int]($w * 0.85)
    $rowsHit = @()
    for ($y = 0; $y -lt $h; $y++) {
        $colored = 0; $samples = 0
        for ($x = $x0; $x -lt $x1; $x += 4) {
            $p = $bmp.GetPixel($x, $y); $samples++
            $isWhite = ($p.R -gt 235 -and $p.G -gt 235 -and $p.B -gt 235)
            $isGrey  = ([Math]::Abs($p.R - $p.G) -lt 12 -and [Math]::Abs($p.G - $p.B) -lt 12)
            if (-not $isWhite -and -not $isGrey) { $colored++ }
        }
        if ($samples -gt 0 -and ($colored / $samples) -gt 0.5) { $rowsHit += $y }
    }
    $bmp.Dispose()
    if ($rowsHit.Count -lt 6) { return $null }
    # keep the LAST contiguous band (newly inserted row is the selected one)
    $bands = @(); $start = $rowsHit[0]; $prev = $rowsHit[0]
    foreach ($y in $rowsHit[1..($rowsHit.Count - 1)]) {
        if ($y -ne $prev + 1) { $bands += ,@($start, $prev); $start = $y }
        $prev = $y
    }
    $bands += ,@($start, $prev)
    $band = $bands | Where-Object { ($_[1] - $_[0]) -ge 6 } | Select-Object -Last 1
    if (-not $band) { return $null }
    return [int]($r.Y + ($band[0] + $band[1]) / 2)
}

function Set-CellText([string]$text) {
    # paste into the cell being edited, then commit with Tab: it stays on the
    # same row (Enter would move to the next row - so the next "Insert After"
    # would land in the wrong place, and phase order can't be changed
    # afterwards in Revit - or could trigger the dialog's OK button)
    [System.Windows.Forms.Clipboard]::SetText($text)
    Start-Sleep -Milliseconds 80
    [System.Windows.Forms.SendKeys]::SendWait('^a')
    [System.Windows.Forms.SendKeys]::SendWait('^v')
    Start-Sleep -Milliseconds 120
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Start-Sleep -Milliseconds 150
}

# ---------------------------------------------------------------------------
try {
    Set-Content -LiteralPath $LogFile -Value "ADA Tools - Phasing dialog automation ($Mode)" -Encoding UTF8
    Log "Revit PID $RevitPid - waiting for the Phasing dialog (max $TimeoutSec s)"

    $found = $null
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while (-not $found -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        try { $found = Find-PhasingDialog } catch { }
    }
    if (-not $found) {
        Log 'Phasing dialog not found.'
        Popup "The Phasing dialog was not detected within $TimeoutSec s.`nNothing was changed.`n`nLog: $LogFile" 48
        exit 1
    }

    $dialog = $found.Dialog
    Log ('Dialog: ' + (Describe $dialog))
    Log '--- UI Automation tree ---'
    Dump $dialog 0 6
    Log '--- end of tree ---'

    if ($Mode -eq 'probe') {
        Popup "Phasing dialog structure saved to:`n$LogFile`n`nSend this file to improve the automation."
        exit 0
    }

    $names = @(Get-Content -LiteralPath $NamesFile -Encoding UTF8 | Where-Object { $_.Trim() -ne '' })
    Log ("{0} phase name(s) to create" -f $names.Count)

    $hwnd = [IntPtr]$dialog.Current.NativeWindowHandle
    $grid = Find-Grid $dialog
    if ($grid) { Log ('Grid: ' + (Describe $grid)) } else { Log 'Grid: not found' }

    Log ('Insert After: ' + (Describe $found.InsertAfter))
    $created = 0; $failed = @()

    foreach ($name in $names) {
        [void][AdaNative]::SetForegroundWindow($hwnd)
        $before = @()
        if ($grid) { $before = @(Get-NameCells $grid | ForEach-Object { $_[1] }) }

        $how = Press $found.InsertAfter
        Log ("Insert After pressed ({0})" -f $how)
        Start-Sleep -Milliseconds 400

        $done = $false
        # strategy 1: the grid exposes editable cells -> set the new row's name
        if ($grid) {
            foreach ($cell in (Get-NameCells $grid)) {
                if ($before -notcontains $cell[1] -and $cell[1] -ne '') {
                    try {
                        $vp = $cell[0].GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
                        $vp.SetValue($name)
                        $done = $true
                        Log ("[UIA] '{0}' -> '{1}'" -f $cell[1], $name)
                    } catch { Log ("[UIA] SetValue failed: " + $_.Exception.Message) }
                    break
                }
            }
        }
        # strategy 2: click the Name cell of the highlighted (new) row and paste
        if (-not $done -and $grid) {
            $y = Find-SelectedRowY $grid
            if ($y) {
                $r = $grid.Current.BoundingRectangle
                $x = [int]($r.X + $r.Width * 0.30)
                [void][AdaNative]::SetForegroundWindow($hwnd)
                [AdaNative]::Click($x, $y)
                Start-Sleep -Milliseconds 250
                [AdaNative]::Click($x, $y)
                Start-Sleep -Milliseconds 250
                Set-CellText $name
                Start-Sleep -Milliseconds 200
                $done = $true
                Log ("[CLICK] row at y={0}, x={1} -> '{2}'" -f $y, $x, $name)
            } else {
                Log "[CLICK] selected row not found for '$name'"
            }
        }

        if ($done) { $created++ } else { $failed += $name }
    }

    Log ("Done: {0} named, {1} failed" -f $created, $failed.Count)
    $msg = "{0} phase(s) inserted and named in the Phasing dialog.`n" -f $created
    if ($failed.Count) { $msg += "{0} could not be named (see log).`n" -f $failed.Count }
    $msg += "`nCHECK the list, then click OK to keep the phases (Cancel discards everything).`nThen run Import Planning again to fill 'Contractor'.`n`nLog: $LogFile"
    Popup $msg
}
catch {
    Log ('ERROR: ' + $_.Exception.Message)
    Popup ("The automation stopped with an error:`n" + $_.Exception.Message + "`n`nClick Cancel in the Phasing dialog to discard partial changes.`nLog: $LogFile") 16
    exit 1
}
