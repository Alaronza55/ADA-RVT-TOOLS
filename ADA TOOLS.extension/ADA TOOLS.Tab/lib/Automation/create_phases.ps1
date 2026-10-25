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
    [switch]$Quiet,
    [switch]$IgnoreGridControl   # test switch: behave as if the grid were not exposed
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

# Caption of the "After" button of the dialog's "Insert" group, per Revit UI
# language (exact match, '&' accelerators ignored)
$INSERT_AFTER = @('After', 'Insert After', 'Après', 'Apres', 'Insérer après',
                  'Danach', 'Nach', 'Dopo', 'Después', 'Despues', 'Na', 'Po')

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
            if ($btn) {
                # the Phasing dialog is nested inside Revit's main window:
                # work in the window that actually owns the button
                return @{ Dialog = (Get-OwnerWindow $btn); InsertAfter = $btn }
            }
        }
    }
    return $null
}

function Get-OwnerWindow($el) {
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $p = $walker.GetParent($el)
    while ($null -ne $p) {
        if ($p.Current.ControlType -eq $CT::Window) { return $p }
        $p = $walker.GetParent($p)
    }
    return $el
}

function Find-Grid($dialog, $afterButton) {
    # biggest data-like control that does NOT contain the After button
    # (that excludes the dialog itself and the tab page around the grid)
    $ar = $afterButton.Current.BoundingRectangle
    $ax = $ar.X + $ar.Width / 2; $ay = $ar.Y + $ar.Height / 2

    # Revit's phase grid is a "GXWND" control (seen in the Revit 2024 dialog)
    $gx = $dialog.FindFirst($Tree::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'GXWND')))
    if ($gx) { return $gx }

    $best = $null; $bestArea = 0
    $types = @($CT::DataGrid, $CT::Table, $CT::List, $CT::Custom, $CT::Pane)
    foreach ($el in $dialog.FindAll($Tree::Descendants, $True_)) {
        $c = $el.Current
        if ($types -notcontains $c.ControlType) { continue }
        $r = $c.BoundingRectangle
        if ($r.Width -le 0 -or $r.Height -le 0) { continue }
        if ($ax -ge $r.X -and $ax -le $r.X + $r.Width -and
            $ay -ge $r.Y -and $ay -le $r.Y + $r.Height) { continue }
        if ($r.X + $r.Width -gt $ar.X) { continue }   # grid is left of the buttons
        $area = $r.Width * $r.Height
        if ($area -gt $bestArea) { $best = $el; $bestArea = $area }
    }
    return $best
}

function Get-FallbackGridRect($dialog, $afterButton) {
    # Revit layout: grid on the left, Insert Before/After on the right,
    # OK/Cancel at the bottom
    $dr = $dialog.Current.BoundingRectangle
    $ar = $afterButton.Current.BoundingRectangle
    $left = $dr.X + $dr.Width * 0.03
    $right = $ar.X - $dr.Width * 0.03
    $top = $ar.Y - $ar.Height * 1.6
    $bottom = $dr.Y + $dr.Height * 0.85
    return New-Object System.Windows.Rect($left, $top, ($right - $left), ($bottom - $top))
}

function Find-SelectedBand($r) {
    # screen-capture the grid area and return the FIRST highlighted (selected)
    # row band as @(topY, bottomY) in screen coordinates, or $null
    $w = [int]$r.Width; $h = [int]$r.Height
    if ($w -lt 20 -or $h -lt 20) { return $null }
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen([int]$r.X, [int]$r.Y, 0, 0, $bmp.Size)
    $g.Dispose()

    # a row is "selected" when most of it is painted (Revit: black, other
    # grids: blue) - normal rows are white with sparse text, the header and
    # grid lines are light grey / 1 px
    $x0 = [int]($w * 0.15); $x1 = [int]($w * 0.85)
    $rowsHit = @()
    for ($y = 0; $y -lt $h; $y++) {
        $painted = 0; $samples = 0
        for ($x = $x0; $x -lt $x1; $x += 4) {
            $p = $bmp.GetPixel($x, $y); $samples++
            if (-not ($p.R -gt 200 -and $p.G -gt 200 -and $p.B -gt 200)) { $painted++ }
        }
        if ($samples -gt 0 -and ($painted / $samples) -gt 0.6) { $rowsHit += $y }
    }
    $bmp.Dispose()
    if ($rowsHit.Count -lt 6) { return $null }
    $bands = @(); $start = $rowsHit[0]; $prev = $rowsHit[0]
    foreach ($y in $rowsHit[1..($rowsHit.Count - 1)]) {
        if ($y -ne $prev + 1) { $bands += ,@($start, $prev); $start = $y }
        $prev = $y
    }
    $bands += ,@($start, $prev)
    $band = $bands | Where-Object { ($_[1] - $_[0]) -ge 6 } | Select-Object -First 1
    if (-not $band) { return $null }
    return @([int]($r.Y + $band[0]), [int]($r.Y + $band[1]))
}

function Escape-SendKeys([string]$text) {
    $out = ''
    foreach ($ch in $text.ToCharArray()) {
        if ('+^%~(){}[]'.IndexOf($ch) -ge 0) { $out += '{' + $ch + '}' } else { $out += $ch }
    }
    return $out
}

function Type-CellText([string]$text) {
    # The Name cell is current (not in edit mode): typing the first character
    # starts editing and REPLACES the cell content; the rest is pasted into
    # the edit box. (No Ctrl+A: outside edit mode it would select every cell
    # of the grid and Ctrl+V would paste into all of them.)
    [System.Windows.Forms.SendKeys]::SendWait((Escape-SendKeys $text.Substring(0, 1)))
    Start-Sleep -Milliseconds 120
    if ($text.Length -gt 1) {
        [System.Windows.Forms.Clipboard]::SetText($text.Substring(1))
        Start-Sleep -Milliseconds 60
        [System.Windows.Forms.SendKeys]::SendWait('^v')
        Start-Sleep -Milliseconds 120
    }
}

# ---------------------------------------------------------------------------
try {
    Add-Content -LiteralPath $LogFile -Value "ADA Tools - Phasing dialog automation ($Mode) - helper v4 (anchor row, reverse insert)" -Encoding UTF8
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
    $grid = if ($IgnoreGridControl) { $null } else { Find-Grid $dialog $found.InsertAfter }
    if ($grid -and $grid.Current.BoundingRectangle.Height -gt 60) {
        $gridRect = $grid.Current.BoundingRectangle
        Log ('Grid: ' + (Describe $grid))
    } else {
        $grid = $null
        $gridRect = Get-FallbackGridRect $dialog $found.InsertAfter
        Log ('Grid: not exposed - using the area left of the buttons: {0},{1},{2},{3}' -f
             [int]$gridRect.X, [int]$gridRect.Y, [int]$gridRect.Width, [int]$gridRect.Height)
    }

    Log ('Insert After: ' + (Describe $found.InsertAfter))

    # The selected row when the dialog opens (normally "Existing") is the
    # ANCHOR. Every phase is inserted right after it, so the new row is always
    # the row just below the anchor - a known, visible position, whatever the
    # grid's selection / scrolling does afterwards. Names are therefore
    # inserted in REVERSE order to end up in planning order.
    [void][AdaNative]::SetForegroundWindow($hwnd)
    Start-Sleep -Milliseconds 300
    $band = Find-SelectedBand $gridRect
    if (-not $band) {
        Log 'No selected row found in the grid - nothing was done.'
        Popup "Could not find the selected phase in the Phasing dialog.`nNothing was changed - select 'Existing' in the list and run Import Planning again.`n`nLog: $LogFile" 48
        exit 1
    }
    $anchorY = [int](($band[0] + $band[1]) / 2)
    $rowH = ($band[1] - $band[0]) + 2
    $newRowY = $anchorY + $rowH
    $headerX = [int]($gridRect.X + $gridRect.Width * 0.025)   # row-number column
    $nameX = [int]($gridRect.X + $gridRect.Width * 0.17)      # Name column (~5-30 %)
    Log ("Anchor row: y={0} (band {1}-{2}), row height {3}, new row y={4}, header x={5}, name x={6}" -f
         $anchorY, $band[0], $band[1], $rowH, $newRowY, $headerX, $nameX)

    $created = 0
    $ordered = @($names); [array]::Reverse($ordered)
    foreach ($name in $ordered) {
        [void][AdaNative]::SetForegroundWindow($hwnd)
        # select the anchor row (also commits the previous name being edited)
        [AdaNative]::Click($headerX, $anchorY)
        Start-Sleep -Milliseconds 200

        $how = Press $found.InsertAfter
        Start-Sleep -Milliseconds 400

        # make the new row's Name cell current, then type the name
        [void][AdaNative]::SetForegroundWindow($hwnd)
        [AdaNative]::Click($nameX, $newRowY)
        Start-Sleep -Milliseconds 250
        Type-CellText $name
        $created++
        Log ("[{0}] inserted ({1}) and typed '{2}'" -f $created, $how, $name)
    }
    # commit the last name
    [AdaNative]::Click($headerX, $anchorY)
    Start-Sleep -Milliseconds 200

    Log ("Done: {0} phase(s) inserted and named" -f $created)
    $msg = "{0} phase(s) inserted and named in the Phasing dialog.`n" -f $created
    $msg += "`nCHECK the list, then click OK to keep the phases (Cancel discards everything).`nThen run Import Planning again to fill 'Contractor'.`n`nLog: $LogFile"
    Popup $msg
}
catch {
    Log ('ERROR: ' + $_.Exception.Message)
    Popup ("The automation stopped with an error:`n" + $_.Exception.Message + "`n`nClick Cancel in the Phasing dialog to discard partial changes.`nLog: $LogFile") 16
    exit 1
}
