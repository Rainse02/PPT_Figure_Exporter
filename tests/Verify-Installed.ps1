# This check loads the installed add-in but does not click controls.
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Close PowerPoint before the isolated installed-add-in check.'
}
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $app.Visible = -1
    $presentation = $app.Presentations.Add(-1)
    $addin = $app.COMAddIns.Item('PPTFigureExporter.Connect')
    if (-not $addin.Connect) {
        throw 'PowerPoint did not load PPTFigureExporter.Connect.'
    }
    $window = [System.Windows.Automation.AutomationElement]::FromHandle((Get-Process POWERPNT).MainWindowHandle)
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, 'Figure Export')
    $tab = $null
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        $tab = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
        if ($null -ne $tab) { break }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($null -eq $tab) {
        throw 'COM connection succeeded, but the Figure Export ribbon tab was not found.'
    }
    Write-Output ('COM connected: ' + $addin.Connect)
    Write-Output ('Ribbon tab: ' + $tab.Current.Name)
    Write-Output 'PASS: installed add-in loads and creates its Ribbon tab in PowerPoint.'
}
finally {
    if ($null -ne $presentation) {
        # This is the empty document created by this invocation.
        $presentation.Saved = -1
        $presentation.Close()
    }
    if ($null -ne $app) {
        $app.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
}
