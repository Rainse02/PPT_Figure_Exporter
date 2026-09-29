#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$Json,
    [switch]$ForBuild,
    [switch]$ForInstall,
    [switch]$SkipNative,
    [switch]$RequireInstaller,
    [switch]$Strict
)

$ErrorActionPreference = 'SilentlyContinue'

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$checks = New-Object System.Collections.Generic.List[object]

function Add-EnvironmentCheck {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Pass', 'Warning', 'Fail', 'Info')]
        [string]$Status,
        [string]$Details,
        [string]$Path,
        [string]$Version,
        [bool]$Required = $false
    )

    [void]$checks.Add([pscustomobject][ordered]@{
            Name     = $Name
            Status   = $Status
            Required = $Required
            Version  = $Version
            Path     = $Path
            Details  = $Details
        })
}

function Get-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $null
}

function Get-CommandPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $command = Get-Command $Name -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -ne $command) {
        return $command.Source
    }

    return $null
}

function Get-FileVersionText {
    param(
        [string]$Path
    )

    if (-not $Path) {
        return $null
    }

    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return $null
    }

    return $item.VersionInfo.FileVersion
}

function ConvertTo-Version {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $match = [regex]::Match($Text, '(\d+)\.(\d+)\.(\d+)\.(\d+)')
    if (-not $match.Success) {
        return $null
    }

    try {
        $versionString = '{0}.{1}.{2}.{3}' -f `
            $match.Groups[1].Value, `
            $match.Groups[2].Value, `
            $match.Groups[3].Value, `
            $match.Groups[4].Value
        return [System.Version]::Parse($versionString)
    } catch {
        return $null
    }
}

function Get-PowerPointInfo {
    $configPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $config = Get-ItemProperty -LiteralPath $configPath -ErrorAction SilentlyContinue
    $platform = $null
    $productIds = $null
    $installationPath = $null

    if ($null -ne $config) {
        $platform = [string]$config.Platform
        $productIds = [string]$config.ProductReleaseIds
        $installationPath = [string]$config.InstallationPath
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    if ($installationPath) {
        $candidates.Add((Join-Path $installationPath 'root\Office16\POWERPNT.EXE'))
        $candidates.Add((Join-Path $installationPath 'Office16\POWERPNT.EXE'))
    }

    $candidates.Add('C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE')
    $candidates.Add('C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE')
    $candidates.Add('C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE')
    $candidates.Add('C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE')

    $path = Get-ExistingPath -Candidates $candidates.ToArray()
    $versionText = Get-FileVersionText -Path $path

    return [pscustomobject][ordered]@{
        Path             = $path
        VersionText      = $versionText
        Version          = ConvertTo-Version -Text $versionText
        Platform         = $platform
        ProductReleaseIds = $productIds
        InstallationPath = $installationPath
    }
}

function Get-MsBuildPath {
    $candidates = @(
        'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\amd64\MSBuild.exe',
        'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe',
        'C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe',
        (Get-CommandPath -Name 'msbuild.exe')
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $null
}

function Get-IsccPath {
    $candidates = @(
        (Join-Path $ProjectRoot '.build\tools\innosetup\ISCC.exe'),
        (Join-Path $ProjectRoot '.build\tools\innosetup673\ISCC.exe'),
        (Join-Path $ProjectRoot '.build\tools\InnoSetup\ISCC.exe'),
        (Join-Path $ProjectRoot '.build\tools\iscc\ISCC.exe'),
        (Get-CommandPath -Name 'ISCC.exe'),
        'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
        'C:\Program Files\Inno Setup 6\ISCC.exe'
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $null
}

function Get-WindowsBuildNumber {
    $key = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    if ($null -eq $key) {
        return $null
    }

    try {
        return [int]$key.CurrentBuild
    } catch {
        return $null
    }
}

$osKey = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$osName = [string]$osKey.ProductName
$osBuild = Get-WindowsBuildNumber
$osIs64Bit = [Environment]::Is64BitOperatingSystem

if ($osIs64Bit) {
    Add-EnvironmentCheck -Name '64-bit Windows' -Status Pass -Details "$osName build $osBuild" -Required $true
} else {
    Add-EnvironmentCheck -Name '64-bit Windows' -Status Fail -Details "$osName build $osBuild" -Required $true
}

$powerPoint = Get-PowerPointInfo
$officeIsX64 = $powerPoint.Platform -eq 'x64'
if ($powerPoint.Path -and $officeIsX64) {
    Add-EnvironmentCheck `
        -Name 'PowerPoint x64' `
        -Status Pass `
        -Details "ProductReleaseIds=$($powerPoint.ProductReleaseIds)" `
        -Path $powerPoint.Path `
        -Version $powerPoint.VersionText `
        -Required $true
} elseif ($powerPoint.Path) {
    Add-EnvironmentCheck `
        -Name 'PowerPoint x64' `
        -Status Fail `
        -Details "PowerPoint was found but Office Click-to-Run Platform is '$($powerPoint.Platform)'" `
        -Path $powerPoint.Path `
        -Version $powerPoint.VersionText `
        -Required $true
} else {
    Add-EnvironmentCheck `
        -Name 'PowerPoint x64' `
        -Status Fail `
        -Details 'A 64-bit PowerPoint executable was not found.' `
        -Required $true
}

$minimumSvgBuild = [System.Version]::Parse('16.0.16130.2302')
if ($null -ne $powerPoint.Version) {
    if ($powerPoint.Version -ge $minimumSvgBuild) {
        Add-EnvironmentCheck `
            -Name 'PowerPoint SVG compatibility' `
            -Status Info `
            -Details 'PowerPoint meets the historical SVG build 16.0.16130.2302. Runtime export should still verify that Shape.Export returns SVG on the target build.' `
            -Version $powerPoint.VersionText `
            -Required $false
    } else {
        Add-EnvironmentCheck `
            -Name 'PowerPoint SVG compatibility' `
            -Status Info `
            -Details 'This PowerPoint build is below the historical SVG build 16.0.16130.2302. Installation is allowed, but runtime export must verify SVG because older builds may return another format.' `
            -Version $powerPoint.VersionText `
            -Required $false
    }
} else {
    Add-EnvironmentCheck `
        -Name 'PowerPoint SVG compatibility' `
        -Status Warning `
        -Details 'PowerPoint version could not be read.' `
        -Required $false
}

$officeIs2021 = $powerPoint.ProductReleaseIds -match '(?i)2021|LTSC'
if ($officeIs2021) {
    Add-EnvironmentCheck `
        -Name 'Office release compatibility' `
        -Status Info `
        -Details 'Office LTSC 2021 x64 is allowed by the installation gate; runtime export must verify SVG on the target build.' `
        -Required $false
} else {
    Add-EnvironmentCheck `
        -Name 'Office release compatibility' `
        -Status Info `
        -Details "Detected ProductReleaseIds=$($powerPoint.ProductReleaseIds)." `
        -Required $false
}

$comClsidKey = 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID'
$comClsid = Get-ItemProperty -LiteralPath $comClsidKey -ErrorAction SilentlyContinue
if ($null -ne $comClsid -and $comClsid.'(default)') {
    Add-EnvironmentCheck `
        -Name 'PowerPoint COM registration' `
        -Status Pass `
        -Details "PowerPoint.Application CLSID $($comClsid.'(default)') is registered." `
        -Required $true
} else {
    Add-EnvironmentCheck `
        -Name 'PowerPoint COM registration' `
        -Status Fail `
        -Details 'PowerPoint.Application COM registration was not found.' `
        -Required $true
}

if ($ForBuild) {
    $extensibilityInteropPath = Get-ExistingPath -Candidates @(
        "$env:SystemRoot\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll"
    )
    $officeInteropPath = Get-ExistingPath -Candidates @(
        "$env:SystemRoot\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL"
    )

    if ($extensibilityInteropPath -and $officeInteropPath) {
        Add-EnvironmentCheck `
            -Name 'Office interop assemblies' `
            -Status Pass `
            -Details 'Official Extensibility and Office interop assemblies are available for compilation. EmbedInteropTypes keeps them out of the target installation.' `
            -Path "$extensibilityInteropPath; $officeInteropPath" `
            -Required $true
    } else {
        Add-EnvironmentCheck `
            -Name 'Office interop assemblies' `
            -Status Fail `
            -Details 'The official Extensibility and Office interop assemblies required for compilation were not found.' `
            -Path "$extensibilityInteropPath; $officeInteropPath" `
            -Required $true
    }
}

$processes = @(Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue)
if ($processes.Count -eq 0) {
    Add-EnvironmentCheck `
        -Name 'PowerPoint process state' `
        -Status Pass `
        -Details 'POWERPNT.EXE is not running.' `
        -Required $ForInstall
} elseif ($ForInstall) {
    Add-EnvironmentCheck `
        -Name 'PowerPoint process state' `
        -Status Fail `
        -Details 'POWERPNT.EXE is running. Close PowerPoint before installing, upgrading, or uninstalling.' `
        -Required $true
} else {
    Add-EnvironmentCheck `
        -Name 'PowerPoint process state' `
        -Status Warning `
        -Details 'POWERPNT.EXE is running. Installation checks will block while it remains open.' `
        -Required $false
}

$netKey = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
$netRelease = 0
if ($null -ne $netKey) {
    $netRelease = [int]$netKey.Release
}

if ($netRelease -ge 528040) {
    Add-EnvironmentCheck `
        -Name '.NET Framework 4.8' `
        -Status Pass `
        -Details "Release key $netRelease meets 4.8." `
        -Path ([string]$netKey.InstallPath) `
        -Version ([string]$netKey.Version) `
        -Required $ForBuild
} elseif ($ForInstall) {
    Add-EnvironmentCheck `
        -Name '.NET Framework 4.8' `
        -Status Fail `
        -Details "Release key $netRelease is below 528040. Install .NET Framework 4.8 separately from https://dotnet.microsoft.com/download/dotnet-framework/net48 before installing." `
        -Required $true
} else {
    Add-EnvironmentCheck `
        -Name '.NET Framework 4.8' `
        -Status Fail `
        -Details "Release key $netRelease is below 528040." `
        -Required $ForBuild
}

$msbuildPath = Get-MsBuildPath
if ($msbuildPath) {
    Add-EnvironmentCheck `
        -Name 'MSBuild' `
        -Status Pass `
        -Details 'MSBuild is available for the .NET Framework frontend.' `
        -Path $msbuildPath `
        -Version (Get-FileVersionText -Path $msbuildPath) `
        -Required $ForBuild
} else {
    Add-EnvironmentCheck `
        -Name 'MSBuild' `
        -Status Fail `
        -Details 'MSBuild.exe was not found.' `
        -Required $ForBuild
}

$vcToolPath = Get-ExistingPath -Candidates @(
    'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Tools\MSVC\14.44.35207\bin\Hostx64\x64\cl.exe',
    'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Tools\MSVC\14.44.35207\bin\Hostx64\x64\link.exe'
)
if ($vcToolPath) {
    Add-EnvironmentCheck `
        -Name 'MSVC linker' `
        -Status Pass `
        -Details 'The Visual C++ x64 toolchain is installed.' `
        -Path $vcToolPath `
        -Version (Get-FileVersionText -Path $vcToolPath) `
        -Required $ForBuild
} else {
    Add-EnvironmentCheck `
        -Name 'MSVC linker' `
        -Status Warning `
        -Details 'The fixed VS Build Tools MSVC path was not found. Rust MSVC builds may still resolve a toolchain through rustup or another VS installation.' `
        -Required $false
}

$sdkRoot = 'C:\Program Files (x86)\Windows Kits\10'
if (Test-Path -LiteralPath (Join-Path $sdkRoot 'bin')) {
    Add-EnvironmentCheck `
        -Name 'Windows SDK' `
        -Status Pass `
        -Details 'Windows Kits 10 bin/include/lib directories are present.' `
        -Path $sdkRoot `
        -Required $ForBuild
} else {
    Add-EnvironmentCheck `
        -Name 'Windows SDK' `
        -Status Warning `
        -Details 'Windows Kits 10 was not found at the standard path.' `
        -Required $ForBuild
}

$cargoPath = Get-ExistingPath -Candidates @(
    (Get-CommandPath -Name 'cargo.exe'),
    'C:\Users\admin\.cargo\bin\cargo.exe'
)
$rustcPath = Get-ExistingPath -Candidates @(
    (Get-CommandPath -Name 'rustc.exe'),
    'C:\Users\admin\.cargo\bin\rustc.exe'
)
if ($SkipNative) {
    $rustStatus = 'Info'
    $rustDetails = 'Native build was explicitly skipped.'
} elseif ($cargoPath -and $rustcPath) {
    $rustStatus = 'Pass'
    $rustDetails = 'rustc and cargo are available for the MSVC native build.'
} else {
    $rustStatus = 'Fail'
    $rustDetails = 'rustc and cargo were not both found.'
}
Add-EnvironmentCheck `
    -Name 'Rust MSVC toolchain' `
    -Status $rustStatus `
    -Details $rustDetails `
    -Path $cargoPath `
    -Required ($ForBuild -and (-not $SkipNative))

$installerPath = Get-IsccPath
if ($installerPath) {
    Add-EnvironmentCheck `
        -Name 'Inno Setup compiler' `
        -Status Pass `
        -Details 'ISCC.exe is available.' `
        -Path $installerPath `
        -Version (Get-FileVersionText -Path $installerPath) `
        -Required $RequireInstaller
} elseif ($RequireInstaller) {
    Add-EnvironmentCheck `
        -Name 'Inno Setup compiler' `
        -Status Fail `
        -Details 'ISCC.exe was not found. Use build.ps1 -BootstrapInno or install Inno Setup locally.' `
        -Required $true
} else {
    Add-EnvironmentCheck `
        -Name 'Inno Setup compiler' `
        -Status Warning `
        -Details 'ISCC.exe was not found. Packaging requires a local Inno Setup compiler.' `
        -Required $false
}

if ($ForBuild) {
    $frontendProject = Join-Path $ProjectRoot 'frontend\PPTFigureExporter.csproj'
    if (Test-Path -LiteralPath $frontendProject -PathType Leaf) {
        Add-EnvironmentCheck `
            -Name 'Frontend project' `
            -Status Pass `
            -Details 'The frontend project file is present.' `
            -Path $frontendProject `
            -Required $true
    } else {
        Add-EnvironmentCheck `
            -Name 'Frontend project' `
            -Status Fail `
            -Details 'frontend\PPTFigureExporter.csproj was not found.' `
            -Path $frontendProject `
            -Required $true
    }

    $cargoManifest = Join-Path $ProjectRoot 'native\Cargo.toml'
    if (Test-Path -LiteralPath $cargoManifest -PathType Leaf) {
        Add-EnvironmentCheck `
            -Name 'Native project' `
            -Status Pass `
            -Details 'The Rust manifest is present.' `
            -Path $cargoManifest `
            -Required (-not $SkipNative)
    } elseif (-not $SkipNative) {
        Add-EnvironmentCheck `
            -Name 'Native project' `
            -Status Fail `
            -Details 'native\Cargo.toml was not found.' `
            -Path $cargoManifest `
            -Required $true
    }

    $cargoLock = Join-Path $ProjectRoot 'native\Cargo.lock'
    if ($SkipNative) {
        Add-EnvironmentCheck `
            -Name 'Cargo.lock' `
            -Status Info `
            -Details 'Native build was skipped; Cargo.lock was not required.' `
            -Path $cargoLock `
            -Required $false
    } elseif (Test-Path -LiteralPath $cargoLock -PathType Leaf) {
        Add-EnvironmentCheck `
            -Name 'Cargo.lock' `
            -Status Pass `
            -Details 'Locked native builds are enabled.' `
            -Path $cargoLock `
            -Required $true
    } else {
        Add-EnvironmentCheck `
            -Name 'Cargo.lock' `
            -Status Fail `
            -Details 'native\Cargo.lock is required for cargo build --locked.' `
            -Path $cargoLock `
            -Required $true
    }
}

$passCount = @($checks | Where-Object { $_.Status -eq 'Pass' }).Count
$warningCount = @($checks | Where-Object { $_.Status -eq 'Warning' }).Count
$failCount = @($checks | Where-Object { $_.Status -eq 'Fail' }).Count
$requiredFailures = @(
    $checks | Where-Object { $_.Required -and $_.Status -eq 'Fail' }
).Count

$summaryObject = [ordered]@{}
[void]$summaryObject.Add('GeneratedAt', (Get-Date).ToString('o'))
[void]$summaryObject.Add('ProjectRoot', $ProjectRoot)
[void]$summaryObject.Add('ForBuild', [bool]$ForBuild)
[void]$summaryObject.Add('ForInstall', [bool]$ForInstall)
[void]$summaryObject.Add('Checks', @($checks.ToArray()))
[void]$summaryObject.Add('PassCount', $passCount)
[void]$summaryObject.Add('WarningCount', $warningCount)
[void]$summaryObject.Add('FailCount', $failCount)
[void]$summaryObject.Add('RequiredFailures', $requiredFailures)

if ($Json) {
    $jsonText = ConvertTo-Json -InputObject $summaryObject -Depth 6
    Write-Output $jsonText
} else {
    $checks |
        Format-Table Name, Status, Required, Version, Path, Details -AutoSize |
        Out-String -Width 240 |
        Write-Output

    Write-Output ('Pass={0} Warning={1} Fail={2} RequiredFailures={3}' -f `
        $passCount, $warningCount, $failCount, $requiredFailures)
}

if ($Strict -and $requiredFailures -gt 0) {
    throw "Environment check failed with $requiredFailures required failure(s)."
}
