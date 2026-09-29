#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$SkipNative,
    [switch]$SkipFrontend,
    [switch]$SkipInstaller,
    [switch]$BootstrapInno,
    [switch]$KeepStaging,
    [switch]$EnvironmentOnly,
    [switch]$RepackageOnly,
    [string]$ProjectPath,
    [string]$NativeManifestPath,
    [string]$NativeTargetDirectory = '.build\native-release-target',
    [string]$NativeOutputPath,
    [string]$DotNetOfflineInstaller,
    [string]$Version = '1.0.0.0'
)

$ErrorActionPreference = 'Stop'

if ($RepackageOnly) {
    $SkipNative = $true
    $SkipFrontend = $true
}

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$CheckScript = Join-Path $ProjectRoot 'scripts\check-environment.ps1'
$InstallerScript = Join-Path $ProjectRoot 'installer\PPTFigureExporter.iss'
$DistRoot = Join-Path $ProjectRoot 'dist'
$DistApp = Join-Path $DistRoot 'app'
$DistPrereq = Join-Path $DistRoot 'prereq'
$DistLicenses = Join-Path $DistRoot 'licenses'
$ReferencePackageVersion = '1.0.3'
$ReferencePackageName = "Microsoft.NETFramework.ReferenceAssemblies.net48.$ReferencePackageVersion.nupkg"
$ReferencePackagePath = Join-Path $ProjectRoot ".build\tools\$ReferencePackageName"
$ReferenceAssembliesPath = Join-Path $ProjectRoot '.build\tools\net48-reference\build\.NETFramework\v4.8'

if (-not $ProjectPath) {
    $ProjectPath = Join-Path $ProjectRoot 'frontend\PPTFigureExporter.csproj'
} elseif (-not [System.IO.Path]::IsPathRooted($ProjectPath)) {
    $ProjectPath = Join-Path $ProjectRoot $ProjectPath
}

if (-not $NativeManifestPath) {
    $NativeManifestPath = Join-Path $ProjectRoot 'native\Cargo.toml'
} elseif (-not [System.IO.Path]::IsPathRooted($NativeManifestPath)) {
    $NativeManifestPath = Join-Path $ProjectRoot $NativeManifestPath
}

if (-not [System.IO.Path]::IsPathRooted($NativeTargetDirectory)) {
    $NativeTargetDirectory = Join-Path $ProjectRoot $NativeTargetDirectory
}
if ($NativeOutputPath) {
    if (-not [System.IO.Path]::IsPathRooted($NativeOutputPath)) {
        $NativeOutputPath = Join-Path $ProjectRoot $NativeOutputPath
    }
} else {
    $NativeOutputPath = Join-Path $NativeTargetDirectory 'release\figure-export-core.exe'
}

function Write-BuildStep {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "==> $Message" -ForegroundColor Cyan
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

function Get-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $null
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

function Invoke-InnoBootstrap {
    $toolsRoot = Join-Path $ProjectRoot '.build\tools'
    $setupPath = Join-Path $toolsRoot 'innosetup-setup.exe'
    $installRoot = Join-Path $toolsRoot 'innosetup'
    $bootstrapTemp = Join-Path $ProjectRoot '.build\temp\innosetup'
    $downloadUri = 'https://github.com/jrsoftware/issrc/releases/download/is-6_7_3/innosetup-6.7.3.exe'

    New-Item -ItemType Directory -Force -Path $toolsRoot | Out-Null

    if (-not (Test-Path -LiteralPath $setupPath -PathType Leaf)) {
        Write-BuildStep 'Downloading the official Inno Setup installer'
        Invoke-WebRequest `
            -Uri $downloadUri `
            -OutFile $setupPath `
            -UseBasicParsing
    }

    $signature = Get-AuthenticodeSignature -FilePath $setupPath
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch '(?i)Jordan Russell|JRSoftware|Pyrsys B\.V\.') {
        throw 'The downloaded Inno Setup installer did not pass the Authenticode signer check.'
    }

    if (-not (Test-Path -LiteralPath (Join-Path $installRoot 'ISCC.exe') -PathType Leaf)) {
        Write-BuildStep 'Installing portable Inno Setup into the project directory'
        New-Item -ItemType Directory -Force -Path $installRoot | Out-Null
        New-Item -ItemType Directory -Force -Path $bootstrapTemp | Out-Null

        $quotedInstallRoot = $installRoot.Replace('"', '\"')
        $quotedLogPath = (Join-Path $bootstrapTemp 'innosetup-install.log').Replace('"', '\"')
        $argumentList = @(
            '/PORTABLE=1',
            '/CURRENTUSER',
            '/NOICONS',
            '/VERYSILENT',
            '/SUPPRESSMSGBOXES',
            '/NORESTART',
            '/SP-',
            "/DIR=`"$quotedInstallRoot`"",
            "/LOG=`"$quotedLogPath`""
        )

        $previousTemp = [Environment]::GetEnvironmentVariable('TEMP', 'Process')
        $previousTmp = [Environment]::GetEnvironmentVariable('TMP', 'Process')
        try {
            [Environment]::SetEnvironmentVariable('TEMP', $bootstrapTemp, 'Process')
            [Environment]::SetEnvironmentVariable('TMP', $bootstrapTemp, 'Process')
            $process = Start-Process `
                -FilePath $setupPath `
                -ArgumentList $argumentList `
                -WorkingDirectory $ProjectRoot `
                -WindowStyle Hidden `
                -Wait `
                -PassThru
        } finally {
            [Environment]::SetEnvironmentVariable('TEMP', $previousTemp, 'Process')
            [Environment]::SetEnvironmentVariable('TMP', $previousTmp, 'Process')
        }

        if ($null -eq $process -or $process.ExitCode -ne 0) {
            $exitCode = if ($null -eq $process) { 'unknown' } else { $process.ExitCode }
            throw "Portable Inno Setup bootstrap failed with exit code $exitCode."
        }
    }

    $isccPath = Get-IsccPath
    if (-not $isccPath) {
        throw 'Portable Inno Setup bootstrap completed without producing ISCC.exe.'
    }

    return $isccPath
}

function Invoke-EnvironmentCheck {
    Write-BuildStep 'Checking the local build environment'

    $checkParameters = @{
        Json             = $true
        ForBuild         = [bool](-not $RepackageOnly)
        SkipNative       = [bool]($SkipNative -or $RepackageOnly)
        RequireInstaller = [bool](-not $SkipInstaller)
    }

    $output = @(& $CheckScript @checkParameters)
    $jsonText = $output -join [Environment]::NewLine
    try {
        $report = $jsonText | ConvertFrom-Json
    } catch {
        throw "Environment check did not return valid JSON: $($_.Exception.Message)"
    }

    if ($report.RequiredFailures -gt 0) {
        $failedChecks = @(
            $report.Checks | Where-Object { $_.Required -and $_.Status -eq 'Fail' }
        )
        $names = ($failedChecks | ForEach-Object { $_.Name }) -join ', '
        throw "Environment check failed: $names"
    }

    return $report
}

function Ensure-ReferenceAssemblies {
    $marker = Join-Path $ReferenceAssembliesPath 'mscorlib.dll'
    if (Test-Path -LiteralPath $marker -PathType Leaf) {
        return
    }

    $toolsRoot = Join-Path $ProjectRoot '.build\tools'
    New-Item -ItemType Directory -Force -Path $toolsRoot | Out-Null
    $packageUri = "https://api.nuget.org/v3-flatcontainer/microsoft.netframework.referenceassemblies.net48/$ReferencePackageVersion/$ReferencePackageName"
    if (-not (Test-Path -LiteralPath $ReferencePackagePath -PathType Leaf)) {
        Write-BuildStep 'Downloading Microsoft .NET Framework 4.8 reference assemblies from NuGet'
        Invoke-WebRequest -Uri $packageUri -OutFile $ReferencePackagePath -UseBasicParsing
    }

    $referenceRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $ReferenceAssembliesPath))
    New-Item -ItemType Directory -Force -Path $referenceRoot | Out-Null
    Expand-Archive -LiteralPath $ReferencePackagePath -DestinationPath $referenceRoot -Force
    if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) {
        throw "The reference assemblies package did not provide $marker"
    }
}

function Invoke-NativeBuild {
    if (-not (Test-Path -LiteralPath $NativeManifestPath -PathType Leaf)) {
        throw "Rust manifest was not found: $NativeManifestPath"
    }

    $lockPath = Join-Path (Split-Path -Parent $NativeManifestPath) 'Cargo.lock'
    if (-not (Test-Path -LiteralPath $lockPath -PathType Leaf)) {
        throw "Cargo.lock is required for the locked native build: $lockPath"
    }

    $cargoCandidates = @(
        (Get-CommandPath -Name 'cargo.exe'),
        'C:\Users\admin\.cargo\bin\cargo.exe'
    )
    $cargoPath = $null
    foreach ($cargoCandidate in $cargoCandidates) {
        if ($cargoCandidate -and (Test-Path -LiteralPath $cargoCandidate -PathType Leaf)) {
            $cargoPath = (Resolve-Path -LiteralPath $cargoCandidate).Path
            break
        }
    }
    if (-not $cargoPath) {
        throw 'cargo.exe was not found.'
    }

    Write-BuildStep 'Building the Rust export core with cargo --locked --release'
    $nativeProjectDirectory = Split-Path -Parent $NativeManifestPath
    Push-Location $nativeProjectDirectory
    try {
        & $cargoPath build --locked --release --manifest-path $NativeManifestPath --target-dir $NativeTargetDirectory | ForEach-Object { Write-Host $_ }
        if ($LASTEXITCODE -ne 0) {
            throw "Rust build failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }

    if (-not (Test-Path -LiteralPath $NativeOutputPath -PathType Leaf)) {
        throw "Rust build completed without producing $NativeOutputPath"
    }

    return $NativeOutputPath
}

function Invoke-FrontendBuild {
    if (-not (Test-Path -LiteralPath $ProjectPath -PathType Leaf)) {
        throw "Frontend project was not found: $ProjectPath"
    }

    $msbuildPath = Get-MsBuildPath
    if (-not $msbuildPath) {
        throw 'MSBuild.exe was not found.'
    }

    Write-BuildStep 'Building the .NET Framework frontend with MSBuild'
    $msbuildArguments = @(
        $ProjectPath,
        '/t:Rebuild',
        '/p:Configuration=Release',
        '/p:Platform=x64',
        "/p:FrameworkPathOverride=$ReferenceAssembliesPath",
        '/nologo',
        '/v:minimal'
    )
    & $msbuildPath @msbuildArguments | ForEach-Object { Write-Host $_ }
    if ($LASTEXITCODE -ne 0) {
        throw "Frontend MSBuild failed with exit code $LASTEXITCODE."
    }

    $projectDirectory = Split-Path -Parent $ProjectPath
    $dll = Get-ChildItem -LiteralPath $projectDirectory -Recurse -File -Filter 'PPTFigureExporter.dll' |
        Where-Object { $_.FullName -match '\\bin\\' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -eq $dll) {
        throw 'MSBuild completed without producing PPTFigureExporter.dll under a bin directory.'
    }

    return $dll.FullName
}

function Copy-FrontendOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DllPath
    )

    $outputDirectory = Split-Path -Parent $DllPath
    $files = Get-ChildItem -LiteralPath $outputDirectory -File |
        Where-Object { $_.Extension -notin @('.pdb', '.xml') }
    foreach ($file in $files) {
        Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $DistApp $file.Name) -Force
    }

    $requiredDll = Join-Path $DistApp 'PPTFigureExporter.dll'
    if (-not (Test-Path -LiteralPath $requiredDll -PathType Leaf)) {
        throw "The frontend output did not contain $requiredDll"
    }
}

function Copy-PackageDocuments {
    $projectLicense = Join-Path $ProjectRoot 'LICENSE'
    if (-not (Test-Path -LiteralPath $projectLicense -PathType Leaf)) {
        throw 'The project Apache-2.0 LICENSE file is required for packaging.'
    }
    $notices = Join-Path $ProjectRoot 'docs\THIRD_PARTY_NOTICES.txt'
    if (-not (Test-Path -LiteralPath $notices -PathType Leaf)) {
        throw 'Third-party notices are required. Run scripts\generate-notices.ps1 before packaging.'
    }
    $licenseCandidates = @(
        (Join-Path $ProjectRoot 'LICENSE'),
        (Join-Path $ProjectRoot 'LICENSE.txt'),
        (Join-Path $ProjectRoot 'LICENSE.md'),
        (Join-Path $ProjectRoot 'docs\THIRD_PARTY_NOTICES.txt'),
        (Join-Path $ProjectRoot 'docs\THIRD_PARTY_NOTICES.md')
    )

    foreach ($candidate in $licenseCandidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            Copy-Item -LiteralPath $candidate -Destination (Join-Path $DistLicenses (Split-Path -Leaf $candidate)) -Force
        }
    }

    $readme = Join-Path $ProjectRoot 'README.md'
    if (Test-Path -LiteralPath $readme -PathType Leaf) {
        Copy-Item -LiteralPath $readme -Destination (Join-Path $DistApp 'README.md') -Force
    }
}

function Assert-DistPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    $distResolved = (Resolve-Path -LiteralPath $DistRoot).Path.TrimEnd('\') + '\'
    $targetResolved = (Resolve-Path -LiteralPath $Path).Path.TrimEnd('\') + '\'
    if (-not $targetResolved.StartsWith($distResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove a path outside the project dist directory: $targetResolved"
    }
}

function Clear-DistApp {
    param(
        [switch]$PreserveNative
    )

    if (-not (Test-Path -LiteralPath $DistApp)) {
        return
    }

    Assert-DistPath -Path $DistApp
    if (-not $PreserveNative) {
        Remove-Item -LiteralPath $DistApp -Recurse -Force
        return
    }

    $files = Get-ChildItem -LiteralPath $DistApp -File
    foreach ($file in $files) {
        if ($file.Name -ne 'figure-export-core.exe') {
            Remove-Item -LiteralPath $file.FullName -Force
        }
    }

    $directories = Get-ChildItem -LiteralPath $DistApp -Directory
    foreach ($directory in $directories) {
        Remove-Item -LiteralPath $directory.FullName -Recurse -Force
    }
}

function Resolve-DotNetOfflineInstaller {
    if ([string]::IsNullOrWhiteSpace($DotNetOfflineInstaller)) {
        return $null
    }

    $source = $DotNetOfflineInstaller
    if (-not [System.IO.Path]::IsPathRooted($source)) {
        $source = Join-Path $ProjectRoot $source
    }

    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "The explicitly supplied .NET Framework 4.8 sidecar was not found: $source"
    }

    $resolvedSource = (Resolve-Path -LiteralPath $source).Path
    $resolvedPrereqRoot = [System.IO.Path]::GetFullPath($DistPrereq).TrimEnd('\') + '\'
    if ($resolvedSource.StartsWith($resolvedPrereqRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'The explicit .NET sidecar source must be outside dist\prereq so staging cleanup cannot remove the input.'
    }

    return $resolvedSource
}

function Copy-DotNetOfflineInstaller {
    if ($null -eq $dotNetOfflineSource) {
        return $null
    }

    $source = $dotNetOfflineSource
    $signature = Get-AuthenticodeSignature -FilePath $source
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch '(?i)Microsoft') {
        throw 'The explicitly supplied .NET Framework offline installer did not pass the Microsoft Authenticode signer check.'
    }

    New-Item -ItemType Directory -Force -Path $DistPrereq | Out-Null
    $destination = Join-Path $DistPrereq 'dotNet48-x86-x64-AllOS-ENU.exe'
    $sourceResolved = (Resolve-Path -LiteralPath $source).Path
    $destinationResolved = $null
    if (Test-Path -LiteralPath $destination -PathType Leaf) {
        $destinationResolved = (Resolve-Path -LiteralPath $destination).Path
    }

    if ($sourceResolved -ne $destinationResolved) {
        Copy-Item -LiteralPath $source -Destination $destination -Force
    }

    return $destination
}

if ($EnvironmentOnly) {
    Invoke-EnvironmentCheck | ConvertTo-Json -Depth 8
    return
}

if ($BootstrapInno) {
    [void](Invoke-InnoBootstrap)
}

$report = Invoke-EnvironmentCheck
if (-not $RepackageOnly) {
    Ensure-ReferenceAssemblies
}

if (-not $SkipInstaller -and -not (Test-Path -LiteralPath $InstallerScript -PathType Leaf)) {
    throw "Installer script was not found: $InstallerScript"
}

$dotNetOfflineSource = Resolve-DotNetOfflineInstaller

New-Item -ItemType Directory -Force -Path $DistRoot | Out-Null
$stagedNativePath = Join-Path $DistApp 'figure-export-core.exe'
$preserveStagedNative = $SkipNative -and (Test-Path -LiteralPath $stagedNativePath -PathType Leaf)
if (-not $RepackageOnly -and -not $SkipFrontend -and (Test-Path -LiteralPath $DistApp)) {
    Clear-DistApp -PreserveNative:$preserveStagedNative
}
New-Item -ItemType Directory -Force -Path $DistApp | Out-Null

foreach ($stagingPath in @($DistPrereq, $DistLicenses)) {
    if (Test-Path -LiteralPath $stagingPath) {
        Assert-DistPath -Path $stagingPath
        Remove-Item -LiteralPath $stagingPath -Recurse -Force
    }

    New-Item -ItemType Directory -Force -Path $stagingPath | Out-Null
}

$nativeOutput = $null
if ($RepackageOnly) {
    $nativeOutput = $stagedNativePath
    if (-not (Test-Path -LiteralPath $nativeOutput -PathType Leaf)) {
        throw "RepackageOnly requires a validated native core at $nativeOutput"
    }
} elseif (-not $SkipNative) {
    $nativeOutput = Invoke-NativeBuild
} elseif (Test-Path -LiteralPath $stagedNativePath -PathType Leaf) {
    $nativeOutput = $stagedNativePath
} else {
    $nativeOutput = $NativeOutputPath
    if (-not (Test-Path -LiteralPath $nativeOutput -PathType Leaf)) {
        Write-Warning "Native build was skipped and no existing core was found at $nativeOutput"
        $nativeOutput = $null
    }
}

$frontendDll = $null
if ($RepackageOnly) {
    $frontendDll = Join-Path $DistApp 'PPTFigureExporter.dll'
    if (-not (Test-Path -LiteralPath $frontendDll -PathType Leaf)) {
        throw "RepackageOnly requires a validated frontend DLL at $frontendDll"
    }
} elseif (-not $SkipFrontend) {
    $frontendDll = Invoke-FrontendBuild
} else {
    $frontendDll = Join-Path $DistApp 'PPTFigureExporter.dll'
    if (-not (Test-Path -LiteralPath $frontendDll -PathType Leaf)) {
        $projectDirectory = Split-Path -Parent $ProjectPath
        $frontendDllCandidate = Get-ChildItem -LiteralPath $projectDirectory -Recurse -File -Filter 'PPTFigureExporter.dll' |
            Where-Object { $_.FullName -match '\\bin\\' } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if ($null -ne $frontendDllCandidate) {
            Copy-FrontendOutput -DllPath $frontendDllCandidate.FullName
        }
    }
    if (-not (Test-Path -LiteralPath $frontendDll -PathType Leaf)) {
        $frontendDll = $null
    }
}

if ($frontendDll) {
    if (-not $SkipFrontend -and -not $RepackageOnly) {
        Copy-FrontendOutput -DllPath $frontendDll
    }
} elseif (-not $SkipInstaller) {
    throw 'A frontend DLL is required for packaging.'
}

$nativeDestination = Join-Path $DistApp 'figure-export-core.exe'
if ($nativeOutput) {
    $sourceResolved = (Resolve-Path -LiteralPath $nativeOutput).Path
    $destinationResolved = $null
    if (Test-Path -LiteralPath $nativeDestination -PathType Leaf) {
        $destinationResolved = (Resolve-Path -LiteralPath $nativeDestination).Path
    }

    if ($sourceResolved -ne $destinationResolved) {
        Copy-Item -LiteralPath $nativeOutput -Destination $nativeDestination -Force
    }
} elseif (-not $SkipInstaller) {
    throw 'A native core executable is required for packaging.'
}

Copy-PackageDocuments
$dotNetSidecar = Copy-DotNetOfflineInstaller

$manifest = [ordered]@{
    Product          = 'PPTFigureExporter'
    Version          = $Version
    BuiltAt          = (Get-Date).ToString('o')
    Platform         = 'x64'
    Frontend         = 'PPTFigureExporter.dll'
    NativeCore       = 'figure-export-core.exe'
    DotNetOffline    = $false
    DotNetDependency = 'external'
    DotNetSidecar    = if ($dotNetSidecar) {
        'prereq\' + (Split-Path -Leaf $dotNetSidecar)
    } else {
        $null
    }
}
$manifestPath = Join-Path $DistApp 'build-manifest.json'
$manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

if (-not $SkipInstaller) {
    $isccPath = Get-IsccPath
    if (-not $isccPath) {
        throw 'ISCC.exe was not found. Re-run with -BootstrapInno or install Inno Setup locally.'
    }

    Write-BuildStep 'Compiling the Inno Setup installer'
    $outputDirectory = $DistRoot
    $isccArguments = @(
        '/Q',
        "/O$outputDirectory",
        "/DAppVersion=$Version",
        $InstallerScript
    )
    & $isccPath @isccArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Inno Setup compilation failed with exit code $LASTEXITCODE."
    }

    $setupPath = Join-Path $DistRoot 'PPTFigureExporter_Setup.exe'
    if (-not (Test-Path -LiteralPath $setupPath -PathType Leaf)) {
        throw "Inno Setup completed without producing $setupPath"
    }
}

if (-not $KeepStaging) {
    Write-Host "Staging output: $DistApp" -ForegroundColor Green
}

Write-Host 'Build completed.' -ForegroundColor Green
