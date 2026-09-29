[CmdletBinding()]
param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [string]$OutputPath = '',
    [string]$TargetTriple = 'x86_64-pc-windows-msvc',
    [string]$CargoHome = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $ProjectRoot 'docs\THIRD_PARTY_NOTICES.txt'
}

$manifestPath = Join-Path $ProjectRoot 'native\Cargo.toml'
$lockPath = Join-Path $ProjectRoot 'native\Cargo.lock'

if (-not (Test-Path -LiteralPath $manifestPath)) {
    throw "Cargo manifest was not found: $manifestPath"
}

if (-not (Test-Path -LiteralPath $lockPath)) {
    throw "Cargo.lock was not found: $lockPath"
}

if ([string]::IsNullOrWhiteSpace($CargoHome)) {
    $repositoryCargoHome = Join-Path $ProjectRoot '.build\cargo-home'
    if (Test-Path -LiteralPath $repositoryCargoHome) {
        $CargoHome = (Resolve-Path -LiteralPath $repositoryCargoHome).Path
    }
    elseif (-not [string]::IsNullOrWhiteSpace($env:CARGO_HOME)) {
        $CargoHome = $env:CARGO_HOME
    }
}

if (-not [string]::IsNullOrWhiteSpace($CargoHome)) {
    $CargoHome = (Resolve-Path -LiteralPath $CargoHome).Path
    $env:CARGO_HOME = $CargoHome
}

function Get-Metadata {
    param(
        [string]$Manifest,
        [string]$Platform
    )

    $lines = @(& cargo metadata --manifest-path $Manifest --locked --offline --format-version 1 --filter-platform $Platform 2>$null)
    if ($LASTEXITCODE -ne 0 -or $lines.Count -eq 0) {
        return $null
    }

    try {
        return ($lines -join [Environment]::NewLine) | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Get-LockRecords {
    param(
        [string]$Path
    )

    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $records = New-Object 'System.Collections.Generic.List[object]'
    $blocks = [regex]::Matches($text, '(?ms)^\[\[package\]\]\r?\n(.*?)(?=^\[\[package\]\]|\z)')

    foreach ($block in $blocks) {
        $body = $block.Groups[1].Value
        $nameMatch = [regex]::Match($body, '(?m)^name\s*=\s*"([^"]+)"')
        $versionMatch = [regex]::Match($body, '(?m)^version\s*=\s*"([^"]+)"')
        $sourceMatch = [regex]::Match($body, '(?m)^source\s*=\s*"([^"]+)"')
        $checksumMatch = [regex]::Match($body, '(?m)^checksum\s*=\s*"([^"]+)"')

        if (-not $nameMatch.Success -or -not $versionMatch.Success) {
            continue
        }

        $records.Add([pscustomobject]@{
            Name = $nameMatch.Groups[1].Value
            Version = $versionMatch.Groups[1].Value
            Source = if ($sourceMatch.Success) { $sourceMatch.Groups[1].Value } else { $null }
            Checksum = if ($checksumMatch.Success) { $checksumMatch.Groups[1].Value } else { $null }
        })
    }

    return $records
}

function Get-RegistryRoot {
    param(
        [string]$CargoHomeValue
    )

    if ([string]::IsNullOrWhiteSpace($CargoHomeValue)) {
        return $null
    }

    $root = Join-Path $CargoHomeValue 'registry\src'
    if (-not (Test-Path -LiteralPath $root)) {
        return $null
    }

    return Get-ChildItem -LiteralPath $root -Directory |
        Where-Object { $_.Name -like 'index.crates.io-*' } |
        Select-Object -First 1
}

function Get-PackageDirectory {
    param(
        [object]$Package,
        [string]$RegistryRoot
    )

    if ($null -ne $Package.manifest_path) {
        $directory = Split-Path -Parent $Package.manifest_path
        if (Test-Path -LiteralPath $directory) {
            return $directory
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($RegistryRoot)) {
        $candidate = Join-Path $RegistryRoot ($Package.name + '-' + $Package.version)
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Read-Text {
    param(
        [string]$Path
    )

    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $bytes = $bytes[3..($bytes.Length - 1)]
    }

    return ([Text.Encoding]::UTF8.GetString($bytes)).
        Replace("`r`n", "`n").
        Replace("`r", "`n").
        TrimEnd()
}

function Get-Digest {
    param(
        [string]$Text
    )

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-SourceFiles {
    param(
        [string]$PackageDirectory,
        [string]$PackageName
    )

    $paths = @(
        Get-ChildItem -LiteralPath $PackageDirectory -File |
            Where-Object { $_.Name -match '(?i)^(license|licence|copying|notice)([-_.].*|$)' } |
            Select-Object -ExpandProperty FullName
    )

    if ($PackageName -eq 'libwebp-sys') {
        $vendor = Join-Path $PackageDirectory 'vendor'
        if (Test-Path -LiteralPath $vendor) {
            $paths += @(
                Get-ChildItem -LiteralPath $vendor -Recurse -File |
                    Where-Object { $_.Name -match '(?i)^(license|licence|copying|notice|authors|patents)([-_.].*|$)' } |
                    Select-Object -ExpandProperty FullName
            )
        }
    }

    return @($paths | Sort-Object -Unique)
}

function Get-IjgText {
    param(
        [string]$PackageDirectory
    )

    $sourcePath = Join-Path $PackageDirectory 'src\fdct.rs'
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        return $null
    }

    $source = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::UTF8)
    $match = [regex]::Match($source, '(?s)^/\*\r?\n(.*?)\r?\n \*/')
    if (-not $match.Success) {
        return $null
    }

    $lines = $match.Groups[1].Value -split '\r?\n'
    $notice = foreach ($line in $lines) {
        if ($line.StartsWith(' *')) {
            $line.Substring(2).TrimStart()
        }
        else {
            $line
        }
    }

    return ($notice -join [char]10).TrimEnd()
}

function Get-SpdxIds {
    param(
        [string]$Expression
    )

    $known = @('0BSD', 'Apache-2.0', 'BSD-2-Clause', 'BSD-3-Clause', 'BSL-1.0', 'IJG', 'MIT', 'Unicode-3.0', 'Unlicense', 'Zlib')
    $result = New-Object 'System.Collections.Generic.List[string]'
    foreach ($id in $known) {
        if ($Expression -match ('(?<![A-Za-z0-9.-])' + [regex]::Escape($id) + '(?![A-Za-z0-9.-])')) {
            $result.Add($id)
        }
    }

    return @($result)
}

function Get-CanonicalPath {
    param(
        [string]$SpdxId,
        [string]$RegistryRoot
    )

    $relativePath = switch ($SpdxId) {
        '0BSD' { 'adler2-2.0.1\LICENSE-0BSD' }
        'Apache-2.0' { 'adler2-2.0.1\LICENSE-APACHE' }
        'BSD-2-Clause' { 'arrayref-0.3.9\LICENSE' }
        'BSD-3-Clause' { 'tiny-skia-0.11.4\LICENSE' }
        'BSL-1.0' { 'ryu-1.0.23\LICENSE-BOOST' }
        'MIT' { 'adler2-2.0.1\LICENSE-MIT' }
        'Unicode-3.0' { 'unicode-ident-1.0.26\LICENSE-UNICODE' }
        'Unlicense' { 'byteorder-lite-0.1.0\UNLICENSE' }
        'Zlib' { 'bytemuck-1.25.2\LICENSE-ZLIB' }
        default { $null }
    }

    if ($null -eq $relativePath) {
        return $null
    }

    $candidate = Join-Path $RegistryRoot $relativePath
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    return $null
}

$metadata = Get-Metadata -Manifest $manifestPath -Platform $TargetTriple
$lockRecords = Get-LockRecords -Path $lockPath
$lockByKey = @{}
foreach ($record in $lockRecords) {
    $lockByKey[$record.Name + '|' + $record.Version] = $record
}

$registryRootObject = Get-RegistryRoot -CargoHomeValue $CargoHome
$registryRoot = if ($null -ne $registryRootObject) { $registryRootObject.FullName } else { $null }
$packages = New-Object 'System.Collections.Generic.List[object]'

if ($null -ne $metadata -and $null -ne $metadata.resolve) {
    $nodesById = @{}
    foreach ($node in $metadata.resolve.nodes) {
        $nodesById[$node.id] = $node
    }

    $reachable = New-Object 'System.Collections.Generic.HashSet[string]'
    $queue = New-Object 'System.Collections.Generic.Queue[string]'
    $queue.Enqueue($metadata.resolve.root)

    while ($queue.Count -gt 0) {
        $id = $queue.Dequeue()
        if (-not $reachable.Add($id)) {
            continue
        }

        $node = $nodesById[$id]
        if ($null -eq $node) {
            continue
        }

        foreach ($dependency in @($node.deps)) {
            if ($null -ne $dependency.pkg) {
                $queue.Enqueue($dependency.pkg)
            }
        }
    }

    foreach ($package in $metadata.packages) {
        if ($package.id -eq $metadata.resolve.root -or -not $reachable.Contains($package.id) -or [string]::IsNullOrWhiteSpace($package.source)) {
            continue
        }

        $packages.Add($package)
    }
}
else {
    if ($null -eq $registryRootObject) {
        throw 'cargo metadata failed and no crates.io registry source directory was found.'
    }

    foreach ($record in $lockRecords) {
        if ([string]::IsNullOrWhiteSpace($record.Source)) {
            continue
        }

        $directory = Join-Path $registryRoot ($record.Name + '-' + $record.Version)
        $manifest = Join-Path $directory 'Cargo.toml'
        if (-not (Test-Path -LiteralPath $manifest)) {
            continue
        }

        $manifestText = [IO.File]::ReadAllText($manifest, [Text.Encoding]::UTF8)
        $licenseMatch = [regex]::Match($manifestText, '(?m)^license\s*=\s*"([^"]+)"')
        $repositoryMatch = [regex]::Match($manifestText, '(?m)^repository\s*=\s*"([^"]+)"')

        $packages.Add([pscustomobject]@{
            id = $record.Name + '|' + $record.Version
            name = $record.Name
            version = $record.Version
            license = if ($licenseMatch.Success) { $licenseMatch.Groups[1].Value } else { $null }
            license_file = $null
            repository = if ($repositoryMatch.Success) { $repositoryMatch.Groups[1].Value } else { $null }
            source = $record.Source
            manifest_path = $manifest
        })
    }
}

if ($packages.Count -eq 0) {
    throw 'No resolved third-party crates were found.'
}

$noticeByDigest = @{}
$noticeRecords = New-Object 'System.Collections.Generic.List[object]'

function Register-Notice {
    param(
        [string]$Text,
        [string]$Origin
    )

    $digest = Get-Digest -Text $Text
    if ($noticeByDigest.ContainsKey($digest)) {
        $record = $noticeByDigest[$digest]
        $record.Origins.Add($Origin)
        return $record
    }

    $record = [pscustomobject]@{
        Id = $null
        Digest = $digest
        Text = $Text
        Origins = New-Object 'System.Collections.Generic.List[string]'
    }
    $record.Origins.Add($Origin)
    $noticeByDigest[$digest] = $record
    $noticeRecords.Add($record)
    return $record
}

foreach ($package in ($packages | Sort-Object name,version)) {
    $key = $package.name + '|' + $package.version
    $lockRecord = $lockByKey[$key]
    $directory = Get-PackageDirectory -Package $package -RegistryRoot $registryRoot
    if ($null -eq $directory) {
        throw "Package source directory was not found for $key"
    }

    $references = New-Object 'System.Collections.Generic.List[object]'
    $sourceFiles = Get-SourceFiles -PackageDirectory $directory -PackageName $package.name
    if ($null -ne $package.license_file -and -not [string]::IsNullOrWhiteSpace($package.license_file)) {
        $declared = Join-Path $directory $package.license_file
        if (Test-Path -LiteralPath $declared) {
            $sourceFiles += $declared
        }
    }

    foreach ($file in @($sourceFiles | Sort-Object -Unique)) {
        $text = Read-Text -Path $file
        $relative = $file.Substring($directory.Length).TrimStart([char[]]"\/")
        $record = Register-Notice -Text $text -Origin ($package.name + '-' + $package.version + '/' + $relative)
        $references.Add([pscustomobject]@{
            Origin = $relative
            Record = $record
        })
    }

    if ($package.name -eq 'jpeg-encoder') {
        $ijgText = Get-IjgText -PackageDirectory $directory
        if (-not [string]::IsNullOrWhiteSpace($ijgText)) {
            $record = Register-Notice -Text $ijgText -Origin ($package.name + '-' + $package.version + '/src/fdct.rs (IJG notice block)')
            $references.Add([pscustomobject]@{
                Origin = 'src/fdct.rs (IJG notice block)'
                Record = $record
            })
        }
    }

    $expression = [string]$package.license
    if ([string]::IsNullOrWhiteSpace($expression)) {
        throw "License expression is missing for $key"
    }

    foreach ($spdxId in (Get-SpdxIds -Expression $expression)) {
        if ($references.Count -gt 0) { continue }
        $hasReference = $false
        foreach ($reference in $references) {
            if ($reference.Origin -match ('(?i)' + [regex]::Escape($spdxId))) {
                $hasReference = $true
                break
            }
        }

        if ($hasReference) {
            continue
        }

        if ($spdxId -eq 'IJG' -and $package.name -eq 'jpeg-encoder') {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($registryRoot)) {
            throw "Cannot locate canonical license source for $spdxId"
        }

        $canonicalPath = Get-CanonicalPath -SpdxId $spdxId -RegistryRoot $registryRoot
        if ($null -eq $canonicalPath) {
            throw "Cannot locate canonical license source for $spdxId"
        }

        $canonicalText = Read-Text -Path $canonicalPath
        $record = Register-Notice -Text $canonicalText -Origin ('SPDX ' + $spdxId + ' canonical text, sourced from ' + $canonicalPath.Substring($registryRoot.Length).TrimStart([char[]]"\/"))
        $references.Add([pscustomobject]@{
            Origin = 'SPDX ' + $spdxId + ' canonical text'
            Record = $record
        })
    }

    if ($references.Count -eq 0) {
        throw "No license or notice text was collected for $key ($expression)"
    }

    $package | Add-Member -NotePropertyName NoticeReferences -NotePropertyValue $references
    $checksum = if ($null -ne $lockRecord) { $lockRecord.Checksum } else { $null }
    $package | Add-Member -NotePropertyName LockChecksum -NotePropertyValue $checksum
    $archive = 'https://static.crates.io/crates/' + $package.name + '/' + $package.name + '-' + $package.version + '.crate'
    $package | Add-Member -NotePropertyName SourceArchive -NotePropertyValue $archive
}

$noticeIndex = 1
foreach ($record in $noticeRecords) {
    $record.Id = 'N' + $noticeIndex.ToString('000')
    $noticeIndex++
}

$builder = New-Object Text.StringBuilder
[void]$builder.AppendLine('PPT Figure Exporter — Third-Party Notices')
[void]$builder.AppendLine('')
[void]$builder.AppendLine('This file is generated from native/Cargo.lock and Cargo metadata.')
[void]$builder.AppendLine(('The resolved target is {0}.' -f $TargetTriple))
[void]$builder.AppendLine('The crate versions and checksums below are fixed by Cargo.lock.')
[void]$builder.AppendLine('All listed crates are unmodified upstream crates.io sources.')
[void]$builder.AppendLine('Project-owned code is licensed under Apache-2.0; see the separate LICENSE file.')
[void]$builder.AppendLine('Third-party components retain the licenses and notices reproduced below.')
[void]$builder.AppendLine('Build tools and installer components are outside this runtime dependency list.')
[void]$builder.AppendLine('')
[void]$builder.AppendLine('Resolved third-party packages')
[void]$builder.AppendLine('==============================')
[void]$builder.AppendLine('')

foreach ($package in ($packages | Sort-Object name,version)) {
    [void]$builder.AppendLine(('Package: {0} {1}' -f $package.name, $package.version))
    [void]$builder.AppendLine(('License expression: {0}' -f $package.license))
    if ([string]::IsNullOrWhiteSpace($package.repository)) {
        [void]$builder.AppendLine('Repository: not declared in the crate manifest')
    }
    else {
        [void]$builder.AppendLine(('Repository: {0}' -f $package.repository))
    }
    [void]$builder.AppendLine(('Source archive: {0}' -f $package.SourceArchive))
    if ([string]::IsNullOrWhiteSpace($package.LockChecksum)) {
        [void]$builder.AppendLine('Cargo.lock checksum: unavailable')
    }
    else {
        [void]$builder.AppendLine(('Cargo.lock checksum: {0}' -f $package.LockChecksum))
    }
    [void]$builder.AppendLine('License and notice text:')
    foreach ($reference in $package.NoticeReferences) {
        [void]$builder.AppendLine(('  [{0}] {1}' -f $reference.Record.Id, $reference.Origin))
    }
    [void]$builder.AppendLine('')
}

[void]$builder.AppendLine('Consolidated original license and notice texts')
[void]$builder.AppendLine('=============================================')
[void]$builder.AppendLine('')
[void]$builder.AppendLine('Identical normalized texts are included once; package mappings above remain complete.')
[void]$builder.AppendLine('')

foreach ($record in $noticeRecords) {
    [void]$builder.AppendLine(('----- {0} -----' -f $record.Id))
    [void]$builder.AppendLine('Origin records:')
    foreach ($origin in ($record.Origins | Sort-Object -Unique)) {
        [void]$builder.AppendLine(('  {0}' -f $origin))
    }
    [void]$builder.AppendLine('')
    [void]$builder.AppendLine($record.Text)
    [void]$builder.AppendLine('')
}

$outputDirectory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$encoding = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText($OutputPath, $builder.ToString(), $encoding)
Write-Output ('Wrote {0} package notices and {1} consolidated texts to {2}' -f $packages.Count, $noticeRecords.Count, (Resolve-Path -LiteralPath $OutputPath).Path)
