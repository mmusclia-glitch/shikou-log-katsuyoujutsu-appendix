param(
    [Parameter(Mandatory = $true)]
    [string]$Source,

    [Parameter(Mandatory = $true)]
    [string]$Since)


$ErrorActionPreference = 'Stop'


function Normalize-InputText([string]$text) {
    if ($null -eq $text) {
        return ''
    }

    $result = $text.Trim()

    while ($result.Length -ge 2 -and $result.StartsWith('"') -and $result.EndsWith('"')) {
        $result = $result.Substring(1, $result.Length - 2).Trim()
    }

    return [Environment]::ExpandEnvironmentVariables($result)
}

function Resolve-ValidatedPath([string]$pathText, [ValidateSet('Leaf', 'Container')] [string]$PathType) {
    $normalized = Normalize-InputText $pathText

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    try {
        $resolved = Resolve-Path -LiteralPath $normalized -ErrorAction Stop
        foreach ($item in $resolved) {
            $providerPath = $item.ProviderPath
            if ($PathType -eq 'Leaf' -and (Test-Path -LiteralPath $providerPath -PathType Leaf)) {
                return [System.IO.Path]::GetFullPath($providerPath)
            }
            if ($PathType -eq 'Container' -and (Test-Path -LiteralPath $providerPath -PathType Container)) {
                return [System.IO.Path]::GetFullPath($providerPath)
            }
        }
    }
    catch {
    }

    if ($PathType -eq 'Leaf' -and (Test-Path -LiteralPath $normalized -PathType Leaf)) {
        return [System.IO.Path]::GetFullPath($normalized)
    }

    if ($PathType -eq 'Container' -and (Test-Path -LiteralPath $normalized -PathType Container)) {
        return [System.IO.Path]::GetFullPath($normalized)
    }

    return $null
}

function Write-PathDiagnostic([string]$label, [string]$rawValue, [ValidateSet('Leaf', 'Container')] [string]$PathType) {
    $normalized = Normalize-InputText $rawValue

    Write-Host ('[{0}] の入力値を確認してください。' -f $label)
    Write-Host ('  入力値(生)     : [{0}]' -f $rawValue)
    Write-Host ('  入力値(整形後) : [{0}]' -f $normalized)

    $parent = ''
    $leaf = ''

    try {
        $parent = Split-Path -Path $normalized -Parent
        $leaf = Split-Path -Path $normalized -Leaf
    }
    catch {
    }

    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Write-Host ('  親フォルダ     : {0}' -f $parent)
        Write-Host ('  親フォルダ存在 : {0}' -f (Test-Path -LiteralPath $parent -PathType Container))
    }

    if (-not [string]::IsNullOrWhiteSpace($leaf)) {
        Write-Host ('  末尾要素       : {0}' -f $leaf)
    }

    if ($PathType -eq 'Leaf' -and -not [string]::IsNullOrWhiteSpace($parent) -and (Test-Path -LiteralPath $parent -PathType Container)) {
        $nearby = Get-ChildItem -LiteralPath $parent -File -Filter '*.xlsx' -ErrorAction SilentlyContinue |
            Sort-Object Name |
            Select-Object -First 10 -ExpandProperty Name

        if ($nearby.Count -gt 0) {
            Write-Host '  同一フォルダの .xlsx 一覧（先頭10件）:'
            foreach ($name in $nearby) {
                Write-Host ('    - {0}' -f $name)
            }
        }
    }

    Write-Host ''
}

function Write-Info([string]$msg) {
    Write-Host $msg
}

function Normalize-DateString([datetime]$dt) {
    return $dt.ToString('yyyy-MM-dd HH:mm:ss')
}

function Get-SafeFolderName([string]$text, [int]$MaxLength = 120) {
    if ([string]::IsNullOrWhiteSpace($text)) {
        return '_'
    }

    $result = $text

    foreach ($c in ([System.IO.Path]::GetInvalidFileNameChars() + [System.IO.Path]::GetInvalidPathChars())) {
        $result = $result.Replace([string]$c, '_')
    }

    $result = $result -replace '[\\\/:\*\?"<>\|]', '_'
    $result = $result -replace '[\x00-\x1F]', '_'
    $result = $result.Trim()
    $result = $result.TrimEnd([char[]]@('.', ' '))

    if ([string]::IsNullOrWhiteSpace($result)) {
        $result = '_'
    }

    if ($result -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') {
        $result = '_' + $result
    }

    if ($MaxLength -gt 0 -and $result.Length -gt $MaxLength) {
        $result = $result.Substring(0, $MaxLength)
        $result = $result.TrimEnd([char[]]@('.', ' '))
        if ([string]::IsNullOrWhiteSpace($result)) {
            $result = '_'
        }
    }

    return $result
}

function Get-ShortHash([string]$text) {
    if ($null -eq $text) {
        $text = ''
    }

    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
        $hashBytes = $sha1.ComputeHash($bytes)
        return (([System.BitConverter]::ToString($hashBytes) -replace '-', '').Substring(0, 10))
    }
    finally {
        if ($sha1 -ne $null) {
            $sha1.Dispose()
        }
    }
}

function Get-NormalizedFullPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }

    $full = [System.IO.Path]::GetFullPath($Path)
    return $full.TrimEnd([char[]]@('\', '/'))
}

function Test-IsSameOrChildPath([string]$Path, [string]$BasePath) {
    try {
        $candidate = Get-NormalizedFullPath $Path
        $base = Get-NormalizedFullPath $BasePath

        if ([string]::IsNullOrWhiteSpace($candidate) -or [string]::IsNullOrWhiteSpace($base)) {
            return $false
        }

        if ([string]::Equals($candidate, $base, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }

        return $candidate.StartsWith(($base + [System.IO.Path]::DirectorySeparatorChar), [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch {
        return $false
    }
}

function Test-IsExcludedPath([string]$CandidatePath, [string[]]$ExcludeRoots, [string[]]$ExcludeFiles) {
    $candidateFull = ''
    try {
        $candidateFull = [System.IO.Path]::GetFullPath($CandidatePath)
    }
    catch {
        return $false
    }

    foreach ($root in $ExcludeRoots) {
        if (-not [string]::IsNullOrWhiteSpace($root) -and (Test-IsSameOrChildPath -Path $candidateFull -BasePath $root)) {
            return $true
        }
    }

    foreach ($filePath in $ExcludeFiles) {
        if (-not [string]::IsNullOrWhiteSpace($filePath)) {
            try {
                if ([string]::Equals($candidateFull, [System.IO.Path]::GetFullPath($filePath), [System.StringComparison]::OrdinalIgnoreCase)) {
                    return $true
                }
            }
            catch {
            }
        }
    }

    return $false
}

function New-UniqueOutputFilePath([string]$Directory, [string]$PreferredFileName, [string]$IdentityKey, [int]$MaxPathLength = 235) {
    New-Item -ItemType Directory -Force -Path $Directory | Out-Null

    $preferredSafe = Normalize-InputText $PreferredFileName
    if ([string]::IsNullOrWhiteSpace($preferredSafe)) {
        $preferredSafe = '_'
    }

    $preferredSafe = $preferredSafe -replace '[\\\/:\*\?"<>\|]', '_'
    $preferredSafe = $preferredSafe.Trim()
    $preferredSafe = $preferredSafe.TrimEnd([char[]]@('.', ' '))
    if ([string]::IsNullOrWhiteSpace($preferredSafe)) {
        $preferredSafe = '_'
    }

    $directPath = Join-Path $Directory $preferredSafe
    if ($directPath.Length -le $MaxPathLength -and -not (Test-Path -LiteralPath $directPath)) {
        return $directPath
    }

    $ext = [System.IO.Path]::GetExtension($preferredSafe)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($preferredSafe)
    $baseName = Get-SafeFolderName $baseName 80
    $hash = Get-ShortHash $IdentityKey
    $seq = 1

    while ($true) {
        $seqText = if ($seq -eq 1) { '' } else { '__{0:D2}' -f $seq }
        $tail = '__{0}{1}{2}' -f $hash, $seqText, $ext
        $minPossibleLength = $Directory.Length + 1 + 8 + $tail.Length

        if ($minPossibleLength -gt $MaxPathLength) {
            throw ('出力先フォルダのパスが長すぎます。Directory=[{0}] Length={1}' -f $Directory, $Directory.Length)
        }

        $allowedBaseLength = $MaxPathLength - $Directory.Length - 1 - $tail.Length
        if ($allowedBaseLength -lt 8) {
            $allowedBaseLength = 8
        }

        $trimmedBase = if ($baseName.Length -gt $allowedBaseLength) { $baseName.Substring(0, $allowedBaseLength) } else { $baseName }
        $candidateName = '{0}{1}' -f $trimmedBase, $tail
        $candidatePath = Join-Path $Directory $candidateName

        if (-not (Test-Path -LiteralPath $candidatePath)) {
            return $candidatePath
        }

        $seq++
    }
}

function Copy-ItemWithDiagnostic([string]$SourcePath, [string]$DestinationPath) {
    try {
        $destParent = Split-Path -Parent $DestinationPath
        if (-not [string]::IsNullOrWhiteSpace($destParent)) {
            New-Item -ItemType Directory -Force -Path $destParent | Out-Null
        }

        Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    }
    catch {
        $destParent = Split-Path -Parent $DestinationPath
        throw ("Copy-Item に失敗しました.`n  SourceExists={0}`n  SourceLength={1}`n  Source={2}`n  DestParentExists={3}`n  DestLength={4}`n  Dest={5}`n  OriginalError={6}" -f `
            (Test-Path -LiteralPath $SourcePath -PathType Leaf), `
            $SourcePath.Length, `
            $SourcePath, `
            (Test-Path -LiteralPath $destParent -PathType Container), `
            $DestinationPath.Length, `
            $DestinationPath, `
            $_.Exception.Message)
    }
}

function Move-ItemWithDiagnostic([string]$SourcePath, [string]$DestinationPath) {
    try {
        $destParent = Split-Path -Parent $DestinationPath
        if (-not [string]::IsNullOrWhiteSpace($destParent)) {
            New-Item -ItemType Directory -Force -Path $destParent | Out-Null
        }

        Move-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    }
    catch {
        $destParent = Split-Path -Parent $DestinationPath
        throw ("Move-Item に失敗しました.`n  SourceExists={0}`n  SourceLength={1}`n  Source={2}`n  DestParentExists={3}`n  DestLength={4}`n  Dest={5}`n  OriginalError={6}" -f `
            (Test-Path -LiteralPath $SourcePath -PathType Leaf), `
            $SourcePath.Length, `
            $SourcePath, `
            (Test-Path -LiteralPath $destParent -PathType Container), `
            $DestinationPath.Length, `
            $DestinationPath, `
            $_.Exception.Message)
    }
}

function Store-DuplicateItemWithDiagnostic([string]$SourcePath, [string]$DestinationPath, [ValidateSet('Copy', 'Move')] [string]$Action) {
    if ($Action -eq 'Move') {
        Move-ItemWithDiagnostic -SourcePath $SourcePath -DestinationPath $DestinationPath
    }
    else {
        Copy-ItemWithDiagnostic -SourcePath $SourcePath -DestinationPath $DestinationPath
    }
}

function Get-RelativePathSafe([string]$BasePath, [string]$TargetPath) {
    try {
        $base = [System.IO.Path]::GetFullPath($BasePath)
        if (-not $base.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
            $base += [System.IO.Path]::DirectorySeparatorChar
        }

        $baseUri = New-Object System.Uri($base)
        $targetUri = New-Object System.Uri([System.IO.Path]::GetFullPath($TargetPath))
        $relativeUri = $baseUri.MakeRelativeUri($targetUri)
        $relative = [System.Uri]::UnescapeDataString($relativeUri.ToString())

        return ($relative -replace '/', '\')
    }
    catch {
        return ''
    }
}

function Get-FileCreateTime([System.IO.FileInfo]$File) {
    try {
        return $File.CreationTime
    }
    catch {
        return $File.LastWriteTime
    }
}

Write-Host '============================================='
Write-Host 'Index Pickup Script v5.0.0'
Write-Host '============================================='
Write-Host ''

# ------------------------------------------------------------
# 入力チェック
# ------------------------------------------------------------
$Source = Normalize-InputText $Source
$Since  = Normalize-InputText $Since

$sourceFull = Resolve-ValidatedPath -pathText $Source -PathType Container
if ($null -eq $sourceFull) {
    Write-PathDiagnostic -label 'SOURCE' -rawValue $Source -PathType Container
    throw ('SOURCEフォルダが存在しません: {0}' -f $Source)
}


try {
    $sinceDate = [datetime]::ParseExact($Since, 'yyyy-MM-dd HH:mm', $null)
}
catch {
    throw ('更新日時の形式が不正です。入力値=[{0}] 例: 2026-03-01 00:00' -f $Since)
}

Write-Info ('Source : {0}' -f $sourceFull)
Write-Info ('Since  : {0}' -f (Normalize-DateString $sinceDate))
Write-Info 'FixedFields : ファイル名（拡張子付き）, 作成日時, 更新日時, フルパス, 概要'
Write-Info 'DuplicateHandling : LatestOnly (older same-name files are ignored)'
Write-Host ''

# ------------------------------------------------------------
# 出力フォルダ
# ------------------------------------------------------------
$runTime = Get-Date -Format 'yyyyMMdd_HHmmss'
$baseDir = Join-Path $PSScriptRoot ('output\IndexRun_{0}' -f $runTime)
$pickupDir = Join-Path $baseDir 'pickup_files'

New-Item -ItemType Directory -Force -Path $baseDir | Out-Null
New-Item -ItemType Directory -Force -Path $pickupDir | Out-Null

$excludeRoots = New-Object System.Collections.Generic.List[string]
$excludeFiles = New-Object System.Collections.Generic.List[string]

$scriptRootFull = [System.IO.Path]::GetFullPath($PSScriptRoot)

if ((Test-IsSameOrChildPath -Path $scriptRootFull -BasePath $sourceFull) -and
    -not [string]::Equals((Get-NormalizedFullPath $scriptRootFull), (Get-NormalizedFullPath $sourceFull), [System.StringComparison]::OrdinalIgnoreCase)) {
    [void]$excludeRoots.Add($scriptRootFull)
}

if (Test-IsSameOrChildPath -Path $baseDir -BasePath $sourceFull) {
    [void]$excludeRoots.Add($baseDir)
}

foreach ($candidateFile in @(
    (Join-Path $PSScriptRoot 'IndexPickup_v4_5.ps1'),
    (Join-Path $PSScriptRoot 'IndexPickup_v4_5_2.ps1'),
    (Join-Path $PSScriptRoot 'IndexPickup_v4_5_3.ps1'),
    (Join-Path $PSScriptRoot 'IndexPickup_v4_5_4.ps1'),
    (Join-Path $PSScriptRoot 'IndexPickup_v4_6_0.ps1'),
    (Join-Path $PSScriptRoot 'IndexPickup_v4_6_2.ps1'),
    (Join-Path $PSScriptRoot 'IndexPickup_v5_0_0.ps1'),
    (Join-Path $PSScriptRoot 'Start_IndexPickup_v4_5.bat'),
    (Join-Path $PSScriptRoot 'Start_IndexPickup_v4_5_2.bat'),
    (Join-Path $PSScriptRoot 'Start_IndexPickup_v4_5_3.bat'),
    (Join-Path $PSScriptRoot 'Start_IndexPickup_v4_5_4.bat'),
    (Join-Path $PSScriptRoot 'Start_IndexPickup_v4_6_0.bat'),
    (Join-Path $PSScriptRoot 'Start_IndexPickup_v4_6_1_fixed.bat'),
    (Join-Path $PSScriptRoot 'Start_IndexPickup_v4_6_2.bat'),
    (Join-Path $PSScriptRoot 'Start_IndexPickup_v5_0_0.bat')
)) {
    if (-not [string]::IsNullOrWhiteSpace($candidateFile) -and (Test-IsSameOrChildPath -Path $candidateFile -BasePath $sourceFull)) {
        [void]$excludeFiles.Add([System.IO.Path]::GetFullPath($candidateFile))
    }
}

# ------------------------------------------------------------
# 固定項目使用（公式インデックスは参照しない）
# ------------------------------------------------------------
Write-Info '公式インデックスは参照しません。固定項目で処理を継続します。'
$registeredLastNum = 0
$registeredLastID = 'NONE'

if ($excludeRoots.Count -gt 0 -or $excludeFiles.Count -gt 0) {
    Write-Host ''
    Write-Info 'SOURCE 配下の制御ファイル／出力フォルダは自動除外します。'

    foreach ($root in ($excludeRoots | Sort-Object -Unique)) {
        Write-Info ('除外フォルダ       : {0}' -f $root)
    }

    foreach ($filePath in ($excludeFiles | Sort-Object -Unique)) {
        Write-Info ('除外ファイル       : {0}' -f $filePath)
    }
}

Write-Host ''

# ------------------------------------------------------------
# ファイル抽出
# ------------------------------------------------------------
Write-Info '更新日時条件に一致するファイルを抽出中...'

$allFiles = Get-ChildItem -LiteralPath $sourceFull -File -Recurse |
    Where-Object {
        $_.LastWriteTime -ge $sinceDate -and
        -not (Test-IsExcludedPath -CandidatePath $_.FullName -ExcludeRoots $excludeRoots.ToArray() -ExcludeFiles $excludeFiles.ToArray())
    }

Write-Info ('抽出件数（重複整理前） : {0}' -f $allFiles.Count)

# ------------------------------------------------------------
# 同名ファイルの代表選択
# 更新日時が最も新しい 1 件のみ採用
# 同一更新日時の同名ファイルは追加の tie-break を行わず、最初に見つかったものを採用
# ------------------------------------------------------------
$grouped = $allFiles | Group-Object -Property Name

$manifestRows = New-Object System.Collections.Generic.List[object]
$copyLogRows = New-Object System.Collections.Generic.List[object]

$currentIDNum = $registeredLastNum
$selectedFileNamesInOrder = New-Object System.Collections.Generic.List[string]

foreach ($group in ($grouped | Sort-Object Name)) {

    $selected = $null
    foreach ($candidate in $group.Group) {
        if ($null -eq $selected) {
            $selected = $candidate
            continue
        }

        if ($candidate.LastWriteTime -gt $selected.LastWriteTime) {
            $selected = $candidate
        }
    }

    if ($null -eq $selected) {
        continue
    }

    $duplicateGroupName = $group.Name
    $sameLatestCount = ($group.Group | Where-Object { $_.LastWriteTime -eq $selected.LastWriteTime }).Count

    $selectionReason = ''
    if ($group.Count -eq 1) {
        $selectionReason = 'Only file in duplicate group. Selected.'
    }
    elseif ($sameLatestCount -gt 1) {
        $selectionReason = 'Selected by LastWriteTime descending. Same latest timestamp ties are not further resolved.'
    }
    else {
        $selectionReason = 'Selected by LastWriteTime descending.'
    }

    $destPath = Join-Path $pickupDir $selected.Name
    Copy-ItemWithDiagnostic -SourcePath $selected.FullName -DestinationPath $destPath

    $currentIDNum++
    $newID = 'IDX-{0:D6}' -f $currentIDNum
    [void]$selectedFileNamesInOrder.Add([System.IO.Path]::GetFileName($destPath))

    $manifestRows.Add([PSCustomObject]@{
        FileName        = $selected.Name
        CreateTime      = (Normalize-DateString (Get-FileCreateTime $selected))
        LastWriteTime   = (Normalize-DateString $selected.LastWriteTime)
        Size            = $selected.Length
        FullPath        = $selected.FullName
        Summary         = ''
        RelativePath    = (Get-RelativePathSafe -BasePath $sourceFull -TargetPath $selected.FullName)
        CopiedPath      = $destPath
        StoredPath      = $destPath
        StorageAction   = 'Copy'
        SelectionStatus = 'selected'
        DuplicateGroup  = $duplicateGroupName
        SelectionRank   = 1
        SelectionReason = $selectionReason
    }) | Out-Null

    $copyLogRows.Add([PSCustomObject]@{
        ID              = $newID
        FileName        = $selected.Name
        SourcePath      = $selected.FullName
        CopiedPath      = $destPath
        StoredPath      = $destPath
        StorageAction   = 'Copy'
        SelectionStatus = 'selected'
        DuplicateGroup  = $duplicateGroupName
        SelectionRank   = 1
        LastWriteTime   = (Normalize-DateString $selected.LastWriteTime)
        Size            = $selected.Length
    }) | Out-Null
}

# ------------------------------------------------------------
# selected の採番レンジ
# ------------------------------------------------------------
$selectedCount = $manifestRows.Count
$startID = if ($selectedCount -gt 0) { 'IDX-{0:D6}' -f ($registeredLastNum + 1) } else { 'NONE' }
$endID   = if ($selectedCount -gt 0) { 'IDX-{0:D6}' -f ($registeredLastNum + $selectedCount) } else { 'NONE' }

$outputLimit = $selectedCount

$registeredLastIDText = $registeredLastID
$startIDText          = $startID
$endIDText            = $endID

# ------------------------------------------------------------
# CSV出力
# ------------------------------------------------------------
$manifestPath = Join-Path $baseDir 'pickup_manifest.csv'
$copyLogPath = Join-Path $baseDir 'copy_log.csv'
$runCardPath = Join-Path $baseDir 'RUN_card_auto_v5_0_0.txt'

$manifestRows |
    Select-Object FileName, CreateTime, LastWriteTime, FullPath, Summary, Size, RelativePath, CopiedPath, StoredPath, StorageAction, SelectionStatus, DuplicateGroup, SelectionRank, SelectionReason |
    Export-Csv -LiteralPath $manifestPath -Encoding UTF8 -NoTypeInformation

$copyLogRows |
    Export-Csv -LiteralPath $copyLogPath -Encoding UTF8 -NoTypeInformation

# ------------------------------------------------------------
# RUNカード出力（厳密フォーマット）
# ------------------------------------------------------------
$runCard = New-Object System.Collections.Generic.List[string]
$runCard.Add('RUN_CARD_VERSION=1') | Out-Null
$runCard.Add(('GENERATED_AT={0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))) | Out-Null
$runCard.Add(('SOURCE={0}' -f $sourceFull)) | Out-Null
$runCard.Add(('SINCE={0}' -f (Normalize-DateString $sinceDate))) | Out-Null
$runCard.Add('INDEX=NONE') | Out-Null
$runCard.Add('DUPLICATE_ACTION=LatestOnly') | Out-Null
$runCard.Add(('REGISTERED_LAST_ID={0}' -f $registeredLastIDText)) | Out-Null
$runCard.Add(('OUTPUT_LIMIT={0}' -f $outputLimit)) | Out-Null
$runCard.Add(('SELECTED_COUNT={0}' -f $selectedCount)) | Out-Null
$runCard.Add('DUPLICATE_OLD_COUNT=0') | Out-Null
$runCard.Add(('START_ID={0}' -f $startIDText)) | Out-Null
$runCard.Add(('END_ID={0}' -f $endIDText)) | Out-Null
$runCard.Add('[FILES]') | Out-Null

foreach ($name in $selectedFileNamesInOrder) {
    $runCard.Add($name) | Out-Null
}

$runCard.Add('[/FILES]') | Out-Null
$runCard | Set-Content -LiteralPath $runCardPath -Encoding UTF8

# ------------------------------------------------------------
# 完了表示
# ------------------------------------------------------------
Write-Host ''
Write-Host 'DONE'
Write-Host ('OUTPUT_FOLDER={0}' -f $baseDir)
Write-Host ('MANIFEST_FILE={0}' -f $manifestPath)
Write-Host ('RUN_CARD_FILE={0}' -f $runCardPath)
Write-Host ('PICKUP_DIR={0}' -f $pickupDir)
Write-Host ('SELECTED_COUNT={0}' -f $selectedCount)
Write-Host ('REGISTERED_LAST_ID={0}' -f $registeredLastIDText)
Write-Host ('START_ID={0}' -f $startIDText)
Write-Host ('END_ID={0}' -f $endIDText)
Write-Host ''
Write-Host 'CHATGPT_STAGE_1'
Write-Host '1) Upload pickup_manifest.csv.'
Write-Host '2) Upload RUN card.'
Write-Host 'Then wait for the prompt instruction.'
Write-Host 'After that, upload all files listed in RUN card [FILES].'
Write-Host ''
