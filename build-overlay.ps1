param(
    [string]$Tag = "ppt-word-gen:0.8.0",
    [string]$WslDistro = "Ubuntu"
)

$shellMajor = $PSVersionTable.PSVersion.Major
if ($shellMajor -lt 7) {
    $pwsh = (Get-Command pwsh.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    & $pwsh -NoProfile -File $PSCommandPath -Tag $Tag -WslDistro $WslDistro
    exit $LASTEXITCODE
}

$ErrorActionPreference = "Stop"
$pythonExe = (Get-Command python.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source

$buildContext = (& wsl.exe -d $WslDistro -u root -- mktemp -d /root/ppt-build-XXXXXX).Trim()
if ($LASTEXITCODE -ne 0 -or $buildContext -notmatch '^/root/ppt-build-[A-Za-z0-9]+$') {
    throw "Failed to create or validate WSL build directory: $buildContext"
}

$sourceFiles = @(
    "requirements.txt",
    "static/demo.html",
    "assets/word_templates/cid629-joint-simulation-v1.5.docx",
    "skills/ai-simulation-report/SKILL.md",
    "skills/ai-simulation-report/agents/openai.yaml",
    "skills/ai-simulation-report/references/report-contract.md",
    "skills/ai-simulation-report/references/validation-rules.md",
    "Dockerfile"
)
$sourceFiles += Get-ChildItem (Join-Path $PSScriptRoot "ppt_word_gen") -File -Filter "*.py" | ForEach-Object {
    "ppt_word_gen/$($_.Name)"
}
$sourceFiles += Get-ChildItem (Join-Path $PSScriptRoot ".docker-wheels") -File | ForEach-Object {
    ".docker-wheels/$($_.Name)"
}
$sourceFiles += Get-ChildItem (Join-Path $PSScriptRoot ".word-wheels") -File | ForEach-Object {
    ".word-wheels/$($_.Name)"
}
$sourceFiles += Get-ChildItem (Join-Path $PSScriptRoot ".mcp-wheels") -File | ForEach-Object {
    ".mcp-wheels/$($_.Name)"
}
$transferFile = $null

try {
    $transferFile = New-TemporaryFile
    $transferWindowsPath = $transferFile.FullName.Replace("\", "/")
    $wslTransferPath = (& wsl.exe -d $WslDistro -- wslpath -a $transferWindowsPath).Trim()
    if ($LASTEXITCODE -ne 0 -or -not $wslTransferPath.StartsWith("/mnt/")) {
        throw "Failed to resolve transfer file in WSL: $wslTransferPath"
    }

    foreach ($relativePath in $sourceFiles) {
        $sourcePath = Join-Path $PSScriptRoot $relativePath
        $linuxPath = $relativePath.Replace("\", "/")
        $destination = "$buildContext/$linuxPath"
        $linuxParent = [IO.Path]::GetDirectoryName($linuxPath).Replace("\", "/")
        if ($linuxParent) {
            & wsl.exe -d $WslDistro -u root -- mkdir -p "$buildContext/$linuxParent"
            if ($LASTEXITCODE -ne 0) { throw "Failed to create build directory: $linuxParent" }
        }
        # 企业文件保护会让 .NET 读到 wheel 的封装字节；Python 文件流可复制解密后的真实内容。
        & $pythonExe (Join-Path $PSScriptRoot "copy_decrypted.py") $sourcePath $transferFile.FullName
        $transferFile.Refresh()
        if ($LASTEXITCODE -ne 0 -or $transferFile.Length -eq 0) {
            throw "Failed to read source file through Python: $relativePath"
        }
        & wsl.exe -d $WslDistro -u root -- cp -- "$wslTransferPath" "$destination"
        if ($LASTEXITCODE -ne 0) { throw "Failed to transfer source file: $relativePath" }
    }

    & wsl.exe -d $WslDistro -u root -- bash -lc "cd '$buildContext' && docker build --build-arg OFFLINE_INSTALL=1 -f Dockerfile -t '$Tag' ."
    if ($LASTEXITCODE -ne 0) { throw "Failed to build Docker image: $Tag" }

    Write-Host "[OK] Image built: $Tag"
}
finally {
    if ($buildContext -match '^/root/ppt-build-[A-Za-z0-9]+$') {
        & wsl.exe -d $WslDistro -u root -- rm -rf -- "$buildContext"
    }
    if ($null -ne $transferFile) {
        Remove-Item -LiteralPath $transferFile.FullName -Force -ErrorAction SilentlyContinue
    }
}
