# ============================================================
#  吉客云货品档案抽查技能 · 一键安装脚本
#  仓库：https://github.com/237130014/merchandising-tools
#  用法：irm https://raw.githubusercontent.com/237130014/merchandising-tools/main/install.ps1 | iex
# ============================================================

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$RepoOwner = '237130014'
$RepoName  = 'merchandising-tools'
$SkillName = 'jackyun-goods-field-audit'
$Branch    = 'main'

# ---- 目标目录（可用环境变量 MT_SKILLS_DIR 覆盖，便于测试）----
if ($env:MT_SKILLS_DIR -and $env:MT_SKILLS_DIR.Trim() -ne '') {
    $SkillsRoot = $env:MT_SKILLS_DIR
} else {
    $SkillsRoot = Join-Path $env:USERPROFILE '.workbuddy\skills'
}
$Dest = Join-Path $SkillsRoot $SkillName

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$tmpRoot = Join-Path $env:TEMP ("jy-install-" + $stamp)
$zipPath = Join-Path $tmpRoot 'pkg.zip'

function Write-Step($msg) { Write-Host ("  -> " + $msg) -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host ("  OK  " + $msg) -ForegroundColor Green }
function Write-Warn2($msg){ Write-Host ("  !!  " + $msg) -ForegroundColor Yellow }

Write-Host ''
Write-Host '========================================' -ForegroundColor White
Write-Host ' 吉客云货品档案抽查技能 · 一键安装' -ForegroundColor White
Write-Host '========================================' -ForegroundColor White
Write-Host ''

try {
    # ---------- 1. 下载 ----------
    Write-Step "下载技能包 ($RepoOwner/$RepoName@$Branch)"
    New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
    $zipUrl = "https://github.com/$RepoOwner/$RepoName/archive/refs/heads/$Branch.zip"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    $sizeKb = [math]::Round((Get-Item $zipPath).Length / 1KB, 1)
    Write-Ok "已下载 $sizeKb KB"

    # ---------- 2. 解压 ----------
    Write-Step '解压'
    $expandDir = Join-Path $tmpRoot 'x'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $expandDir)
    Write-Ok '解压完成'

    # ---------- 3. 定位技能目录（找含 SKILL.md 的那个）----------
    Write-Step "定位 $SkillName"
    $candidates = Get-ChildItem -Path $expandDir -Recurse -Directory -Filter $SkillName -ErrorAction SilentlyContinue
    $srcDir = $null
    foreach ($c in $candidates) {
        if (Test-Path (Join-Path $c.FullName 'SKILL.md')) { $srcDir = $c.FullName; break }
    }
    if (-not $srcDir) { throw "压缩包里没找到有效的 $SkillName/SKILL.md" }
    Write-Ok "找到：$srcDir"

    # ---------- 4. 落位（旧版先备份）----------
    Write-Step "安装到 $Dest"
    if (-not (Test-Path $SkillsRoot)) {
        New-Item -ItemType Directory -Path $SkillsRoot -Force | Out-Null
        Write-Ok "已创建技能目录 $SkillsRoot"
    }
    if (Test-Path $Dest) {
        $bak = "$Dest.bak-$stamp"
        Move-Item -Path $Dest -Destination $bak -Force
        Write-Warn2 "检测到旧版本，已备份为 $(Split-Path $bak -Leaf)"
    }
    Copy-Item -Path $srcDir -Destination $Dest -Recurse -Force
    Write-Ok '文件已复制'

    # ---------- 5. 校验层级 ----------
    Write-Step '校验目录层级'
    $skillMd = Join-Path $Dest 'SKILL.md'
    if (-not (Test-Path $skillMd)) {
        throw "校验失败：$skillMd 不存在（可能多套了一层目录）"
    }
    $fileCount = (Get-ChildItem -Path $Dest -Recurse -File).Count
    Write-Ok "SKILL.md 就位，共 $fileCount 个文件"

    # ---------- 6. 检查 openpyxl ----------
    Write-Step '检查 Python 依赖 openpyxl'
    $pyExe = $null

    # 6.1 优先 WorkBuddy 托管 venv
    $venvCandidates = @(
        (Join-Path $env:USERPROFILE '.workbuddy\binaries\python\envs\default\Scripts\python.exe'),
        (Join-Path $env:USERPROFILE '.workbuddy\binaries\python\envs\default\bin\python')
    )
    foreach ($v in $venvCandidates) {
        if (-not $pyExe -and (Test-Path $v)) { $pyExe = $v }
    }

    # 6.2 托管 versions 目录里的 python.exe
    if (-not $pyExe) {
        $verRoot = Join-Path $env:USERPROFILE '.workbuddy\binaries\python\versions'
        if (Test-Path $verRoot) {
            $found = Get-ChildItem -Path $verRoot -Recurse -Filter 'python.exe' -ErrorAction SilentlyContinue |
                     Where-Object { $_.FullName -notmatch '\\Scripts\\' } |
                     Select-Object -First 1
            if ($found) { $pyExe = $found.FullName }
        }
    }

    # 6.3 系统 python
    if (-not $pyExe) {
        $cmd = Get-Command python -ErrorAction SilentlyContinue
        if ($cmd) { $pyExe = $cmd.Source }
    }

    if (-not $pyExe) {
        Write-Warn2 '没找到 Python，跳过依赖检查（技能首次运行时会自行处理）'
    } else {
        Write-Ok "Python：$pyExe"
        $probe = & $pyExe -c "import openpyxl,sys;print(openpyxl.__version__)" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Ok "openpyxl 已就绪（$probe）"
        } else {
            Write-Step '正在安装 openpyxl（首次约 10 秒）'
            & $pyExe -m pip install --quiet --disable-pip-version-check openpyxl 2>&1 | Out-Null
            $probe2 = & $pyExe -c "import openpyxl;print(openpyxl.__version__)" 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Ok "openpyxl 安装成功（$probe2）"
            } else {
                Write-Warn2 '自动安装失败，请稍后手动执行：pip install openpyxl'
            }
        }
    }

    # ---------- 7. 完成 ----------
    Write-Host ''
    Write-Host '========================================' -ForegroundColor Green
    Write-Host ' 安装成功' -ForegroundColor Green
    Write-Host '========================================' -ForegroundColor Green
    Write-Host ''
    Write-Host "  技能位置：$Dest"
    Write-Host ''
    Write-Host '  下一步：完全退出 WorkBuddy 再重新打开' -ForegroundColor Yellow
    Write-Host '         （右下角托盘图标也要退出，不是只关窗口）' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  重开后新建对话，把货品表拖进去，发：' -ForegroundColor White
    Write-Host '         用吉客云货品档案抽查技能检查这个表' -ForegroundColor White
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Host '========================================' -ForegroundColor Red
    Write-Host ' 安装失败' -ForegroundColor Red
    Write-Host '========================================' -ForegroundColor Red
    Write-Host ("  原因：" + $_.Exception.Message) -ForegroundColor Red
    Write-Host ''
    Write-Host '  常见处理：' -ForegroundColor Yellow
    Write-Host '   1) 网络不通 / 公司代理拦截 -> 换手机热点重试'
    Write-Host '   2) 目标目录被占用 -> 先退出 WorkBuddy 再装'
    Write-Host '   3) 仍失败 -> 手动下载 ZIP：'
    Write-Host "      https://github.com/$RepoOwner/$RepoName" -ForegroundColor Cyan
    Write-Host ''
}
finally {
    if (Test-Path $tmpRoot) {
        Remove-Item -Path $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
