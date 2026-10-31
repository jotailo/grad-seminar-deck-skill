param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [string]$OutDir = "",
    [int]$Width = 1680
)

# 用本机 PowerPoint 把 pptx 逐页导出为 PNG，供「目检」环节逐张查看。
# 用法：
#   powershell -NoProfile -ExecutionPolicy Bypass -File render_deck.ps1 -Pptx "D:\deck\组会汇报.pptx"
#   powershell -NoProfile -ExecutionPolicy Bypass -File render_deck.ps1 -Pptx "...pptx" -OutDir "D:\deck\render"
# 依赖：本机安装 Microsoft PowerPoint（Windows）。没有 PowerPoint 时脚本明确报错——
# 此时不要跳过目检，改为人工逐页看，并在交付时写明「未做渲染目检」。
#
# 为什么先把 pptx 复制到临时目录再导出：
# PowerPoint 的 COM 实例会按文件路径复用已打开的文档。用同一路径反复生成/导出时，
# 可能导出上一次的内容（看着没问题，其实已是旧版）。复制成唯一文件名可强制重新读取。

$ErrorActionPreference = "Stop"
$src = (Resolve-Path -LiteralPath $Pptx).Path
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = Join-Path (Split-Path -Parent $src) "render"
}
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
# 清掉上一次的 PNG，避免新旧混在一起看错页
Get-ChildItem -LiteralPath $OutDir -Filter *.png -File -ErrorAction SilentlyContinue | Remove-Item -Force

$stamp = Get-Date -Format "yyyyMMddHHmmssfff"
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("deckrender_" + $stamp)
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
$tmpPptx = Join-Path $tmpRoot ("deck_" + $stamp + ".pptx")
Copy-Item -LiteralPath $src -Destination $tmpPptx -Force
$tmpOut = Join-Path $tmpRoot "png"
New-Item -ItemType Directory -Path $tmpOut -Force | Out-Null

$app = $null
$pres = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
}
catch {
    Write-Error "未找到 PowerPoint（COM 不可用），无法渲染目检图。请人工打开逐页检查，或在装了 PowerPoint 的机器上重跑。"
    exit 3
}

try {
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $app.Presentations.Open($tmpPptx, $true, $false, $false)
    $h = [int]($Width * 9 / 16)
    $pres.Export($tmpOut, "PNG", $Width, $h)
    Get-ChildItem -LiteralPath $tmpOut -Filter *.png -File | Copy-Item -Destination $OutDir -Force
    $count = @(Get-ChildItem -LiteralPath $OutDir -Filter *.png -File).Count
    Write-Output "[ok] 已导出 $count 张到 $OutDir"
    Write-Output "[next] 逐张看图：内容是否被遮挡、图注是否完整、字是否清晰、有无溢出与空页"
}
catch {
    Write-Error "导出失败：$($_.Exception.Message)"
    exit 4
}
finally {
    if ($pres) { $pres.Close() | Out-Null }
    if ($app) { $app.Quit() | Out-Null }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    [GC]::Collect()
    if (Test-Path -LiteralPath $tmpRoot) { Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
