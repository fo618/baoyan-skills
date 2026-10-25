# Windows 内置 OCR（WinRT）—— 读取微信截图里的文字
#
# 为什么需要它：微信「置顶」状态不写进任何本地数据库（session.db 的 SessionTable
# 没有 top/pin 字段），只能从屏幕上读。而 AI 模型读不了 png，所以必须过一层 OCR。
# 本机没有 pip 能装成的 OCR（rapidocr-onnxruntime 在 venv 里 onnxruntime.dll 加载失败，
# winsdk 需要 Rust 编译），但 Windows 自带的 WinRT OCR 可以（中文简体 zh-Hans-CN 可用）。
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File ocr_screen.ps1 <图片绝对路径>
#   建议先把截图复制到纯 ASCII 路径（如 %TEMP%\wx.png）再传进来，避免中文路径编码问题。
#   输出：每行 "[y=像素 x=像素] 文字"，可直接按 y 排序还原版面。
#
# 配合 scripts/capture_wechat.py 使用：先截屏，再 OCR。

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]

function Await($WinRtTask, $ResultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}

$imgPath = $args[0]
Write-Output "IMG=$imgPath"

$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]

Write-Output "--- available OCR languages ---"
try {
    $langs = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
    foreach ($l in $langs) { Write-Output ("LANG: " + $l.LanguageTag + " / " + $l.DisplayName) }
} catch { Write-Output ("langs err: " + $_.Exception.Message) }

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($imgPath)) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

$engine = $null
foreach ($tag in @('zh-Hans-CN','zh-Hans','zh-CN')) {
    try {
        $lang = New-Object Windows.Globalization.Language($tag)
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
        if ($engine) { Write-Output "USING LANG $tag"; break }
    } catch {}
}
if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if (-not $engine) { Write-Output "NO OCR ENGINE"; exit 1 }

$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
Write-Output "--- OCR LINES ---"
foreach ($line in $result.Lines) {
    $ys = ($line.Words | ForEach-Object { $_.BoundingRect.Y } | Measure-Object -Minimum).Minimum
    $xs = ($line.Words | ForEach-Object { $_.BoundingRect.X } | Measure-Object -Minimum).Minimum
    Write-Output ("[y=" + [int]$ys + " x=" + [int]$xs + "] " + $line.Text)
}
