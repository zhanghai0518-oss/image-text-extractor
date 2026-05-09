# Windows OCR Script for Electron App
# Usage: powershell -ExecutionPolicy Bypass -File win-ocr.ps1 -ImagePath "path/to/image.png"

param(
  [Parameter(Mandatory=$true)]
  [string]$ImagePath
)

# 检查文件是否存在
if (-not (Test-Path $ImagePath)) {
  Write-Output "ERROR: 文件不存在"
  exit 1
}

try {
  # 加载必要的Windows Runtime类型
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  
  # 定义异步等待辅助函数
  $awaitTask = {
    param($task)
    $task.AsTask().Wait()
    return $task.AsTask().Result
  }
  
  # 获取Windows OCR引擎
  $ocrEngineType = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
  $ocrEngine = $ocrEngineType::TryCreateFromUserProfileLanguages()
  
  if ($ocrEngine -eq $null) {
    Write-Output "ERROR: OCR引擎不可用，请安装中文语言包"
    exit 1
  }
  
  # 读取图片文件
  $storageFileType = [Windows.Storage.StorageFile, Windows.Foundation, ContentType=WindowsRuntime]
  $getFileTask = $storageFileType::GetFileFromPathAsync($ImagePath)
  $file = Invoke-Command -ScriptBlock $awaitTask -ArgumentList $getFileTask
  
  # 打开图片流
  $openTask = $file.OpenAsync([Windows.Storage.FileAccessMode]::Read)
  $stream = Invoke-Command -ScriptBlock $awaitTask -ArgumentList $openTask
  
  # 解码图片
  $decoderType = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime]
  $decoderTask = $decoderType::CreateAsync($stream)
  $decoder = Invoke-Command -ScriptBlock $awaitTask -ArgumentList $decoderTask
  
  # 获取SoftwareBitmap
  $bitmapTask = $decoder.GetSoftwareBitmapAsync()
  $bitmap = Invoke-Command -ScriptBlock $awaitTask -ArgumentList $bitmapTask
  
  # 执行OCR
  $ocrTask = $ocrEngine.RecognizeAsync($bitmap)
  $result = Invoke-Command -ScriptBlock $awaitTask -ArgumentList $ocrTask
  
  # 输出结果
  Write-Output $result.Text
  
  # 关闭流
  $stream.Close()
  
} catch {
  Write-Output "ERROR: $($_.Exception.Message)"
  exit 1
}