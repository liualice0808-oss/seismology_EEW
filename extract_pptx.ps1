<#
.SYNOPSIS
    自動遍歷並提取當前目錄下的 PPTX 檔案文字與圖片，並輸出結構化 JSON。
.DESCRIPTION
    本腳本會搜尋當前目錄下的所有簡報檔案 (.pptx)，解開 OpenXML 結構提取投影片文字與備忘錄，
    並自動解析各投影片的圖資關聯，將投影片圖片提取至 .\gradio_lite_app\assets\ 目錄，
    最後編譯出一個結構化的 slides.json 檔案。
.USAGE
    在 PowerShell 中執行：
    powershell -ExecutionPolicy Bypass -File .\extract_pptx.ps1
#>

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$currentDir = Get-Location
$destDir = "$currentDir\gradio_lite_app"
$assetsDir = "$destDir\assets"

if (-not (Test-Path $assetsDir)) {
    New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null
}

# 搜尋當前目錄下的所有 PPTX
$files = Get-ChildItem -Path $currentDir -Filter "*.pptx" | Select-Object -ExpandProperty Name

if ($files.Count -eq 0) {
    Write-Host "⚠️ 在當前目錄中未找到任何 .pptx 檔案！請確保簡報與此腳本放在同一個目錄中。" -ForegroundColor Yellow
    Exit
}

$slidesList = @()

foreach ($file in $files) {
    $pptxPath = "$currentDir\$file"
    $docTitle = $file -replace '\.pptx$', ''
    # 清理檔案名稱中的特殊字元，避免路徑錯誤
    $sanitizedDoc = $docTitle -replace '[^a-zA-Z0-9_\u4e00-\u9fa5]', '_'
    
    Write-Host "🔍 正在提取簡報的文字與圖片：$file" -ForegroundColor Cyan
    
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($pptxPath)
        
        # 1. 讀取所有講者備忘錄
        $notesContent = @{}
        $notesEntries = $zip.Entries | Where-Object { $_.FullName -like "ppt/notesSlides/notesSlide*.xml" }
        foreach ($entry in $notesEntries) {
            $numStr = $entry.Name -replace 'notesSlide', '' -replace '\.xml', ''
            if ($numStr -and $numStr -match '^\d+$') {
                $slideNum = [int]$numStr
                
                $stream = $entry.Open()
                $reader = New-Object System.IO.StreamReader($stream)
                $xmlText = $reader.ReadToEnd()
                $reader.Close()
                $stream.Close()
                
                $matches = [regex]::Matches($xmlText, "<a:t[^>]*>(.*?)</a:t>")
                $text = ($matches | ForEach-Object { $_.Groups[1].Value.Trim() } | Where-Object { $_ -ne "" }) -join " "
                $notesContent[$slideNum] = $text
            }
        }
        
        # 2. 讀取所有投影片
        $slideEntries = $zip.Entries | Where-Object { $_.FullName -like "ppt/slides/slide*.xml" } | Sort-Object {
            $numStr = $_.Name -replace 'slide', '' -replace '\.xml', ''
            [int]$numStr
        }
        
        foreach ($slideEntry in $slideEntries) {
            $numStr = $slideEntry.Name -replace 'slide', '' -replace '\.xml', ''
            $slideNum = [int]$numStr
            
            $stream = $slideEntry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $xmlText = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
            
            $matches = [regex]::Matches($xmlText, "<a:t[^>]*>(.*?)</a:t>")
            $slideText = ($matches | ForEach-Object { $_.Groups[1].Value.Trim() } | Where-Object { $_ -ne "" }) -join " "
            
            # 3. 讀取投影片關係檔以尋找圖片 (Relationship Mapping)
            $relPath = "ppt/slides/_rels/slide$($slideNum).xml.rels"
            $relEntry = $zip.Entries | Where-Object { $_.FullName -eq $relPath }
            
            $slideImages = @()
            if ($relEntry) {
                $relStream = $relEntry.Open()
                $relReader = New-Object System.IO.StreamReader($relStream)
                $relXml = $relReader.ReadToEnd()
                $relReader.Close()
                $relStream.Close()
                
                # 正則表達式抓取媒體目標路徑
                $imgMatches = [regex]::Matches($relXml, 'Target="\.\./media/([^"]+)"')
                foreach ($imgMatch in $imgMatches) {
                    $mediaName = $imgMatch.Groups[1].Value
                    $mediaFullPath = "ppt/media/$mediaName"
                    $mediaEntry = $zip.Entries | Where-Object { $_.FullName -eq $mediaFullPath }
                    
                    if ($mediaEntry) {
                        # 生成在 assets 資料夾中的唯一檔名，防止不同簡報圖片衝突
                        $outImgName = "${sanitizedDoc}_slide${slideNum}_${mediaName}"
                        $outImgPath = "$assetsDir\$outImgName"
                        
                        # 實體提取檔案至 assets 目錄
                        if (-not (Test-Path $outImgPath)) {
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($mediaEntry, $outImgPath, $true)
                        }
                        
                        # 保存相對路徑
                        $slideImages += "assets/$outImgName"
                    }
                }
            }
            
            $nText = ""
            if ($notesContent.ContainsKey($slideNum)) {
                $nText = $notesContent[$slideNum]
            }
            
            $slidesList += [PSCustomObject]@{
                doc_title = $docTitle
                slide_id = "slide$slideNum"
                slide_num = $slideNum
                content = $slideText
                notes = $nText
                images = $slideImages
            }
        }
        
        $zip.Dispose()
    } catch {
        Write-Host "❌ 處理 $file 時發生錯誤: $_" -ForegroundColor Red
    }
}

# 4. 輸出編譯後的結構化 slides.json 檔案
$json = $slidesList | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText("$destDir\slides.json", $json, [System.Text.Encoding]::UTF8)

Write-Host "✅ 成功提取完成！" -ForegroundColor Green
Write-Host "📂 投影片圖片已存至：$assetsDir" -ForegroundColor Green
Write-Host "📄 知識庫已輸出為：$destDir\slides.json" -ForegroundColor Green
Write-Host "🚀 總共處理了 $($slidesList.Count) 頁投影片。" -ForegroundColor Green
