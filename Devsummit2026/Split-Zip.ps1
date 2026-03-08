# ZIPファイルを指定サイズごとに分割する
# 使い方:
# 1. $InFilePath を分割したいZIPのパスに変更
# 2. PowerShellでこのスクリプトを実行

$SplitSize = 24MB
$InFilePath = 'C:\DevConference\DevConference2026.zip'

if (-not (Test-Path -LiteralPath $InFilePath)) {
    Write-Error "対象ファイルが見つかりません: $InFilePath"
    exit 1
}

$Folder = Split-Path -Path $InFilePath -Parent
$FileName = Split-Path -Path $InFilePath -Leaf

$inStream = [System.IO.File]::OpenRead($InFilePath)

try {
    $partNumber = 1
    $buffer = New-Object byte[] $SplitSize

    while (($bytesRead = $inStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
        $outFilePath = "{0}.{1:D2}" -f $InFilePath, $partNumber
        $outStream = [System.IO.File]::Create($outFilePath)

        try {
            $outStream.Write($buffer, 0, $bytesRead)
        }
        finally {
            $outStream.Dispose()
        }

        Write-Host "作成: $outFilePath"
        $partNumber++
    }
}
finally {
    $inStream.Dispose()
}

$joinBatPath = Join-Path $Folder ($FileName + ".join.bat")
$joinBatContent = @"
@echo off
copy /b "$FileName.01" + "$FileName.02" "$FileName.restored.tmp"
if errorlevel 1 goto :error

set i=3
:loop
if exist "$FileName.0%i%" (
    copy /b "$FileName.restored.tmp" + "$FileName.0%i%" "$FileName.restored2.tmp" >nul
    move /y "$FileName.restored2.tmp" "$FileName.restored.tmp" >nul
    set /a i+=1
    goto loop
)

if exist "$FileName.%i%" (
    copy /b "$FileName.restored.tmp" + "$FileName.%i%" "$FileName.restored2.tmp" >nul
    move /y "$FileName.restored2.tmp" "$FileName.restored.tmp" >nul
    set /a i+=1
    goto loop
)

move /y "$FileName.restored.tmp" "$FileName" >nul
echo 復元完了: $FileName
goto :eof

:error
echo 復元に失敗しました
exit /b 1
"@

Set-Content -LiteralPath $joinBatPath -Value $joinBatContent -Encoding Default

Write-Host ""