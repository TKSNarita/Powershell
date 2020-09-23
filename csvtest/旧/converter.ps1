$tmp = "C:\Users\tksna\OneDrive\デスクトップ\csvtest\test.xlsx" #変換元のExcelデータ

$objExcel = New-Object -ComObject Excel.Application #オブジェクトを変数に設定

$path = (resolve-path -path $tmp).path

$savePath = $tmp -replace ".xlsx",".csv" #変換（コンバート）後のファイル名をCSVに設定

Remove-Item $savePath #既に同名のCSVファイルがあれば削除

Start-Sleep 1

$objworkbook=$objExcel.Workbooks.Open($tmp) #Excelデータをオブジェクト変数に設定

$objworkbook.SaveAs($savepath,6) #６はCSVファイルとして保存するコード

$objworkbook.Close($false) #ファイルを閉じる