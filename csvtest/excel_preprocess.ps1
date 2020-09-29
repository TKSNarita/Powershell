
# 日付を変数に格納
$logTime = (Get-Date).ToString("yyyy-MM-dd-hh-mm")

# ログファイル格納場所を変数に格納
$logpath = "C:\Users\tksna\OneDrive\デスクトップ\csvtest\log"

# ログファイル名を格納
$logname = $logpath + "\" + "log$logTime.log"

# ログ出力開始宣言
Start-Transcript $logname

echo "$logtime：ExcelのCSV変換を開始します"

# Excelを操作宣言
$excel = New-Object -ComObject Excel.Application

# 可視化する
$excel.Visible = $False

# 対象ファイルを変数に格納
$tmp = "C:\Users\tksna\OneDrive\デスクトップ\csvtest\test.xlsx"

# 対象ファイルを別名で保存する
# $tmp2 = "C:\Users\tksna\OneDrive\デスクトップ\cvtest2\test.xlsx"
# ファイルのコピーと、読み取り専用属性を無効に設定
# Copy-Item -Path $tmp -Destination $tmp2;
# Set-ItemProperty -Path $tmp2 -Name IsReadOnly -Value $false;


# 変換（コンバート）後のファイル名を変数に格納
$savePath = "C:\Users\tksna\OneDrive\デスクトップ\csvtest\finaltest.csv"

#変換（コンバート）後のファイル名をCSVに設定
# $path = (resolve-path -path $tmp).path
# $savePath = $tmp -replace ".xlsx",".csv" 

#同名のCSVファイルがあれば削除する
Remove-Item $savePath 

# Excelファイルを開く
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbooks.open
$book = $excel.Workbooks.Open($tmp, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing,$True)

# ワークシートを番号で指定し接続
$sheet = $excel.Worksheets.Item(1)

# シートの5行目に1行挿入する
# $sheet.Rows.item(5).Insert()

# シートの5列目に1列挿入する
# $sheet.Columns.item(1).Insert()

# シートの1列目を削除
$sheet.Rows.item(1).Delete()

# シートの1行目を削除
$sheet.Columns.item(1).Delete()

# シートのエリアを指定して削除
$sheet.Range("H1:M13").Delete()

# シートのエリアを指定して削除
$sheet.Range("A5:M13").Delete()

# 使用している行数を取得する
#$ROW = $sheet.UsedRange.Rows.Count

# 使用している列数を取得する
#$COL = $sheet.UsedRange.Columns.Count

# メッセージボックスで変数2つの内容を表示
#Add-Type -Assembly System.Windows.Forms
#[System.Windows.Forms.MessageBox]::Show("行数は $ROW です。`n列数は $COL です。", "結果")

#６はCSVファイルとして保存するコード
$book.SaveAs($savepath,6)

# 上書き保存
#$book.Save()

# Excelを閉じる
$excel.Quit()

echo "Excelを閉じました"

# プロセスを解放する
$excel = $null

echo "Excelのプロセス解放完了"

Stop-Transcript

[GC]::Collect()