# Excel操作　行・列の操作編


# Excelを操作する為の宣言
$excel = New-Object -ComObject Excel.Application

# 可視化する
$excel.Visible = $false

# 既存のExcelファイルを開く
$book = $excel.Workbooks.Open("C:\Users\tksna\OneDrive\デスクトップ\csvtest\test.xlsx")

# ワークシートを番号で指定し、接続する
$sheet = $excel.Worksheets.Item(1)

# シートの5行目に1行挿入する
$sheet.Rows.item(3).Insert()

# シートの3列目に1列挿入する
$sheet.Columns.item(2).Insert()

# 使用している行数を取得する
$ROW = $sheet.UsedRange.Rows.Count

# 使用している列数を取得する
$COL = $sheet.UsedRange.Columns.Count

# メッセージボックスで変数2つの内容を表示
Add-Type -Assembly System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("行数は $ROW です。`n列数は $COL です。", "結果")

# 上書き保存
$book.Save()

# Excelを閉じる
$excel.Quit()

# プロセスを解放する
$excel = $null
[GC]::Collect()