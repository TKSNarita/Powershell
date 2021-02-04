#現在の場所を記憶する
$path = (Convert-Path .)

#探索するカレントディレクトリを入力する
echo ""
$input = Read-Host "探索する対象ディレクトリのパスを入力してください。`r`nVドライブの場合「V:\」と入力してください"
$out = "探索対象のディレクトリは:" + $input
Write-Output $out

echo ""
echo "ただいまフォルダリスト作成中のためお待ちください"

#探索するカレントディレクトリに移動する
cd $input

#カレントディレクトリ配下（サブディレクトリ含む)を探索してdirlist.csvにエクスポート
Get-ChildItem * -Recurse | Where-Object { $_.PSIsContainer } | Select-Object Name, FullName, LastWriteTime | Export-Csv -Encoding UTF8 $path/dirlist.csv

#dirlist.csvの行数を取得してrowに代入
$row = (Get-Content .\dirlist.csv).Length

#dirlist.csvをリスト形式で変数dirlに代入
$dirl = Import-Csv .\dirlist.csv



#フォルダパスの更新日時が最新10個のファイルリスト抽出する
echo ""
echo "これからフォルダ内のファイルリスト取得を開始します"

for ($i=0; $i -lt $row-2; $i++){

    Get-ChildItem $dirl[$i].FullName -File -Recurse | Sort-Object -Property LastWriteTime | Select-Object -Last 3 | Select-Object Name,FullName,Length,LastWriteTime | Export-Csv -Encoding Default $path/filename_last10master.csv -append
    Get-ChildItem $dirl[$i].FullName -File -Recurse | Sort-Object -Property Length | Select-Object -Last 1 | Select-Object Name,FullName,Length,LastWriteTime | Export-Csv -Encoding Default $path/filename_volume1master.csv -append      

    echo ""
    $com = '【進捗】' + ($i+1) + '/' + ($row-2) + '　ファイルリスト取得中'
    echo $com
    }

echo ""
echo "ファイルリストfilename_volume1master.csv と filename_last10master.csv　のエクスポートが完了しました"