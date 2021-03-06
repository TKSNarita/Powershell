#行数取得
$row = (Get-Content .\nettest.csv).Length

#行列を変数に代入
$line = Import-Csv .\nettest.csv

for ($j=1; $j -lt 4; $j++){

    for ($i=0; $i -lt $row; $i++){

    $date = Get-Date -Format "yyyy-MMdd-HHmmss"
    $n1 = $line[$i].block + "_block_" + $j + "回目"

    echo $n1 | Tee-Object -FilePath trace.log
    echo $date | Tee-Object -FilePath trace.log
    pathping -n $line[$i].ip | Tee-Object -FilePath trace.log

    echo "---------------------------------------------------------------------------------------" | Tee-Object -FilePath trace.log

    }

}

pause