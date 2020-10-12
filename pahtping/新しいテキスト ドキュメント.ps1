#行数取得
$row = (Get-Content .\nettest.csv).Length

#行列を変数に代入
$line = Import-Csv .\nettest.csv

for ($i=0; $i -lt $row; $i++){

$date = Get-Date -Format "yyyy-MMdd-HHmmss"
$n1 = $line[$i+1].block + "block_1回目"

echo $n1 >> .\trace.log
echo $date >> .\trace.log
pathping -n $line[$i+1].ip >>trace.log

$date = Get-Date -Format "yyyy-MMdd-HHmmss"
$n2 = $line[$i+1].block + "block_2回目"

echo $n2 >> .\trace.log
echo $date >> .\trace.log
pathping -n $line[$i+1].ip >>trace.log

$date = Get-Date -Format "yyyy-MMdd-HHmmss"
$n3 = $line[$i+1].block + "block_3回目"

echo $n3 >> .\trace.log
echo $date >> .\trace.log
pathping -n $line[$i+1].ip >>trace.log

}

pause