#�s���擾
$row = (Get-Content .\nettest.csv).Length

#�s���ϐ��ɑ��
$line = Import-Csv .\nettest.csv

for ($i=0; $i -lt $row; $i++){

$date = Get-Date -Format "yyyy-MMdd-HHmmss"
$n1 = $line[$i+1].block + "block_1���"

echo $n1 >> .\trace.log
echo $date >> .\trace.log
pathping -n $line[$i+1].ip >>trace.log

$date = Get-Date -Format "yyyy-MMdd-HHmmss"
$n2 = $line[$i+1].block + "block_2���"

echo $n2 >> .\trace.log
echo $date >> .\trace.log
pathping -n $line[$i+1].ip >>trace.log

$date = Get-Date -Format "yyyy-MMdd-HHmmss"
$n3 = $line[$i+1].block + "block_3���"

echo $n3 >> .\trace.log
echo $date >> .\trace.log
pathping -n $line[$i+1].ip >>trace.log

}

pause