$row = (Get-Content .\nettest.csv).Length

$line = Import-Csv .\nettest.csv

for ($i=0; $i -lt 3; $i++){
$date = Get-Date -Format "yyyy-MMdd-HHmmss"
$n1 = $line[$i+1].block + "block_1‰ñ–Ú"

echo $n1 >> .\trace.log
echo $date >> .\trace.log

$n2 = $line[$i+1].block + "block_2‰ñ–Ú"

echo $n2 >> .\trace.log
echo $date >> .\trace.log

$n3 = $line[$i+1].block + "block_3‰ñ–Ú"

echo $n3 >> .\trace.log
echo $date >> .\trace.log

}

pause