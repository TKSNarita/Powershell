#�s���擾
$row = (Get-Content .\nettest.csv).Length

#�s���ϐ��ɑ��
$line = Import-Csv .\nettest.csv

for ($j=1; $j -lt 4; $j++){

    for ($i=0; $i -lt $row; $i++){

    $date = Get-Date -Format "yyyy-MMdd-HHmmss"
    $n1 = $line[$i].block + "_block_" + $j + "���"

    echo $n1 | Tee-Object -FilePath trace.log
    echo $date | Tee-Object -FilePath trace.log
    pathping -n $line[$i].ip | Tee-Object -FilePath trace.log

    echo "---------------------------------------------------------------------------------------" | Tee-Object -FilePath trace.log

    }

}

pause