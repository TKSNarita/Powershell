$tmp = "C:\Users\tksna\OneDrive\�f�X�N�g�b�v\csvtest\test.xlsx" #�ϊ�����Excel�f�[�^

$objExcel = New-Object -ComObject Excel.Application #�I�u�W�F�N�g��ϐ��ɐݒ�

$path = (resolve-path -path $tmp).path

$savePath = $tmp -replace ".xlsx",".csv" #�ϊ��i�R���o�[�g�j��̃t�@�C������CSV�ɐݒ�

Remove-Item $savePath #���ɓ�����CSV�t�@�C��������΍폜

Start-Sleep 1

$objworkbook=$objExcel.Workbooks.Open($tmp) #Excel�f�[�^���I�u�W�F�N�g�ϐ��ɐݒ�

$objworkbook.SaveAs($savepath,6) #�U��CSV�t�@�C���Ƃ��ĕۑ�����R�[�h

$objworkbook.Close($false) #�t�@�C�������