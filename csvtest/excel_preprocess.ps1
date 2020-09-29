
# ���t��ϐ��Ɋi�[
$logTime = (Get-Date).ToString("yyyy-MM-dd-hh-mm")

# ���O�t�@�C���i�[�ꏊ��ϐ��Ɋi�[
$logpath = "C:\Users\tksna\OneDrive\�f�X�N�g�b�v\csvtest\log"

# ���O�t�@�C�������i�[
$logname = $logpath + "\" + "log$logTime.log"

# ���O�o�͊J�n�錾
Start-Transcript $logname

echo "$logtime�FExcel��CSV�ϊ����J�n���܂�"

# Excel�𑀍�錾
$excel = New-Object -ComObject Excel.Application

# ��������
$excel.Visible = $False

# �Ώۃt�@�C����ϐ��Ɋi�[
$tmp = "C:\Users\tksna\OneDrive\�f�X�N�g�b�v\csvtest\test.xlsx"

# �Ώۃt�@�C����ʖ��ŕۑ�����
# $tmp2 = "C:\Users\tksna\OneDrive\�f�X�N�g�b�v\cvtest2\test.xlsx"
# �t�@�C���̃R�s�[�ƁA�ǂݎ���p�����𖳌��ɐݒ�
# Copy-Item -Path $tmp -Destination $tmp2;
# Set-ItemProperty -Path $tmp2 -Name IsReadOnly -Value $false;


# �ϊ��i�R���o�[�g�j��̃t�@�C������ϐ��Ɋi�[
$savePath = "C:\Users\tksna\OneDrive\�f�X�N�g�b�v\csvtest\finaltest.csv"

#�ϊ��i�R���o�[�g�j��̃t�@�C������CSV�ɐݒ�
# $path = (resolve-path -path $tmp).path
# $savePath = $tmp -replace ".xlsx",".csv" 

#������CSV�t�@�C��������΍폜����
Remove-Item $savePath 

# Excel�t�@�C�����J��
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbooks.open
$book = $excel.Workbooks.Open($tmp, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing,$True)

# ���[�N�V�[�g��ԍ��Ŏw�肵�ڑ�
$sheet = $excel.Worksheets.Item(1)

# �V�[�g��5�s�ڂ�1�s�}������
# $sheet.Rows.item(5).Insert()

# �V�[�g��5��ڂ�1��}������
# $sheet.Columns.item(1).Insert()

# �V�[�g��1��ڂ��폜
$sheet.Rows.item(1).Delete()

# �V�[�g��1�s�ڂ��폜
$sheet.Columns.item(1).Delete()

# �V�[�g�̃G���A���w�肵�č폜
$sheet.Range("H1:M13").Delete()

# �V�[�g�̃G���A���w�肵�č폜
$sheet.Range("A5:M13").Delete()

# �g�p���Ă���s�����擾����
#$ROW = $sheet.UsedRange.Rows.Count

# �g�p���Ă���񐔂��擾����
#$COL = $sheet.UsedRange.Columns.Count

# ���b�Z�[�W�{�b�N�X�ŕϐ�2�̓��e��\��
#Add-Type -Assembly System.Windows.Forms
#[System.Windows.Forms.MessageBox]::Show("�s���� $ROW �ł��B`n�񐔂� $COL �ł��B", "����")

#�U��CSV�t�@�C���Ƃ��ĕۑ�����R�[�h
$book.SaveAs($savepath,6)

# �㏑���ۑ�
#$book.Save()

# Excel�����
$excel.Quit()

echo "Excel����܂���"

# �v���Z�X���������
$excel = $null

echo "Excel�̃v���Z�X�������"

Stop-Transcript

[GC]::Collect()