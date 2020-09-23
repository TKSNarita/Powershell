# Excel�𑀍�錾
$excel = New-Object -ComObject Excel.Application

# ��������
$excel.Visible = $False

# �Ώۃt�@�C����ϐ��Ɋi�[
$tmp = "C:\Users\tksna\OneDrive\�f�X�N�g�b�v\csvtest\test.xlsx"

# �ϊ��i�R���o�[�g�j��̃t�@�C������ϐ��Ɋi�[
$savePath = "C:\Users\tksna\OneDrive\�f�X�N�g�b�v\csvtest\finaltest.csv"

#�ϊ��i�R���o�[�g�j��̃t�@�C������CSV�ɐݒ�
# $path = (resolve-path -path $tmp).path
# $savePath = $tmp -replace ".xlsx",".csv" 

#������CSV�t�@�C��������΍폜����
Remove-Item $savePath 

# Excel�t�@�C�����J��
$book = $excel.Workbooks.Open($tmp)

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

# �v���Z�X���������
$excel = $null
[GC]::Collect()