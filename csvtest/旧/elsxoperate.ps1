# Excel����@�s�E��̑����


# Excel�𑀍삷��ׂ̐錾
$excel = New-Object -ComObject Excel.Application

# ��������
$excel.Visible = $false

# ������Excel�t�@�C�����J��
$book = $excel.Workbooks.Open("C:\Users\tksna\OneDrive\�f�X�N�g�b�v\csvtest\test.xlsx")

# ���[�N�V�[�g��ԍ��Ŏw�肵�A�ڑ�����
$sheet = $excel.Worksheets.Item(1)

# �V�[�g��5�s�ڂ�1�s�}������
$sheet.Rows.item(3).Insert()

# �V�[�g��3��ڂ�1��}������
$sheet.Columns.item(2).Insert()

# �g�p���Ă���s�����擾����
$ROW = $sheet.UsedRange.Rows.Count

# �g�p���Ă���񐔂��擾����
$COL = $sheet.UsedRange.Columns.Count

# ���b�Z�[�W�{�b�N�X�ŕϐ�2�̓��e��\��
Add-Type -Assembly System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("�s���� $ROW �ł��B`n�񐔂� $COL �ł��B", "����")

# �㏑���ۑ�
$book.Save()

# Excel�����
$excel.Quit()

# �v���Z�X���������
$excel = $null
[GC]::Collect()