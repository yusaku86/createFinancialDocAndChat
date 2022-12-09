VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProfit 
   Caption         =   "���㎑���쐬"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11625
   OleObjectBlob   =   "frmProfit.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'/**
 '* ���C���v���V�[�W��(�`���b�g���M�p���㎑���쐬)
'**/
Private Sub cmdEnter_Click()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '// ���͂��ꂽ�l�̃o���f�[�V����
    If validate = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    '// ���H����t�@�C�����K�؂Ȃ��̂��m�F
    If validateFile = False Then
        MsgBox "�w�肵���t�@�C�����K�؂ł͂���܂���B", vbExclamation, "���㎑���쐬"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Call createSheetIfNotExist("document_type", ThisWorkbook)
    
    '// �`���b�g���M���Ɏ����̎�ނ��Ж��E�Ώی��Ȃǂ𔻕ʂ��邽�߂ɃV�[�g�̒l��ύX
    With ThisWorkbook.Sheets("document_type")
        .Cells(1, 1).Value = "profit"
        .Cells(1, 2).Value = cmbCompany.Value
        .Cells(1, 3).Value = cmbProfitType.Value
        .Cells(1, 4).Value = txtYear.Value & "�N" & cmbMonth.Value
    End With
        
    ThisWorkbook.Sheets("document_type").Visible = False
    
    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value)
    
    '// �\���H�J�n(�s�v�ȗ���폜���A�ݕ��̂ݒ��o���č��v�̍s�쐬�E���ёւ�)
    Call ProcessChart
        
    '// �r���ݒ�ƃZ���̌����A�^�C�g���ݒ�
    Call RuleLine
    
    '// PDF�o��
    Dim pdfName As String
    pdfName = cmbCompany & cmbProfitType.Value & " �����ʔ���ꗗ�\.pdf"
    
    Call ExportPDF(pdfName)
        
    MsgBox "PDF�o�͂��������܂����B�f�X�N�g�b�v���m�F���Ă��������B", vbInformation, "���㎑���쐬"
    
    targetFile.Close False
    
    Set targetFile = Nothing
    Unload Me

End Sub

'// ���͂��ꂽ�l�̃o���f�[�V����
Private Function validate() As Boolean

    validate = False

    '// ��Ж������͂���Ă��邩
    If cmbCompany.Value = "" Then
        MsgBox "��Ж���I�����Ă��������B", vbQuestion, "���㎑���쐬"
        Exit Function
         
    '// �Ώ۔N�����͂���Ă��邩
    ElseIf txtYear.Value = "" Then
        MsgBox "�Ώ۔N����͂��Ă��������B", vbQuestion, "���㎑���쐬"
        Exit Function
        
    '// �Ώ۔N�ɐ��������͂���Ă��邩
    ElseIf IsNumeric(txtYear.Value) = False Then
        MsgBox "�Ώ۔N�ɂ͐�������͂��Ă��������B", vbQuestion, "���㎑���쐬"
        Exit Function
        
    '// �Ώی����I������Ă��邩
    ElseIf cmbMonth.Value = "" Then
        MsgBox "�Ώی���I�����Ă��������B", vbQuestion, "���㎑���쐬"
        Exit Function
    
    '// ���H�t�@�C���������͂���Ă��邩
    ElseIf txtFileName.Value = "" Then
        MsgBox "���H�t�@�C��������͂��Ă��������B"
        Exit Function
    End If
    
    validate = True

End Function

'// ���H����t�@�C���Ƃ��đI�����ꂽ�t�@�C�����K�؂Ȃ��̂��m�F
Private Function validateFile() As Boolean

    validateFile = False

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value, ReadOnly:=True)
    
    With targetFile.Sheets(1)
    
        If .Cells(1, 1).Value = "����" _
            And .Cells(1, 2).Value = "�R�[�h" _
            And .Cells(1, 3).Value = "�Ȗ�" _
            And .Cells(1, 4).Value = "�R�[�h" _
            And .Cells(1, 5).Value = "�⏕�Ȗ�" _
            And .Cells(1, 6).Value = "�R�[�h" _
            And .Cells(1, 7).Value = "�����" Then
            
            validateFile = True
        End If
    End With
    
    targetFile.Close False
    Set targetFile = Nothing

End Function

'/**
 '* ���[�U�[�t�H�[���N�����̏���
'**/
Private Sub UserForm_Initialize()
 
 '// �R���{�{�b�N�X�ɉ�Ж��ƌ���ǉ�
    With cmbCompany
        .AddItem "�R�݉^����"
        .AddItem "�R�݉^����YMܰ��"
        .AddItem "��YCL"
        .AddItem "���CYM��ݽ���l���c�Ə�"
        .AddItem "���CYM��ݽ���{�Љc�Ə�"
    End With

    Dim i As Long
    
    For i = 4 To 12
        cmbMonth.AddItem i & "��"
    Next
    
    For i = 1 To 3
        cmbMonth.AddItem i & "��"
    Next
    
    '�N�̃f�t�H���g�l��ݒ�
    txtYear.text = Year(Now)
    
    txtFileName.Locked = True
  
End Sub

'// �Q�Ƃ��������Ƃ��̏��� �� �_�C�A���O��\�����ĉ��H����t�@�C����I��
Private Sub cmdDialog_Click()

    Dim wsh As New WshShell

    Dim filename As String: filename = selectFile("���H����t�@�C����I�����Ă��������B", wsh.SpecialFolders(4) & "\", "Excel�t�@�C��", "*.csv;*.xlsx")
    
    Set wsh = Nothing
    
    If filename = "" Then
        Exit Sub
    End If
    
    txtFileName.Locked = False
    
    Dim fso As New FileSystemObject
    
    txtFileName.Value = fso.GetFileName(filename)
    Me.txtHiddenFileFullPath.Value = filename
    
    Set fso = Nothing

    txtFileName.Locked = True

End Sub

'// ��Ж����ύX���ꂽ���̏���
Private Sub cmbCompany_Change()

    cmbProfitType.Clear

    If cmbCompany.Value = "�R�݉^����" Then
        cmbProfitType.AddItem "�y�^������z"
        cmbProfitType.AddItem "�y�q�ɔ���z"
    
    ElseIf cmbCompany.Value = "�R�݉^����YMܰ��" Then
        cmbProfitType.AddItem "�y�C�������z"
        cmbProfitType.AddItem "�y�ԗ��̔������z"
        cmbProfitType.AddItem "�y���ި����Ď����z"
    
    ElseIf cmbCompany.Value = "��YCL" Then
        cmbProfitType.AddItem "�y�^�������E�ɓ���Ǝ����z"
        cmbProfitType.Value = "�y�^�������E�ɓ���Ǝ����z"
    
    ElseIf cmbCompany.Value = "���CYM��ݽ���l���c�Ə�" Then
        cmbProfitType.AddItem "�y�^������E�y������z"
        cmbProfitType.Value = "�y�^������E�y������z"
    
    ElseIf cmbCompany.Value = "���CYM��ݽ���{�Љc�Ə�" Then
        cmbProfitType.AddItem "�y�^������z"
        cmbProfitType.Value = "�y�^������z"
    End If
    
End Sub

'/**
 '* �\���H
'**/
Private Sub ProcessChart()
    
    '�s�v�ȗ���폜���A�ݕ��̂ݒ��o
    Range("A:B, D:E").Delete xlToRight
    
    With Cells(1, 1)
        .AutoFilter 4, "<>�ݕ�"
        .CurrentRegion.Offset(1).Delete xlUp
        .AutoFilter
    End With
    
    Columns(17).Cut
    Cells(1, 4).Select
    ActiveSheet.Paste
    
    '// ���z�����͂���Ă��Ȃ����̗�폜
    Dim lastcolumn As Long: lastcolumn = Cells(2, Columns.Count).End(xlToLeft).Column
    Range(Columns(lastcolumn + 1), Columns(Cells(1, Columns.Count).End(xlToLeft).Column + 1)).Delete xlToLeft
    
    '// ���v�s�쐬
    Rows(2).Insert xlDown
    Cells(2, 1).Value = "���v"
    
    Dim i As Integer
    
    For i = 4 To lastcolumn
        Cells(2, i).Value = Application.WorksheetFunction.Sum(Range(Cells(2, i), Cells(Cells(Rows.Count, 1).End(xlUp).Row, i)))
        Cells(2, i).NumberFormatLocal = "#,###"
    Next
    
    '���בւ�
    With ActiveSheet.Sort
        With .SortFields
            .Clear
            .Add Key:=Cells(3, 4), Order:=xlDescending
        End With
        
        .Header = xlNo
        .SetRange Range(Cells(3, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, lastcolumn))
        .Apply
    End With
    
    Range(Columns(1), Columns(lastcolumn)).EntireColumn.AutoFit

End Sub

'/**
 '* �r���������A�^�C�g���ݒ�
'**/
Private Sub RuleLine()
    
    Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    
    Range(Cells(3, 1), Cells(3, Columns.Count).End(xlToLeft)).Borders(xlEdgeTop).LineStyle = xlDouble
    Rows("1:2").Insert xlDown
    
    With Cells(1, 1)
        .Value = cmbCompany.Value & cmbProfitType.Value
        .Font.Bold = True
    End With
    
    With Cells(2, 1)
        .Value = SetPeriod
        .Font.Bold = True
    End With
    
    With Cells(1, 4)
        .Value = "�����ʔ���ꗗ�\(�P��:�~)"
        .Font.Bold = True
    End With
    
    With Range(Cells(1, 4), Cells(1, Columns.Count).End(xlToLeft))
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    With Range(Cells(4, 1), Cells(4, 3))
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    Range(Cells(3, 1), Cells(3, Columns.Count).End(xlToLeft)).HorizontalAlignment = xlCenter
    
End Sub

'/**
 '* ��Ђ��ƂɑΏۂ̊��Ԃ𔻒�
'**/
Private Function SetPeriod() As String
    
    '���CYM�̏ꍇ
    If InStr(1, cmbCompany.Value, "���CYM") > 0 Then
        If Val(Replace(cmbMonth.Value, "��", "")) < 4 Then
            SetPeriod = txtYear.Value - 1 & "�N6���`" & cmbMonth.Value
        Else
            SetPeriod = txtYear.Value & "�N6���`" & cmbMonth.Value
        End If
    '����ȊO
    Else
        If Val(Replace(cmbMonth.Value, "��", "")) < 4 Then
            SetPeriod = txtYear.Value - 1 & "�N4���`" & cmbMonth.Value
        Else
            SetPeriod = txtYear.Value & "�N4���`" & cmbMonth.Value
        End If
    End If
    
End Function

'/**
 '* PDF�o��
'**/
Private Sub ExportPDF(ByVal filename As String)

    '// ����ݒ�
    With ActiveSheet.PageSetup
        .Zoom = False
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .CenterHorizontally = True
    End With
    
    '// �f�X�N�g�b�v�p�X���擾����PDF�o��
    Dim wsh As New WshShell

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=wsh.SpecialFolders(4) & "\" & filename
    
    Set wsh = Nothing

End Sub

'/**
 '*�u����v�����������̏���
'**/

Private Sub cmdCancel_Click()
    Unload Me
End Sub
