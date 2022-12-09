VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCost 
   Caption         =   "�o����쐬"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6420
   OleObjectBlob   =   "frmCost.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// �Ǘ���v�p�������쐬����t�H�[��
Option Explicit

'// �o����쐬(���C���v���V�[�W��)
Private Sub cmdEnter_Click()
    
    Application.ScreenUpdating = False
    
    '// ���͂��ꂽ�l�̃o���f�[�V����
    If validate = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    '// ���H����t�@�C���Ƃ��Ďw�肳�ꂽ�t�@�C�����K�؂Ȃ��̂��m�F
    If validateFile = False Then
        MsgBox "�w�肵���t�@�C�����K�؂ł͂���܂���B", vbExclamation, "�o����쐬"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Call createSheetIfNotExist("document_type", ThisWorkbook)
    
    '// �`���b�g���M���Ɏ����̎�ނ��Ж��E�Ώی��Ȃǂ𔻕ʂ��邽�߂ɃV�[�g�̒l��ύX
    With ThisWorkbook.Sheets("document_type")
        .Cells(1, 1).Value = "cost"
        .Cells(1, 2).Value = cmbCompany.Value
        .Cells(1, 3).Value = "�y��ʌo��z"
        .Cells(1, 4).Value = txtYear.Value & "�N" & cmbMonth.Value
    End With
        
    ThisWorkbook.Sheets("document_type").Visible = False
    
    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value)
    
    '// �\���H
    Call ProcessChart
    
    '// ���H����csv�t�@�C����Excel�t�@�C���Ƃ��ĕۑ�
    Dim wsh As New WshShell
    
    ActiveWorkbook.SaveAs wsh.SpecialFolders(4) & "\" & cmbCompany.Value & txtYear.Value & "�N" & cmbMonth.Value & "�o���.xlsx", xlOpenXMLWorkbook
    ActiveWorkbook.Close False
    
    Set wsh = Nothing
    
    MsgBox "�������������܂����B", vbInformation, "�o����쐬"
    
    Unload Me

End Sub

'// ���͂��ꂽ�l�̃o���f�[�V����
Private Function validate() As Boolean

    validate = False

    '// ��Ж����I������Ă��邩
    If Me.cmbCompany.Value = "" Then
        MsgBox "��Ж���I�����Ă��������B", vbQuestion, "�o����쐬"
        Exit Function
        
    '// �v��N�����͂���Ă��邩
    ElseIf Me.txtYear.Value = "" Then
        MsgBox "�v��N����͂��Ă��������B", vbQuestion, "�o����쐬"
        Exit Function
    
    '// �v��N�ɐ��������͂���Ă��邩
    ElseIf IsNumeric(Me.txtYear.Value) = False Then
        MsgBox "�v��N�ɂ͐�������͂��Ă��������B", vbQuestion, "�o����쐬"
        Exit Function
        
    '// �v�㌎���I������Ă��邩
    ElseIf Me.cmbMonth.Value = "" Then
        MsgBox "�v�㌎����͂��Ă��������B", vbQuestion, "�o����쐬"
        Exit Function
    
    '// ���H�t�@�C�������͂���Ă��邩
    ElseIf Me.txtFileName.Value = "" Then
        MsgBox "���H�t�@�C����I�����Ă��������B", vbQuestion, "�o����쐬"
        Exit Function
    End If
    
    validate = True

End Function

'// ���H����t�@�C���Ƃ��Ďw�肳�ꂽ�t�@�C�����K�؂Ȃ��̂��m�F
Private Function validateFile() As Boolean

    validateFile = False

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value, ReadOnly:=True)
    
    With targetFile.Sheets(1)
    
        If .Cells(1, 1).Value = "���t" _
            And .Cells(1, 2).Value = "�ԍ�" _
            And .Cells(1, 3).Value = "�؜�/�`��" _
            And .Cells(1, 4).Value = "�ؕ�����ȖڃR�[�h" _
            And .Cells(1, 5).Value = "�ؕ�����Ȗږ�" _
            And .Cells(1, 6).Value = "�ؕ��⏕�ȖڃR�[�h" _
            And .Cells(1, 7).Value = "�ؕ��⏕�Ȗږ�" _
            And .Cells(1, 8).Value = "�ؕ��E�v" _
            And .Cells(1, 9).Value = "�ؕ������R�[�h" _
            And .Cells(1, 10).Value = "�ؕ�����於" _
            And .Cells(1, 11).Value = "�ؕ�����R�[�h" _
            And .Cells(1, 12).Value = "�ؕ����喼" _
            And .Cells(1, 13).Value = "�ؕ��ŋ�R�[�h" _
            And .Cells(1, 14).Value = "�ؕ��ŋ敪" _
            And .Cells(1, 15).Value = "�ؕ����z" _
            And .Cells(1, 16).Value = "�ؕ������" _
            And .Cells(1, 17).Value = "�ݕ�����ȖڃR�[�h" _
            And .Cells(1, 18).Value = "�ݕ�����Ȗږ�" _
            And .Cells(1, 19).Value = "�ݕ��⏕�ȖڃR�[�h" _
            And .Cells(1, 20).Value = "�ݕ��⏕�Ȗږ�" _
            And .Cells(1, 21).Value = "�ݕ��E�v" _
            And .Cells(1, 22).Value = "�ݕ������R�[�h" _
            And .Cells(1, 23).Value = "�ݕ�����於" _
            And .Cells(1, 24).Value = "�ݕ�����R�[�h" _
            And .Cells(1, 25).Value = "�ݕ����喼" Then
            
            If .Cells(1, 26).Value = "�ݕ��ŋ�R�[�h" _
                And .Cells(1, 27).Value = "�ݕ��ŋ敪" _
                And .Cells(1, 28).Value = "�ݕ����z" _
                And .Cells(1, 29).Value = "�ݕ������" _
                And .Cells(1, 30).Value = "���͌����" Then
            
                validateFile = True
            End If
        End If
    End With
    
    targetFile.Close False
    Set targetFile = Nothing

End Function

'�\���H
Private Sub ProcessChart()
        
    '// �R�݉^���̏ꍇ��2127:����YM���폜
    If cmbCompany.Value = "�R�݉^����" Then
        Call delete2127
    End If
        
    '// �s�v�ȗ���폜
    Range("B:C,M:M,Q:AD").Delete xlToLeft

    Dim lastRow As Integer: lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '// ���傪���ʂ܂��͋󗓂̂��̂��폜
    With Cells(1, 1)
        .AutoFilter 9, "0", xlOr, ""
        .CurrentRegion.Resize(.CurrentRegion.Rows.Count - 1).Offset(1, 0).Delete xlUp
        .AutoFilter
    End With
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '// �Ŕ����z��쐬
    Cells(1, 14).Value = "�Ŕ����z"
    Cells(2, 14).Formula = "=L2-M2"
    
    Cells(2, 14).AutoFill Range(Cells(2, 14), Cells(lastRow, 14))
    
    Range(Cells(2, 14), Cells(lastRow, 14)).Copy
    Cells(2, 14).PasteSpecial xlPasteValues
    Columns(14).NumberFormatLocal = "#,###"
    
    '// �r���ݒ�
    Range(Cells(1, 1), Cells(lastRow, 14)).Borders.LineStyle = xlContinuous
    
    Columns("A:N").EntireColumn.AutoFit
    
    '// ����R�[�h�ŏ����ɕ��ёւ�
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, 9), Order:=xlAscending
        .SetRange Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column))
        .Header = xlYes
        .Apply
    End With
    
End Sub

'// 2127:YM�����폜
Private Sub delete2127()

    With Cells(1, 1)
        .AutoFilter 4, "2127"
        .CurrentRegion.Offset(1).Delete xlUp
        .AutoFilter
        
        .AutoFilter 17, "2127"
        .CurrentRegion.Offset(1).Delete xlUp
        .AutoFilter
    End With
    
End Sub


'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()
    
    With cmbCompany
        .AddItem "�R�݉^����"
        .AddItem "�R�݉^����YMܰ��"
        .AddItem "��YCL"
    End With
    
    Dim i As Integer
    For i = 4 To 12
        cmbMonth.AddItem i & "��"
    Next
    For i = 1 To 3
        cmbMonth.AddItem i & "��"
    Next
    
    txtYear.Value = Year(Now)
    
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

'// ������������Ƃ��̏���
Private Sub cmdCancel_Click()
    
    Unload Me

End Sub

