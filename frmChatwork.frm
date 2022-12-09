VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChatwork 
   Caption         =   "���M���e����"
   ClientHeight    =   8535.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12435
   OleObjectBlob   =   "frmChatwork.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmChatwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// �`���b�g���M�̂��߂̃t�H�[��
Option Explicit

'/**
 '* ���C���v���O����(�`���b�g�ő�����e�ݒ�&�`���b�g���M)
'**/
Private Sub cmdEnter_Click()

    '// �o���f�[�V����
    If validate = False Then: Exit Sub
    
    '// hhtp�ʐM����url
    Dim roomId As String: roomId = Split(cmbRoom.Value, ":")(0)
    
    '// API�g�[�N��
    Dim apiToken As String: apiToken = Sheets("�`���b�g���[�N").Cells(7, 4).Value
    
    '/**
     '* �`���b�g���M
    '**/
    Dim cc As New ChatWorkController
    
    '// �`���b�g���M�p�̕��͍쐬
    Dim message As String: message = cc.createChatWorkText(Split(Me.txtMentionList.Value, vbCrLf), Me.cmbCompany.Value & ":" & Me.cmbType.Value, Me.txtMessage.Value)
    
    Dim result As Boolean
 
    '// ���b�Z�[�W�̂ݑ��M����ꍇ
    If Me.txtFile.Value = "" Then
        result = cc.sendMessage(message, roomId, apiToken)
    
    '// �t�@�C���ƃ��b�Z�[�W�𑗐M����ꍇ
    Else
        result = cc.sendMessageWithFile(message, Me.txtHiddenFileFullPath.Value, roomId, apiToken)
    End If
    
    If result = True Then
        MsgBox "���M���������܂����B", vbInformation, ThisWorkbook.Name
    Else
        MsgBox "���M�ł��܂���ł����B", vbExclamation, ThisWorkbook.Name
    End If
      
    Call clearControls
      
End Sub

'/**
 '* �o���f�[�V����
'**/
Private Function validate() As Boolean

    '// ��Ж������͂���Ă��邩
    If cmbCompany.Value = "" Then
        MsgBox "��Ж���I�����Ă��������B", vbQuestion, "�`���b�g���M"
        cmbCompany.SetFocus
        validate = False
        Exit Function
    End If
    
    '// �����̎�ނ����͂���Ă��邩
    If cmbType.Value = "" Then
        MsgBox "�����̎�ނ�I�����Ă��������B", vbQuestion, "�`���b�g���M"
        cmbType.SetFocus
        validate = False
        Exit Function
    End If
    
    '// ���M��O���[�v���I������Ă��邩
    If cmbRoom.Value = "" Then
        MsgBox "���M��O���[�v��I�����Ă��������B", vbQuestion, "�`���b�g���M"
        cmbRoom.SetFocus
        validate = False
        Exit Function
    End If
          
    '// ���b�Z�[�W�����͂���Ă��邩
    If txtMessage.Value = "" Then
        If MsgBox("���b�Z�[�W�����͂���Ă��܂��񂪁A���M���Ă�낵���ł���?", vbQuestion + vbYesNo, "�`���b�g���M") = vbNo Then
            validate = False
            Exit Function
        End If
    End If
    
    validate = True
    
End Function

'/**
 '* �e�L�X�g�{�b�N�X�ƃR���{�{�b�N�X�̒l�N���A
'**/
Private Sub clearControls()

    Dim myControl As Control
    
    For Each myControl In Me.Controls
        If myControl.Name Like "txt*" Or myControl.Name Like "cmb*" Then
            myControl.Value = ""
        End If
    Next
    
End Sub

'/**
 '* ���[�U�[�t�H�[���N�����̐ݒ�
'**/
Private Sub UserForm_Initialize()
                
    '// �`���b�g�O���[�v���̑I�����ǉ�
    Dim i As Long
   
    For i = 7 To ThisWorkbook.Sheets("�`���b�g���[�N").Cells(Rows.Count, 5).End(xlUp).Row
        cmbRoom.AddItem ThisWorkbook.Sheets("�`���b�g���[�N").Cells(i, 5).Value
    Next
    
    '// ���M���胊�X�g�ǉ�
    For i = 7 To ThisWorkbook.Sheets("�`���b�g���[�N").Cells(Rows.Count, 6).End(xlUp).Row
        cmbMention.AddItem ThisWorkbook.Sheets("�`���b�g���[�N").Cells(i, 6).Value
    Next
    
    '// ��Ж��̑I�����ǉ�
    With cmbCompany
        .AddItem "�R�݉^����"
        .AddItem "�R�݉^����YMܰ��"
        .AddItem "��YCL"
        .AddItem "���CYM��ݽ���l���c�Ə�"
        .AddItem "���CYM��ݽ���{�Љc�Ə�"
        
        .Value = ThisWorkbook.Sheets("document_type").Cells(1, 2).Value
    End With
    
    '// �����̎�ނ̕ύX
    cmbType.Value = ThisWorkbook.Sheets("document_type").Cells(1, 3).Value
    
End Sub

'/**
 '* ��Ж��̒l���ύX���ꂽ���̏���
'**/
Private Sub cmbCompany_Change()
    
   '�����̃^�C�v�̑I�����ǉ�&�l�ݒ�
    cmbType.Clear
        
    '// �o����̏ꍇ �� �����̎�ނ��y��ʌo��ɂ���z
    If ThisWorkbook.Sheets("document_type").Cells(1, 1).Value = "cost" Then
        cmbType.Value = "�y��ʌo��z"
    End If
    
    '// �R�݉^���̏ꍇ
    If cmbCompany.Value = "�R�݉^����" Then
        cmbType.AddItem "�y�^������z"
        cmbType.AddItem "�y�q�ɔ���z"
        cmbType.AddItem "�y��ʌo��z"
    
    '// YMܰ���̏ꍇ
    ElseIf cmbCompany.Value = "�R�݉^����YMܰ��" Then
        cmbType.AddItem "�y�C�������z"
        cmbType.AddItem "�y�ԗ��̔������z"
        cmbType.AddItem "�y���ި����Ď����z"
        cmbType.AddItem "�y��ʌo��z"

    '// YCL�̏ꍇ
    ElseIf cmbCompany.Value = "��YCL" Then
        cmbType.AddItem "�y�^�������E�ɓ���Ǝ����z"
        cmbType.AddItem "�y��ʌo��z"
      
    '// ���CYM�l���̏ꍇ
    ElseIf cmbCompany.Value = "���CYM��ݽ���l���c�Ə�" Then
        cmbType.AddItem "�y�^������E�y������z"
    
    '// ���CYM�{�Ђ̏ꍇ
    ElseIf cmbCompany.Value = "���CYM��ݽ���{�Љc�Ə�" Then
        cmbType.AddItem "�y�^������z"
    End If
    
End Sub

'/**
 '* ��ނ��ύX���ꂽ���̏���
'**/
Private Sub cmbType_Change()
    
    If cmbType.Value = "" Then
        Exit Sub
    End If
    
    '/**
     '* �f�X�N�g�b�v�̃t�@�C���̒��ɑ��t�Ώۂ̃t�@�C���������txtFile�̒l�ɐݒ肷��
    '**/
    Dim targetFileName As String
    
    '// �o����̏ꍇ
    If Me.cmbType.Value = "�y��ʌo��z" Then
        targetFileName = cmbCompany.Value & Year(Sheets("document_type").Cells(1, 4).Value) & "�N" & Month(Sheets("document_type").Cells(1, 4).Value) & "��" & "�o���.xlsx"
    
    '// ���㎑���̏ꍇ
    Else
        targetFileName = cmbCompany.Value & cmbType.Value & " �����ʔ���ꗗ�\.pdf"
    End If
    
    Dim fso As New FileSystemObject
    Dim wsh As New WshShell
    
    If fso.FileExists(wsh.SpecialFolders(4) & "\" & targetFileName) Then
        Me.txtFile.Locked = False
        
        Me.txtFile.Value = targetFileName
        Me.txtHiddenFileFullPath.Value = wsh.SpecialFolders(4) & "\" & targetFileName
    
        Me.txtFile.Locked = True
    End If
    
    Set fso = Nothing
    Set wsh = Nothing
    
    '/**
     '* �`���b�g�̃��b�Z�[�W�ݒ�
    '**/

    '// �o����̏ꍇ
    If Me.cmbType.Value = "�y��ʌo��z" Then
        txtMessage.Value = "�����l�ł��B" & Format(ThisWorkbook.Sheets("document_type").Cells(1, 4).Value, "yyyy�Nm��") _
                         & "����ʌo��(�Ǘ���v�p����)��Y�t�������܂��B" & vbLf _
                         & "���m�F���肢�������܂��B"
    '// ���㎑���̏ꍇ
    Else
        txtMessage.Value = "�y�񍐁z" & cmbCompany.Value & vbLf _
                         & Format(ThisWorkbook.Sheets("document_type").Cells(1, 4).Value, "yyyy�Nm��") _
                         & "�����_�ł�" & cmbType.Value & "�͓Y�t�̒ʂ�ł��B������͂���܂���B" & vbLf _
                         & "���m�F���肢�������܂��B"
    End If
    
End Sub

'/**
 '* �ǉ����������Ƃ��̏���
'**/

Private Sub cmdAdd_Click()
    
    If cmbMention.Value = "" Then
        Exit Sub
    End If
    
    '// ���Ɉ��悪�ǉ�����Ă����甲����
    If InStr(1, txtMentionList.Value, cmbMention.Value) > 0 Then
        cmbMention.Value = ""
        Exit Sub
    End If
    
        
    If txtMentionList.Value = "" Then
        txtMentionList.Value = cmbMention.Value
    Else
        txtMentionList.Value = txtMentionList.Value & vbLf & cmbMention.Value
    End If
    
    cmbMention.Value = ""

End Sub

'// ���Z�b�g���������Ƃ��̏���
Private Sub cmdReset_Click()

    txtMentionList.Value = ""

End Sub

'// �Q�Ƃ��������Ƃ��̏���
Private Sub cmdDialog_Click()

    Dim wsh As Object: Set wsh = CreateObject("Wscript.Shell")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim attachedFileName As String: attachedFileName = selectFile("�Y�t�t�@�C���I��", wsh.SpecialFolders(4) & "\", "Excel�t�@�C���EPDF", "*.xlsx;*.pdf;*.csv")
    
    If attachedFileName <> "" Then
        txtFile.Locked = False
    
        txtFile.Value = fso.GetFileName(attachedFileName)
        txtHiddenFileFullPath.Value = attachedFileName
    
        txtFile.Locked = True
    End If
    
    Set wsh = Nothing
    Set fso = Nothing
    
End Sub

'// �t�@�C�������Z�b�g�����������̏���
Private Sub cmdClearFile_Click()

    Me.txtFile.Locked = False
    
    Me.txtFile.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    Me.txtFile.Locked = True

End Sub

'// ������������Ƃ��̏���
Private Sub cmdCancel_Click()
    
    Unload Me

End Sub


