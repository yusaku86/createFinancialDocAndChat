Attribute VB_Name = "common"
Option Explicit

'// ���㎑���쐬�̃t�H�[���N��
Public Sub openFormProfit()

    frmProfit.Show

End Sub

'// �o����쐬�̃t�H�[���N��
Public Sub openFormCost()

    frmCost.Show

End Sub

'// �`���b�g���[�N���M�̃t�H�[���N��
Public Sub openFormChatwork()

    If Sheets("�`���b�g���[�N").Cells(7, 4).Value = "" Then
        MsgBox "API�g�[�N���̐ݒ肪����Ă��܂���B" & vbLf & "�V�[�g�u�`���b�g���[�N�v��API�g�[�N���̐ݒ�����Ă��������B", vbQuestion, ThisWorkbook.Name
        Exit Sub
    End If
    
    frmChatwork.Show

End Sub
