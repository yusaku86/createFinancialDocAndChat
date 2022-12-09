Attribute VB_Name = "common"
Option Explicit

'// 売上資料作成のフォーム起動
Public Sub openFormProfit()

    frmProfit.Show

End Sub

'// 経費資料作成のフォーム起動
Public Sub openFormCost()

    frmCost.Show

End Sub

'// チャットワーク送信のフォーム起動
Public Sub openFormChatwork()

    If Sheets("チャットワーク").Cells(7, 4).Value = "" Then
        MsgBox "APIトークンの設定がされていません。" & vbLf & "シート「チャットワーク」でAPIトークンの設定をしてください。", vbQuestion, ThisWorkbook.Name
        Exit Sub
    End If
    
    frmChatwork.Show

End Sub
