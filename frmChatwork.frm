VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChatwork 
   Caption         =   "送信内容入力"
   ClientHeight    =   8535.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12435
   OleObjectBlob   =   "frmChatwork.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmChatwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// チャット送信のためのフォーム
Option Explicit

'/**
 '* メインプログラム(チャットで送る内容設定&チャット送信)
'**/
Private Sub cmdEnter_Click()

    '// バリデーション
    If validate = False Then: Exit Sub
    
    '// hhtp通信するurl
    Dim roomId As String: roomId = Split(cmbRoom.Value, ":")(0)
    
    '// APIトークン
    Dim apiToken As String: apiToken = Sheets("チャットワーク").Cells(7, 4).Value
    
    '/**
     '* チャット送信
    '**/
    Dim cc As New ChatWorkController
    
    '// チャット送信用の文章作成
    Dim message As String: message = cc.createChatWorkText(Split(Me.txtMentionList.Value, vbCrLf), Me.cmbCompany.Value & ":" & Me.cmbType.Value, Me.txtMessage.Value)
    
    Dim result As Boolean
 
    '// メッセージのみ送信する場合
    If Me.txtFile.Value = "" Then
        result = cc.sendMessage(message, roomId, apiToken)
    
    '// ファイルとメッセージを送信する場合
    Else
        result = cc.sendMessageWithFile(message, Me.txtHiddenFileFullPath.Value, roomId, apiToken)
    End If
    
    If result = True Then
        MsgBox "送信が完了しました。", vbInformation, ThisWorkbook.Name
    Else
        MsgBox "送信できませんでした。", vbExclamation, ThisWorkbook.Name
    End If
      
    Call clearControls
      
End Sub

'/**
 '* バリデーション
'**/
Private Function validate() As Boolean

    '// 会社名が入力されているか
    If cmbCompany.Value = "" Then
        MsgBox "会社名を選択してください。", vbQuestion, "チャット送信"
        cmbCompany.SetFocus
        validate = False
        Exit Function
    End If
    
    '// 資料の種類が入力されているか
    If cmbType.Value = "" Then
        MsgBox "資料の種類を選択してください。", vbQuestion, "チャット送信"
        cmbType.SetFocus
        validate = False
        Exit Function
    End If
    
    '// 送信先グループが選択されているか
    If cmbRoom.Value = "" Then
        MsgBox "送信先グループを選択してください。", vbQuestion, "チャット送信"
        cmbRoom.SetFocus
        validate = False
        Exit Function
    End If
          
    '// メッセージが入力されているか
    If txtMessage.Value = "" Then
        If MsgBox("メッセージが入力されていませんが、送信してよろしいですか?", vbQuestion + vbYesNo, "チャット送信") = vbNo Then
            validate = False
            Exit Function
        End If
    End If
    
    validate = True
    
End Function

'/**
 '* テキストボックスとコンボボックスの値クリア
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
 '* ユーザーフォーム起動時の設定
'**/
Private Sub UserForm_Initialize()
                
    '// チャットグループ名の選択肢追加
    Dim i As Long
   
    For i = 7 To ThisWorkbook.Sheets("チャットワーク").Cells(Rows.Count, 5).End(xlUp).Row
        cmbRoom.AddItem ThisWorkbook.Sheets("チャットワーク").Cells(i, 5).Value
    Next
    
    '// 送信相手リスト追加
    For i = 7 To ThisWorkbook.Sheets("チャットワーク").Cells(Rows.Count, 6).End(xlUp).Row
        cmbMention.AddItem ThisWorkbook.Sheets("チャットワーク").Cells(i, 6).Value
    Next
    
    '// 会社名の選択肢追加
    With cmbCompany
        .AddItem "山岸運送㈱"
        .AddItem "山岸運送㈱YMﾜｰｸｽ"
        .AddItem "㈱YCL"
        .AddItem "東海YMﾄﾗﾝｽ㈱浜松営業所"
        .AddItem "東海YMﾄﾗﾝｽ㈱本社営業所"
        
        .Value = ThisWorkbook.Sheets("document_type").Cells(1, 2).Value
    End With
    
    '// 資料の種類の変更
    cmbType.Value = ThisWorkbook.Sheets("document_type").Cells(1, 3).Value
    
End Sub

'/**
 '* 会社名の値が変更された時の処理
'**/
Private Sub cmbCompany_Change()
    
   '資料のタイプの選択肢追加&値設定
    cmbType.Clear
        
    '// 経費資料の場合 ⇒ 資料の種類を【一般経費にする】
    If ThisWorkbook.Sheets("document_type").Cells(1, 1).Value = "cost" Then
        cmbType.Value = "【一般経費】"
    End If
    
    '// 山岸運送の場合
    If cmbCompany.Value = "山岸運送㈱" Then
        cmbType.AddItem "【運送売上】"
        cmbType.AddItem "【倉庫売上】"
        cmbType.AddItem "【一般経費】"
    
    '// YMﾜｰｸｽの場合
    ElseIf cmbCompany.Value = "山岸運送㈱YMﾜｰｸｽ" Then
        cmbType.AddItem "【修理収入】"
        cmbType.AddItem "【車両販売収入】"
        cmbType.AddItem "【ﾎﾞﾃﾞｨﾌﾟﾘﾝﾄ収入】"
        cmbType.AddItem "【一般経費】"

    '// YCLの場合
    ElseIf cmbCompany.Value = "㈱YCL" Then
        cmbType.AddItem "【運賃収入・庫内作業収入】"
        cmbType.AddItem "【一般経費】"
      
    '// 東海YM浜松の場合
    ElseIf cmbCompany.Value = "東海YMﾄﾗﾝｽ㈱浜松営業所" Then
        cmbType.AddItem "【運送売上・軽油売上】"
    
    '// 東海YM本社の場合
    ElseIf cmbCompany.Value = "東海YMﾄﾗﾝｽ㈱本社営業所" Then
        cmbType.AddItem "【運送売上】"
    End If
    
End Sub

'/**
 '* 種類が変更された時の処理
'**/
Private Sub cmbType_Change()
    
    If cmbType.Value = "" Then
        Exit Sub
    End If
    
    '/**
     '* デスクトップのファイルの中に送付対象のファイルがあればtxtFileの値に設定する
    '**/
    Dim targetFileName As String
    
    '// 経費資料の場合
    If Me.cmbType.Value = "【一般経費】" Then
        targetFileName = cmbCompany.Value & Year(Sheets("document_type").Cells(1, 4).Value) & "年" & Month(Sheets("document_type").Cells(1, 4).Value) & "月" & "経費資料.xlsx"
    
    '// 売上資料の場合
    Else
        targetFileName = cmbCompany.Value & cmbType.Value & " 取引先別売上一覧表.pdf"
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
     '* チャットのメッセージ設定
    '**/

    '// 経費資料の場合
    If Me.cmbType.Value = "【一般経費】" Then
        txtMessage.Value = "お疲れ様です。" & Format(ThisWorkbook.Sheets("document_type").Cells(1, 4).Value, "yyyy年m月") _
                         & "分一般経費(管理会計用資料)を添付いたします。" & vbLf _
                         & "ご確認お願いいたします。"
    '// 売上資料の場合
    Else
        txtMessage.Value = "【報告】" & cmbCompany.Value & vbLf _
                         & Format(ThisWorkbook.Sheets("document_type").Cells(1, 4).Value, "yyyy年m月") _
                         & "末時点での" & cmbType.Value & "は添付の通りです。未回収はありません。" & vbLf _
                         & "ご確認お願いいたします。"
    End If
    
End Sub

'/**
 '* 追加を押したときの処理
'**/

Private Sub cmdAdd_Click()
    
    If cmbMention.Value = "" Then
        Exit Sub
    End If
    
    '// 既に宛先が追加されていたら抜ける
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

'// リセットを押したときの処理
Private Sub cmdReset_Click()

    txtMentionList.Value = ""

End Sub

'// 参照を押したときの処理
Private Sub cmdDialog_Click()

    Dim wsh As Object: Set wsh = CreateObject("Wscript.Shell")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim attachedFileName As String: attachedFileName = selectFile("添付ファイル選択", wsh.SpecialFolders(4) & "\", "Excelファイル・PDF", "*.xlsx;*.pdf;*.csv")
    
    If attachedFileName <> "" Then
        txtFile.Locked = False
    
        txtFile.Value = fso.GetFileName(attachedFileName)
        txtHiddenFileFullPath.Value = attachedFileName
    
        txtFile.Locked = True
    End If
    
    Set wsh = Nothing
    Set fso = Nothing
    
End Sub

'// ファイルをリセットを押した時の処理
Private Sub cmdClearFile_Click()

    Me.txtFile.Locked = False
    
    Me.txtFile.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    Me.txtFile.Locked = True

End Sub

'// 閉じるを押したときの処理
Private Sub cmdCancel_Click()
    
    Unload Me

End Sub


