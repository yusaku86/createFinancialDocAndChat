Attribute VB_Name = "chatwork"
Option Explicit

'プログラム0-2｜定数設定
Const adTypeBinary = 1
Const adTypeText = 2
  
'// チャットにメッセージとファイルを添付
Function UploadFileWithMessageToChatwork(ByVal apiToken As String, ByVal roomId As String, ByVal title As String, ByVal message As String, ByRef mentions As Variant, ByVal filePath As String) As Boolean
    
    Dim i As Integer, mention As String
    
    '// 宛先の文章作成 「To:000000テストさん」の形になる
    For i = 0 To UBound(mentions)
        mention = mention & "[To:" & Split(mentions(i), ":")(0) & "]" & Split(mentions(i), ":")(1) & "さん" & vbLf
    Next
    
    Dim text As String: text = mention & vbLf & "[info][title]" & title & "[/title]" & message & "[/info]"
    Dim url As String: url = "https://api.chatwork.com/v2/rooms/" & roomId & "/files"
    
    Dim myStream As New ADODB.stream
    myStream.Open
 
    Dim MyFSO As New FileSystemObject
    
    '// 添付ファイル名をエンコード
    Dim fileUrl As String
    fileUrl = MyFSO.GetFileName(filePath)
    fileUrl = Application.WorksheetFunction.EncodeURL(fileUrl)
    
    '// ファイルタイプ
    Dim fileType As String: fileType = "application/octet-stream"
    
    '// HTTPリクエストで使用するboudary(境界線)
    Dim httpBoundary As String: httpBoundary = createBoundary
    
    '/**
     '* HTTPリクエストのボディ作成
    '**/
    
    '// 添付ファイルのリクエスト作成
    Call createHttpRequestOfFile(myStream, httpBoundary, fileUrl, filePath)
        
    '// メッセージ部分のリクエスト作成
    Call createHttpRequestOfMessage(myStream, httpBoundary, text)
     
    '// HTTPリクエストの終了部分を作成
    Call createHttpFooter(myStream, httpBoundary)
    
    myStream.Position = 0
    myStream.Type = adTypeBinary
    
    '/**
     '* HTTPリクエスト実行
    '**/
    Dim xmlHttp As New XMLHTTP60
    
    With xmlHttp
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & httpBoundary
        .setRequestHeader "X-ChatWorkToken", apiToken
        .send myStream.Read()
    End With
    
    '// 通知の実行結果を取得
    If InStr(xmlHttp.responseText, "file_id") > 0 Then
        UploadFileWithMessageToChatwork = True
    Else
        UploadFileWithMessageToChatwork = False
    End If
    
    Set myStream = Nothing
    Set xmlHttp = Nothing

End Function
 
'// 添付ファイルのHTTPリクエスト作成
Private Sub createHttpRequestOfFile(ByRef myStream As ADODB.stream, ByVal boundary As String, ByVal fileUrl As String, ByVal filePath)

    '// ストリームのキャラセットとタイプ変更
    Call changeCharsetAndType(myStream, adTypeText, "shift_jis")
    
    '// ヘッダー部分作成
    Dim httpRequest As String
    
    httpRequest = "--" & boundary & vbLf _
            & "Content-Disposition: form-data; name=""file""; filename*=utf-8''" & fileUrl & vbLf _
            & "Content-Type:application/octet-stream" & vbLf & vbLf
 
    myStream.WriteText httpRequest
 
    '/**
     '* 添付ファイルをバイナリデータ化
    '**/
    
    changeCharsetAndType myStream, adTypeBinary
 
    '// 新しいストリームに添付ファイルを読み込み、読み込んだ内容を元のストリームに追加する
    Dim secondStream As New ADODB.stream
    secondStream.Type = adTypeBinary
    secondStream.Open
    secondStream.LoadFromFile filePath

    myStream.Write secondStream.Read()
 
    secondStream.Close
    Set secondStream = Nothing
     
End Sub
 
'// メッセージのHTTPリクエスト作成
Private Sub createHttpRequestOfMessage(ByRef myStream As ADODB.stream, ByVal boundary As String, ByVal text As String)
      
    changeCharsetAndType myStream, adTypeText, "UTF-8"
    
    Dim httpRequest As String
    
    httpRequest = vbCrLf & "--" & boundary & vbLf _
                & "Content-Disposition: form-data; name=""message""" + vbLf + vbLf _
                & text + vbCrLf

    myStream.WriteText httpRequest

End Sub
 
'// HTTPリクエストの終了部分を作る
Private Function createHttpFooter(ByRef myStream As ADODB.stream, ByVal boundary As String) As Boolean
    
    changeCharsetAndType myStream, adTypeText, "shift_jis"
    myStream.WriteText vbLf & "--" & boundary & "--" & vbLf
 
End Function
  
'/**
 '* データの文字コードとタイプを変更
 '* @params stream データを書き込むストリーム
 '* @params adType targetStreamのタイプ(テキストかバイナリか)
 '* @params char   変更する文字コード
'**/
Private Sub changeCharsetAndType(ByRef targetStream As ADODB.stream, ByVal adType As Long, Optional ByVal char As String)
    
    Dim currentPosition As Long: currentPosition = targetStream.Position
    
    targetStream.Position = 0
    
    targetStream.Type = adType
    
    If char <> "" Then
        targetStream.Charset = char
    End If
    
    targetStream.Position = currentPosition
 
End Sub
 
'HTTPリクエストとして渡すためのデータの境界(boundary)を設定
Private Function createBoundary() As String
     
    '// HTTPリクエストで使用するデータの境界を作成(初回のみ)
    Dim multipartChars As String: multipartChars = "-_1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim boundary As String: boundary = "--------------------"
 
    Dim i As Long
    Dim point As Long

    For i = 1 To 16
        Randomize
        point = Int(Len(multipartChars) * Rnd + 1)
        boundary = boundary + Mid(multipartChars, point, 1)
    Next

    createBoundary = boundary
 
End Function
