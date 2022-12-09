Attribute VB_Name = "chatwork"
Option Explicit

'�v���O����0-2�b�萔�ݒ�
Const adTypeBinary = 1
Const adTypeText = 2
  
'// �`���b�g�Ƀ��b�Z�[�W�ƃt�@�C����Y�t
Function UploadFileWithMessageToChatwork(ByVal apiToken As String, ByVal roomId As String, ByVal title As String, ByVal message As String, ByRef mentions As Variant, ByVal filePath As String) As Boolean
    
    Dim i As Integer, mention As String
    
    '// ����̕��͍쐬 �uTo:000000�e�X�g����v�̌`�ɂȂ�
    For i = 0 To UBound(mentions)
        mention = mention & "[To:" & Split(mentions(i), ":")(0) & "]" & Split(mentions(i), ":")(1) & "����" & vbLf
    Next
    
    Dim text As String: text = mention & vbLf & "[info][title]" & title & "[/title]" & message & "[/info]"
    Dim url As String: url = "https://api.chatwork.com/v2/rooms/" & roomId & "/files"
    
    Dim myStream As New ADODB.stream
    myStream.Open
 
    Dim MyFSO As New FileSystemObject
    
    '// �Y�t�t�@�C�������G���R�[�h
    Dim fileUrl As String
    fileUrl = MyFSO.GetFileName(filePath)
    fileUrl = Application.WorksheetFunction.EncodeURL(fileUrl)
    
    '// �t�@�C���^�C�v
    Dim fileType As String: fileType = "application/octet-stream"
    
    '// HTTP���N�G�X�g�Ŏg�p����boudary(���E��)
    Dim httpBoundary As String: httpBoundary = createBoundary
    
    '/**
     '* HTTP���N�G�X�g�̃{�f�B�쐬
    '**/
    
    '// �Y�t�t�@�C���̃��N�G�X�g�쐬
    Call createHttpRequestOfFile(myStream, httpBoundary, fileUrl, filePath)
        
    '// ���b�Z�[�W�����̃��N�G�X�g�쐬
    Call createHttpRequestOfMessage(myStream, httpBoundary, text)
     
    '// HTTP���N�G�X�g�̏I���������쐬
    Call createHttpFooter(myStream, httpBoundary)
    
    myStream.Position = 0
    myStream.Type = adTypeBinary
    
    '/**
     '* HTTP���N�G�X�g���s
    '**/
    Dim xmlHttp As New XMLHTTP60
    
    With xmlHttp
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & httpBoundary
        .setRequestHeader "X-ChatWorkToken", apiToken
        .send myStream.Read()
    End With
    
    '// �ʒm�̎��s���ʂ��擾
    If InStr(xmlHttp.responseText, "file_id") > 0 Then
        UploadFileWithMessageToChatwork = True
    Else
        UploadFileWithMessageToChatwork = False
    End If
    
    Set myStream = Nothing
    Set xmlHttp = Nothing

End Function
 
'// �Y�t�t�@�C����HTTP���N�G�X�g�쐬
Private Sub createHttpRequestOfFile(ByRef myStream As ADODB.stream, ByVal boundary As String, ByVal fileUrl As String, ByVal filePath)

    '// �X�g���[���̃L�����Z�b�g�ƃ^�C�v�ύX
    Call changeCharsetAndType(myStream, adTypeText, "shift_jis")
    
    '// �w�b�_�[�����쐬
    Dim httpRequest As String
    
    httpRequest = "--" & boundary & vbLf _
            & "Content-Disposition: form-data; name=""file""; filename*=utf-8''" & fileUrl & vbLf _
            & "Content-Type:application/octet-stream" & vbLf & vbLf
 
    myStream.WriteText httpRequest
 
    '/**
     '* �Y�t�t�@�C�����o�C�i���f�[�^��
    '**/
    
    changeCharsetAndType myStream, adTypeBinary
 
    '// �V�����X�g���[���ɓY�t�t�@�C����ǂݍ��݁A�ǂݍ��񂾓��e�����̃X�g���[���ɒǉ�����
    Dim secondStream As New ADODB.stream
    secondStream.Type = adTypeBinary
    secondStream.Open
    secondStream.LoadFromFile filePath

    myStream.Write secondStream.Read()
 
    secondStream.Close
    Set secondStream = Nothing
     
End Sub
 
'// ���b�Z�[�W��HTTP���N�G�X�g�쐬
Private Sub createHttpRequestOfMessage(ByRef myStream As ADODB.stream, ByVal boundary As String, ByVal text As String)
      
    changeCharsetAndType myStream, adTypeText, "UTF-8"
    
    Dim httpRequest As String
    
    httpRequest = vbCrLf & "--" & boundary & vbLf _
                & "Content-Disposition: form-data; name=""message""" + vbLf + vbLf _
                & text + vbCrLf

    myStream.WriteText httpRequest

End Sub
 
'// HTTP���N�G�X�g�̏I�����������
Private Function createHttpFooter(ByRef myStream As ADODB.stream, ByVal boundary As String) As Boolean
    
    changeCharsetAndType myStream, adTypeText, "shift_jis"
    myStream.WriteText vbLf & "--" & boundary & "--" & vbLf
 
End Function
  
'/**
 '* �f�[�^�̕����R�[�h�ƃ^�C�v��ύX
 '* @params stream �f�[�^���������ރX�g���[��
 '* @params adType targetStream�̃^�C�v(�e�L�X�g���o�C�i����)
 '* @params char   �ύX���镶���R�[�h
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
 
'HTTP���N�G�X�g�Ƃ��ēn�����߂̃f�[�^�̋��E(boundary)��ݒ�
Private Function createBoundary() As String
     
    '// HTTP���N�G�X�g�Ŏg�p����f�[�^�̋��E���쐬(����̂�)
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
