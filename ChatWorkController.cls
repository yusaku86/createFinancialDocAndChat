VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChatWorkController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// �`���b�g���[�N�𑀍삷��N���X
Option Explicit

'/**
 '* �`���b�g�Ƀ��b�Z�[�W�݂̂𑗐M
'**/
Public Function sendMessage(ByVal message As String, ByVal roomId, ByVal apiToken As String) As Boolean

    Dim roomUrl As String: roomUrl = "https://api.chatwork.com/v2/rooms/" & roomId & "/messages"

    Dim xmlHttp As New XMLHTTP60
    
    With xmlHttp
        .Open "POST", roomUrl, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "X-ChatWorkToken", apiToken
        .send "body=" & message
    End With
    
    '// �ʒm�̎��s���ʂ��擾
    If InStr(xmlHttp.responseText, "message_id") > 0 Then
        sendMessage = True
    Else
        sendMessage = False
    End If
    
    Set xmlHttp = Nothing
    
End Function

'/**
 '* �`���b�g�ɑ��M���郁�b�Z�[�W���쐬
'**/
Public Function createChatWorkText(ByVal mentions As Variant, ByVal title As String, ByVal message As String) As String

    Dim i As Long
    
    Dim mention As String
    
    '// ����쐬
    For i = 0 To UBound(mentions)
        If mention = "" Then
            mention = "[To:" & Split(mentions(i), ":")(0) & "]" & Split(mentions(i), ":")(1) & "����"
        Else
            mention = mention & vbLf & "[To:" & Split(mentions(i), ":")(0) & "]" & Split(mentions(i), ":")(1) & "����"
        End If
    Next
    
    Dim text As String
    
    '// ���b�Z�[�W���󔒂̏ꍇ
    If message = "" Then
        text = mention
    Else
        text = mention & "[info][title]" & title & "[/title]" & message & "[/info]"
    End If
    
    createChatWorkText = text

End Function

'==============================================================================================================================
'/**
 '* �`���b�g�Ƀt�@�C���ƃ��b�Z�[�W�𑗐M
'**/
Public Function sendMessageWithFile(ByVal message As String, ByVal filePath As String, ByVal roomId As String, ByVal apiToken As String) As Boolean
    
    '// HTTP���N�G�X�g�Ŏg�p����f�[�^�̋��E��(boundary)
    Dim httpBoundary As String: httpBoundary = createBoundary()
    
    Dim myStream As New ADODB.stream
    myStream.Open
    
    '// �Y�t�t�@�C����URL�G���R�[�h
    Dim fso As New FileSystemObject
    
    Dim fileUrl As String: fileUrl = WorksheetFunction.EncodeURL(fso.GetFileName(filePath))
    Set fso = Nothing
    
    '/**
     '* HTTP�̃��N�G�X�g�{�f�B�쐬
    '**/
    
    '// �Y�t�t�@�C���̃��N�G�X�g�쐬
    '// createHttpRequestOfFile [�������ރX�g���[��], [HTTP�ʐM�Ŏg�p���鋫�E��], [�Y�t�t�@�C������URL�G���R�[�h��������], [�Y�t�t�@�C���p�X]
    Call createHttpRequestOfFile(myStream, httpBoundary, fileUrl, filePath)
        
    '// ���b�Z�[�W�����̃��N�G�X�g�쐬
    If message <> "" Then
        '// createHttpRequestOfMessage [�������ރX�g���[��], [HTTP�ʐM�Ŏg�p���鋫�E��], [���M���郁�b�Z�[�W]
        Call createHttpRequestOfMessage(myStream, httpBoundary, message)
    End If
    
    '// HTTP���N�G�X�g�̏I���������쐬
    '// createHttpFooter [�������ރX�g���[��], [Http�ʐM�Ŏg�p���鋫�E��]
    Call createHttpFooter(myStream, httpBoundary)
    
    myStream.Position = 0
    myStream.Type = adTypeBinary
    
    '// HTTP�ʐM������url
    Dim roomUrl As String: roomUrl = "https://api.chatwork.com/v2/rooms/" & roomId & "/files"
    
    '/**
     '* HTTP���N�G�X�g���s
    '**/
    Dim xmlHttp As New XMLHTTP60

    With xmlHttp
        .Open "POST", roomUrl, False
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & httpBoundary
        .setRequestHeader "X-ChatWorkToken", apiToken
        .send myStream.Read()
    End With
    
    '// �ʒm�̎��s���ʂ��擾
    If InStr(xmlHttp.responseText, "file_id") > 0 Then
        sendMessageWithFile = True
    Else
        sendMessageWithFile = False
    End If
    
    Set myStream = Nothing
    Set xmlHttp = Nothing

End Function

'// HTTP���N�G�X�g�Ƃ��ēn�����߂̃f�[�^�̋��E(boundary)��ݒ�
Private Function createBoundary() As String
     
    '// HTTP���N�G�X�g�Ŏg�p����f�[�^�̋��E���쐬
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

'// �Y�t�t�@�C����HTTP���N�G�X�g�쐬
Private Sub createHttpRequestOfFile(ByRef myStream As ADODB.stream, ByVal boundary As String, ByVal fileUrl As String, ByVal filePath)

    '// �X�g���[���̃L�����Z�b�g�ƃ^�C�v�ύX
    Call changeCharsetAndType(myStream, adTypeText, "shift_jis")
    
    '// �w�b�_�[�����쐬
    Dim httpRequest As String
    
    httpRequest = "--" & boundary & vbCrLf _
            & "Content-Disposition: form-data; name=""file""; filename*=utf-8''" & fileUrl & vbCrLf _
            & "Content-Type:application/octet-stream" & vbCrLf & vbCrLf
 
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
Private Sub createHttpRequestOfMessage(ByRef myStream As ADODB.stream, ByVal boundary As String, ByVal message As String)
      
    changeCharsetAndType myStream, adTypeText, "UTF-8"
    
    Dim httpRequest As String
    
    httpRequest = vbCrLf & "--" & boundary & vbCrLf _
                & "Content-Disposition: form-data; name=""message""" + vbCrLf + vbCrLf _
                & message + vbCrLf

    myStream.WriteText httpRequest

End Sub

'// HTTP���N�G�X�g�̏I�����������
Private Function createHttpFooter(ByRef myStream As ADODB.stream, ByVal boundary As String) As Boolean
    
    changeCharsetAndType myStream, adTypeText, "shift_jis"
    myStream.WriteText vbCrLf & "--" & boundary & "--" & vbCrLf
 
End Function

'/**
 '* �f�[�^�̕����R�[�h�ƃ^�C�v��ύX
 '* @params stream �f�[�^���������ރX�g���[��
 '* @params adType targetStream�̃^�C�v(�e�L�X�g���o�C�i����)
 '* @params char   �ύX���镶���R�[�h
'**/
Private Sub changeCharsetAndType(ByRef targetStream As ADODB.stream, ByVal adType As Long, Optional ByVal char As String)
    
    '// �X�g���[���̃^�C�v�╶���R�[�h�̓|�W�V������0�ɂ��Ȃ��ƕύX�ł��Ȃ�����
    '// ���݂̃|�W�V������ϐ��Ɋi�[���A�^�C�v�ƕ����R�[�h��ύX����Ƃɖ߂�
    Dim currentPosition As Long: currentPosition = targetStream.Position
    
    targetStream.Position = 0
    
    targetStream.Type = adType
    
    If char <> "" Then
        targetStream.Charset = char
    End If
    
    targetStream.Position = currentPosition
 
End Sub