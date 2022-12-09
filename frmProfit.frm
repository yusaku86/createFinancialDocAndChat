VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProfit 
   Caption         =   "売上資料作成"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11625
   OleObjectBlob   =   "frmProfit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'/**
 '* メインプロシージャ(チャット送信用売上資料作成)
'**/
Private Sub cmdEnter_Click()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '// 入力された値のバリデーション
    If validate = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    '// 加工するファイルが適切なものか確認
    If validateFile = False Then
        MsgBox "指定したファイルが適切ではありません。", vbExclamation, "売上資料作成"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Call createSheetIfNotExist("document_type", ThisWorkbook)
    
    '// チャット送信時に資料の種類や会社名・対象月などを判別するためにシートの値を変更
    With ThisWorkbook.Sheets("document_type")
        .Cells(1, 1).Value = "profit"
        .Cells(1, 2).Value = cmbCompany.Value
        .Cells(1, 3).Value = cmbProfitType.Value
        .Cells(1, 4).Value = txtYear.Value & "年" & cmbMonth.Value
    End With
        
    ThisWorkbook.Sheets("document_type").Visible = False
    
    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value)
    
    '// 表加工開始(不要な列を削除し、貸方のみ抽出して合計の行作成・並び替え)
    Call ProcessChart
        
    '// 罫線設定とセルの結合、タイトル設定
    Call RuleLine
    
    '// PDF出力
    Dim pdfName As String
    pdfName = cmbCompany & cmbProfitType.Value & " 取引先別売上一覧表.pdf"
    
    Call ExportPDF(pdfName)
        
    MsgBox "PDF出力が完了しました。デスクトップを確認してください。", vbInformation, "売上資料作成"
    
    targetFile.Close False
    
    Set targetFile = Nothing
    Unload Me

End Sub

'// 入力された値のバリデーション
Private Function validate() As Boolean

    validate = False

    '// 会社名が入力されているか
    If cmbCompany.Value = "" Then
        MsgBox "会社名を選択してください。", vbQuestion, "売上資料作成"
        Exit Function
         
    '// 対象年が入力されているか
    ElseIf txtYear.Value = "" Then
        MsgBox "対象年を入力してください。", vbQuestion, "売上資料作成"
        Exit Function
        
    '// 対象年に数字が入力されているか
    ElseIf IsNumeric(txtYear.Value) = False Then
        MsgBox "対象年には数字を入力してください。", vbQuestion, "売上資料作成"
        Exit Function
        
    '// 対象月が選択されているか
    ElseIf cmbMonth.Value = "" Then
        MsgBox "対象月を選択してください。", vbQuestion, "売上資料作成"
        Exit Function
    
    '// 加工ファイル名が入力されているか
    ElseIf txtFileName.Value = "" Then
        MsgBox "加工ファイル名を入力してください。"
        Exit Function
    End If
    
    validate = True

End Function

'// 加工するファイルとして選択されたファイルが適切なものか確認
Private Function validateFile() As Boolean

    validateFile = False

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value, ReadOnly:=True)
    
    With targetFile.Sheets(1)
    
        If .Cells(1, 1).Value = "部門" _
            And .Cells(1, 2).Value = "コード" _
            And .Cells(1, 3).Value = "科目" _
            And .Cells(1, 4).Value = "コード" _
            And .Cells(1, 5).Value = "補助科目" _
            And .Cells(1, 6).Value = "コード" _
            And .Cells(1, 7).Value = "取引先" Then
            
            validateFile = True
        End If
    End With
    
    targetFile.Close False
    Set targetFile = Nothing

End Function

'/**
 '* ユーザーフォーム起動時の処理
'**/
Private Sub UserForm_Initialize()
 
 '// コンボボックスに会社名と月を追加
    With cmbCompany
        .AddItem "山岸運送㈱"
        .AddItem "山岸運送㈱YMﾜｰｸｽ"
        .AddItem "㈱YCL"
        .AddItem "東海YMﾄﾗﾝｽ㈱浜松営業所"
        .AddItem "東海YMﾄﾗﾝｽ㈱本社営業所"
    End With

    Dim i As Long
    
    For i = 4 To 12
        cmbMonth.AddItem i & "月"
    Next
    
    For i = 1 To 3
        cmbMonth.AddItem i & "月"
    Next
    
    '年のデフォルト値を設定
    txtYear.text = Year(Now)
    
    txtFileName.Locked = True
  
End Sub

'// 参照を押したときの処理 ⇒ ダイアログを表示して加工するファイルを選択
Private Sub cmdDialog_Click()

    Dim wsh As New WshShell

    Dim filename As String: filename = selectFile("加工するファイルを選択してください。", wsh.SpecialFolders(4) & "\", "Excelファイル", "*.csv;*.xlsx")
    
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

'// 会社名が変更された時の処理
Private Sub cmbCompany_Change()

    cmbProfitType.Clear

    If cmbCompany.Value = "山岸運送㈱" Then
        cmbProfitType.AddItem "【運送売上】"
        cmbProfitType.AddItem "【倉庫売上】"
    
    ElseIf cmbCompany.Value = "山岸運送㈱YMﾜｰｸｽ" Then
        cmbProfitType.AddItem "【修理収入】"
        cmbProfitType.AddItem "【車両販売収入】"
        cmbProfitType.AddItem "【ﾎﾞﾃﾞｨﾌﾟﾘﾝﾄ収入】"
    
    ElseIf cmbCompany.Value = "㈱YCL" Then
        cmbProfitType.AddItem "【運賃収入・庫内作業収入】"
        cmbProfitType.Value = "【運賃収入・庫内作業収入】"
    
    ElseIf cmbCompany.Value = "東海YMﾄﾗﾝｽ㈱浜松営業所" Then
        cmbProfitType.AddItem "【運送売上・軽油売上】"
        cmbProfitType.Value = "【運送売上・軽油売上】"
    
    ElseIf cmbCompany.Value = "東海YMﾄﾗﾝｽ㈱本社営業所" Then
        cmbProfitType.AddItem "【運送売上】"
        cmbProfitType.Value = "【運送売上】"
    End If
    
End Sub

'/**
 '* 表加工
'**/
Private Sub ProcessChart()
    
    '不要な列を削除し、貸方のみ抽出
    Range("A:B, D:E").Delete xlToRight
    
    With Cells(1, 1)
        .AutoFilter 4, "<>貸方"
        .CurrentRegion.Offset(1).Delete xlUp
        .AutoFilter
    End With
    
    Columns(17).Cut
    Cells(1, 4).Select
    ActiveSheet.Paste
    
    '// 金額が入力されていない月の列削除
    Dim lastcolumn As Long: lastcolumn = Cells(2, Columns.Count).End(xlToLeft).Column
    Range(Columns(lastcolumn + 1), Columns(Cells(1, Columns.Count).End(xlToLeft).Column + 1)).Delete xlToLeft
    
    '// 合計行作成
    Rows(2).Insert xlDown
    Cells(2, 1).Value = "合計"
    
    Dim i As Integer
    
    For i = 4 To lastcolumn
        Cells(2, i).Value = Application.WorksheetFunction.Sum(Range(Cells(2, i), Cells(Cells(Rows.Count, 1).End(xlUp).Row, i)))
        Cells(2, i).NumberFormatLocal = "#,###"
    Next
    
    '並べ替え
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
 '* 罫線を引き、タイトル設定
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
        .Value = "取引先別売上一覧表(単位:円)"
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
 '* 会社ごとに対象の期間を判定
'**/
Private Function SetPeriod() As String
    
    '東海YMの場合
    If InStr(1, cmbCompany.Value, "東海YM") > 0 Then
        If Val(Replace(cmbMonth.Value, "月", "")) < 4 Then
            SetPeriod = txtYear.Value - 1 & "年6月～" & cmbMonth.Value
        Else
            SetPeriod = txtYear.Value & "年6月～" & cmbMonth.Value
        End If
    'それ以外
    Else
        If Val(Replace(cmbMonth.Value, "月", "")) < 4 Then
            SetPeriod = txtYear.Value - 1 & "年4月～" & cmbMonth.Value
        Else
            SetPeriod = txtYear.Value & "年4月～" & cmbMonth.Value
        End If
    End If
    
End Function

'/**
 '* PDF出力
'**/
Private Sub ExportPDF(ByVal filename As String)

    '// 印刷設定
    With ActiveSheet.PageSetup
        .Zoom = False
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .CenterHorizontally = True
    End With
    
    '// デスクトップパスを取得してPDF出力
    Dim wsh As New WshShell

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=wsh.SpecialFolders(4) & "\" & filename
    
    Set wsh = Nothing

End Sub

'/**
 '*「閉じる」を押した時の処理
'**/

Private Sub cmdCancel_Click()
    Unload Me
End Sub
