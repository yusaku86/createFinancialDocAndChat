VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCost 
   Caption         =   "経費資料作成"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6420
   OleObjectBlob   =   "frmCost.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// 管理会計用資料を作成するフォーム
Option Explicit

'// 経費資料作成(メインプロシージャ)
Private Sub cmdEnter_Click()
    
    Application.ScreenUpdating = False
    
    '// 入力された値のバリデーション
    If validate = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    '// 加工するファイルとして指定されたファイルが適切なものか確認
    If validateFile = False Then
        MsgBox "指定したファイルが適切ではありません。", vbExclamation, "経費資料作成"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Call createSheetIfNotExist("document_type", ThisWorkbook)
    
    '// チャット送信時に資料の種類や会社名・対象月などを判別するためにシートの値を変更
    With ThisWorkbook.Sheets("document_type")
        .Cells(1, 1).Value = "cost"
        .Cells(1, 2).Value = cmbCompany.Value
        .Cells(1, 3).Value = "【一般経費】"
        .Cells(1, 4).Value = txtYear.Value & "年" & cmbMonth.Value
    End With
        
    ThisWorkbook.Sheets("document_type").Visible = False
    
    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value)
    
    '// 表加工
    Call ProcessChart
    
    '// 加工したcsvファイルをExcelファイルとして保存
    Dim wsh As New WshShell
    
    ActiveWorkbook.SaveAs wsh.SpecialFolders(4) & "\" & cmbCompany.Value & txtYear.Value & "年" & cmbMonth.Value & "経費資料.xlsx", xlOpenXMLWorkbook
    ActiveWorkbook.Close False
    
    Set wsh = Nothing
    
    MsgBox "処理が完了しました。", vbInformation, "経費資料作成"
    
    Unload Me

End Sub

'// 入力された値のバリデーション
Private Function validate() As Boolean

    validate = False

    '// 会社名が選択されているか
    If Me.cmbCompany.Value = "" Then
        MsgBox "会社名を選択してください。", vbQuestion, "経費資料作成"
        Exit Function
        
    '// 計上年が入力されているか
    ElseIf Me.txtYear.Value = "" Then
        MsgBox "計上年を入力してください。", vbQuestion, "経費資料作成"
        Exit Function
    
    '// 計上年に数字が入力されているか
    ElseIf IsNumeric(Me.txtYear.Value) = False Then
        MsgBox "計上年には数字を入力してください。", vbQuestion, "経費資料作成"
        Exit Function
        
    '// 計上月が選択されているか
    ElseIf Me.cmbMonth.Value = "" Then
        MsgBox "計上月を入力してください。", vbQuestion, "経費資料作成"
        Exit Function
    
    '// 加工ファイルが入力されているか
    ElseIf Me.txtFileName.Value = "" Then
        MsgBox "加工ファイルを選択してください。", vbQuestion, "経費資料作成"
        Exit Function
    End If
    
    validate = True

End Function

'// 加工するファイルとして指定されたファイルが適切なものか確認
Private Function validateFile() As Boolean

    validateFile = False

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value, ReadOnly:=True)
    
    With targetFile.Sheets(1)
    
        If .Cells(1, 1).Value = "日付" _
            And .Cells(1, 2).Value = "番号" _
            And .Cells(1, 3).Value = "証憑/伝番" _
            And .Cells(1, 4).Value = "借方勘定科目コード" _
            And .Cells(1, 5).Value = "借方勘定科目名" _
            And .Cells(1, 6).Value = "借方補助科目コード" _
            And .Cells(1, 7).Value = "借方補助科目名" _
            And .Cells(1, 8).Value = "借方摘要" _
            And .Cells(1, 9).Value = "借方取引先コード" _
            And .Cells(1, 10).Value = "借方取引先名" _
            And .Cells(1, 11).Value = "借方部門コード" _
            And .Cells(1, 12).Value = "借方部門名" _
            And .Cells(1, 13).Value = "借方税区コード" _
            And .Cells(1, 14).Value = "借方税区分" _
            And .Cells(1, 15).Value = "借方金額" _
            And .Cells(1, 16).Value = "借方消費税" _
            And .Cells(1, 17).Value = "貸方勘定科目コード" _
            And .Cells(1, 18).Value = "貸方勘定科目名" _
            And .Cells(1, 19).Value = "貸方補助科目コード" _
            And .Cells(1, 20).Value = "貸方補助科目名" _
            And .Cells(1, 21).Value = "貸方摘要" _
            And .Cells(1, 22).Value = "貸方取引先コード" _
            And .Cells(1, 23).Value = "貸方取引先名" _
            And .Cells(1, 24).Value = "貸方部門コード" _
            And .Cells(1, 25).Value = "貸方部門名" Then
            
            If .Cells(1, 26).Value = "貸方税区コード" _
                And .Cells(1, 27).Value = "貸方税区分" _
                And .Cells(1, 28).Value = "貸方金額" _
                And .Cells(1, 29).Value = "貸方消費税" _
                And .Cells(1, 30).Value = "入力元画面" Then
            
                validateFile = True
            End If
        End If
    End With
    
    targetFile.Close False
    Set targetFile = Nothing

End Function

'表加工
Private Sub ProcessChart()
        
    '// 山岸運送の場合は2127:未払YMを削除
    If cmbCompany.Value = "山岸運送㈱" Then
        Call delete2127
    End If
        
    '// 不要な列を削除
    Range("B:C,M:M,Q:AD").Delete xlToLeft

    Dim lastRow As Integer: lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '// 部門が共通または空欄のものを削除
    With Cells(1, 1)
        .AutoFilter 9, "0", xlOr, ""
        .CurrentRegion.Resize(.CurrentRegion.Rows.Count - 1).Offset(1, 0).Delete xlUp
        .AutoFilter
    End With
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '// 税抜金額列作成
    Cells(1, 14).Value = "税抜金額"
    Cells(2, 14).Formula = "=L2-M2"
    
    Cells(2, 14).AutoFill Range(Cells(2, 14), Cells(lastRow, 14))
    
    Range(Cells(2, 14), Cells(lastRow, 14)).Copy
    Cells(2, 14).PasteSpecial xlPasteValues
    Columns(14).NumberFormatLocal = "#,###"
    
    '// 罫線設定
    Range(Cells(1, 1), Cells(lastRow, 14)).Borders.LineStyle = xlContinuous
    
    Columns("A:N").EntireColumn.AutoFit
    
    '// 部門コードで昇順に並び替え
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, 9), Order:=xlAscending
        .SetRange Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column))
        .Header = xlYes
        .Apply
    End With
    
End Sub

'// 2127:YM未払削除
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


'// フォーム起動時の処理
Private Sub UserForm_Initialize()
    
    With cmbCompany
        .AddItem "山岸運送㈱"
        .AddItem "山岸運送㈱YMﾜｰｸｽ"
        .AddItem "㈱YCL"
    End With
    
    Dim i As Integer
    For i = 4 To 12
        cmbMonth.AddItem i & "月"
    Next
    For i = 1 To 3
        cmbMonth.AddItem i & "月"
    Next
    
    txtYear.Value = Year(Now)
    
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

'// 閉じるを押したときの処理
Private Sub cmdCancel_Click()
    
    Unload Me

End Sub

