VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "認定農業者受付業務システム"
   ClientHeight    =   9030.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13040
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CommandButton1_Click()
'①　認定処理初年（月）に実行
    Dim tmp As Integer, sr As Integer, sr0 As Integer, i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer, ReturnValue As Integer, Value As Integer
    Dim LastRow As Integer, LastRow2 As Integer, LastRow3 As Integer
    Dim OpenFileName As String, FileName As String, Path As String, SetFile As String
    Dim TmpChar As String, myFileName As String, SaveDir As String, tmpFileName As String
    Dim twbS00 As Worksheet, twbS03 As Worksheet, twbS01 As Worksheet, twbS10 As Worksheet, twbS11 As Worksheet, twbS12 As Worksheet
    Dim iwbS10 As Worksheet, iwbS11 As Worksheet, iwbS12 As Worksheet
    Dim ThisWbook As Workbook
    Dim Flag As Boolean
    Dim StringObject As Range, StringObject0 As Range
    Set twbS00 = Worksheets("フォーム呼出")
    Set twbS01 = Worksheets("0_日程表")
    Set twbS03 = Worksheets("2_申請・認定日")
    Set twbS10 = Worksheets("目次-認定者")
    Set twbS11 = Worksheets("目次-辞退者")
    Set twbS12 = Worksheets("目次-再発行")
On Error GoTo myError
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    With twbS10
    
'        初年度（回）登録の確認
        tmp = MsgBox("本年度初回登録ですか？", vbYesNo + vbQuestion, "確認")
        If tmp = vbYes Then
            If twbS00.Range("J1") = "●" Then
                MsgBox "初回ではありません処理を終了します。" & vbCrLf & "「フォーム呼出」シートのJ1セルを参照して下さい。"
                Unload UserForm5
                Exit Sub
            End If
            
'            初期化
            Flag = True
            twbS00.Range("J1") = "●"
            twbS10.Range("A3") = 1
            twbS11.Range("A3") = 1
            twbS12.Range("A3") = 1
            twbS00.Range("J3:M14").ClearContents
            
'            フォルダの作成
'            MsgBox "本業務の月数入力をお願いします。"
'            UserForm6.Show
'            SaveDir = ActiveWorkbook.Path & "\" & Format(twbS00.Range("G1"), "[$-ja-JP]ge") & "." & index
'            If Dir(SaveDir, vbDirectory) = "" Then
'                MkDir SaveDir
'                MkDir SaveDir & "\バックアップフォルダ"
'                MsgBox "同じ階層に" & Format(twbS00.Range("G1"), "[$-ja-JP]ge") & "." & index & "という保存フォルダを作成しました。"
'            End If
        Else
            Flag = False
            If twbS00.Range("J1") <> "●" Then
                MsgBox "初回ではありません処理を終了します。" & vbCrLf & "「フォーム呼出」シートのJ1セルを参照して下さい。"
                
                Unload UserForm5
                Exit Sub
            Else
            
'            初回ではなかった場合に前回更新・新規・辞退等のリスト（目次）を取込む。
                tmp = MsgBox("前回までの、目次を取込ますので前回の申請書一覧を読み込んでください。", vbYesNo + vbQuestion, "確認")
                If tmp = vbYes Then
                    OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,10申請者一覧表*.xls?")
                    If OpenFileName <> "False" Then
                         SetFile = OpenFileName
                    Else
                        MsgBox "キャンセルされました", vbCritical
                        Unload Me
                        Exit Sub
                    End If
                    Application.ScreenUpdating = False
                    Workbooks.Open FileName:=SetFile, ReadOnly:=True, UpdateLinks:=0
                    Set ImportWbook = Workbooks.Open(Path & SetFile)
                    Set iwbS10 = ImportWbook.Worksheets("目次-認定者")
                    Set iwbS11 = ImportWbook.Worksheets("目次-辞退者")
                    Set iwbS12 = ImportWbook.Worksheets("目次-再発行")
                    
                    LastRow = iwbS10.Cells(iwbS10.Rows.Count, 4).End(xlUp).Row
                    LastRow2 = iwbS11.Cells(iwbS11.Rows.Count, 4).End(xlUp).Row
                    LastRow3 = iwbS12.Cells(iwbS12.Rows.Count, 4).End(xlUp).Row
                    
                    iwbS10.Range("A3" & ":" & "E" & LastRow).Copy
                    twbS10.Range("A3").PasteSpecial xlPasteValues
                    iwbS11.Range("A3" & ":" & "E" & LastRow2).Copy
                    twbS11.Range("A3").PasteSpecial xlPasteValues
                    iwbS12.Range("A3" & ":" & "E" & LastRow3).Copy
                    twbS12.Range("A3").PasteSpecial xlPasteValues
                    Application.CutCopyMode = False
                Else
                    MsgBox "処理を中断します"
                    Unload Me
                End If
                ImportWbook.Close
                Application.DisplayAlerts = True
                Application.ScreenUpdating = True
                
            End If
        End If
        
'        実施記録
        twbS00.Range("A2").Interior.ColorIndex = 44
        twbS00.Range("A2").Font.ColorIndex = 1
        twbS00.Range("E2") = "" & Now & ""
        Unload Me
        MsgBox "処理が完了しました。続いてデータベースの取込を行って下さい。"
    End With
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
    Unload UserForm5
End Sub

Private Sub CommandButton2_Click()
'②　最新データベース取込
    Dim tmp As Integer, sr As Integer, LastRow As Integer, i As Integer, j As Integer
    Dim OpenFileName As String, FileName As String, Path As String, SetFile As String
    Dim ThisWbook As Workbook, ImportWbook As Workbook
    Dim twbS05 As Worksheet, iwbS01 As Worksheet, twbS00 As Worksheet, iwbS02 As Worksheet
    Set ThisWbook = ActiveWorkbook
    Set twbS05 = ThisWbook.Worksheets("data")
    Set twbS00 = ThisWbook.Worksheets("フォーム呼出")
On Error GoTo myError
    With twbS05
        Application.DisplayAlerts = False
        
'        データベースの選択（取込）
        tmp = MsgBox("認定農業者データベース（最新版）を選択して下さい", vbYesNo + vbQuestion, "確認")
        If tmp = vbYes Then
            OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,認定農業者データ*.xls?")
            If OpenFileName <> "False" Then
                 SetFile = OpenFileName
            Else
                MsgBox "キャンセルされました", vbCritical
                Unload Me
                Exit Sub
            End If
            Application.ScreenUpdating = False
            Workbooks.Open FileName:=SetFile, ReadOnly:=True, UpdateLinks:=0
            Set ImportWbook = Workbooks.Open(Path & SetFile)
            Set iwbS01 = ImportWbook.Worksheets("data (個人)")
            Set iwbS02 = ImportWbook.Worksheets("data (法人)")
            .Range("A4:CA1000").ClearContents
            
'            data (個人)シートのコピー
            LastRow = iwbS01.Cells(iwbS01.Rows.Count, 7).End(xlUp).Row
            iwbS01.Range("A5" & ":" & "CA" & LastRow).Copy
            .Range("A4").PasteSpecial xlPasteValues
            
'            data (法人)シートのコピー
            LastRow = iwbS02.Cells(iwbS02.Rows.Count, 7).End(xlUp).Row
            sr = .Cells(.Rows.Count, 7).End(xlUp).Row + 1
            
            iwbS02.Range("A5" & ":" & "A" & LastRow).Copy
            .Range("A" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("B5" & ":" & "F" & LastRow).Copy
            .Range("C" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("G5" & ":" & "G" & LastRow).Copy
            .Range("M" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("H5" & ":" & "H" & LastRow).Copy
            .Range("N" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("I5" & ":" & "I" & LastRow).Copy
            .Range("T" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("J5" & ":" & "J" & LastRow).Copy
            .Range("U" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("R5" & ":" & "U" & LastRow).Copy
            .Range("P" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("K5" & ":" & "K" & LastRow).Copy
            .Range("H" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("L5" & ":" & "L" & LastRow).Copy
            .Range("I" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("M5" & ":" & "M" & LastRow).Copy
            .Range("J" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("N5" & ":" & "N" & LastRow).Copy
            .Range("K" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("BO5" & ":" & "CY" & LastRow).Copy
            .Range("AP" & sr).PasteSpecial xlPasteValues
            iwbS02.Range("Q5" & ":" & "U" & LastRow).Copy
            .Range("O" & sr).PasteSpecial xlPasteValues

            LastRow = .Cells(.Rows.Count, 7).End(xlUp).Row + 1
            
'            新規追加分の色分け
            .Range("A" & LastRow & ":" & "CB" & LastRow + 10).Interior.ColorIndex = 0
            .Range("A" & LastRow & ":" & "CB" & LastRow + 10).Interior.ColorIndex = 34
            
'            実施記録
            .Range("F2") = "実行した日付は、" & Now & "　です。"
            .Range("A3:CA1000").AutoFilter
            
'            各月の認定農業者数の取得
            twbS00.Range("I3") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 4, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 4, 30))
            twbS00.Range("I4") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 5, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 5, 31))
            twbS00.Range("I5") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 6, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 6, 30))
            twbS00.Range("I6") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 7, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 7, 31))
            twbS00.Range("I7") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 8, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 8, 31))
            twbS00.Range("I8") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 9, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 9, 30))
            twbS00.Range("I9") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 10, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 10, 31))
            twbS00.Range("I10") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 11, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 11, 30))
            twbS00.Range("I11") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), 12, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), 12, 31))
            twbS00.Range("I12") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now) + 1, 1, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now) + 1, 1, 31))
            twbS00.Range("I13") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now) + 1, 2, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now) + 1, 2, 29))
            twbS00.Range("I14") = WorksheetFunction.CountIfs( _
                    twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now) + 1, 3, 1), _
                    twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now) + 1, 3, 31))

'            実施記録
            twbS00.Range("A3").Interior.ColorIndex = 44
            twbS00.Range("A3").Font.ColorIndex = 1
            twbS00.Range("E3") = "" & Now & ""
            
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            Application.CutCopyMode = False
            ImportWbook.Close False
            Unload Me
        Else
            MsgBox "処理を中断します"
            Unload Me
        End If
        
'        認定農業者総数取得
        twbS00.Select
        If twbS00.Range("L1") = "" Then
            twbS00.Range("G2") = WorksheetFunction.CountIf(twbS05.Range("G4:G1000"), "<>")
            twbS00.Range("L1") = "●"
        End If
    End With
    MsgBox "データベースの取り込みが完了しました。dataシートにて確認できます。"
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
    Unload Me
End Sub

Private Sub CommandButton3_Click()
'③　月毎の抽出
    Dim DateFrom As String, DateTo As String
    Dim tmp As Variant, TempNumber As Variant
    Dim InputValue As Integer, Value As Integer, i As Integer
    Dim LastRow, SetRow As Integer
    Dim twbS01 As Worksheet, twbS05 As Worksheet, twbS06 As Worksheet, twbS00 As Worksheet
On Error GoTo myError
    Set twbS01 = Worksheets("0_日程表")
    Set twbS05 = Worksheets("data")
    Set twbS06 = Worksheets("月別抽出")
    Set twbS00 = Worksheets("フォーム呼出")
    
'   初期化　不要なセル値の削除
    twbS01.Range("D2:D101").ClearContents
    twbS01.Range("E3:F101").ClearContents
    twbS01.Range("H3:J101").ClearContents
    twbS01.Range("U2:U101").ClearContents
    
'    各月の認定農業者終期期限の数取得
'    If twbS00.Range("I3") = "" Then
'        For i = 3 To 14
'            tmp = WorksheetFunction.CountIfs( _
'            twbS05.Range("BC4:BC503"), ">=" & DateSerial(Year(Now), Month(Now) + (i - 3), 1), _
'            twbS05.Range("BC4:BC503"), "<=" & DateSerial(Year(Now), Month(Now) + (i - 2), 0))
'            If twbS00.Cells(i, 9) = "" Then
'                twbS00.Cells(i, 9) = tmp
'            End If
'        Next i
'    End If

    DateFrom = Me.TextBox1.Text
    DateTo = Me.TextBox2.Text
    
'    入力された期間の抽出
    With twbS05
        twbS06.Range("A4:CA1000").clear
        .Range("BC4").AutoFilter 53, ">=" & DateFrom, xlAnd, "<=" & DateTo
        If WorksheetFunction.Subtotal(3, .Range("G:G")) < 2 Then
            MsgBox DateFrom & "から" & DateTo & "のデータは存在しません!" & vbCrLf & "確認の上再度実行して下さい。", vbInformation
            twbS06.Range("A4:CA1000").clear
            .Range("BC4").AutoFilter
            Exit Sub
        End If
        SetRow = .Range("A" & Rows.Count).End(xlUp).Row
        twbS06.Range("A4:CA4").ClearContents
        .Range("A4:CA4" & SetRow).Copy twbS06.Range("A4")
        .Range("BC4").AutoFilter
    End With
        Unload UserForm5
        twbS06.Select

        twbS06.Range("G4" & ":" & "G" & twbS06.Range("G" & Rows.Count).End(xlUp).Row).Copy
        twbS01.Range("D2").PasteSpecial xlPasteValues
        SetRow = twbS01.Range("D" & Rows.Count).End(xlUp).Row + 1
        twbS01.Range(SetRow & ":101").Delete
        
        Application.CutCopyMode = False
        
        MsgBox "抽出が完了しました。" & vbCrLf & "抽出件数は、" & WorksheetFunction.Subtotal(2, twbS06.Columns(1)) & "件です。" _
        & Space(10) & "抽出列は、BC列です。" & vbCrLf & "抽出月のデータが間違いないか確認してください。" _
        & vbCrLf & vbCrLf & "抽出したお名前は、0_日程表にコピーしました。"
        twbS00.Select
        
'        実施記録
        twbS00.Range("A4").Interior.ColorIndex = 44
        twbS00.Range("A4").Font.ColorIndex = 1
        twbS00.Range("E4") = "" & Now & ""
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
    Unload UserForm5
End Sub

Private Sub CommandButton4_Click()
'④　当該更新者データ（申請書）作成
    Dim tmp As Integer, sr As Integer, LastRow As Integer, SetRow As Integer, SetDate As Integer
    Dim i As Integer, SearchValue As Integer, LastRow2 As Integer, k As Integer, j As Integer
    Dim OpenFileName As String, FileName As String, Path As String, SetFile As String, ThisWbookPass As String, SetFileName As String
    Dim SetName As String, ReturnValue As String, myFileName As String, TmpChar As String, SaveDir As String
    Dim FoundValue As Object
    Dim ThisWbook As Workbook, ImportWbook As Workbook, ImportWbook2 As Workbook
    Dim twbS01 As Worksheet, twbS00 As Worksheet, twbS08 As Worksheet, iwbS01 As Worksheet, iwbS04 As Worksheet, iwbS05 As Worksheet
    Dim iwbS11 As Worksheet, iwbS14 As Worksheet, iwbS15 As Worksheet, iwbS08 As Worksheet
    Dim Flag As Boolean
    Set ThisWbook = ActiveWorkbook
    Set twbS01 = ThisWbook.Worksheets("0_日程表")
    Set twbS00 = ThisWbook.Worksheets("フォーム呼出")
    Set twbS08 = ThisWbook.Worksheets("月別抽出")
On Error GoTo myError
    Application.DisplayAlerts = False
    With twbS01
    
'        申請書を作成保存するフォルダの作成
        tmp = MsgBox("申請書を保存するフォルダを同じ階層に作成します。", vbYesNo + vbQuestion, "確認")
        If tmp = vbYes Then
            Flag = True
            SaveDir = ActiveWorkbook.Path & "\00認定農業者データ"
            If Dir(SaveDir, vbDirectory) = "" Then
                MkDir SaveDir
                MsgBox "このファイルと同じ階層に「00認定農業者データ」という保存フォルダを作成しました。"
            End If
        Else
            Flag = False
            MsgBox "すでにフォルダは作成されているか、作成されていなければ、手動でフォルダを作成して下さい。"
        End If
        
'        申請書リストの人数（行）取得
        LastRow = twbS01.Cells(twbS01.Rows.Count, 4).End(xlUp).Row
        Unload Me
        
'        原本ファイルの取込
        tmp = MsgBox("2-1 認定申請書 原本をセットして下さい。", vbYesNo + vbQuestion, "確認")
        If tmp = vbYes Then
            OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,2-1 認定申請書*.xls?")
            If OpenFileName <> "False" Then
                 SetFile = OpenFileName
            Else
                MsgBox "キャンセルされました", vbCritical
                Exit Sub
            End If
            
            Workbooks.Open FileName:=SetFile, ReadOnly:=False, UpdateLinks:=0
            Set ImportWbook = Workbooks.Open(Path & SetFile)
'            Windows(ImportWbook.Name).Visible = False
            Set iwbS01 = ImportWbook.Worksheets("入力シート")
            Set iwbS04 = ImportWbook.Worksheets("経営指標")
            Set iwbS05 = ImportWbook.Worksheets("生産施設")
            Set iwbS08 = ImportWbook.Worksheets("Record")
            If iwbS01.Range("M5") <> "漢字" Then
                MsgBox "取込んだファイルが違います。最初から手続きお願いします。", vbCritical
                ImportWbook.Close
                Exit Sub
            End If
            
            
            MsgBox "認定申請書（原本）取り込みが完了しました。" & vbCrLf & "続いて更新される方の申請書を読み込んでください。"
        Else
            MsgBox "キャンセルされました", vbCritical
            Exit Sub
        End If
        
        Application.ScreenUpdating = False
        ImportWbook.Application.ScreenUpdating = False
        
'        旧申請書の取込
        
        For i = 2 To LastRow
'        For i = 14 To LastRow  'テスト用
            If .Cells(i, 25) = "" Then
Label2:
                tmp = MsgBox(.Cells(i, 4).Value & " 様のファイルを選択してください。", vbYesNo + vbQuestion, "確認")
                
                If tmp = vbYes Then
                    OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*認定申請書*.xls?")
                    If OpenFileName <> "False" Then
                        SetFile = OpenFileName
                    Else
                        MsgBox "キャンセルされました", vbCritical
                        ImportWbook.Close False
                        Exit Sub
                    End If
                Else
                    MsgBox "キャンセルされました", vbCritical
                    ImportWbook.Close False
                    Unload Me
                    Exit Sub
                End If
                Workbooks.Open FileName:=SetFile, ReadOnly:=False, UpdateLinks:=0
                Set ImportWbook2 = Workbooks.Open(Path & SetFile)
'               　Windows(ImportWbook2.Name).Visible = False
                Set iwbS11 = ImportWbook2.Worksheets("入力シート")
                Set iwbS14 = ImportWbook2.Worksheets("経営指標")
                Set iwbS15 = ImportWbook2.Worksheets("生産施設")

                SetName = ""  '名前の初期化を行う
                
'                法人名または個人名の取得
                If iwbS11.Range("D5") = "" Then
                    SetName = iwbS11.Range("M5")
                    
                Else
                    SetName = iwbS11.Range("D5")
                End If
                If .Cells(i, 4).Value <> SetName Then
                    MsgBox "セットしたファイル" & SetName & "様ではありません。再度選択して下さい。"
                    ImportWbook2.Close
                    GoTo Label2
                End If
                
'               データの取込（コピー）※経営指標はコピーしない！
'                iwbS04.Unprotect 'シートの保護解除
'                iwbS14.Range("A181:N204").Copy
'                iwbS04.Range("A181").PasteSpecial xlPasteValuesAndNumberFormats
'                iwbS04.Protect 'シートの保護
                iwbS05.Unprotect
                iwbS15.Range("A4:F16").Copy
                iwbS05.Range("A4").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS05.Protect
    
                iwbS01.Range("D2") = iwbS11.Range("D2")
                iwbS01.Range("D3") = iwbS11.Range("D3")
                iwbS01.Range("D4") = iwbS11.Range("D4")
                iwbS01.Range("M4") = iwbS11.Range("M4")
    
                If IsNull(iwbS11.Range("D5")) Then
                    iwbS01.Range("M5") = iwbS11.Range("M5")
                    iwbS01.Range("M5").SetPhonetic
                    iwbS01.Range("M5").Phonetic.Visible = True
                    
                Else
                    iwbS01.Range("M5") = iwbS11.Range("M5")
                    iwbS01.Range("M5").SetPhonetic
                    iwbS01.Range("M5").Phonetic.Visible = True
                    iwbS01.Range("D5") = iwbS11.Range("D5")
                    iwbS01.Range("D5").SetPhonetic
                    iwbS01.Range("D5").Phonetic.Visible = True
                    iwbS01.Range("M7") = iwbS11.Range("M7")
                    iwbS01.Range("D8") = iwbS11.Range("D8")
                    iwbS01.Range("M8") = iwbS11.Range("M8")
                    
                End If
                
'                電話番号と郵便番号の取得
                LastRow2 = twbS08.Cells(twbS08.Rows.Count, 7).End(xlUp).Row
                For k = 2 To LastRow2
                    If SetName = twbS08.Cells(k, 7) Then
                        iwbS01.Range("D9") = twbS08.Cells(k, 16)
                        iwbS01.Range("M9") = twbS08.Cells(k, 18)
                        If twbS08.Cells(k, 19) <> "" Then
                            iwbS01.Range("M11") = twbS08.Cells(k, 19)
                        End If
                    End If
                Next
                
'                住所等の取得
                iwbS01.Range("D10") = iwbS11.Range("D10")
                If iwbS11.Range("AA2") <> "" Then
                    iwbS01.Range("AA2") = iwbS11.Range("AA2")
                End If
                iwbS01.Range("AB2") = iwbS11.Range("AB2")
                
'                構成員・役員
                iwbS11.Range("U6:U14").Copy
                iwbS01.Range("U6").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("Z5:Z14").Copy
                iwbS01.Range("Z5").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AB5:AB14").Copy
                iwbS01.Range("AB5").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AF5:AF14").Copy
                iwbS01.Range("AF5").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AH6:AH14").Copy
                iwbS01.Range("AH6").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AJ6:AJ14").Copy
                iwbS01.Range("AJ6").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AL6:AL14").Copy
                iwbS01.Range("AL6").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AN5:AN14").Copy
                iwbS01.Range("AN5").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AP5:AP14").Copy
                iwbS01.Range("AP5").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AR5:AR14").Copy
                iwbS01.Range("AR5").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AT5:AT14").Copy
                iwbS01.Range("AT5").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AP15:AP17").Copy
                iwbS01.Range("AP15").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AT15:AT17").Copy
                iwbS01.Range("AT15").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("L21:L26").Copy
                iwbS01.Range("L21").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("P21:P26").Copy
                iwbS01.Range("P21").PasteSpecial xlPasteValuesAndNumberFormats
                
'                農業生産施設
                iwbS11.Range("W21:W26").Copy
                iwbS01.Range("W21").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AF21:AF26").Copy
                iwbS01.Range("AF21").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AJ21:AJ26").Copy
                iwbS01.Range("AJ21").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AN21:AN26").Copy
                iwbS01.Range("AN21").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AR21:AR26").Copy
                iwbS01.Range("AR21").PasteSpecial xlPasteValuesAndNumberFormats
                
'                加工・販売事業
                iwbS11.Range("D30:D35").Copy
                iwbS01.Range("D30").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("L30:L35").Copy
                iwbS01.Range("L30").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("P30:P35").Copy
                iwbS01.Range("P30").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("T30:T35").Copy
                iwbS01.Range("T30").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("X30:X35").Copy
                iwbS01.Range("X30").PasteSpecial xlPasteValuesAndNumberFormats
                
'                その他の特記事項
                iwbS11.Range("AC29:AC31").Copy
                iwbS01.Range("AC29").PasteSpecial xlPasteValuesAndNumberFormats
                
'                生産項目
                iwbS11.Range("B42:B53").Copy
                iwbS01.Range("B42").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("H42:H53").Copy
                iwbS01.Range("H42").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("L42:L53").Copy
                iwbS01.Range("L42").PasteSpecial xlPasteValuesAndNumberFormats
                
'                5年後の項目
                iwbS11.Range("AB42:AB53").Copy
                iwbS01.Range("AB42").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AF42:AF53").Copy
                iwbS01.Range("AF42").PasteSpecial xlPasteValuesAndNumberFormats
    
'                作業受託
                iwbS11.Range("B59:B64").Copy
                iwbS01.Range("B59").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("E59:E64").Copy
                iwbS01.Range("E59").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("K59:K64").Copy
                iwbS01.Range("K59").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("S59:S64").Copy
                iwbS01.Range("S59").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("V59:V64").Copy
                iwbS01.Range("V59").PasteSpecial xlPasteValuesAndNumberFormats
                
'                農業機械
                iwbS11.Range("AI59:AI70").Copy
                iwbS01.Range("AI59").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AP59:AP70").Copy
                iwbS01.Range("AP59").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AS59:AS70").Copy
                iwbS01.Range("AS59").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AR59:AR70").Copy
                iwbS01.Range("AR59").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AU59:AU70").Copy
                iwbS01.Range("AU59").PasteSpecial xlPasteValuesAndNumberFormats
                
'                現状と目標・措置
                iwbS11.Range("BA74:BA81").Copy
                iwbS01.Range("BA74").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("BA83:BA90").Copy
                iwbS01.Range("BA83").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("BA92:BA99").Copy
                iwbS01.Range("BA92").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("BA101:BA108").Copy
                iwbS01.Range("BA101").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("BA110:BA126").Copy
                iwbS01.Range("BA110").PasteSpecial xlPasteValuesAndNumberFormats
                
'                5年前の機械施設
                iwbS11.Range("AJ74:AU86").Copy
                iwbS01.Range("AJ74").PasteSpecial xlPasteValuesAndNumberFormats
                
                Application.CutCopyMode = False
                
'                №の取得
                For j = 2 To LastRow
                    If SetName = .Cells(j, 4) Then
                       k = .Cells(j, 1).Value
                    End If
                Next
                ImportWbook2.Close
'                ファイルの保存　氏名と№を組み合わせたファイル名を作成
                SetFileName = Right("0" & k, 2) & "認定申請書" & Format(Date, "[$-ja-JP]ge") & "（" & SetName & " ）.xlsm"
                ImportWbook.SaveAs FileName:=SaveDir & "\" & Right("0" & k, 2) & "認定申請書" & Format(Date, "[$-ja-JP]ge") & "（" & SetName & " ）"
                
            End If
        Next
        Application.ScreenUpdating = True
        ImportWbook.Application.ScreenUpdating = True
'        Windows(SetFileName).Visible = True
        Application.ActiveWorkbook.Close
        Application.DisplayAlerts = True
        MsgBox "申請書の新規作成が完了しました。"
        
'        実施登録
        twbS00.Range("A5").Interior.ColorIndex = 44
        twbS00.Range("A5").Font.ColorIndex = 1
        twbS00.Range("E5") = "" & Now & ""
    End With
    Exit Sub
myError:
    MsgBox "選択したファイルが違います！処理を終了します。", vbExclamation
End Sub

Private Sub CommandButton5_Click()
'⑤　新規登録者データ（申請書）作成
    UserForm2.Show
End Sub
Private Sub CommandButton19_Click()
'⑦ 案内時に送付する書類（経営状況調査表）作成
    Dim tmp As Integer, sr As Integer, LastRow As Integer, SetRow As Integer, SetDate As Integer
    Dim i As Integer, SearchValue As Integer, LastRow2 As Integer, k As Integer, j As Integer, fileCount As Integer
    Dim arrayI() As Variant
    Dim OpenFileName As String, FileName As String, FileName2 As String, FileName3 As String, Path As String, SetFile As String, ThisWbookPass As String, SetFileName As String
    Dim SetName As String, ReturnValue As String, myFileName As String, TmpChar As String, SaveDir As String, SaveDir2 As String, SaveDir3 As String, FileNames As String, FileNames2 As String, FileNames3 As String
    Dim FoundValue As Object
    Dim fso As FileSystemObject
    Dim ThisWbook As Workbook, ImportWbook As Workbook, ImportWbook2 As Workbook, ImportWbook3 As Workbook
    Dim twbS01 As Worksheet, twbS00 As Worksheet, twbS08 As Worksheet, iwbS01 As Worksheet, iwbS04 As Worksheet, iwbS05 As Worksheet
    Dim iwbS11 As Worksheet, iwbS14 As Worksheet, iwbS15 As Worksheet
    Dim iwbS06 As Worksheet, iwbS07 As Worksheet, iwbS08 As Worksheet, iwbS09 As Worksheet
    Dim Flag As Boolean
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
    Dim oFso As FileSystemObject
    Set ThisWbook = ActiveWorkbook
    Set twbS01 = ThisWbook.Worksheets("0_日程表")
    Set twbS00 = ThisWbook.Worksheets("フォーム呼出")
On Error GoTo myError
    
    Unload Me
'    処理時間測定
    startTime = Timer
    
    SaveDir = ActiveWorkbook.Path & "\00認定農業者データ"
    SaveDir2 = ActiveWorkbook.Path

'    フォルダ内にあるファイルの自動読み込み
    '--- ファイルシステムオブジェクト ---'
    Set fso = CreateObject("Scripting.FileSystemObject")
    '--- ファイル数を格納する変数 ---'
    fileCount = fso.GetFolder(SaveDir).Files.Count
    
    '    送付・準備書類フォルダの作成
    SaveDir = ActiveWorkbook.Path & "\00認定農業者データ"
    SaveDir2 = SaveDir & "\送付・準備書類"
    If (oFso Is Nothing) Then
        Set oFso = CreateObject("Scripting.FileSystemObject")
    End If
    If (oFso.FolderExists(SaveDir2) = False) Then
        '// フォルダが存在しない
        MsgBox "「送付・準備書類」という名前のフォルダを新規作成します"
        MkDir SaveDir & "\送付・準備書類"
        SaveDir2 = SaveDir & "\送付・準備書類"
    Else
    End If
    
'    経営状況調査表、フォルダの作成
    FileName3 = SaveDir2 & "\経営状況調査表フォルダ"
    If (oFso Is Nothing) Then
        Set oFso = CreateObject("Scripting.FileSystemObject")
    End If
    If (oFso.FolderExists(FileName3) = False) Then
        '// フォルダが存在しない
        MsgBox "「経営状況調査表フォルダ」という名前のフォルダを新規作成します"
        MkDir SaveDir2 & "\経営状況調査表フォルダ"
    Else
    End If

    For i = 1 To fileCount
'    For i = 1 To 2 'テストコード
    
'    ワード文書印刷（まとめて印刷する場合）

'        Dim SaveDir As String
'        Dim wd As Object
'        SaveDir = ActiveWorkbook.Path & "\送付・準備書類\"
'        Set wd = CreateObject("Word.application")
'        wd.Visible = True
'        wd.documents.Open FileName:=SaveDir & "①　制度概要（改.doc" '印刷したい文書
'        wd.ActiveDocument.PrintOut Background:=False
'
'        wd.documents.Open FileName:=SaveDir & "②　新 意向確認書（令和）.doc" '印刷したい文書
'        wd.ActiveDocument.PrintOut Background:=False
'
'        wd.Quit
'        Set wd = Nothing

'     ワード文書印刷ここまで

        If i <= 9 Then
            TmpChar = "0" & CStr(i)
        Else
            TmpChar = i
        End If
        FileNames = Dir(SaveDir & "\" & TmpChar & "認定申請書R*")
        FileNames2 = Dir(SaveDir2 & "\④　経営状況調査表*")
        
'        ⑤　新認定計画申請書が指定の場所になかった場合は手動で取込。送付・準備書類フォルダにコピー
        If FileNames2 = "" Then
            tmp = MsgBox("④　経営状況調査表" & vbCrLf & "ファイルを選択してください。", vbYesNo + vbQuestion, "確認")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,④　経営状況調査表*.xls?")
                If OpenFileName <> "False" Then
                    FileCopy OpenFileName, SaveDir2 & "\" & "④　経営状況調査表.xlsx"
                    FileNames2 = Dir(SaveDir2 & "\" & "④　経営状況調査表*")
                Else
                    MsgBox "キャンセルされました", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "キャンセルされました", vbCritical
                ImportWbook.Close False
                Unload Me
                Exit Sub
            End If
        Else
        End If

        Application.DisplayAlerts = False
        Workbooks.Open FileName:=SaveDir & "\" & FileNames, ReadOnly:=True, UpdateLinks:=0
        Set ImportWbook = Workbooks.Open(SaveDir & "\" & FileNames)
        Set iwbS01 = ImportWbook.Worksheets("入力シート")
        Set iwbS04 = ImportWbook.Worksheets("簡易版")
        
'        現状と目標の数値データ入力
        Workbooks.Open FileName:=SaveDir2 & "\" & FileNames2, ReadOnly:=False, UpdateLinks:=0
        Set ImportWbook2 = Workbooks.Open(SaveDir2 & "\" & FileNames2)
        Set iwbS05 = ImportWbook2.Worksheets("経営状況調査表")
        
        If iwbS01.Range("D5") = "" Then
            SetName = iwbS01.Range("M5")
        Else
            SetName = iwbS01.Range("D5")
        End If
        
'        個人情報
        iwbS05.Range("B6") = SetName
        iwbS05.Range("G6").Value = iwbS01.Range("D10").Value
        iwbS05.Range("L6").Value = Right(iwbS01.Range("M9").Value, 7)
        iwbS05.Range("N6").Value = iwbS01.Range("M11").Value
        
'        構成員
        iwbS05.Range("B11:B20").Value = iwbS01.Range("U5:U14").Value
        iwbS05.Range("G11:G20").Value = iwbS01.Range("AF5:AF14").Value
        iwbS05.Range("I11:I20").Value = iwbS01.Range("AB5:AB14").Value
        iwbS05.Range("L11:L20").Value = iwbS01.Range("AP5:AP14").Value
        
'        雇用
        iwbS05.Range("B26").Value = iwbS01.Range("AP15").Value
        iwbS05.Range("G26").Value = iwbS01.Range("AP16").Value
        iwbS05.Range("K26").Value = iwbS01.Range("AP17").Value
        
'        生産施設
        iwbS05.Range("B31:B36").Value = iwbS01.Range("W21:W26").Value
        iwbS05.Range("H31:H36").Value = iwbS01.Range("AF21:AF26").Value
        iwbS05.Range("J31:J36").Value = iwbS01.Range("AJ21:AJ26").Value
        
'        農業機器
        iwbS05.Range("B41:B52").Value = iwbS01.Range("AJ74:AJ85").Value
        iwbS05.Range("I41:I52").Value = iwbS01.Range("AP74:AP85").Value
        iwbS05.Range("K41:K52").Value = iwbS01.Range("AS74:AS85").Value
        
'        農業生産（耕種）
        iwbS05.Range("B57:B62").Value = iwbS01.Range("B42:B47").Value
        iwbS05.Range("H57:H62").Value = iwbS01.Range("H42:H47").Value
        iwbS05.Range("J57:J62").Value = iwbS01.Range("L42:L47").Value
        
'        農業生産（畜産）
        iwbS05.Range("B64:B69").Value = iwbS01.Range("B48:B53").Value
        iwbS05.Range("H64:H69").Value = iwbS01.Range("H48:H53").Value
        iwbS05.Range("J64:J69").Value = iwbS01.Range("L48:L53").Value
        
'        作業受託
        iwbS05.Range("B74:B79").Value = iwbS01.Range("B59:B64").Value
        iwbS05.Range("D74:D79").Value = iwbS01.Range("E59:E64").Value
        iwbS05.Range("L74:L79").Value = iwbS01.Range("K59:K64").Value
        
'        受託（販売まで委託）
        iwbS05.Range("B84:B85").Value = iwbS01.Range("D70:D71").Value
        iwbS05.Range("D84:D85").Value = iwbS01.Range("G70:G71").Value
        iwbS05.Range("I84:I85").Value = iwbS01.Range("L70:L71").Value
        iwbS05.Range("K84:K85").Value = iwbS01.Range("P70:P71").Value
        
'        加工販売
        iwbS05.Range("B90:B95").Value = iwbS01.Range("D30:D35").Value
        iwbS05.Range("I90:I95").Value = iwbS01.Range("L30:L35").Value
        iwbS05.Range("K90:K95").Value = iwbS01.Range("P30:P35").Value
        
'        現状と目標の数値データの印刷
        iwbS05.PrintOut

'        聞き取りデータの保存
        SetFileName = FileName3 & "\" & TmpChar & SetName & "④　経営状況調査表.xlsx"
        ImportWbook2.SaveAs FileName:=SetFileName
        ImportWbook2.Close
        ImportWbook.Close
        Application.DisplayAlerts = True
    Next
    
'    処理時間結果
    endTime = Timer
    processTime = endTime - startTime
    MsgBox "印刷が終了しました。時間は：" & processTime & "秒です。"
    
'   実施登録
    twbS00.Range("A8").Interior.ColorIndex = 44
    twbS00.Range("A8").Font.ColorIndex = 1
    twbS00.Range("E8") = "" & Now & ""

Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
    
End Sub

Private Sub CommandButton6_Click()
'⑥ 案内時に送付する書類（アンケート）作成
    Dim tmp As Integer, sr As Integer, LastRow As Integer, SetRow As Integer, SetDate As Integer
    Dim i As Integer, SearchValue As Integer, LastRow2 As Integer, k As Integer, j As Integer, fileCount As Integer
    Dim arrayI() As Variant
    Dim OpenFileName As String, FileName As String, FileName2 As String, FileName3 As String, Path As String, SetFile As String, ThisWbookPass As String, SetFileName As String
    Dim SetName As String, ReturnValue As String, myFileName As String, TmpChar As String, SaveDir As String, SaveDir2 As String, FileNames As String, FileNames2 As String, FileNames3 As String
    Dim memberName As String, TempNumber As String
    Dim FoundValue As Object
    Dim ThisWbook As Workbook, ImportWbook As Workbook, ImportWbook2 As Workbook, ImportWbook3 As Workbook
    Dim twbS01 As Worksheet, twbS00 As Worksheet, twbS02 As Worksheet, iwbS01 As Worksheet, iwbS04 As Worksheet, iwbS05 As Worksheet
    Dim iwbS11 As Worksheet, iwbS14 As Worksheet, iwbS15 As Worksheet
    Dim iwbS06 As Worksheet, iwbS07 As Worksheet, iwbS08 As Worksheet, iwbS09 As Worksheet, iwbS10 As Worksheet
    Dim Flag As Boolean
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
    Dim oFso As FileSystemObject
    Dim fso As FileSystemObject
    Set ThisWbook = ActiveWorkbook
    Set twbS01 = ThisWbook.Worksheets("0_日程表")
    Set twbS00 = ThisWbook.Worksheets("フォーム呼出")
On Error GoTo myError
    Unload Me
    
'    処理時間測定
    startTime = Timer
    
'    送付・準備書類フォルダの作成
    SaveDir = ActiveWorkbook.Path & "\00認定農業者データ"
    SaveDir2 = SaveDir & "\送付・準備書類"
    If (oFso Is Nothing) Then
        Set oFso = CreateObject("Scripting.FileSystemObject")
    End If
    If (oFso.FolderExists(SaveDir2) = False) Then
        '// フォルダが存在しない
        MsgBox "「送付・準備書類」という名前のフォルダを新規作成します"
        MkDir SaveDir & "\送付・準備書類"
        SaveDir2 = SaveDir & "\送付・準備書類"
    Else
    End If

'    アンケート、フォルダの作成
    FileName3 = SaveDir2 & "\アンケートフォルダ"
    If (oFso Is Nothing) Then
        Set oFso = CreateObject("Scripting.FileSystemObject")
    End If
    If (oFso.FolderExists(FileName3) = False) Then
        '// フォルダが存在しない
        MsgBox "「アンケートフォルダ」という名前のフォルダを新規作成します"
        MkDir SaveDir2 & "\アンケートフォルダ"
    Else
    End If
    
'    フォルダ内にあるファイルの自動読み込み
    '--- ファイルシステムオブジェクト ---'
    Set fso = CreateObject("Scripting.FileSystemObject")
    '--- ファイル数を格納する変数 ---'
    fileCount = fso.GetFolder(SaveDir).Files.Count
    For i = 1 To fileCount
'    For i = 34 To fileCount 'テストコード
        If i <= 9 Then
            TempNumber = "0" & CStr(i)
        Else
            TempNumber = i
        End If
        
'        作成した認定申請書の取込
        FileNames = Dir(SaveDir & "\*" & TempNumber & "*認定申請書R*")
        Application.DisplayAlerts = False
        Workbooks.Open FileName:=SaveDir & "\" & FileNames, ReadOnly:=True, UpdateLinks:=0
        Set ImportWbook = Workbooks.Open(SaveDir & "\" & FileNames)
        Set iwbS01 = ImportWbook.Worksheets("入力シート")
        Set iwbS04 = ImportWbook.Worksheets("審査表")
        Set iwbS05 = ImportWbook.Worksheets("簡易版")

'        アンケート用紙の取込
        FileNames2 = Dir(SaveDir2 & "\" & "③　新様式B13・14*")
        
'        アンケート用紙が指定の場所になかった場合は手動で取込。送付・準備書類フォルダにコピー
        If FileNames2 = "" Then
            tmp = MsgBox("③　新様式B13・14「農業経営改善計画の達成状況等について（アンケート）」【R3.3 末】" & vbCrLf & "ファイルを選択してください。", vbYesNo + vbQuestion, "確認")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,③　新様式B13・14*.xls?")
                If OpenFileName <> "False" Then
                    FileCopy OpenFileName, SaveDir2 & "\" & "③　新様式B13・14「農業経営改善計画の達成状況等について（アンケート）」【R3.3 末】.xlsx"
                    FileNames2 = Dir(SaveDir2 & "\" & "③　新様式B13・14*")
                Else
                    MsgBox "キャンセルされました", vbCritical
                    ImportWbook.Close False
                    Exit Sub
                End If
            Else
                MsgBox "キャンセルされました", vbCritical
                ImportWbook.Close False
                Unload Me
                Exit Sub
            End If
        Else
        End If
        
'        読み込んだアンケートファイルにデータをコピーする
        Workbooks.Open FileName:=SaveDir2 & "\" & FileNames2, ReadOnly:=False, UpdateLinks:=0
        Set ImportWbook2 = Workbooks.Open(SaveDir2 & "\" & FileNames2)
        Set iwbS09 = ImportWbook2.Worksheets("達成状況（新様式）")
        Set iwbS10 = ImportWbook2.Worksheets("新規認定（新様式）")
        
'        新規か更新かの判定
        If iwbS01.Range("D4") = "新規" Then
'        新規受付者のデータコピー
            If iwbS01.Range("D5") = "" Then
                SetName = iwbS01.Range("M5")
            Else
                SetName = iwbS01.Range("D5")
            End If
            With iwbS01
                iwbS10.Range("J9").Value = SetName
                TmpChar = .Range("D10").Value
                TmpChar = Replace(TmpChar, "天草市", "")
                iwbS10.Range("S13").Value = TmpChar
            End With
            iwbS10.Activate
            iwbS10.PrintOut
        Else
'        更新受付者のデータコピー
            If iwbS01.Range("D5") = "" Then
                SetName = iwbS01.Range("M5")
            Else
                SetName = iwbS01.Range("D5")
            End If
            With iwbS01
                iwbS09.Range("J8").Value = SetName
                TmpChar = .Range("D10").Value
                TmpChar = Replace(TmpChar, "天草市", "")
                iwbS09.Range("S12").Value = TmpChar
                
                If iwbS05.Range("BE22").Value = "" Then
                    iwbS09.Range("M16").Value = ""
                Else
                    iwbS09.Range("M16").Value = iwbS05.Range("BE22").Value
                End If
                
                If iwbS05.Range("CX53").Value = 0 Then
                    iwbS09.Range("G31").Value = 0
                Else
                    iwbS09.Range("G31").Value = iwbS05.Range("CX53").Value - 1
                End If
                
                If .Range("B16").Value = "単一経営" Then
                    iwbS09.Range("S64").Value = .Range("B14").Value
                Else
                    For j = 4 To 17
                        If .Range("B14").Value = 1 Then
                            iwbS09.Range("S66").Value = .Cells(j, 58).Value
                        ElseIf .Range("B14").Value = 2 Then
                            iwbS09.Range("W66").Value = .Cells(j, 58).Value
                        End If
                    Next
                End If
                
                If iwbS05.Range("CX53").Value = 1 Then
                    iwbS09.Range("J84").Value = "〇"
                Else
                    iwbS09.Range("S84").Value = "〇"
                End If
                
                If .Range("X54").Value < 1000000 Then
                    iwbS09.Range("K87").Value = "〇"
                ElseIf .Range("X54").Value > 1000000 And .Range("X54").Value < 2000000 Then
                    iwbS09.Range("K88").Value = "〇"
                ElseIf .Range("X54").Value > 2000000 And .Range("X54").Value < 3000000 Then
                    iwbS09.Range("K89").Value = "〇"
                ElseIf .Range("X54").Value > 3000000 And .Range("X54").Value < 4000000 Then
                    iwbS09.Range("K90").Value = "〇"
                ElseIf .Range("X54").Value > 4000000 And .Range("X54").Value < 5000000 Then
                    iwbS09.Range("K91").Value = "〇"
                ElseIf .Range("X54").Value > 5000000 And .Range("X54").Value < 6000000 Then
                    iwbS09.Range("K92").Value = "〇"
                ElseIf .Range("X54").Value > 6000000 And .Range("X54").Value < 7000000 Then
                    iwbS09.Range("K93").Value = "〇"
                ElseIf .Range("X54").Value > 7000000 And .Range("X54").Value < 8000000 Then
                    iwbS09.Range("K94").Value = "〇"
                ElseIf .Range("X54").Value > 8000000 And .Range("X54").Value < 9000000 Then
                    iwbS09.Range("K95").Value = "〇"
                ElseIf .Range("X54").Value > 9000000 And .Range("X54").Value < 10000000 Then
                    iwbS09.Range("K96").Value = "〇"
                ElseIf .Range("X54").Value > 10000000 And .Range("X54").Value < 15000000 Then
                    iwbS09.Range("K97").Value = "〇"
                ElseIf .Range("X54").Value > 150000000 And .Range("X54").Value < 30000000 Then
                    iwbS09.Range("K98").Value = "〇"
                ElseIf .Range("X54").Value > 300000000 Then
                    iwbS09.Range("K99").Value = "〇"
                End If
                
                If .Range("AR54").Value < 1000000 Then
                    iwbS09.Range("O87").Value = "〇"
                ElseIf .Range("AR54").Value > 1000000 And .Range("AR54").Value < 2000000 Then
                    iwbS09.Range("O88").Value = "〇"
                ElseIf .Range("AR54").Value > 2000000 And .Range("AR54").Value < 3000000 Then
                    iwbS09.Range("O89").Value = "〇"
                ElseIf .Range("AR54").Value > 3000000 And .Range("AR54").Value < 4000000 Then
                    iwbS09.Range("O90").Value = "〇"
                ElseIf .Range("AR54").Value > 4000000 And .Range("AR54").Value < 5000000 Then
                    iwbS09.Range("O91").Value = "〇"
                ElseIf .Range("AR54").Value > 5000000 And .Range("AR54").Value < 6000000 Then
                    iwbS09.Range("O92").Value = "〇"
                ElseIf .Range("AR54").Value > 6000000 And .Range("AR54").Value < 7000000 Then
                    iwbS09.Range("O93").Value = "〇"
                ElseIf .Range("AR54").Value > 7000000 And .Range("AR54").Value < 8000000 Then
                    iwbS09.Range("O94").Value = "〇"
                ElseIf .Range("AR54").Value > 8000000 And .Range("AR54").Value < 9000000 Then
                    iwbS09.Range("O95").Value = "〇"
                ElseIf .Range("AR54").Value > 9000000 And .Range("AR54").Value < 10000000 Then
                    iwbS09.Range("O96").Value = "〇"
                ElseIf .Range("AR54").Value > 10000000 And .Range("AR54").Value < 15000000 Then
                    iwbS09.Range("O97").Value = "〇"
                ElseIf .Range("AR54").Value > 150000000 And .Range("AR54").Value < 30000000 Then
                    iwbS09.Range("O98").Value = "〇"
                ElseIf .Range("AR54").Value > 300000000 Then
                    iwbS09.Range("O99").Value = "〇"
                End If
    
                iwbS09.Range("H105").Value = iwbS05.Range("BS59").Value
                iwbS09.Range("N105").Value = iwbS05.Range("CE59").Value
                
                iwbS09.Range("O111").Value = .Range("L21").Value + .Range("L22").Value + .Range("L23").Value
                iwbS09.Range("O112").Value = .Range("L24").Value + .Range("L25").Value + .Range("L26").Value
                
                iwbS09.Range("V111").Value = .Range("P21").Value + .Range("P22").Value + .Range("P23").Value
                iwbS09.Range("V112").Value = .Range("P24").Value + .Range("P25").Value + .Range("P26").Value
                                    
                For j = 3 To 17
                    If .Cells(j, 58) = 1 Then
                        iwbS09.Range("E117").Value = .Cells(j, 59)
                    ElseIf .Cells(j, 58) = 2 Then
                        iwbS09.Range("E118").Value = .Cells(j, 59)
                    Else
                    End If
                Next
                
                LastRow = .Cells(.Rows.Count, 4).End(xlUp).Row
                
                For k = 2 To LastRow
                    If SetName = twbS01.Cells(k, 4) Then
                        If twbS01.Cells(k, 28) = "" Then
                            iwbS09.Range("M16").Value = ""
                        Else
                            iwbS09.Range("J25").Value = twbS01.Cells(k, 28)
                        End If
                        iwbS09.Range("I28").Value = twbS01.Cells(k, 27)
                    End If
                Next
                
                
            End With
            iwbS09.Activate
            iwbS09.PrintOut
        End If
        
        SetFileName = FileName3 & "\" & TempNumber & " " & SetName & "05_【R3.3 末】アンケート様式B13・B14（修正版）.xlsx"
        ImportWbook2.SaveAs FileName:=SetFileName
        ImportWbook2.Close
        
        ImportWbook.Close
        Application.DisplayAlerts = True
    Next
    
'    処理時間結果
    endTime = Timer
    processTime = endTime - startTime
    MsgBox "印刷が終了しました。時間は：" & processTime & "秒です。"
    
'   実施登録
    twbS00.Range("A7").Interior.ColorIndex = 44
    twbS00.Range("A7").Font.ColorIndex = 1
    twbS00.Range("E7") = "" & Now & ""
Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
End Sub
Private Sub CommandButton18_Click()
'⑧　封筒印刷　角2
    Dim sr As Integer, LastRow As Integer, Value As Integer
    Dim ThisWbook As Workbook
    Dim twbS08 As Worksheet, twbS01 As Worksheet, twbS00 As Worksheet
    Dim tmp As VbMsgBoxResult
    Set twbS08 = Worksheets("再認定案内印刷")
    Set twbS01 = Worksheets("0_日程表")
    Set twbS00 = Worksheets("フォーム呼出")
On Error GoTo myError
    MsgBox "印刷時にプリンター設定" & vbLf & "（印刷を行いたいプリンターを通常使用するプリンターに）" & vbLf & "しておいて下さい。", vbExclamation
    tmp = MsgBox("印刷したい方の「0_日程表のU列」にチェックを入れて置いて下さい。" & vbCrLf & "印刷プレビューを表示します。確認後、左上の×で1人ずつ確認してください。", vbYesNo)
    Unload Me
    With twbS01
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        Value = WorksheetFunction.CountA(twbS01.Range("U:U")) - 1
        If Value = 0 Then
            MsgBox "印刷したい方のU列にチェックを入れて再実行して下さい。", vbExclamation
            Exit Sub
        End If
        LastRow = .Cells(.Rows.Count, 3).End(xlUp).Row
        twbS08.Range("L4").NumberFormatLocal = "@"
        twbS08.Range("R4").NumberFormatLocal = "@"
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "●" Then
'                If .Cells(sr, 4) = "" Or .Cells(sr, 6) = "" Or .Cells(sr, 8) = "" Or .Cells(sr, 9) = "" Or .Cells(sr, 10) = "" Then
'                    MsgBox "データ未入力箇所があります。", vbExclamation
'                    Exit Sub
'                End If
                twbS08.Range("M4").Value = .Cells(sr, 14).Value
                twbS08.Range("N4").Value = .Cells(sr, 15).Value
                twbS08.Range("O4").Value = .Cells(sr, 16).Value
                twbS08.Range("R4").Value = .Cells(sr, 17).Value
                twbS08.Range("S4").Value = .Cells(sr, 18).Value
                twbS08.Range("T4").Value = .Cells(sr, 19).Value
                twbS08.Range("U4").Value = .Cells(sr, 20).Value
                
                twbS08.Range("F12").Value = .Cells(sr, 13).Value
                twbS08.Range("B15").Value = .Cells(sr, 4).Value
                
'                twbS08.Range("G26").Value = .Cells(sr, 6).Value
'                twbS08.Range("M26").Value = .Cells(sr, 7).Value
'                twbS08.Range("O26").Value = .Cells(sr, 8).Value
'                twbS08.Range("G28").Value = .Cells(sr, 9).Value
'                twbS08.Range("N28").Value = .Cells(sr, 10).Value
                twbS08.PrintPreview
            End If
        Next sr
        Application.ScreenUpdating = True
        tmp = MsgBox("プレビューを完了しました。続いて印刷します。よろしいですか？", vbYesNo)
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "●" Then
                twbS08.Range("M4").Value = .Cells(sr, 14).Value
                twbS08.Range("N4").Value = .Cells(sr, 15).Value
                twbS08.Range("O4").Value = .Cells(sr, 16).Value
                twbS08.Range("R4").Value = .Cells(sr, 17).Value
                twbS08.Range("S4").Value = .Cells(sr, 18).Value
                twbS08.Range("T4").Value = .Cells(sr, 19).Value
                twbS08.Range("U4").Value = .Cells(sr, 20).Value
                
                twbS08.Range("F12").Value = .Cells(sr, 13).Value
                twbS08.Range("B15").Value = .Cells(sr, 4).Value
                
'                twbS08.Range("G26").Value = .Cells(sr, 6).Value
'                twbS08.Range("M26").Value = .Cells(sr, 7).Value
'                twbS08.Range("O26").Value = .Cells(sr, 8).Value
'                twbS08.Range("G28").Value = .Cells(sr, 9).Value
'                twbS08.Range("N28").Value = .Cells(sr, 10).Value
                twbS08.PrintOut
            End If
        Next sr
    End With
    Application.ScreenUpdating = True
    twbS00.Range("A9").Interior.ColorIndex = 44
    twbS00.Range("A9").Font.ColorIndex = 1
    twbS00.Range("E9") = "" & Now & ""
    MsgBox "印刷を完了しました。"
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
End Sub

Private Sub CommandButton17_Click()
'⑧　通知案内封筒印刷（認定案内）　（長3）

    Dim sr As Integer, LastRow As Integer, Value As Integer
    Dim ThisWbook As Workbook
    Dim twbS12 As Worksheet, twbS01 As Worksheet, twbS00 As Worksheet
    Dim tmp As VbMsgBoxResult
    Set twbS12 = Worksheets("長３印刷 (案内)")
    Set twbS01 = Worksheets("0_日程表")
    Set twbS00 = Worksheets("フォーム呼出")
On Error GoTo myError
    MsgBox "印刷時にプリンター設定" & vbLf & "（印刷を行いたいプリンターを通常使用するプリンターに）" & vbLf & "しておいて下さい。", vbExclamation
    tmp = MsgBox("印刷したい方の「0_日程表のU列」にチェックを入れて置いて下さい。" & vbCrLf & "印刷プレビューを表示します。確認後、左上の×で1人ずつ確認してください。", vbYesNo)
    Unload Me
    With twbS01
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 3).End(xlUp).Row
        Value = WorksheetFunction.CountA(twbS01.Range("U:U")) - 1
        If Value = 0 Then
            MsgBox "印刷したい方のU列にチェックを入れて再実行して下さい。", vbExclamation
            Exit Sub
        End If
        twbS12.Range("L4").NumberFormatLocal = "@"
        twbS12.Range("R4").NumberFormatLocal = "@"
        
'        印刷プレビュー
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "●" Then
                twbS12.Range("E1").Value = .Cells(sr, 14).Value
                twbS12.Range("F1").Value = .Cells(sr, 15).Value
                twbS12.Range("G1").Value = .Cells(sr, 16).Value
                twbS12.Range("H1").Value = .Cells(sr, 17).Value
                twbS12.Range("I1").Value = .Cells(sr, 18).Value
                twbS12.Range("J1").Value = .Cells(sr, 19).Value
                twbS12.Range("K1").Value = .Cells(sr, 20).Value
                
                twbS12.Range("C3").Value = .Cells(sr, 13).Value
                twbS12.Range("C5").Value = .Cells(sr, 4).Value
                twbS12.PrintPreview
            End If
        Next sr
        Application.ScreenUpdating = True
        
'        プレビューの印刷
        tmp = MsgBox("プレビューを完了しました。続いて印刷します。よろしいですか？", vbYesNo)
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "●" Then
                twbS12.Range("E1").Value = .Cells(sr, 14).Value
                twbS12.Range("F1").Value = .Cells(sr, 15).Value
                twbS12.Range("G1").Value = .Cells(sr, 16).Value
                twbS12.Range("H1").Value = .Cells(sr, 17).Value
                twbS12.Range("I1").Value = .Cells(sr, 18).Value
                twbS12.Range("J1").Value = .Cells(sr, 19).Value
                twbS12.Range("K1").Value = .Cells(sr, 20).Value
                
                twbS12.Range("C3").Value = .Cells(sr, 13).Value
                twbS12.Range("C5").Value = .Cells(sr, 4).Value
                twbS12.PrintOut
            End If
        Next sr
    End With
    Application.ScreenUpdating = True
    
 '   実施登録
    twbS00.Range("A9").Interior.ColorIndex = 44
    twbS00.Range("A9").Font.ColorIndex = 1
    twbS00.Range("E9") = "" & Now & ""
    MsgBox "印刷を完了しました。"
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation

End Sub

Private Sub CommandButton7_Click()
'⑩　通知案内封筒印刷（更新受付日時案内）　（長3）

    Dim sr As Integer, LastRow As Integer, Value As Integer
    Dim ThisWbook As Workbook
    Dim twbS12 As Worksheet, twbS01 As Worksheet, twbS00 As Worksheet
    Dim tmp As VbMsgBoxResult
    Set twbS12 = Worksheets("長３印刷")
    Set twbS01 = Worksheets("0_日程表")
    Set twbS00 = Worksheets("フォーム呼出")
On Error GoTo myError
    MsgBox "印刷時にプリンター設定" & vbLf & "（印刷を行いたいプリンターを通常使用するプリンターに）" & vbLf & "しておいて下さい。", vbExclamation
    tmp = MsgBox("印刷したい方の「0_日程表のU列」にチェックを入れて置いて下さい。" & vbCrLf & "印刷プレビューを表示します。確認後、左上の×で1人ずつ確認してください。", vbYesNo)
    Unload Me
    With twbS01
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 3).End(xlUp).Row
        Value = WorksheetFunction.CountA(twbS01.Range("U:U")) - 1
        If Value = 0 Then
            MsgBox "印刷したい方のU列にチェックを入れて再実行して下さい。", vbExclamation
            Exit Sub
        End If
        twbS12.Range("L4").NumberFormatLocal = "@"
        twbS12.Range("R4").NumberFormatLocal = "@"
        
'        印刷プレビュー
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "●" Then
                If .Cells(sr, 4) = "" Or .Cells(sr, 6) = "" Or .Cells(sr, 8) = "" Or .Cells(sr, 9) = "" Or .Cells(sr, 10) = "" Then
                    MsgBox "データ未入力箇所があります。", vbExclamation
                    Exit Sub
                End If
                twbS12.Range("E1").Value = .Cells(sr, 14).Value
                twbS12.Range("F1").Value = .Cells(sr, 15).Value
                twbS12.Range("G1").Value = .Cells(sr, 16).Value
                twbS12.Range("H1").Value = .Cells(sr, 17).Value
                twbS12.Range("I1").Value = .Cells(sr, 18).Value
                twbS12.Range("J1").Value = .Cells(sr, 19).Value
                twbS12.Range("K1").Value = .Cells(sr, 20).Value
                
                twbS12.Range("C3").Value = .Cells(sr, 13).Value
                twbS12.Range("C5").Value = .Cells(sr, 4).Value
                twbS12.Range("T7").Value = .Cells(sr, 6).Value
                twbS12.Range("X7").Value = .Cells(sr, 7).Value
                twbS12.Range("AA7").Value = .Cells(sr, 8).Value
                twbS12.Range("S8").Value = .Cells(sr, 9).Value
                twbS12.Range("Y8").Value = .Cells(sr, 10).Value
                twbS12.PrintPreview
            End If
        Next sr
        Application.ScreenUpdating = True
        
'        プレビュー印刷
        tmp = MsgBox("プレビューを完了しました。続いて印刷します。よろしいですか？", vbYesNo)
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "●" Then
                twbS12.Range("E1").Value = .Cells(sr, 14).Value
                twbS12.Range("F1").Value = .Cells(sr, 15).Value
                twbS12.Range("G1").Value = .Cells(sr, 16).Value
                twbS12.Range("H1").Value = .Cells(sr, 17).Value
                twbS12.Range("I1").Value = .Cells(sr, 18).Value
                twbS12.Range("J1").Value = .Cells(sr, 19).Value
                twbS12.Range("K1").Value = .Cells(sr, 20).Value
                
                twbS12.Range("C3").Value = .Cells(sr, 13).Value
                twbS12.Range("C5").Value = .Cells(sr, 4).Value
                twbS12.Range("T7").Value = .Cells(sr, 6).Value
                twbS12.Range("X7").Value = .Cells(sr, 7).Value
                twbS12.Range("AA7").Value = .Cells(sr, 8).Value
                twbS12.Range("S8").Value = .Cells(sr, 9).Value
                twbS12.Range("Y8").Value = .Cells(sr, 10).Value
                twbS12.PrintOut
            End If
        Next sr
    End With
    Application.ScreenUpdating = True
    
'   実施登録
    twbS00.Range("A11").Interior.ColorIndex = 44
    twbS00.Range("A11").Font.ColorIndex = 1
    twbS00.Range("E11") = "" & Now & ""
    MsgBox "印刷を完了しました。"
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
End Sub
Private Sub CommandButton16_Click()
'（使用しない！）
'⑧事前聞き取りデータ印刷
'簡易版3ページ分印刷　＆　④新認定計画申請書と⑤　目標と措置の文例集印刷　＆　追加印刷（制度概要・意向確認書）
    Dim tmp As Integer, sr As Integer, LastRow As Integer, SetRow As Integer, SetDate As Integer
    Dim i As Integer, SearchValue As Integer, LastRow2 As Integer, k As Integer, j As Integer, fileCount As Integer
    Dim arrayI() As Variant
    Dim OpenFileName As String, FileName As String, FileName2 As String, FileName3 As String, Path As String, SetFile As String, ThisWbookPass As String, SetFileName As String
    Dim SetName As String, ReturnValue As String, myFileName As String, TmpChar As String, SaveDir As String, SaveDir2 As String, FileNames As String, FileNames2 As String, FileNames3 As String
    Dim FoundValue As Object
    Dim fso As FileSystemObject
    Dim ThisWbook As Workbook, ImportWbook As Workbook, ImportWbook2 As Workbook, ImportWbook3 As Workbook
    Dim twbS01 As Worksheet, twbS00 As Worksheet, twbS08 As Worksheet, iwbS01 As Worksheet, iwbS04 As Worksheet, iwbS05 As Worksheet
    Dim iwbS11 As Worksheet, iwbS14 As Worksheet, iwbS15 As Worksheet
    Dim iwbS06 As Worksheet, iwbS07 As Worksheet, iwbS08 As Worksheet, iwbS09 As Worksheet
    Dim Flag As Boolean
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
    Set ThisWbook = ActiveWorkbook
    Set twbS01 = ThisWbook.Worksheets("0_日程表")
    Set twbS00 = ThisWbook.Worksheets("フォーム呼出")
On Error GoTo myError
    
    Unload Me
'    処理時間測定
    startTime = Timer
    
    SaveDir = ActiveWorkbook.Path & "\00認定農業者データ"
    SaveDir2 = ActiveWorkbook.Path

'    フォルダ内にあるファイルの自動読み込み
    '--- ファイルシステムオブジェクト ---'
    Set fso = CreateObject("Scripting.FileSystemObject")
    '--- ファイル数を格納する変数 ---'
    fileCount = fso.GetFolder(SaveDir).Files.Count
    For i = 1 To fileCount
    
'    ワード文書印刷（まとめて印刷する場合）

'        Dim SaveDir As String
'        Dim wd As Object
'        SaveDir = ActiveWorkbook.Path & "\送付・準備書類\"
'        Set wd = CreateObject("Word.application")
'        wd.Visible = True
'        wd.documents.Open FileName:=SaveDir & "①　制度概要（改.doc" '印刷したい文書
'        wd.ActiveDocument.PrintOut Background:=False
'
'        wd.documents.Open FileName:=SaveDir & "②　新 意向確認書（令和）.doc" '印刷したい文書
'        wd.ActiveDocument.PrintOut Background:=False
'
'        wd.Quit
'        Set wd = Nothing

'     ワード文書印刷ここまで

        If i <= 9 Then
            TmpChar = "0" & CStr(i)
        Else
            TmpChar = i
        End If
        FileNames = Dir(SaveDir & "\*" & TmpChar & "認定申請書R*")
        FileNames2 = Dir(ActiveWorkbook.Path & "\⑤　新認定計画申請書（下書き用数値のみ）*")
        FileNames3 = Dir(ActiveWorkbook.Path & "\⑥　目標と措置の文例集*")
        
'        ⑤　新認定計画申請書が指定の場所になかった場合は手動で取込。送付・準備書類フォルダにコピー
        If FileNames2 = "" Then
            tmp = MsgBox("⑤　新認定計画申請書（下書き用数値のみ）" & vbCrLf & "ファイルを選択してください。", vbYesNo + vbQuestion, "確認")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,⑤　新認定計画申請書（下書き用数値のみ）*.xls?")
                If OpenFileName <> "False" Then
                    FileCopy OpenFileName, SaveDir2 & "\" & "⑤　新認定計画申請書（下書き用数値のみ）.xlsx"
                    FileNames2 = Dir(SaveDir2 & "\" & "⑤　新認定計画申請書（下書き用数値のみ）*")
                Else
                    MsgBox "キャンセルされました", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "キャンセルされました", vbCritical
                ImportWbook.Close False
                Unload Me
                Exit Sub
            End If
        Else
        End If
        
        Application.DisplayAlerts = False
        Workbooks.Open FileName:=SaveDir & "\" & FileNames, ReadOnly:=True, UpdateLinks:=0
        Set ImportWbook = Workbooks.Open(SaveDir & "\" & FileNames)
        Set iwbS01 = ImportWbook.Worksheets("入力シート")
        Set iwbS04 = ImportWbook.Worksheets("簡易版")
        
'        現状と目標の数値データ入力
        Workbooks.Open FileName:=SaveDir2 & "\" & FileNames2, ReadOnly:=False, UpdateLinks:=0
        Set ImportWbook2 = Workbooks.Open(SaveDir2 & "\" & FileNames2)
        Set iwbS05 = ImportWbook2.Worksheets("現状と目標")
        If iwbS01.Range("D5") = "" Then
            SetName = iwbS01.Range("M5")
        Else
            SetName = iwbS01.Range("D5")
        End If
        iwbS05.Range("B5") = SetName
        
        iwbS05.Range("B13:B18").Value = iwbS01.Range("B42:B47").Value
        iwbS05.Range("C13:C18").Value = iwbS01.Range("H42:H47").Value
        iwbS05.Range("D13:D18").Value = iwbS01.Range("L42:L47").Value
        iwbS05.Range("E13:E18").Value = iwbS01.Range("T42:T47").Value

        iwbS05.Range("B24:B29").Value = iwbS01.Range("B48:B53").Value
        iwbS05.Range("C24:C29").Value = iwbS01.Range("H48:H53").Value
        iwbS05.Range("D24:D29").Value = iwbS01.Range("L48:L53").Value
        iwbS05.Range("E24:E29").Value = iwbS01.Range("T48:T53").Value

        For j = 5 To 8
            If iwbS01.Cells(j, 34) = "●" Then
            iwbS05.Cells(j, 8) = iwbS01.Cells(j, 21)
            End If
        Next
        
'        目標と措置の文例集取込
'        ⑥　目標と措置の文例集が指定の場所になかった場合は手動で取込。送付・準備書類フォルダにコピー
        If FileNames3 = "" Then
            tmp = MsgBox("⑥　目標と措置の文例集" & vbCrLf & "ファイルを選択してください。", vbYesNo + vbQuestion, "確認")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,⑥　目標と措置の文例集*.xls?")
                If OpenFileName <> "False" Then
                    FileCopy OpenFileName, SaveDir2 & "\" & "⑥　目標と措置の文例集.xlsm"
                    FileNames3 = Dir(SaveDir2 & "\" & "⑥　目標と措置の文例集*")
                Else
                    MsgBox "キャンセルされました", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "キャンセルされました", vbCritical
                ImportWbook.Close False
                Unload Me
                Exit Sub
            End If
        Else
        End If
        Workbooks.Open FileName:=SaveDir2 & "\" & FileNames3, ReadOnly:=True, UpdateLinks:=0
        Set ImportWbook3 = Workbooks.Open(SaveDir2 & "\" & FileNames3)
        Set iwbS06 = ImportWbook3.Worksheets("③生産方式")
        Set iwbS07 = ImportWbook3.Worksheets("④経営管理")
        Set iwbS08 = ImportWbook3.Worksheets("⑤農業従事")
        Set iwbS09 = ImportWbook3.Worksheets("⑥その他")

'        申請書簡易版の印刷
        iwbS04.PrintOut From:=1, To:=1
        
'        現状と目標の数値データの印刷
        iwbS05.PrintOut
        
'        目標と措置の文例集印刷
        iwbS06.PrintOut
        iwbS07.PrintOut
        iwbS08.PrintOut
        iwbS09.PrintOut

'        聞き取りデータの保存
        SetFileName = SaveDir2 & "\00認定農業者データ\送付・準備書類\アンケートフォルダ\" & TmpChar & SetName & "⑤　新認定計画申請書（下書き用数値のみ）.xlsx"
        ImportWbook2.SaveAs FileName:=SetFileName
        ImportWbook2.Close
        ImportWbook3.Close
        
        ImportWbook.Close
        Application.DisplayAlerts = True
    Next
    
'    処理時間結果
    endTime = Timer
    processTime = endTime - startTime
    MsgBox "印刷が終了しました。時間は：" & processTime & "秒です。"
    
'   実施登録
    twbS00.Range("A9").Interior.ColorIndex = 44
    twbS00.Range("A9").Font.ColorIndex = 1
    twbS00.Range("E9") = "" & Now & ""

Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
End Sub

Private Sub CommandButton8_Click()
'⑨　カレンダー作成（0_日程表のデータより）

Dim Week_Value As Integer, sr As Integer, tmp As Integer, i As Integer, k As Integer, sc As Integer, j As Integer
    Dim arrayDate() As Variant, arrayName() As Variant, arrayNo() As Variant, arrayGroup() As Variant
    Dim arrayTime() As Variant, arrayHole1() As Variant, arrayHole2() As Variant, arrayWeek() As Variant
    Dim arrayTempNumber() As Variant, arrayWeekValue() As Variant, arrayEarlyweek() As Variant
    Dim SetDate As Date
    Dim twbS00 As Worksheet, twbS01 As Worksheet, twbS02 As Worksheet, twbS03 As Worksheet
    Dim ThisWbook As Workbook
    Set twbS00 = Worksheets("フォーム呼出")
    Set twbS01 = Worksheets("カレンダー")
    Set twbS02 = Worksheets("0_日程表")
    Set twbS03 = Worksheets("2_申請・認定日")
On Error GoTo myError
    Unload Me
    tmp = twbS02.Range("F" & Rows.Count).End(xlUp).Row
    
'    0_日程表シートのデータを日付順→時間順で指定範囲のソートを行う
    With twbS02
        .Sort.SortFields.clear
        .Sort.SortFields.Add _
            Key:=ActiveSheet.Cells(1, 6), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Sort.SortFields.Add _
            Key:=ActiveSheet.Cells(1, 8), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        With .Sort
            .SetRange Range(Cells(1, 2), Cells(tmp, 41))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    arrayWeekValue = Array(1, 6, 11, 16, 21)
    arrayTempNumber = Array(1, 12, 23, 34, 45, 56)
    
'    列の値の削除
    twbS01.Range("4:11").ClearContents
    twbS01.Range("15:22").ClearContents
    twbS01.Range("26:33").ClearContents
    twbS01.Range("37:44").ClearContents
    twbS01.Range("48:55").ClearContents
    twbS01.Range("59:66").ClearContents
    
    With twbS01
        Application.ScreenUpdating = False
        SetDate = WorksheetFunction.Small(twbS02.Range("F2" & ":" & "F" & tmp), 1)
        If Weekday(SetDate, 3) = 7 Then
            twbS01.Range("A1") = SetDate
        Else
           twbS01.Range("A1") = SetDate - Weekday(SetDate, 3)
        End If
        
'        データが１行だった場合
        If tmp = 2 Then
            tmp = 3
        End If
        
        arrayDate = twbS02.Range("F2" & ":" & "F" & tmp)
        arrayWeek = twbS02.Range("W2" & ":" & "W" & tmp)
        Week_Value = twbS02.Range("W2")
        arrayNo = twbS02.Range("A2" & ":" & "A" & tmp)
        arrayGroup = twbS02.Range("E2" & ":" & "E" & tmp)
        arrayName = twbS02.Range("D2" & ":" & "D" & tmp)
        arrayTime = twbS02.Range("H2" & ":" & "H" & tmp)
        arrayHole1 = twbS02.Range("I2" & ":" & "I" & tmp)
        arrayHole2 = twbS02.Range("J2" & ":" & "J" & tmp)
        arrayEarlyweek = twbS02.Range("X2" & ":" & "X" & tmp)
        
        j = 1
        For i = LBound(arrayWeek) To UBound(arrayWeek)
            If Week_Value <> arrayWeek(i, 1) Then
                Week_Value = arrayWeek(i, 1)
                twbS01.Cells(arrayTempNumber(j), 1) = arrayEarlyweek(i, 1)
                j = j + 1
            End If
        Next
        
'        0_日程表データをカレンダーへ転記する
        For sr = LBound(arrayTempNumber) To UBound(arrayTempNumber)
            For sc = LBound(arrayWeekValue) To UBound(arrayWeekValue)
                k = arrayTempNumber(sr) + 3
                For i = LBound(arrayDate) To UBound(arrayDate)
                    If twbS01.Cells(arrayTempNumber(sr), arrayWeekValue(sc)) = arrayDate(i, 1) Then
                        twbS01.Cells(k, arrayWeekValue(sc)) = arrayNo(i, 1)
                        twbS01.Cells(k, arrayWeekValue(sc) + 1) = arrayGroup(i, 1)
                        twbS01.Cells(k, arrayWeekValue(sc) + 2) = arrayName(i, 1)
                        twbS01.Cells(k, arrayWeekValue(sc) + 3) = arrayTime(i, 1)
                        twbS01.Cells(k, arrayWeekValue(sc) + 4) = arrayHole1(i, 1) & arrayHole2(i, 1)
                        k = k + 1
                    End If
                Next
            Next
        Next
        
'        0_日程表から2_申請・認定日シートへ氏名転記
        tmp = twbS03.Range("E" & Rows.Count).End(xlUp).Row
        twbS03.Range("E2" & ":" & "E" & tmp).ClearContents
        sr = 3
        For i = LBound(arrayName) To UBound(arrayName)
            twbS03.Cells(sr, 5) = arrayName(i, 1)
            sr = sr + 1
        Next i
        .Select
        MsgBox "完了しました。"
        Application.ScreenUpdating = True
        
'       実施登録
        twbS00.Range("A10").Interior.ColorIndex = 44
        twbS00.Range("A10").Font.ColorIndex = 1
        twbS00.Range("E10") = "" & Now & ""
    End With
    Exit Sub
myError:
    MsgBox "未入力があります！処理を終了します。", vbExclamation
End Sub

Private Sub CommandButton9_Click()
'⑪　所得金額取得（所得金額、営農類型、連名取込）

Dim tmp As Integer, ReturnValue As Integer, TempNumber As Integer, LastRow As Integer, i As Integer
    Dim SetName As String, OpenFileName As String, myFileName As Integer
    Dim FileName As String, Path As String, SetFile As String
    Dim ThisWbook As Workbook, ImportWbook As Workbook
    Dim FoundValue As Range
    Dim twbS00 As Worksheet, twbS03 As Worksheet, twbS04 As Worksheet
    Dim iwbS01 As Worksheet, iwbS02 As Worksheet, iwbS03 As Worksheet
    Dim sr As Integer, sr_Addend As Integer, SetRow_Next As Integer
    Set ThisWbook = ActiveWorkbook
    Set twbS00 = Worksheets("フォーム呼出")
    Set twbS04 = Worksheets("3_一覧兼審査表")
    Set twbS03 = Worksheets("2_申請・認定日")
On Error GoTo myError
    Application.DisplayAlerts = False
    
'    認定番号順に並べ替え
    twbS03.Range("A2").Sort key1:=twbS03.Range("A3"), order1:=xlAscending, Header:=xlYes
    
'    データ入力値以下の50行を削除を行う
    LastRow = twbS03.Cells(twbS03.Rows.Count, 4).End(xlUp).Row
    sr_Addend = LastRow + 1
    SetRow_Next = LastRow + 50
    twbS03.Rows(sr_Addend & ":" & SetRow_Next).Delete
    sr_Addend = LastRow + 2
    SetRow_Next = LastRow + 50
    twbS04.Rows(sr_Addend & ":" & SetRow_Next).Delete
    
'    一覧兼審査表を表示
    twbS04.Select
    Application.Wait Now + TimeValue("00:00:02")
    LastRow = twbS03.Cells(twbS03.Rows.Count, 5).End(xlUp).Row
    Unload UserForm5
    Application.ScreenUpdating = False
    

    For i = 3 To LastRow
Label1:
Label2:
'        未入力チェック
        If twbS03.Cells(i, 2).Value = "" Then
            twbS03.Select
            MsgBox "2_申請・認定日の区分を入力して下さい。"
            Unload Me
            Exit Sub
        End If

'        申請書の取込を行い共同申請名・営農類型・目標所得の値取得
        If twbS03.Cells(i, 2).Value Like "[新規,再認定,変更]*" And twbS04.Cells(i, 11).Value = "" And twbS04.Cells(i, 12).Value = "" Then
            tmp = MsgBox("登録されていない " & twbS03.Cells(i, 5) & " 様の申請書を選択してください。", vbYesNo + vbQuestion, "確認")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
                If OpenFileName <> "False" Then
                     SetFile = OpenFileName
                Else
                    MsgBox "キャンセルされました", vbCritical
                    Unload Me
                    Exit Sub
                End If
                Workbooks.Open FileName:=SetFile, ReadOnly:=True, UpdateLinks:=0
                Set ImportWbook = Workbooks.Open(Path & SetFile)
                Set iwbS01 = ImportWbook.Worksheets("入力シート")
                Set iwbS02 = ImportWbook.Worksheets("審査表")
                Set iwbS03 = ImportWbook.Worksheets("簡易版")
                If iwbS01.Range("D5") = "" Then
                    SetName = iwbS01.Range("M5")
                Else
                    SetName = iwbS01.Range("D5")
                End If
                
'                共同申請があれば取り込む
                If SetName = twbS03.Cells(i, 5) Then
                    Set FoundValue = twbS03.Range("E:E").Find(What:=SetName, LookAt:=xlPart)
                    If Not FoundValue Is Nothing Then
                        sr = twbS03.Range("E:E").Find(SetName).Row
                        If iwbS01.Range("AH6") = "●" Then
                            iwbS01.Range("U6").Copy
                            twbS04.Range("F" & sr).PasteSpecial xlPasteValues
                        End If
                        If iwbS01.Range("AH7") = "●" Then
                            iwbS01.Range("U7").Copy
                            twbS04.Range("G" & sr).PasteSpecial xlPasteValues
                        End If
                        If iwbS01.Range("AH8") = "●" Then
                            iwbS01.Range("U8").Copy
                            twbS04.Range("H" & sr).PasteSpecial xlPasteValues
                        End If
                        If iwbS01.Range("AH9") = "●" Then
                            iwbS01.Range("U9").Copy
                            twbS04.Range("I" & sr).PasteSpecial xlPasteValues
                        End If
    
                        iwbS01.Range("B14").Copy
                        twbS04.Range("J" & sr).PasteSpecial xlPasteValues
                        
'                        所得金額（現状と目標）取得
                        If iwbS03.Range("CX53") = 1 Then
                            If iwbS03.Range("Z56") <> "" Then
                                iwbS03.Range("Z56").Copy
                                twbS04.Range("K" & sr).PasteSpecial xlPasteValuesAndNumberFormats
                            End If
                            If iwbS03.Range("AL56") <> "" Then
                                iwbS03.Range("AL56").Copy
                                twbS04.Range("L" & sr).PasteSpecial xlPasteValuesAndNumberFormats
                            End If
                        Else
'                        総所得と一人当たりの所得金額が異なる場合は、括弧書きで表示する
                            If iwbS03.Range("Z56") <> "" Then
                                twbS04.Range("K" & sr) = Format(iwbS03.Range("Z56"), "0") & vbLf & "（" & Format(iwbS03.Range("Z59"), "0") & "）"
                            End If
                            If iwbS03.Range("AL56") <> "" Then
                                twbS04.Range("L" & sr) = Format(iwbS03.Range("AL56"), "0") & vbLf & "（" & Format(iwbS03.Range("AL59"), "0") & "）"
                            End If
                        End If
                    Else
                        MsgBox "申請者一覧表と一致するデーターがありませんでした。もう一度確認してください。"
                        ImportWbook.Close False
                        GoTo Label1
                        End
                    End If
                Else
                    MsgBox twbS03.Cells(i, 5) & " 様の申請書ではありません。もう一度選択して下さい。"
                    ImportWbook.Close False
                    GoTo Label2
                End If
            Else
                MsgBox "処理を中断します", vbCritical
                Unload Me
                Exit Sub
            End If
            ImportWbook.Close False
        Else
            Flag = True
'             MsgBox twbS03.Cells(i, 5) & " 様のデータは入力されているようです。", vbYesNo + vbQuestion, "確認"
        End If
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    twbS04.Select
    Unload Me
    
'   実施登録
    twbS00.Range("A12").Interior.ColorIndex = 44
    twbS00.Range("A12").Font.ColorIndex = 1
    twbS00.Range("E12") = "" & Now & ""
    If Flag Then
        MsgBox "一部登録されているようです。上書き保存します。"
    Else
        MsgBox "登録完了しました。上書き保存します。"
    End If
    On Error Resume Next
    ActiveWorkbook.Save
    If Err.Number > 0 Then MsgBox "保存されませんでした"
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbCritical
End Sub
Private Sub CommandButton10_Click()
'⑫　一覧兼審査表をJA毎に作成
Dim sr As Integer, LastRow As Integer, tmp_Addend As Integer, i As Integer
    Dim sr_Addend As Integer, SetRow_Next As Integer
    Dim myFileName As String
    Dim Flag As Boolean
    Dim ThisWbook As Workbook
    Dim twbS01 As Worksheet, iwbS00 As Worksheet, twbS00 As Worksheet
On Error GoTo myError
    Set twbS01 = Worksheets("3_一覧兼審査表")
    Set twbS00 = Worksheets("フォーム呼出")
    Unload Me
    With twbS01
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        tmp_Addend = .Cells(.Rows.Count, 4).End(xlUp).Row
        Workbooks.Add
        twbS01.Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        ActiveWorkbook.Sheets("3_一覧兼審査表").Name = "審査一覧表"
        MsgBox "県送付用一覧審査表を保存します。保存場所を指定して下さい。"
        myFileName = Application.GetSaveAsFilename(InitialFileName:="審査一覧表", FileFilter:="Excelブック,*.xlsx")
        If myFileName <> "False" Then
            ActiveWorkbook.SaveAs FileName:=myFileName
            ActiveWorkbook.Close
        Else
            ActiveWorkbook.Close
            Exit Sub
        End If
        
'        ファイルの作成
        Workbooks.Add
        twbS01.Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        ActiveWorkbook.Sheets("3_一覧兼審査表").Name = "JA本渡五和"
        Set iwbS00 = ActiveWorkbook.Worksheets("JA本渡五和")
        iwbS00.Range("A3:N50").ClearContents
        tmp = 3
        
'        JA本渡割り当て
        For i = 3 To tmp_Addend
            If .Cells(i, 4) Like "*本渡北" Or .Cells(i, 4) Like "*栄町" Or .Cells(i, 4) Like "*八幡町" Or .Cells(i, 4) Like "*旭町" Or _
                .Cells(i, 4) Like "*本渡南" Or .Cells(i, 4) Like "*中央新町" Or .Cells(i, 4) Like "*中村町" Or .Cells(i, 4) Like "*本町本" Or _
                .Cells(i, 4) Like "*北浜町" Or .Cells(i, 4) Like "*川原町" Or .Cells(i, 4) Like "*北原町" Or .Cells(i, 4) Like "*本町新休" Or _
                .Cells(i, 4) Like "*今釜町" Or .Cells(i, 4) Like "*古川町" Or .Cells(i, 4) Like "*丸尾町" Or .Cells(i, 4) Like "*本町下河内" Or _
                .Cells(i, 4) Like "*今釜新町" Or .Cells(i, 4) Like "*南新町" Or .Cells(i, 4) Like "*瀬戸町" Or .Cells(i, 4) Like "*城川原" Or _
                .Cells(i, 4) Like "*東浜町" Or .Cells(i, 4) Like "*太田町" Or .Cells(i, 4) Like "*志柿" Or .Cells(i, 4) Like "*御領" Or _
                .Cells(i, 4) Like "*大浜町" Or .Cells(i, 4) Like "*東町" Or .Cells(i, 4) Like "*下浦町" Or .Cells(i, 4) Like "*御領" Or _
                .Cells(i, 4) Like "*城下町" Or .Cells(i, 4) Like "*浄南町" Or .Cells(i, 4) Like "*亀場町亀川" Or .Cells(i, 4) Like "*鬼池" Or _
                .Cells(i, 4) Like "*船之尾町" Or .Cells(i, 4) Like "*山の手町" Or .Cells(i, 4) Like "*亀場町食場" Or .Cells(i, 4) Like "*二江" Or _
                .Cells(i, 4) Like "*浜崎町" Or .Cells(i, 4) Like "*川原新町" Or .Cells(i, 4) Like "*楠浦町" Or .Cells(i, 4) Like "*手野一" Or _
                .Cells(i, 4) Like "*小松原町" Or .Cells(i, 4) Like "*諏訪町" Or .Cells(i, 4) Like "*枦宇土町" Or .Cells(i, 4) Like "*手野二" Or _
                .Cells(i, 4) Like "*港町" Or .Cells(i, 4) Like "*南町" Or .Cells(i, 4) Like "*宮地岳町" Then
                .Rows(i).Copy
                iwbS00.Rows(tmp).PasteSpecial Paste:=xlPasteValues
                tmp = tmp + 1
            End If
        Next i
        LastRow = iwbS00.Cells(iwbS00.Rows.Count, 5).End(xlUp).Row
        sr_Addend = LastRow + 1
        SetRow_Next = LastRow + 50
        iwbS00.Rows(sr_Addend & ":" & SetRow_Next).Delete
        ActiveWorkbook.Sheets("Sheet1").Delete
        Application.CutCopyMode = False
        
'        ファイルの保存
        MsgBox "JA本渡五和を保存します。保存場所を指定して下さい。"
        myFileName = Application.GetSaveAsFilename(InitialFileName:="JA本渡五和", FileFilter:="Excelブック,*.xlsx")

 '        JA本渡五和分データが空だったら、ファイルは作成しない。
        If tmp = 3 Then
            myFileName = False
        End If
        If myFileName <> "False" Then
            ActiveWorkbook.SaveAs FileName:=myFileName
            ActiveWorkbook.Close
        Else
            ActiveWorkbook.Close
            MsgBox "JA本渡五和分データが空だった為、ファイルは削除しました。"
        End If
        
'        ファイルの作成
        Workbooks.Add
        twbS01.Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        ActiveWorkbook.Sheets("3_一覧兼審査表").Name = "JAあまくさ"
        Set iwbS00 = ActiveWorkbook.Worksheets("JAあまくさ")
        iwbS00.Range("A3:N50").ClearContents
        tmp = 3
        
'        JAあまくさ割り当て
        For i = 3 To tmp_Addend
            If .Cells(i, 4) Like "*佐伊津町" Or .Cells(i, 4) Like "*馬場" Or .Cells(i, 4) Like "*河浦*" Or .Cells(i, 4) Like "*大江軍浦" Or _
                .Cells(i, 4) Like "*牛深町" Or .Cells(i, 4) Like "*古川" Or .Cells(i, 4) Like "*今富" Or .Cells(i, 4) Like "*下田南" Or _
                .Cells(i, 4) Like "*久玉町" Or .Cells(i, 4) Like "*湯船原" Or .Cells(i, 4) Like "*崎津" Or .Cells(i, 4) Like "*下田北" Or _
                .Cells(i, 4) Like "*魚貫町" Or .Cells(i, 4) Like "*赤崎" Or .Cells(i, 4) Like "*今富" Or .Cells(i, 4) Like "*高浜南" Or _
                .Cells(i, 4) Like "*二浦町早浦" Or .Cells(i, 4) Like "*須子" Or .Cells(i, 4) Like "*崎津" Or .Cells(i, 4) Like "*高浜北" Or _
                .Cells(i, 4) Like "*二浦町亀浦" Or .Cells(i, 4) Like "*大浦" Or .Cells(i, 4) Like "*立原" Or .Cells(i, 4) Like "*福連木" Or _
                .Cells(i, 4) Like "*深海町" Or .Cells(i, 4) Like "*楠甫" Or .Cells(i, 4) Like "*宮野河内" Or .Cells(i, 4) Like "*高浜北" Or _
                .Cells(i, 4) Like "*浦" Or .Cells(i, 4) Like "*上津浦" Or .Cells(i, 4) Like "*路木" Or .Cells(i, 4) Like "*大江向" Or _
                .Cells(i, 4) Like "*宮田" Or .Cells(i, 4) Like "*下津浦" Or .Cells(i, 4) Like "*久留" Or .Cells(i, 4) Like "*小宮地" Or _
                .Cells(i, 4) Like "*棚底" Or .Cells(i, 4) Like "*小島子" Or .Cells(i, 4) Like "*白木河内" Or .Cells(i, 4) Like "*大多尾" Or _
                .Cells(i, 4) Like "*打田" Or .Cells(i, 4) Like "*大島子" Or .Cells(i, 4) Like "*新合" Or .Cells(i, 4) Like "*大宮地" Or _
                .Cells(i, 4) Like "*河内" Or .Cells(i, 4) Like "*今田" Or .Cells(i, 4) Like "*大江" Or .Cells(i, 4) Like "*碇石" Or _
                .Cells(i, 4) Like "*新和町" Or .Cells(i, 4) Like "*牧島" Or .Cells(i, 4) Like "*横浦" Or .Cells(i, 4) Like "*御所浦町" Or .Cells(i, 4) Like "*中田" Then
                .Rows(i).Copy
                iwbS00.Rows(tmp).PasteSpecial Paste:=xlPasteValues
                tmp = tmp + 1
            End If
        Next i
        LastRow = iwbS00.Cells(iwbS00.Rows.Count, 5).End(xlUp).Row
        sr_Addend = LastRow + 1
        SetRow_Next = LastRow + 50
        iwbS00.Rows(sr_Addend & ":" & SetRow_Next).Delete
        ActiveWorkbook.Sheets("Sheet1").Delete
        Application.CutCopyMode = False

'        ファイルの保存
        MsgBox "JAあまくさを保存します。保存場所を指定して下さい。"
        myFileName = Application.GetSaveAsFilename(InitialFileName:="JAあまくさ", FileFilter:="Excelブック,*.xlsx")
        
'        JAあまくさ分データが空だったら、ファイルは作成しない。
        If tmp = 3 Then
            myFileName = False
        End If
        If myFileName <> "False" Then
            ActiveWorkbook.SaveAs FileName:=myFileName
            ActiveWorkbook.Close
        Else
            ActiveWorkbook.Close
            MsgBox "JAあまくさ分データが空だった為、ファイルは削除しました。"
        End If
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End With
    
'   実施登録
        twbS00.Range("A13").Interior.ColorIndex = 44
        twbS00.Range("A13").Font.ColorIndex = 1
        twbS00.Range("E13") = "" & Now & ""
        MsgBox "ファイルの分割が完了しました。"
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
End Sub

Private Sub CommandButton11_Click()
'⑬　認定証書送付封筒印刷（角２）
Dim sr As Integer, LastRow As Integer, Value As Integer
    Dim ThisWbook As Workbook
    Dim twbS08 As Worksheet, twbS03 As Worksheet, twbS00 As Worksheet
    Dim InputValue As Integer, i As Integer, sRow As Integer, tmpValue As Integer
    Dim saiNintei As Integer, henkou As Integer, shinki As Integer, zitai As Integer
    Dim SetMsg As String
    Dim tmp As VbMsgBoxResult
    Set twbS03 = Worksheets("2_申請・認定日")
    Set twbS08 = Worksheets("認定報告印刷")
    Set twbS00 = Worksheets("フォーム呼出")
On Error GoTo myError
    MsgBox "印刷する前にプリンター設定" & vbLf & "（印刷を行いたいプリンターを通常使用するプリンターに）" & vbLf & "しておいて下さい。", vbExclamation
    tmp = MsgBox("印刷する方の「C列」にチェックを入れておいて下さい。" & vbLf & "プレビュー後、1人ずつ左上の×で確認してください。", vbYesNo)
    Unload Me
    With twbS03
    If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        
'        印刷列チェック
        Value = WorksheetFunction.CountA(.Range("C:C")) - 1
        If Value = 0 Then
            MsgBox "印刷する場合、C列にチェックを入れて実行して下さい。", vbExclamation
            Exit Sub
        End If
        LastRow = .Cells(.Rows.Count, 5).End(xlUp).Row
        
'        区分の数値チェック
        For sr = 3 To LastRow
            If .Cells(sr, 2).Value = "再認定" Then
            saiNintei = saiNintei + 1
            End If
            If .Cells(sr, 2).Value = "変更" Then
            henkou = henkou + 1
            End If
            If .Cells(sr, 2).Value = "新規" Then
            shinki = shinki + 1
            End If
            If .Cells(sr, 2).Value = "辞退" Then
            zitai = zitai + 1
            End If
        Next
        tmpValue = .Cells(.Rows.Count, 11).End(xlUp).Row
        
'        印刷プレビュー
        For sr = 3 To LastRow
            If .Cells(sr, 2).Value <> "辞退" And .Cells(sr, 3).Value = "●" Then
                If Value = 0 Or twbS03.Cells(sr, 5) = "" Or twbS03.Cells(sr, 11) = "" Or twbS03.Cells(sr, 12) = "" Then
                    MsgBox "未入力の箇所があります。確認して再実行して下さい。", vbExclamation
                    Exit Sub
                End If
                twbS08.Range("M4").Value = .Cells(sr, 16).Value
                twbS08.Range("N4").Value = .Cells(sr, 17).Value
                twbS08.Range("O4").Value = .Cells(sr, 18).Value
                twbS08.Range("R4").Value = .Cells(sr, 19).Value
                twbS08.Range("S4").Value = .Cells(sr, 20).Value
                twbS08.Range("T4").Value = .Cells(sr, 21).Value
                twbS08.Range("U4").Value = .Cells(sr, 22).Value
                twbS08.Range("F12").Value = .Cells(sr, 10).Value
                twbS08.Range("B15").Value = .Cells(sr, 5).Value
                twbS08.PrintPreview EnableChanges:=True
            Else
            End If
        Next sr

'        区分の数値入力
        For i = 3 To 14
            If twbS00.Cells(i, 7) = Month(.Cells(tmpValue, 11)) Then
                sRow = twbS00.Cells(i, 7).Row
            End If
        Next i
        twbS00.Range("J" & sRow) = saiNintei
        twbS00.Range("K" & sRow) = shinki
        twbS00.Range("L" & sRow) = zitai
        twbS00.Range("M" & sRow) = henkou
        
'        プレビューの印刷
        Application.ScreenUpdating = True
        tmp = MsgBox("プレビューを完了しました。続いて印刷します。よろしいですか？", vbYesNo)
        If tmp = vbNo Then Exit Sub
            Application.ScreenUpdating = False
            LastRow = .Cells(.Rows.Count, 5).End(xlUp).Row
        For sr = 3 To LastRow
            If .Cells(sr, 2).Value <> "辞退" And .Cells(sr, 3).Value = "●" Then
                twbS08.Range("M4").Value = .Cells(sr, 16).Value
                twbS08.Range("N4").Value = .Cells(sr, 17).Value
                twbS08.Range("O4").Value = .Cells(sr, 18).Value
                twbS08.Range("R4").Value = .Cells(sr, 19).Value
                twbS08.Range("S4").Value = .Cells(sr, 20).Value
                twbS08.Range("T4").Value = .Cells(sr, 21).Value
                twbS08.Range("U4").Value = .Cells(sr, 22).Value
                twbS08.Range("F12").Value = .Cells(sr, 10).Value
                twbS08.Range("B15").Value = .Cells(sr, 5).Value
                twbS08.PrintOut
            Else
            End If
        Next sr
    End With
    Application.ScreenUpdating = True
    
'   実施登録
    twbS00.Range("A15").Interior.ColorIndex = 44
    twbS00.Range("A15").Font.ColorIndex = 1
    twbS00.Range("E15") = "" & Now & ""
    MsgBox "印刷が完了しました。"
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
End Sub

Private Sub CommandButton12_Click()
'⑭　申請書ファイルリネーム
Dim tmp As Integer, sr As Integer, LastRow As Integer, SetRow As Integer, SetDate As Integer
    Dim i As Integer, SearchValue As Integer, LastRow2 As Integer, k As Integer
    Dim OpenFileName As String, FileName As String, Path As String, SetFile As String, ThisWbookPass As String, SetFileName As String
    Dim SetName As String, ReturnValue As String, myFileName As String, TmpChar As String, SaveDir As String
    Dim FoundValue As Object
    Dim Flag As Boolean
    Dim ThisWbook As Workbook, ImportWbook2 As Workbook
    Dim twbS00 As Worksheet, twbS03 As Worksheet, twbS04 As Worksheet, iwbS11 As Worksheet
    Set ThisWbook = ActiveWorkbook
    Set twbS03 = Worksheets("2_申請・認定日")
    Set twbS04 = Worksheets("3_一覧兼審査表")
    Set twbS00 = Worksheets("フォーム呼出")
On Error GoTo myError
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    With twbS03
    LastRow = .Cells(.Rows.Count, 5).End(xlUp).Row
    UserForm6.Show
    
'    完了フォルダの作成
    tmp = MsgBox("申請書を保存するフォルダを同じ階層に作成します。", vbYesNo + vbQuestion, "確認")
        If tmp = vbYes Then
            Flag = True
            SaveDir = ActiveWorkbook.Path & "\00認定農業者データ（" & index & "月期分）完了"
            If Dir(SaveDir, vbDirectory) = "" Then
                MkDir SaveDir
                MsgBox "このファイルと同じ階層に「00認定農業者データ（" & index & "月期分）完了」という保存フォルダを作成しました。"
            End If
        Else
            Flag = False
            MsgBox "自動的に作成されたフォルダ内に保存されますので、フォルダは作成して下さい。"
        End If
        
'    新規と再認定のみファイル名の更新を行う
    For i = 3 To LastRow
        If .Cells(i, 2).Value = "新規" Or .Cells(i, 2).Value = "再認定" Or .Cells(i, 2).Value = "変更" Then
Label1:
            tmp = MsgBox(.Cells(i, 5).Value & " 様のファイルを選択してください。", vbYesNo + vbQuestion, "確認")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
                If OpenFileName <> "False" Then
                    SetFile = OpenFileName
                Else
                    MsgBox "キャンセルされました", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "キャンセルされました", vbCritical
                Unload Me
                Exit Sub
            End If
            Workbooks.Open FileName:=SetFile, ReadOnly:=True, UpdateLinks:=0
            Set ImportWbook2 = Workbooks.Open(Path & SetFile)
            Windows(ImportWbook2.Name).Visible = False
            Set iwbS11 = ImportWbook2.Worksheets("入力シート")
            
            If iwbS11.Range("D5") = "" Then
                SetName = iwbS11.Range("M5")
            Else
                SetName = iwbS11.Range("D5")
            End If
            If .Cells(i, 5).Value <> SetName Then
                MsgBox "セットしたファイル" & SetName & "様ではありません。再度選択して下さい。"
                ImportWbook2.Close
                GoTo Label1
            End If
            
            ImportWbook2.Activate
            
'            MsgBox "ファイル名（認定番号）が間違いないか確認し保存して下さい。"
            myFileName = Format(twbS00.Range("G1"), "[DBNum3][$]e") & "-" & .Cells(i, 1).Value & "認定申請書（" & SetName & " ）" & ".xlsm"
            ImportWbook2.SaveAs FileName:=SaveDir & "\" & myFileName
            Unload Me
            Windows(ImportWbook2.Name).Visible = True
            ImportWbook2.Close
        End If
    Next
        Application.DisplayAlerts = True
        MsgBox "申請書リネームが完了しました。"
        
 '   実施登録
        twbS00.Range("A16").Interior.ColorIndex = 44
        twbS00.Range("A16").Font.ColorIndex = 1
        twbS00.Range("E16") = "" & Now & ""
    End With
    Exit Sub
myError:
    MsgBox "選択したファイルが違います！処理を終了します。", vbExclamation
End Sub

Private Sub CommandButton13_Click()
'⑮　認定農業者データベース登録・削除
UserForm3.Show
End Sub

Private Sub CommandButton14_Click()
'⑯　認定農業者データベース編集（個人・法人）
UserForm4.Show
End Sub

Private Sub CommandButton15_Click()
'⑰　締処理
    Dim tmp As Integer, sr As Integer, sr0 As Integer, i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer, ReturnValue As Integer, Value As Integer
    Dim LastRow As Integer, nextIndex As Integer, nextYear As Integer
    Dim Flag As Boolean
    Dim TmpChar As String, myFileName As String, SaveDir As String, ParentDirName As String
    Dim twbS00 As Worksheet, twbS01 As Worksheet, twbS02 As Worksheet, twbS03 As Worksheet
    Dim twbS04 As Worksheet, twbS05 As Worksheet, twbS06 As Worksheet, twbS10 As Worksheet, twbS11 As Worksheet
    Dim ThisWbook As Workbook, ActiveWbook As Workbook
    Dim StringObject As Range, StringObject0 As Range
    Set ThisWbook = ActiveWorkbook
    Set twbS00 = Worksheets("フォーム呼出")
    Set twbS01 = Worksheets("カレンダー")
    Set twbS02 = Worksheets("0_日程表")
    Set twbS03 = Worksheets("2_申請・認定日")
    Set twbS04 = Worksheets("3_一覧兼審査表")
    Set twbS05 = Worksheets("data")
    Set twbS06 = Worksheets("月別抽出")
    Set twbS10 = Worksheets("目次-認定者")
    Set twbS11 = Worksheets("目次-辞退者")
On Error GoTo myError
    With twbS10
        Application.DisplayAlerts = False
        TmpChar = .Cells(.Rows.Count, 4).End(xlUp)
        Set StringObject = twbS03.Range("A2", "A200")
        
        If TmpChar = "発元又は宛名" Then
            MsgBox "本年度初回登録です。目次シートに登録します。"
            Flag = False
        Else
            MsgBox "本年度2回目以降の登録です。目次シートに登録します。"
            Flag = True
        End If
        
        Value = twbS03.Range("B" & Rows.Count).End(xlUp).Row
        sr = .Range("D" & Rows.Count).End(xlUp).Row + 1
        sr0 = twbS11.Range("D" & Rows.Count).End(xlUp).Row + 1
        For i = 2 To Value
            If WorksheetFunction.CountIf(.Range("D3:D122"), twbS03.Cells(i, 5)) > 0 Then
                MsgBox "同一名が存在します。" & vbCrLf & "重複登録ではありませんか？目次シートとチェックして下さい。" & vbCrLf & "一旦処理を終了します。"
                Unload Me
                Exit Sub
            End If
        Next i
        For i = 3 To Value
            If twbS03.Cells(i, 2).Value Like "[新規,再認定,変更]*" Then
                .Cells(sr, 1) = twbS03.Cells(i, 1)
                .Cells(sr, 3) = twbS03.Cells(i, 12)
                .Cells(sr, 4) = twbS03.Cells(i, 5)
                .Cells(sr, 5) = "農業経営改善計画認定申請書（" & twbS03.Cells(i, 2) & ")"
                sr = sr + 1
            ElseIf twbS03.Cells(i, 2).Value Like "辞退" Then
                twbS11.Cells(sr0, 3) = twbS03.Cells(i, 12)
                twbS11.Cells(sr0, 4) = twbS03.Cells(i, 5)
                twbS11.Cells(sr0, 5) = "農業経営改善計画認定申請書（" & twbS03.Cells(i, 2) & ")"
                sr0 = sr0 + 1
            Else
            End If
        Next i
        
        MsgBox "本業務の月数入力をお願いします。"
        UserForm6.Show
        
'   実施登録
        twbS00.Range("A19").Interior.ColorIndex = 44
        twbS00.Range("A19").Font.ColorIndex = 1
        twbS00.Range("E19") = "" & Now & ""
    End With
    
    Unload UserForm5
    SaveDir = ThisWbook.Path
    myFileName = "10申請者一覧表（" & Format(twbS00.Range("G1"), "[DBNum3][$]ggge") & "年度" & index & "月期）完了"
    ActiveWorkbook.SaveAs FileName:=SaveDir & "\" & myFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    MsgBox myFileName & "　名で保存しました。" & vbCrLf & "目次の認定番号のズレが無いか、各シートを確認して下さい。"

'    次月ファイルを作成する（年度末は作成しない）
    If index <> 3 Then
        MsgBox "次月ファイルを作成します。"
        If index = 12 Then
            nextIndex = 1
            nextYear = Format(twbS00.Range("G1"), "[$-ja-JP]e") + 1
        Else
            nextIndex = index + 1
            nextYear = Format(twbS00.Range("G1"), "[$-ja-JP]e")
        End If
        
        
    
        ParentDirName = Left(ThisWbook.Path, InStrRev(ThisWbook.Path, "\") - 1)

        SaveDir = ParentDirName & "\R" & nextYear & "." & nextIndex
        
        myFileName = "10申請者一覧表（" & Format(twbS00.Range("G1"), "[DBNum3][$]ggge") & "年度" & nextIndex & "月期）"
        If Dir(SaveDir, vbDirectory) = "" Then
            MkDir SaveDir
            MsgBox "一つ上の階層にR" & nextYear & "." & nextIndex & "という保存フォルダを作成し、その中に " & myFileName & "ファイルを作成します。"
        End If
    
        ActiveWorkbook.SaveAs FileName:=SaveDir & "\" & myFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        
        '初期化
        twbS00.Range("A2:A20").Interior.ColorIndex = 56
        twbS00.Range("A2:A20").Font.ColorIndex = 2
        twbS00.Range("E2:E20").ClearContents
        
        twbS01.Range("A4:Y11").ClearContents
        twbS01.Range("A15:Y22").ClearContents
        twbS01.Range("A26:Y33").ClearContents
        twbS01.Range("A37:Y44").ClearContents
        twbS01.Range("A48:Y55").ClearContents
        twbS01.Range("A59:Y66").ClearContents
        twbS01.Range("A1:E1").ClearContents
        twbS01.Range("A12:E12").ClearContents
        twbS01.Range("A23:E23").ClearContents
        twbS01.Range("A34:E34").ClearContents
        twbS01.Range("A45:E45").ClearContents
        twbS01.Range("A56:E56").ClearContents
        
        twbS02.Range("D2:F200").ClearContents
        twbS02.Range("H2:J200").ClearContents
        
        twbS03.Range("A3:C200").ClearContents
        twbS03.Range("E3:E200").ClearContents
        twbS03.Range("H3:H200").ClearContents
        twbS03.Range("K3:L200").ClearContents
        twbS03.Range("W3:W200").ClearContents
        
        twbS04.Range("F3:N200").ClearContents
        
        
        tmp = twbS05.Cells(twbS05.Rows.Count, 7).End(xlUp).Row + 1
        twbS05.Range("A" & tmp & ":" & "CB" & tmp + 10).ClearContents
        
        twbS06.Range("A4:CA200").ClearContents
        twbS06.Range("A4:CA200").Interior.ColorIndex = 0
        
        '認定番号設定
        tmp = twbS10.Cells(twbS10.Rows.Count, 3).End(xlUp).Row + 1
        
        twbS03.Range("A3") = twbS10.Cells(tmp, 1)
        ActiveWorkbook.Save
        Application.DisplayAlerts = True
        MsgBox "保存が完了しました。次月に使用して下さい。"
    Else
        MsgBox "認定農業者更新手続きお疲れさまでした。本年度は終了となります。"
    End If
    Exit Sub
myError:
    MsgBox "予期せぬエラーが発生しました！処理を終了します。", vbExclamation
End Sub


