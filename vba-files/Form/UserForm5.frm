VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "�F��_�ƎҎ�t�Ɩ��V�X�e��"
   ClientHeight    =   9030.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13040
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CommandButton1_Click()
'�@�@�F�菈�����N�i���j�Ɏ��s
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
    Set twbS00 = Worksheets("�t�H�[���ďo")
    Set twbS01 = Worksheets("0_�����\")
    Set twbS03 = Worksheets("2_�\���E�F���")
    Set twbS10 = Worksheets("�ڎ�-�F���")
    Set twbS11 = Worksheets("�ڎ�-���ގ�")
    Set twbS12 = Worksheets("�ڎ�-�Ĕ��s")
On Error GoTo myError
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    With twbS10
    
'        ���N�x�i��j�o�^�̊m�F
        tmp = MsgBox("�{�N�x����o�^�ł����H", vbYesNo + vbQuestion, "�m�F")
        If tmp = vbYes Then
            If twbS00.Range("J1") = "��" Then
                MsgBox "����ł͂���܂��񏈗����I�����܂��B" & vbCrLf & "�u�t�H�[���ďo�v�V�[�g��J1�Z�����Q�Ƃ��ĉ������B"
                Unload UserForm5
                Exit Sub
            End If
            
'            ������
            Flag = True
            twbS00.Range("J1") = "��"
            twbS10.Range("A3") = 1
            twbS11.Range("A3") = 1
            twbS12.Range("A3") = 1
            twbS00.Range("J3:M14").ClearContents
            
'            �t�H���_�̍쐬
'            MsgBox "�{�Ɩ��̌������͂����肢���܂��B"
'            UserForm6.Show
'            SaveDir = ActiveWorkbook.Path & "\" & Format(twbS00.Range("G1"), "[$-ja-JP]ge") & "." & index
'            If Dir(SaveDir, vbDirectory) = "" Then
'                MkDir SaveDir
'                MkDir SaveDir & "\�o�b�N�A�b�v�t�H���_"
'                MsgBox "�����K�w��" & Format(twbS00.Range("G1"), "[$-ja-JP]ge") & "." & index & "�Ƃ����ۑ��t�H���_���쐬���܂����B"
'            End If
        Else
            Flag = False
            If twbS00.Range("J1") <> "��" Then
                MsgBox "����ł͂���܂��񏈗����I�����܂��B" & vbCrLf & "�u�t�H�[���ďo�v�V�[�g��J1�Z�����Q�Ƃ��ĉ������B"
                
                Unload UserForm5
                Exit Sub
            Else
            
'            ����ł͂Ȃ������ꍇ�ɑO��X�V�E�V�K�E���ޓ��̃��X�g�i�ڎ��j���捞�ށB
                tmp = MsgBox("�O��܂ł́A�ڎ����捞�܂��̂őO��̐\�����ꗗ��ǂݍ���ł��������B", vbYesNo + vbQuestion, "�m�F")
                If tmp = vbYes Then
                    OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,10�\���҈ꗗ�\*.xls?")
                    If OpenFileName <> "False" Then
                         SetFile = OpenFileName
                    Else
                        MsgBox "�L�����Z������܂���", vbCritical
                        Unload Me
                        Exit Sub
                    End If
                    Application.ScreenUpdating = False
                    Workbooks.Open FileName:=SetFile, ReadOnly:=True, UpdateLinks:=0
                    Set ImportWbook = Workbooks.Open(Path & SetFile)
                    Set iwbS10 = ImportWbook.Worksheets("�ڎ�-�F���")
                    Set iwbS11 = ImportWbook.Worksheets("�ڎ�-���ގ�")
                    Set iwbS12 = ImportWbook.Worksheets("�ڎ�-�Ĕ��s")
                    
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
                    MsgBox "�����𒆒f���܂�"
                    Unload Me
                End If
                ImportWbook.Close
                Application.DisplayAlerts = True
                Application.ScreenUpdating = True
                
            End If
        End If
        
'        ���{�L�^
        twbS00.Range("A2").Interior.ColorIndex = 44
        twbS00.Range("A2").Font.ColorIndex = 1
        twbS00.Range("E2") = "" & Now & ""
        Unload Me
        MsgBox "�������������܂����B�����ăf�[�^�x�[�X�̎捞���s���ĉ������B"
    End With
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
    Unload UserForm5
End Sub

Private Sub CommandButton2_Click()
'�A�@�ŐV�f�[�^�x�[�X�捞
    Dim tmp As Integer, sr As Integer, LastRow As Integer, i As Integer, j As Integer
    Dim OpenFileName As String, FileName As String, Path As String, SetFile As String
    Dim ThisWbook As Workbook, ImportWbook As Workbook
    Dim twbS05 As Worksheet, iwbS01 As Worksheet, twbS00 As Worksheet, iwbS02 As Worksheet
    Set ThisWbook = ActiveWorkbook
    Set twbS05 = ThisWbook.Worksheets("data")
    Set twbS00 = ThisWbook.Worksheets("�t�H�[���ďo")
On Error GoTo myError
    With twbS05
        Application.DisplayAlerts = False
        
'        �f�[�^�x�[�X�̑I���i�捞�j
        tmp = MsgBox("�F��_�Ǝ҃f�[�^�x�[�X�i�ŐV�Łj��I�����ĉ�����", vbYesNo + vbQuestion, "�m�F")
        If tmp = vbYes Then
            OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,�F��_�Ǝ҃f�[�^*.xls?")
            If OpenFileName <> "False" Then
                 SetFile = OpenFileName
            Else
                MsgBox "�L�����Z������܂���", vbCritical
                Unload Me
                Exit Sub
            End If
            Application.ScreenUpdating = False
            Workbooks.Open FileName:=SetFile, ReadOnly:=True, UpdateLinks:=0
            Set ImportWbook = Workbooks.Open(Path & SetFile)
            Set iwbS01 = ImportWbook.Worksheets("data (�l)")
            Set iwbS02 = ImportWbook.Worksheets("data (�@�l)")
            .Range("A4:CA1000").ClearContents
            
'            data (�l)�V�[�g�̃R�s�[
            LastRow = iwbS01.Cells(iwbS01.Rows.Count, 7).End(xlUp).Row
            iwbS01.Range("A5" & ":" & "CA" & LastRow).Copy
            .Range("A4").PasteSpecial xlPasteValues
            
'            data (�@�l)�V�[�g�̃R�s�[
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
            
'            �V�K�ǉ����̐F����
            .Range("A" & LastRow & ":" & "CB" & LastRow + 10).Interior.ColorIndex = 0
            .Range("A" & LastRow & ":" & "CB" & LastRow + 10).Interior.ColorIndex = 34
            
'            ���{�L�^
            .Range("F2") = "���s�������t�́A" & Now & "�@�ł��B"
            .Range("A3:CA1000").AutoFilter
            
'            �e���̔F��_�ƎҐ��̎擾
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

'            ���{�L�^
            twbS00.Range("A3").Interior.ColorIndex = 44
            twbS00.Range("A3").Font.ColorIndex = 1
            twbS00.Range("E3") = "" & Now & ""
            
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            Application.CutCopyMode = False
            ImportWbook.Close False
            Unload Me
        Else
            MsgBox "�����𒆒f���܂�"
            Unload Me
        End If
        
'        �F��_�Ǝґ����擾
        twbS00.Select
        If twbS00.Range("L1") = "" Then
            twbS00.Range("G2") = WorksheetFunction.CountIf(twbS05.Range("G4:G1000"), "<>")
            twbS00.Range("L1") = "��"
        End If
    End With
    MsgBox "�f�[�^�x�[�X�̎�荞�݂��������܂����Bdata�V�[�g�ɂĊm�F�ł��܂��B"
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
    Unload Me
End Sub

Private Sub CommandButton3_Click()
'�B�@�����̒��o
    Dim DateFrom As String, DateTo As String
    Dim tmp As Variant, TempNumber As Variant
    Dim InputValue As Integer, Value As Integer, i As Integer
    Dim LastRow, SetRow As Integer
    Dim twbS01 As Worksheet, twbS05 As Worksheet, twbS06 As Worksheet, twbS00 As Worksheet
On Error GoTo myError
    Set twbS01 = Worksheets("0_�����\")
    Set twbS05 = Worksheets("data")
    Set twbS06 = Worksheets("���ʒ��o")
    Set twbS00 = Worksheets("�t�H�[���ďo")
    
'   �������@�s�v�ȃZ���l�̍폜
    twbS01.Range("D2:D101").ClearContents
    twbS01.Range("E3:F101").ClearContents
    twbS01.Range("H3:J101").ClearContents
    twbS01.Range("U2:U101").ClearContents
    
'    �e���̔F��_�ƎҏI�������̐��擾
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
    
'    ���͂��ꂽ���Ԃ̒��o
    With twbS05
        twbS06.Range("A4:CA1000").clear
        .Range("BC4").AutoFilter 53, ">=" & DateFrom, xlAnd, "<=" & DateTo
        If WorksheetFunction.Subtotal(3, .Range("G:G")) < 2 Then
            MsgBox DateFrom & "����" & DateTo & "�̃f�[�^�͑��݂��܂���!" & vbCrLf & "�m�F�̏�ēx���s���ĉ������B", vbInformation
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
        
        MsgBox "���o���������܂����B" & vbCrLf & "���o�����́A" & WorksheetFunction.Subtotal(2, twbS06.Columns(1)) & "���ł��B" _
        & Space(10) & "���o��́ABC��ł��B" & vbCrLf & "���o���̃f�[�^���ԈႢ�Ȃ����m�F���Ă��������B" _
        & vbCrLf & vbCrLf & "���o���������O�́A0_�����\�ɃR�s�[���܂����B"
        twbS00.Select
        
'        ���{�L�^
        twbS00.Range("A4").Interior.ColorIndex = 44
        twbS00.Range("A4").Font.ColorIndex = 1
        twbS00.Range("E4") = "" & Now & ""
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
    Unload UserForm5
End Sub

Private Sub CommandButton4_Click()
'�C�@���Y�X�V�҃f�[�^�i�\�����j�쐬
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
    Set twbS01 = ThisWbook.Worksheets("0_�����\")
    Set twbS00 = ThisWbook.Worksheets("�t�H�[���ďo")
    Set twbS08 = ThisWbook.Worksheets("���ʒ��o")
On Error GoTo myError
    Application.DisplayAlerts = False
    With twbS01
    
'        �\�������쐬�ۑ�����t�H���_�̍쐬
        tmp = MsgBox("�\������ۑ�����t�H���_�𓯂��K�w�ɍ쐬���܂��B", vbYesNo + vbQuestion, "�m�F")
        If tmp = vbYes Then
            Flag = True
            SaveDir = ActiveWorkbook.Path & "\00�F��_�Ǝ҃f�[�^"
            If Dir(SaveDir, vbDirectory) = "" Then
                MkDir SaveDir
                MsgBox "���̃t�@�C���Ɠ����K�w�Ɂu00�F��_�Ǝ҃f�[�^�v�Ƃ����ۑ��t�H���_���쐬���܂����B"
            End If
        Else
            Flag = False
            MsgBox "���łɃt�H���_�͍쐬����Ă��邩�A�쐬����Ă��Ȃ���΁A�蓮�Ńt�H���_���쐬���ĉ������B"
        End If
        
'        �\�������X�g�̐l���i�s�j�擾
        LastRow = twbS01.Cells(twbS01.Rows.Count, 4).End(xlUp).Row
        Unload Me
        
'        ���{�t�@�C���̎捞
        tmp = MsgBox("2-1 �F��\���� ���{���Z�b�g���ĉ������B", vbYesNo + vbQuestion, "�m�F")
        If tmp = vbYes Then
            OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,2-1 �F��\����*.xls?")
            If OpenFileName <> "False" Then
                 SetFile = OpenFileName
            Else
                MsgBox "�L�����Z������܂���", vbCritical
                Exit Sub
            End If
            
            Workbooks.Open FileName:=SetFile, ReadOnly:=False, UpdateLinks:=0
            Set ImportWbook = Workbooks.Open(Path & SetFile)
'            Windows(ImportWbook.Name).Visible = False
            Set iwbS01 = ImportWbook.Worksheets("���̓V�[�g")
            Set iwbS04 = ImportWbook.Worksheets("�o�c�w�W")
            Set iwbS05 = ImportWbook.Worksheets("���Y�{��")
            Set iwbS08 = ImportWbook.Worksheets("Record")
            If iwbS01.Range("M5") <> "����" Then
                MsgBox "�捞�񂾃t�@�C�����Ⴂ�܂��B�ŏ�����葱�����肢���܂��B", vbCritical
                ImportWbook.Close
                Exit Sub
            End If
            
            
            MsgBox "�F��\�����i���{�j��荞�݂��������܂����B" & vbCrLf & "�����čX�V�������̐\������ǂݍ���ł��������B"
        Else
            MsgBox "�L�����Z������܂���", vbCritical
            Exit Sub
        End If
        
        Application.ScreenUpdating = False
        ImportWbook.Application.ScreenUpdating = False
        
'        ���\�����̎捞
        
        For i = 2 To LastRow
'        For i = 14 To LastRow  '�e�X�g�p
            If .Cells(i, 25) = "" Then
Label2:
                tmp = MsgBox(.Cells(i, 4).Value & " �l�̃t�@�C����I�����Ă��������B", vbYesNo + vbQuestion, "�m�F")
                
                If tmp = vbYes Then
                    OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*�F��\����*.xls?")
                    If OpenFileName <> "False" Then
                        SetFile = OpenFileName
                    Else
                        MsgBox "�L�����Z������܂���", vbCritical
                        ImportWbook.Close False
                        Exit Sub
                    End If
                Else
                    MsgBox "�L�����Z������܂���", vbCritical
                    ImportWbook.Close False
                    Unload Me
                    Exit Sub
                End If
                Workbooks.Open FileName:=SetFile, ReadOnly:=False, UpdateLinks:=0
                Set ImportWbook2 = Workbooks.Open(Path & SetFile)
'               �@Windows(ImportWbook2.Name).Visible = False
                Set iwbS11 = ImportWbook2.Worksheets("���̓V�[�g")
                Set iwbS14 = ImportWbook2.Worksheets("�o�c�w�W")
                Set iwbS15 = ImportWbook2.Worksheets("���Y�{��")

                SetName = ""  '���O�̏��������s��
                
'                �@�l���܂��͌l���̎擾
                If iwbS11.Range("D5") = "" Then
                    SetName = iwbS11.Range("M5")
                    
                Else
                    SetName = iwbS11.Range("D5")
                End If
                If .Cells(i, 4).Value <> SetName Then
                    MsgBox "�Z�b�g�����t�@�C��" & SetName & "�l�ł͂���܂���B�ēx�I�����ĉ������B"
                    ImportWbook2.Close
                    GoTo Label2
                End If
                
'               �f�[�^�̎捞�i�R�s�[�j���o�c�w�W�̓R�s�[���Ȃ��I
'                iwbS04.Unprotect '�V�[�g�̕ی����
'                iwbS14.Range("A181:N204").Copy
'                iwbS04.Range("A181").PasteSpecial xlPasteValuesAndNumberFormats
'                iwbS04.Protect '�V�[�g�̕ی�
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
                
'                �d�b�ԍ��ƗX�֔ԍ��̎擾
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
                
'                �Z�����̎擾
                iwbS01.Range("D10") = iwbS11.Range("D10")
                If iwbS11.Range("AA2") <> "" Then
                    iwbS01.Range("AA2") = iwbS11.Range("AA2")
                End If
                iwbS01.Range("AB2") = iwbS11.Range("AB2")
                
'                �\�����E����
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
                
'                �_�Ɛ��Y�{��
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
                
'                ���H�E�̔�����
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
                
'                ���̑��̓��L����
                iwbS11.Range("AC29:AC31").Copy
                iwbS01.Range("AC29").PasteSpecial xlPasteValuesAndNumberFormats
                
'                ���Y����
                iwbS11.Range("B42:B53").Copy
                iwbS01.Range("B42").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("H42:H53").Copy
                iwbS01.Range("H42").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("L42:L53").Copy
                iwbS01.Range("L42").PasteSpecial xlPasteValuesAndNumberFormats
                
'                5�N��̍���
                iwbS11.Range("AB42:AB53").Copy
                iwbS01.Range("AB42").PasteSpecial xlPasteValuesAndNumberFormats
                iwbS11.Range("AF42:AF53").Copy
                iwbS01.Range("AF42").PasteSpecial xlPasteValuesAndNumberFormats
    
'                ��Ǝ��
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
                
'                �_�Ƌ@�B
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
                
'                ����ƖڕW�E�[�u
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
                
'                5�N�O�̋@�B�{��
                iwbS11.Range("AJ74:AU86").Copy
                iwbS01.Range("AJ74").PasteSpecial xlPasteValuesAndNumberFormats
                
                Application.CutCopyMode = False
                
'                ���̎擾
                For j = 2 To LastRow
                    If SetName = .Cells(j, 4) Then
                       k = .Cells(j, 1).Value
                    End If
                Next
                ImportWbook2.Close
'                �t�@�C���̕ۑ��@�����Ƈ���g�ݍ��킹���t�@�C�������쐬
                SetFileName = Right("0" & k, 2) & "�F��\����" & Format(Date, "[$-ja-JP]ge") & "�i" & SetName & " �j.xlsm"
                ImportWbook.SaveAs FileName:=SaveDir & "\" & Right("0" & k, 2) & "�F��\����" & Format(Date, "[$-ja-JP]ge") & "�i" & SetName & " �j"
                
            End If
        Next
        Application.ScreenUpdating = True
        ImportWbook.Application.ScreenUpdating = True
'        Windows(SetFileName).Visible = True
        Application.ActiveWorkbook.Close
        Application.DisplayAlerts = True
        MsgBox "�\�����̐V�K�쐬���������܂����B"
        
'        ���{�o�^
        twbS00.Range("A5").Interior.ColorIndex = 44
        twbS00.Range("A5").Font.ColorIndex = 1
        twbS00.Range("E5") = "" & Now & ""
    End With
    Exit Sub
myError:
    MsgBox "�I�������t�@�C�����Ⴂ�܂��I�������I�����܂��B", vbExclamation
End Sub

Private Sub CommandButton5_Click()
'�D�@�V�K�o�^�҃f�[�^�i�\�����j�쐬
    UserForm2.Show
End Sub
Private Sub CommandButton19_Click()
'�F �ē����ɑ��t���鏑�ށi�o�c�󋵒����\�j�쐬
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
    Set twbS01 = ThisWbook.Worksheets("0_�����\")
    Set twbS00 = ThisWbook.Worksheets("�t�H�[���ďo")
On Error GoTo myError
    
    Unload Me
'    �������ԑ���
    startTime = Timer
    
    SaveDir = ActiveWorkbook.Path & "\00�F��_�Ǝ҃f�[�^"
    SaveDir2 = ActiveWorkbook.Path

'    �t�H���_���ɂ���t�@�C���̎����ǂݍ���
    '--- �t�@�C���V�X�e���I�u�W�F�N�g ---'
    Set fso = CreateObject("Scripting.FileSystemObject")
    '--- �t�@�C�������i�[����ϐ� ---'
    fileCount = fso.GetFolder(SaveDir).Files.Count
    
    '    ���t�E�������ރt�H���_�̍쐬
    SaveDir = ActiveWorkbook.Path & "\00�F��_�Ǝ҃f�[�^"
    SaveDir2 = SaveDir & "\���t�E��������"
    If (oFso Is Nothing) Then
        Set oFso = CreateObject("Scripting.FileSystemObject")
    End If
    If (oFso.FolderExists(SaveDir2) = False) Then
        '// �t�H���_�����݂��Ȃ�
        MsgBox "�u���t�E�������ށv�Ƃ������O�̃t�H���_��V�K�쐬���܂�"
        MkDir SaveDir & "\���t�E��������"
        SaveDir2 = SaveDir & "\���t�E��������"
    Else
    End If
    
'    �o�c�󋵒����\�A�t�H���_�̍쐬
    FileName3 = SaveDir2 & "\�o�c�󋵒����\�t�H���_"
    If (oFso Is Nothing) Then
        Set oFso = CreateObject("Scripting.FileSystemObject")
    End If
    If (oFso.FolderExists(FileName3) = False) Then
        '// �t�H���_�����݂��Ȃ�
        MsgBox "�u�o�c�󋵒����\�t�H���_�v�Ƃ������O�̃t�H���_��V�K�쐬���܂�"
        MkDir SaveDir2 & "\�o�c�󋵒����\�t�H���_"
    Else
    End If

    For i = 1 To fileCount
'    For i = 1 To 2 '�e�X�g�R�[�h
    
'    ���[�h��������i�܂Ƃ߂Ĉ������ꍇ�j

'        Dim SaveDir As String
'        Dim wd As Object
'        SaveDir = ActiveWorkbook.Path & "\���t�E��������\"
'        Set wd = CreateObject("Word.application")
'        wd.Visible = True
'        wd.documents.Open FileName:=SaveDir & "�@�@���x�T�v�i��.doc" '�������������
'        wd.ActiveDocument.PrintOut Background:=False
'
'        wd.documents.Open FileName:=SaveDir & "�A�@�V �ӌ��m�F���i�ߘa�j.doc" '�������������
'        wd.ActiveDocument.PrintOut Background:=False
'
'        wd.Quit
'        Set wd = Nothing

'     ���[�h������������܂�

        If i <= 9 Then
            TmpChar = "0" & CStr(i)
        Else
            TmpChar = i
        End If
        FileNames = Dir(SaveDir & "\" & TmpChar & "�F��\����R*")
        FileNames2 = Dir(SaveDir2 & "\�C�@�o�c�󋵒����\*")
        
'        �D�@�V�F��v��\�������w��̏ꏊ�ɂȂ������ꍇ�͎蓮�Ŏ捞�B���t�E�������ރt�H���_�ɃR�s�[
        If FileNames2 = "" Then
            tmp = MsgBox("�C�@�o�c�󋵒����\" & vbCrLf & "�t�@�C����I�����Ă��������B", vbYesNo + vbQuestion, "�m�F")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,�C�@�o�c�󋵒����\*.xls?")
                If OpenFileName <> "False" Then
                    FileCopy OpenFileName, SaveDir2 & "\" & "�C�@�o�c�󋵒����\.xlsx"
                    FileNames2 = Dir(SaveDir2 & "\" & "�C�@�o�c�󋵒����\*")
                Else
                    MsgBox "�L�����Z������܂���", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "�L�����Z������܂���", vbCritical
                ImportWbook.Close False
                Unload Me
                Exit Sub
            End If
        Else
        End If

        Application.DisplayAlerts = False
        Workbooks.Open FileName:=SaveDir & "\" & FileNames, ReadOnly:=True, UpdateLinks:=0
        Set ImportWbook = Workbooks.Open(SaveDir & "\" & FileNames)
        Set iwbS01 = ImportWbook.Worksheets("���̓V�[�g")
        Set iwbS04 = ImportWbook.Worksheets("�ȈՔ�")
        
'        ����ƖڕW�̐��l�f�[�^����
        Workbooks.Open FileName:=SaveDir2 & "\" & FileNames2, ReadOnly:=False, UpdateLinks:=0
        Set ImportWbook2 = Workbooks.Open(SaveDir2 & "\" & FileNames2)
        Set iwbS05 = ImportWbook2.Worksheets("�o�c�󋵒����\")
        
        If iwbS01.Range("D5") = "" Then
            SetName = iwbS01.Range("M5")
        Else
            SetName = iwbS01.Range("D5")
        End If
        
'        �l���
        iwbS05.Range("B6") = SetName
        iwbS05.Range("G6").Value = iwbS01.Range("D10").Value
        iwbS05.Range("L6").Value = Right(iwbS01.Range("M9").Value, 7)
        iwbS05.Range("N6").Value = iwbS01.Range("M11").Value
        
'        �\����
        iwbS05.Range("B11:B20").Value = iwbS01.Range("U5:U14").Value
        iwbS05.Range("G11:G20").Value = iwbS01.Range("AF5:AF14").Value
        iwbS05.Range("I11:I20").Value = iwbS01.Range("AB5:AB14").Value
        iwbS05.Range("L11:L20").Value = iwbS01.Range("AP5:AP14").Value
        
'        �ٗp
        iwbS05.Range("B26").Value = iwbS01.Range("AP15").Value
        iwbS05.Range("G26").Value = iwbS01.Range("AP16").Value
        iwbS05.Range("K26").Value = iwbS01.Range("AP17").Value
        
'        ���Y�{��
        iwbS05.Range("B31:B36").Value = iwbS01.Range("W21:W26").Value
        iwbS05.Range("H31:H36").Value = iwbS01.Range("AF21:AF26").Value
        iwbS05.Range("J31:J36").Value = iwbS01.Range("AJ21:AJ26").Value
        
'        �_�Ƌ@��
        iwbS05.Range("B41:B52").Value = iwbS01.Range("AJ74:AJ85").Value
        iwbS05.Range("I41:I52").Value = iwbS01.Range("AP74:AP85").Value
        iwbS05.Range("K41:K52").Value = iwbS01.Range("AS74:AS85").Value
        
'        �_�Ɛ��Y�i�k��j
        iwbS05.Range("B57:B62").Value = iwbS01.Range("B42:B47").Value
        iwbS05.Range("H57:H62").Value = iwbS01.Range("H42:H47").Value
        iwbS05.Range("J57:J62").Value = iwbS01.Range("L42:L47").Value
        
'        �_�Ɛ��Y�i�{�Y�j
        iwbS05.Range("B64:B69").Value = iwbS01.Range("B48:B53").Value
        iwbS05.Range("H64:H69").Value = iwbS01.Range("H48:H53").Value
        iwbS05.Range("J64:J69").Value = iwbS01.Range("L48:L53").Value
        
'        ��Ǝ��
        iwbS05.Range("B74:B79").Value = iwbS01.Range("B59:B64").Value
        iwbS05.Range("D74:D79").Value = iwbS01.Range("E59:E64").Value
        iwbS05.Range("L74:L79").Value = iwbS01.Range("K59:K64").Value
        
'        ����i�̔��܂ňϑ��j
        iwbS05.Range("B84:B85").Value = iwbS01.Range("D70:D71").Value
        iwbS05.Range("D84:D85").Value = iwbS01.Range("G70:G71").Value
        iwbS05.Range("I84:I85").Value = iwbS01.Range("L70:L71").Value
        iwbS05.Range("K84:K85").Value = iwbS01.Range("P70:P71").Value
        
'        ���H�̔�
        iwbS05.Range("B90:B95").Value = iwbS01.Range("D30:D35").Value
        iwbS05.Range("I90:I95").Value = iwbS01.Range("L30:L35").Value
        iwbS05.Range("K90:K95").Value = iwbS01.Range("P30:P35").Value
        
'        ����ƖڕW�̐��l�f�[�^�̈��
        iwbS05.PrintOut

'        �������f�[�^�̕ۑ�
        SetFileName = FileName3 & "\" & TmpChar & SetName & "�C�@�o�c�󋵒����\.xlsx"
        ImportWbook2.SaveAs FileName:=SetFileName
        ImportWbook2.Close
        ImportWbook.Close
        Application.DisplayAlerts = True
    Next
    
'    �������Ԍ���
    endTime = Timer
    processTime = endTime - startTime
    MsgBox "������I�����܂����B���Ԃ́F" & processTime & "�b�ł��B"
    
'   ���{�o�^
    twbS00.Range("A8").Interior.ColorIndex = 44
    twbS00.Range("A8").Font.ColorIndex = 1
    twbS00.Range("E8") = "" & Now & ""

Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
    
End Sub

Private Sub CommandButton6_Click()
'�E �ē����ɑ��t���鏑�ށi�A���P�[�g�j�쐬
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
    Set twbS01 = ThisWbook.Worksheets("0_�����\")
    Set twbS00 = ThisWbook.Worksheets("�t�H�[���ďo")
On Error GoTo myError
    Unload Me
    
'    �������ԑ���
    startTime = Timer
    
'    ���t�E�������ރt�H���_�̍쐬
    SaveDir = ActiveWorkbook.Path & "\00�F��_�Ǝ҃f�[�^"
    SaveDir2 = SaveDir & "\���t�E��������"
    If (oFso Is Nothing) Then
        Set oFso = CreateObject("Scripting.FileSystemObject")
    End If
    If (oFso.FolderExists(SaveDir2) = False) Then
        '// �t�H���_�����݂��Ȃ�
        MsgBox "�u���t�E�������ށv�Ƃ������O�̃t�H���_��V�K�쐬���܂�"
        MkDir SaveDir & "\���t�E��������"
        SaveDir2 = SaveDir & "\���t�E��������"
    Else
    End If

'    �A���P�[�g�A�t�H���_�̍쐬
    FileName3 = SaveDir2 & "\�A���P�[�g�t�H���_"
    If (oFso Is Nothing) Then
        Set oFso = CreateObject("Scripting.FileSystemObject")
    End If
    If (oFso.FolderExists(FileName3) = False) Then
        '// �t�H���_�����݂��Ȃ�
        MsgBox "�u�A���P�[�g�t�H���_�v�Ƃ������O�̃t�H���_��V�K�쐬���܂�"
        MkDir SaveDir2 & "\�A���P�[�g�t�H���_"
    Else
    End If
    
'    �t�H���_���ɂ���t�@�C���̎����ǂݍ���
    '--- �t�@�C���V�X�e���I�u�W�F�N�g ---'
    Set fso = CreateObject("Scripting.FileSystemObject")
    '--- �t�@�C�������i�[����ϐ� ---'
    fileCount = fso.GetFolder(SaveDir).Files.Count
    For i = 1 To fileCount
'    For i = 34 To fileCount '�e�X�g�R�[�h
        If i <= 9 Then
            TempNumber = "0" & CStr(i)
        Else
            TempNumber = i
        End If
        
'        �쐬�����F��\�����̎捞
        FileNames = Dir(SaveDir & "\*" & TempNumber & "*�F��\����R*")
        Application.DisplayAlerts = False
        Workbooks.Open FileName:=SaveDir & "\" & FileNames, ReadOnly:=True, UpdateLinks:=0
        Set ImportWbook = Workbooks.Open(SaveDir & "\" & FileNames)
        Set iwbS01 = ImportWbook.Worksheets("���̓V�[�g")
        Set iwbS04 = ImportWbook.Worksheets("�R���\")
        Set iwbS05 = ImportWbook.Worksheets("�ȈՔ�")

'        �A���P�[�g�p���̎捞
        FileNames2 = Dir(SaveDir2 & "\" & "�B�@�V�l��B13�E14*")
        
'        �A���P�[�g�p�����w��̏ꏊ�ɂȂ������ꍇ�͎蓮�Ŏ捞�B���t�E�������ރt�H���_�ɃR�s�[
        If FileNames2 = "" Then
            tmp = MsgBox("�B�@�V�l��B13�E14�u�_�ƌo�c���P�v��̒B���󋵓��ɂ��āi�A���P�[�g�j�v�yR3.3 ���z" & vbCrLf & "�t�@�C����I�����Ă��������B", vbYesNo + vbQuestion, "�m�F")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,�B�@�V�l��B13�E14*.xls?")
                If OpenFileName <> "False" Then
                    FileCopy OpenFileName, SaveDir2 & "\" & "�B�@�V�l��B13�E14�u�_�ƌo�c���P�v��̒B���󋵓��ɂ��āi�A���P�[�g�j�v�yR3.3 ���z.xlsx"
                    FileNames2 = Dir(SaveDir2 & "\" & "�B�@�V�l��B13�E14*")
                Else
                    MsgBox "�L�����Z������܂���", vbCritical
                    ImportWbook.Close False
                    Exit Sub
                End If
            Else
                MsgBox "�L�����Z������܂���", vbCritical
                ImportWbook.Close False
                Unload Me
                Exit Sub
            End If
        Else
        End If
        
'        �ǂݍ��񂾃A���P�[�g�t�@�C���Ƀf�[�^���R�s�[����
        Workbooks.Open FileName:=SaveDir2 & "\" & FileNames2, ReadOnly:=False, UpdateLinks:=0
        Set ImportWbook2 = Workbooks.Open(SaveDir2 & "\" & FileNames2)
        Set iwbS09 = ImportWbook2.Worksheets("�B���󋵁i�V�l���j")
        Set iwbS10 = ImportWbook2.Worksheets("�V�K�F��i�V�l���j")
        
'        �V�K���X�V���̔���
        If iwbS01.Range("D4") = "�V�K" Then
'        �V�K��t�҂̃f�[�^�R�s�[
            If iwbS01.Range("D5") = "" Then
                SetName = iwbS01.Range("M5")
            Else
                SetName = iwbS01.Range("D5")
            End If
            With iwbS01
                iwbS10.Range("J9").Value = SetName
                TmpChar = .Range("D10").Value
                TmpChar = Replace(TmpChar, "�V���s", "")
                iwbS10.Range("S13").Value = TmpChar
            End With
            iwbS10.Activate
            iwbS10.PrintOut
        Else
'        �X�V��t�҂̃f�[�^�R�s�[
            If iwbS01.Range("D5") = "" Then
                SetName = iwbS01.Range("M5")
            Else
                SetName = iwbS01.Range("D5")
            End If
            With iwbS01
                iwbS09.Range("J8").Value = SetName
                TmpChar = .Range("D10").Value
                TmpChar = Replace(TmpChar, "�V���s", "")
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
                
                If .Range("B16").Value = "�P��o�c" Then
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
                    iwbS09.Range("J84").Value = "�Z"
                Else
                    iwbS09.Range("S84").Value = "�Z"
                End If
                
                If .Range("X54").Value < 1000000 Then
                    iwbS09.Range("K87").Value = "�Z"
                ElseIf .Range("X54").Value > 1000000 And .Range("X54").Value < 2000000 Then
                    iwbS09.Range("K88").Value = "�Z"
                ElseIf .Range("X54").Value > 2000000 And .Range("X54").Value < 3000000 Then
                    iwbS09.Range("K89").Value = "�Z"
                ElseIf .Range("X54").Value > 3000000 And .Range("X54").Value < 4000000 Then
                    iwbS09.Range("K90").Value = "�Z"
                ElseIf .Range("X54").Value > 4000000 And .Range("X54").Value < 5000000 Then
                    iwbS09.Range("K91").Value = "�Z"
                ElseIf .Range("X54").Value > 5000000 And .Range("X54").Value < 6000000 Then
                    iwbS09.Range("K92").Value = "�Z"
                ElseIf .Range("X54").Value > 6000000 And .Range("X54").Value < 7000000 Then
                    iwbS09.Range("K93").Value = "�Z"
                ElseIf .Range("X54").Value > 7000000 And .Range("X54").Value < 8000000 Then
                    iwbS09.Range("K94").Value = "�Z"
                ElseIf .Range("X54").Value > 8000000 And .Range("X54").Value < 9000000 Then
                    iwbS09.Range("K95").Value = "�Z"
                ElseIf .Range("X54").Value > 9000000 And .Range("X54").Value < 10000000 Then
                    iwbS09.Range("K96").Value = "�Z"
                ElseIf .Range("X54").Value > 10000000 And .Range("X54").Value < 15000000 Then
                    iwbS09.Range("K97").Value = "�Z"
                ElseIf .Range("X54").Value > 150000000 And .Range("X54").Value < 30000000 Then
                    iwbS09.Range("K98").Value = "�Z"
                ElseIf .Range("X54").Value > 300000000 Then
                    iwbS09.Range("K99").Value = "�Z"
                End If
                
                If .Range("AR54").Value < 1000000 Then
                    iwbS09.Range("O87").Value = "�Z"
                ElseIf .Range("AR54").Value > 1000000 And .Range("AR54").Value < 2000000 Then
                    iwbS09.Range("O88").Value = "�Z"
                ElseIf .Range("AR54").Value > 2000000 And .Range("AR54").Value < 3000000 Then
                    iwbS09.Range("O89").Value = "�Z"
                ElseIf .Range("AR54").Value > 3000000 And .Range("AR54").Value < 4000000 Then
                    iwbS09.Range("O90").Value = "�Z"
                ElseIf .Range("AR54").Value > 4000000 And .Range("AR54").Value < 5000000 Then
                    iwbS09.Range("O91").Value = "�Z"
                ElseIf .Range("AR54").Value > 5000000 And .Range("AR54").Value < 6000000 Then
                    iwbS09.Range("O92").Value = "�Z"
                ElseIf .Range("AR54").Value > 6000000 And .Range("AR54").Value < 7000000 Then
                    iwbS09.Range("O93").Value = "�Z"
                ElseIf .Range("AR54").Value > 7000000 And .Range("AR54").Value < 8000000 Then
                    iwbS09.Range("O94").Value = "�Z"
                ElseIf .Range("AR54").Value > 8000000 And .Range("AR54").Value < 9000000 Then
                    iwbS09.Range("O95").Value = "�Z"
                ElseIf .Range("AR54").Value > 9000000 And .Range("AR54").Value < 10000000 Then
                    iwbS09.Range("O96").Value = "�Z"
                ElseIf .Range("AR54").Value > 10000000 And .Range("AR54").Value < 15000000 Then
                    iwbS09.Range("O97").Value = "�Z"
                ElseIf .Range("AR54").Value > 150000000 And .Range("AR54").Value < 30000000 Then
                    iwbS09.Range("O98").Value = "�Z"
                ElseIf .Range("AR54").Value > 300000000 Then
                    iwbS09.Range("O99").Value = "�Z"
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
        
        SetFileName = FileName3 & "\" & TempNumber & " " & SetName & "05_�yR3.3 ���z�A���P�[�g�l��B13�EB14�i�C���Łj.xlsx"
        ImportWbook2.SaveAs FileName:=SetFileName
        ImportWbook2.Close
        
        ImportWbook.Close
        Application.DisplayAlerts = True
    Next
    
'    �������Ԍ���
    endTime = Timer
    processTime = endTime - startTime
    MsgBox "������I�����܂����B���Ԃ́F" & processTime & "�b�ł��B"
    
'   ���{�o�^
    twbS00.Range("A7").Interior.ColorIndex = 44
    twbS00.Range("A7").Font.ColorIndex = 1
    twbS00.Range("E7") = "" & Now & ""
Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
End Sub
Private Sub CommandButton18_Click()
'�G�@��������@�p2
    Dim sr As Integer, LastRow As Integer, Value As Integer
    Dim ThisWbook As Workbook
    Dim twbS08 As Worksheet, twbS01 As Worksheet, twbS00 As Worksheet
    Dim tmp As VbMsgBoxResult
    Set twbS08 = Worksheets("�ĔF��ē����")
    Set twbS01 = Worksheets("0_�����\")
    Set twbS00 = Worksheets("�t�H�[���ďo")
On Error GoTo myError
    MsgBox "������Ƀv�����^�[�ݒ�" & vbLf & "�i������s�������v�����^�[��ʏ�g�p����v�����^�[�Ɂj" & vbLf & "���Ă����ĉ������B", vbExclamation
    tmp = MsgBox("������������́u0_�����\��U��v�Ƀ`�F�b�N�����Ēu���ĉ������B" & vbCrLf & "����v���r���[��\�����܂��B�m�F��A����́~��1�l���m�F���Ă��������B", vbYesNo)
    Unload Me
    With twbS01
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        Value = WorksheetFunction.CountA(twbS01.Range("U:U")) - 1
        If Value = 0 Then
            MsgBox "�������������U��Ƀ`�F�b�N�����čĎ��s���ĉ������B", vbExclamation
            Exit Sub
        End If
        LastRow = .Cells(.Rows.Count, 3).End(xlUp).Row
        twbS08.Range("L4").NumberFormatLocal = "@"
        twbS08.Range("R4").NumberFormatLocal = "@"
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "��" Then
'                If .Cells(sr, 4) = "" Or .Cells(sr, 6) = "" Or .Cells(sr, 8) = "" Or .Cells(sr, 9) = "" Or .Cells(sr, 10) = "" Then
'                    MsgBox "�f�[�^�����͉ӏ�������܂��B", vbExclamation
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
        tmp = MsgBox("�v���r���[���������܂����B�����Ĉ�����܂��B��낵���ł����H", vbYesNo)
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "��" Then
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
    MsgBox "������������܂����B"
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
End Sub

Private Sub CommandButton17_Click()
'�G�@�ʒm�ē���������i�F��ē��j�@�i��3�j

    Dim sr As Integer, LastRow As Integer, Value As Integer
    Dim ThisWbook As Workbook
    Dim twbS12 As Worksheet, twbS01 As Worksheet, twbS00 As Worksheet
    Dim tmp As VbMsgBoxResult
    Set twbS12 = Worksheets("���R��� (�ē�)")
    Set twbS01 = Worksheets("0_�����\")
    Set twbS00 = Worksheets("�t�H�[���ďo")
On Error GoTo myError
    MsgBox "������Ƀv�����^�[�ݒ�" & vbLf & "�i������s�������v�����^�[��ʏ�g�p����v�����^�[�Ɂj" & vbLf & "���Ă����ĉ������B", vbExclamation
    tmp = MsgBox("������������́u0_�����\��U��v�Ƀ`�F�b�N�����Ēu���ĉ������B" & vbCrLf & "����v���r���[��\�����܂��B�m�F��A����́~��1�l���m�F���Ă��������B", vbYesNo)
    Unload Me
    With twbS01
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 3).End(xlUp).Row
        Value = WorksheetFunction.CountA(twbS01.Range("U:U")) - 1
        If Value = 0 Then
            MsgBox "�������������U��Ƀ`�F�b�N�����čĎ��s���ĉ������B", vbExclamation
            Exit Sub
        End If
        twbS12.Range("L4").NumberFormatLocal = "@"
        twbS12.Range("R4").NumberFormatLocal = "@"
        
'        ����v���r���[
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "��" Then
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
        
'        �v���r���[�̈��
        tmp = MsgBox("�v���r���[���������܂����B�����Ĉ�����܂��B��낵���ł����H", vbYesNo)
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "��" Then
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
    
 '   ���{�o�^
    twbS00.Range("A9").Interior.ColorIndex = 44
    twbS00.Range("A9").Font.ColorIndex = 1
    twbS00.Range("E9") = "" & Now & ""
    MsgBox "������������܂����B"
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation

End Sub

Private Sub CommandButton7_Click()
'�I�@�ʒm�ē���������i�X�V��t�����ē��j�@�i��3�j

    Dim sr As Integer, LastRow As Integer, Value As Integer
    Dim ThisWbook As Workbook
    Dim twbS12 As Worksheet, twbS01 As Worksheet, twbS00 As Worksheet
    Dim tmp As VbMsgBoxResult
    Set twbS12 = Worksheets("���R���")
    Set twbS01 = Worksheets("0_�����\")
    Set twbS00 = Worksheets("�t�H�[���ďo")
On Error GoTo myError
    MsgBox "������Ƀv�����^�[�ݒ�" & vbLf & "�i������s�������v�����^�[��ʏ�g�p����v�����^�[�Ɂj" & vbLf & "���Ă����ĉ������B", vbExclamation
    tmp = MsgBox("������������́u0_�����\��U��v�Ƀ`�F�b�N�����Ēu���ĉ������B" & vbCrLf & "����v���r���[��\�����܂��B�m�F��A����́~��1�l���m�F���Ă��������B", vbYesNo)
    Unload Me
    With twbS01
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 3).End(xlUp).Row
        Value = WorksheetFunction.CountA(twbS01.Range("U:U")) - 1
        If Value = 0 Then
            MsgBox "�������������U��Ƀ`�F�b�N�����čĎ��s���ĉ������B", vbExclamation
            Exit Sub
        End If
        twbS12.Range("L4").NumberFormatLocal = "@"
        twbS12.Range("R4").NumberFormatLocal = "@"
        
'        ����v���r���[
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "��" Then
                If .Cells(sr, 4) = "" Or .Cells(sr, 6) = "" Or .Cells(sr, 8) = "" Or .Cells(sr, 9) = "" Or .Cells(sr, 10) = "" Then
                    MsgBox "�f�[�^�����͉ӏ�������܂��B", vbExclamation
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
        
'        �v���r���[���
        tmp = MsgBox("�v���r���[���������܂����B�����Ĉ�����܂��B��낵���ł����H", vbYesNo)
        If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For sr = 2 To LastRow
            If .Cells(sr, 21) Like "��" Then
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
    
'   ���{�o�^
    twbS00.Range("A11").Interior.ColorIndex = 44
    twbS00.Range("A11").Font.ColorIndex = 1
    twbS00.Range("E11") = "" & Now & ""
    MsgBox "������������܂����B"
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
End Sub
Private Sub CommandButton16_Click()
'�i�g�p���Ȃ��I�j
'�G���O�������f�[�^���
'�ȈՔ�3�y�[�W������@���@�C�V�F��v��\�����ƇD�@�ڕW�Ƒ[�u�̕���W����@���@�ǉ�����i���x�T�v�E�ӌ��m�F���j
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
    Set twbS01 = ThisWbook.Worksheets("0_�����\")
    Set twbS00 = ThisWbook.Worksheets("�t�H�[���ďo")
On Error GoTo myError
    
    Unload Me
'    �������ԑ���
    startTime = Timer
    
    SaveDir = ActiveWorkbook.Path & "\00�F��_�Ǝ҃f�[�^"
    SaveDir2 = ActiveWorkbook.Path

'    �t�H���_���ɂ���t�@�C���̎����ǂݍ���
    '--- �t�@�C���V�X�e���I�u�W�F�N�g ---'
    Set fso = CreateObject("Scripting.FileSystemObject")
    '--- �t�@�C�������i�[����ϐ� ---'
    fileCount = fso.GetFolder(SaveDir).Files.Count
    For i = 1 To fileCount
    
'    ���[�h��������i�܂Ƃ߂Ĉ������ꍇ�j

'        Dim SaveDir As String
'        Dim wd As Object
'        SaveDir = ActiveWorkbook.Path & "\���t�E��������\"
'        Set wd = CreateObject("Word.application")
'        wd.Visible = True
'        wd.documents.Open FileName:=SaveDir & "�@�@���x�T�v�i��.doc" '�������������
'        wd.ActiveDocument.PrintOut Background:=False
'
'        wd.documents.Open FileName:=SaveDir & "�A�@�V �ӌ��m�F���i�ߘa�j.doc" '�������������
'        wd.ActiveDocument.PrintOut Background:=False
'
'        wd.Quit
'        Set wd = Nothing

'     ���[�h������������܂�

        If i <= 9 Then
            TmpChar = "0" & CStr(i)
        Else
            TmpChar = i
        End If
        FileNames = Dir(SaveDir & "\*" & TmpChar & "�F��\����R*")
        FileNames2 = Dir(ActiveWorkbook.Path & "\�D�@�V�F��v��\�����i�������p���l�̂݁j*")
        FileNames3 = Dir(ActiveWorkbook.Path & "\�E�@�ڕW�Ƒ[�u�̕���W*")
        
'        �D�@�V�F��v��\�������w��̏ꏊ�ɂȂ������ꍇ�͎蓮�Ŏ捞�B���t�E�������ރt�H���_�ɃR�s�[
        If FileNames2 = "" Then
            tmp = MsgBox("�D�@�V�F��v��\�����i�������p���l�̂݁j" & vbCrLf & "�t�@�C����I�����Ă��������B", vbYesNo + vbQuestion, "�m�F")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,�D�@�V�F��v��\�����i�������p���l�̂݁j*.xls?")
                If OpenFileName <> "False" Then
                    FileCopy OpenFileName, SaveDir2 & "\" & "�D�@�V�F��v��\�����i�������p���l�̂݁j.xlsx"
                    FileNames2 = Dir(SaveDir2 & "\" & "�D�@�V�F��v��\�����i�������p���l�̂݁j*")
                Else
                    MsgBox "�L�����Z������܂���", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "�L�����Z������܂���", vbCritical
                ImportWbook.Close False
                Unload Me
                Exit Sub
            End If
        Else
        End If
        
        Application.DisplayAlerts = False
        Workbooks.Open FileName:=SaveDir & "\" & FileNames, ReadOnly:=True, UpdateLinks:=0
        Set ImportWbook = Workbooks.Open(SaveDir & "\" & FileNames)
        Set iwbS01 = ImportWbook.Worksheets("���̓V�[�g")
        Set iwbS04 = ImportWbook.Worksheets("�ȈՔ�")
        
'        ����ƖڕW�̐��l�f�[�^����
        Workbooks.Open FileName:=SaveDir2 & "\" & FileNames2, ReadOnly:=False, UpdateLinks:=0
        Set ImportWbook2 = Workbooks.Open(SaveDir2 & "\" & FileNames2)
        Set iwbS05 = ImportWbook2.Worksheets("����ƖڕW")
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
            If iwbS01.Cells(j, 34) = "��" Then
            iwbS05.Cells(j, 8) = iwbS01.Cells(j, 21)
            End If
        Next
        
'        �ڕW�Ƒ[�u�̕���W�捞
'        �E�@�ڕW�Ƒ[�u�̕���W���w��̏ꏊ�ɂȂ������ꍇ�͎蓮�Ŏ捞�B���t�E�������ރt�H���_�ɃR�s�[
        If FileNames3 = "" Then
            tmp = MsgBox("�E�@�ڕW�Ƒ[�u�̕���W" & vbCrLf & "�t�@�C����I�����Ă��������B", vbYesNo + vbQuestion, "�m�F")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,�E�@�ڕW�Ƒ[�u�̕���W*.xls?")
                If OpenFileName <> "False" Then
                    FileCopy OpenFileName, SaveDir2 & "\" & "�E�@�ڕW�Ƒ[�u�̕���W.xlsm"
                    FileNames3 = Dir(SaveDir2 & "\" & "�E�@�ڕW�Ƒ[�u�̕���W*")
                Else
                    MsgBox "�L�����Z������܂���", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "�L�����Z������܂���", vbCritical
                ImportWbook.Close False
                Unload Me
                Exit Sub
            End If
        Else
        End If
        Workbooks.Open FileName:=SaveDir2 & "\" & FileNames3, ReadOnly:=True, UpdateLinks:=0
        Set ImportWbook3 = Workbooks.Open(SaveDir2 & "\" & FileNames3)
        Set iwbS06 = ImportWbook3.Worksheets("�B���Y����")
        Set iwbS07 = ImportWbook3.Worksheets("�C�o�c�Ǘ�")
        Set iwbS08 = ImportWbook3.Worksheets("�D�_�Ə]��")
        Set iwbS09 = ImportWbook3.Worksheets("�E���̑�")

'        �\�����ȈՔł̈��
        iwbS04.PrintOut From:=1, To:=1
        
'        ����ƖڕW�̐��l�f�[�^�̈��
        iwbS05.PrintOut
        
'        �ڕW�Ƒ[�u�̕���W���
        iwbS06.PrintOut
        iwbS07.PrintOut
        iwbS08.PrintOut
        iwbS09.PrintOut

'        �������f�[�^�̕ۑ�
        SetFileName = SaveDir2 & "\00�F��_�Ǝ҃f�[�^\���t�E��������\�A���P�[�g�t�H���_\" & TmpChar & SetName & "�D�@�V�F��v��\�����i�������p���l�̂݁j.xlsx"
        ImportWbook2.SaveAs FileName:=SetFileName
        ImportWbook2.Close
        ImportWbook3.Close
        
        ImportWbook.Close
        Application.DisplayAlerts = True
    Next
    
'    �������Ԍ���
    endTime = Timer
    processTime = endTime - startTime
    MsgBox "������I�����܂����B���Ԃ́F" & processTime & "�b�ł��B"
    
'   ���{�o�^
    twbS00.Range("A9").Interior.ColorIndex = 44
    twbS00.Range("A9").Font.ColorIndex = 1
    twbS00.Range("E9") = "" & Now & ""

Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
End Sub

Private Sub CommandButton8_Click()
'�H�@�J�����_�[�쐬�i0_�����\�̃f�[�^���j

Dim Week_Value As Integer, sr As Integer, tmp As Integer, i As Integer, k As Integer, sc As Integer, j As Integer
    Dim arrayDate() As Variant, arrayName() As Variant, arrayNo() As Variant, arrayGroup() As Variant
    Dim arrayTime() As Variant, arrayHole1() As Variant, arrayHole2() As Variant, arrayWeek() As Variant
    Dim arrayTempNumber() As Variant, arrayWeekValue() As Variant, arrayEarlyweek() As Variant
    Dim SetDate As Date
    Dim twbS00 As Worksheet, twbS01 As Worksheet, twbS02 As Worksheet, twbS03 As Worksheet
    Dim ThisWbook As Workbook
    Set twbS00 = Worksheets("�t�H�[���ďo")
    Set twbS01 = Worksheets("�J�����_�[")
    Set twbS02 = Worksheets("0_�����\")
    Set twbS03 = Worksheets("2_�\���E�F���")
On Error GoTo myError
    Unload Me
    tmp = twbS02.Range("F" & Rows.Count).End(xlUp).Row
    
'    0_�����\�V�[�g�̃f�[�^����t�������ԏ��Ŏw��͈͂̃\�[�g���s��
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
    
'    ��̒l�̍폜
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
        
'        �f�[�^���P�s�������ꍇ
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
        
'        0_�����\�f�[�^���J�����_�[�֓]�L����
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
        
'        0_�����\����2_�\���E�F����V�[�g�֎����]�L
        tmp = twbS03.Range("E" & Rows.Count).End(xlUp).Row
        twbS03.Range("E2" & ":" & "E" & tmp).ClearContents
        sr = 3
        For i = LBound(arrayName) To UBound(arrayName)
            twbS03.Cells(sr, 5) = arrayName(i, 1)
            sr = sr + 1
        Next i
        .Select
        MsgBox "�������܂����B"
        Application.ScreenUpdating = True
        
'       ���{�o�^
        twbS00.Range("A10").Interior.ColorIndex = 44
        twbS00.Range("A10").Font.ColorIndex = 1
        twbS00.Range("E10") = "" & Now & ""
    End With
    Exit Sub
myError:
    MsgBox "�����͂�����܂��I�������I�����܂��B", vbExclamation
End Sub

Private Sub CommandButton9_Click()
'�J�@�������z�擾�i�������z�A�c�_�ތ^�A�A���捞�j

Dim tmp As Integer, ReturnValue As Integer, TempNumber As Integer, LastRow As Integer, i As Integer
    Dim SetName As String, OpenFileName As String, myFileName As Integer
    Dim FileName As String, Path As String, SetFile As String
    Dim ThisWbook As Workbook, ImportWbook As Workbook
    Dim FoundValue As Range
    Dim twbS00 As Worksheet, twbS03 As Worksheet, twbS04 As Worksheet
    Dim iwbS01 As Worksheet, iwbS02 As Worksheet, iwbS03 As Worksheet
    Dim sr As Integer, sr_Addend As Integer, SetRow_Next As Integer
    Set ThisWbook = ActiveWorkbook
    Set twbS00 = Worksheets("�t�H�[���ďo")
    Set twbS04 = Worksheets("3_�ꗗ���R���\")
    Set twbS03 = Worksheets("2_�\���E�F���")
On Error GoTo myError
    Application.DisplayAlerts = False
    
'    �F��ԍ����ɕ��בւ�
    twbS03.Range("A2").Sort key1:=twbS03.Range("A3"), order1:=xlAscending, Header:=xlYes
    
'    �f�[�^���͒l�ȉ���50�s���폜���s��
    LastRow = twbS03.Cells(twbS03.Rows.Count, 4).End(xlUp).Row
    sr_Addend = LastRow + 1
    SetRow_Next = LastRow + 50
    twbS03.Rows(sr_Addend & ":" & SetRow_Next).Delete
    sr_Addend = LastRow + 2
    SetRow_Next = LastRow + 50
    twbS04.Rows(sr_Addend & ":" & SetRow_Next).Delete
    
'    �ꗗ���R���\��\��
    twbS04.Select
    Application.Wait Now + TimeValue("00:00:02")
    LastRow = twbS03.Cells(twbS03.Rows.Count, 5).End(xlUp).Row
    Unload UserForm5
    Application.ScreenUpdating = False
    

    For i = 3 To LastRow
Label1:
Label2:
'        �����̓`�F�b�N
        If twbS03.Cells(i, 2).Value = "" Then
            twbS03.Select
            MsgBox "2_�\���E�F����̋敪����͂��ĉ������B"
            Unload Me
            Exit Sub
        End If

'        �\�����̎捞���s�������\�����E�c�_�ތ^�E�ڕW�����̒l�擾
        If twbS03.Cells(i, 2).Value Like "[�V�K,�ĔF��,�ύX]*" And twbS04.Cells(i, 11).Value = "" And twbS04.Cells(i, 12).Value = "" Then
            tmp = MsgBox("�o�^����Ă��Ȃ� " & twbS03.Cells(i, 5) & " �l�̐\������I�����Ă��������B", vbYesNo + vbQuestion, "�m�F")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")
                If OpenFileName <> "False" Then
                     SetFile = OpenFileName
                Else
                    MsgBox "�L�����Z������܂���", vbCritical
                    Unload Me
                    Exit Sub
                End If
                Workbooks.Open FileName:=SetFile, ReadOnly:=True, UpdateLinks:=0
                Set ImportWbook = Workbooks.Open(Path & SetFile)
                Set iwbS01 = ImportWbook.Worksheets("���̓V�[�g")
                Set iwbS02 = ImportWbook.Worksheets("�R���\")
                Set iwbS03 = ImportWbook.Worksheets("�ȈՔ�")
                If iwbS01.Range("D5") = "" Then
                    SetName = iwbS01.Range("M5")
                Else
                    SetName = iwbS01.Range("D5")
                End If
                
'                �����\��������Ύ�荞��
                If SetName = twbS03.Cells(i, 5) Then
                    Set FoundValue = twbS03.Range("E:E").Find(What:=SetName, LookAt:=xlPart)
                    If Not FoundValue Is Nothing Then
                        sr = twbS03.Range("E:E").Find(SetName).Row
                        If iwbS01.Range("AH6") = "��" Then
                            iwbS01.Range("U6").Copy
                            twbS04.Range("F" & sr).PasteSpecial xlPasteValues
                        End If
                        If iwbS01.Range("AH7") = "��" Then
                            iwbS01.Range("U7").Copy
                            twbS04.Range("G" & sr).PasteSpecial xlPasteValues
                        End If
                        If iwbS01.Range("AH8") = "��" Then
                            iwbS01.Range("U8").Copy
                            twbS04.Range("H" & sr).PasteSpecial xlPasteValues
                        End If
                        If iwbS01.Range("AH9") = "��" Then
                            iwbS01.Range("U9").Copy
                            twbS04.Range("I" & sr).PasteSpecial xlPasteValues
                        End If
    
                        iwbS01.Range("B14").Copy
                        twbS04.Range("J" & sr).PasteSpecial xlPasteValues
                        
'                        �������z�i����ƖڕW�j�擾
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
'                        �������ƈ�l������̏������z���قȂ�ꍇ�́A���ʏ����ŕ\������
                            If iwbS03.Range("Z56") <> "" Then
                                twbS04.Range("K" & sr) = Format(iwbS03.Range("Z56"), "0") & vbLf & "�i" & Format(iwbS03.Range("Z59"), "0") & "�j"
                            End If
                            If iwbS03.Range("AL56") <> "" Then
                                twbS04.Range("L" & sr) = Format(iwbS03.Range("AL56"), "0") & vbLf & "�i" & Format(iwbS03.Range("AL59"), "0") & "�j"
                            End If
                        End If
                    Else
                        MsgBox "�\���҈ꗗ�\�ƈ�v����f�[�^�[������܂���ł����B������x�m�F���Ă��������B"
                        ImportWbook.Close False
                        GoTo Label1
                        End
                    End If
                Else
                    MsgBox twbS03.Cells(i, 5) & " �l�̐\�����ł͂���܂���B������x�I�����ĉ������B"
                    ImportWbook.Close False
                    GoTo Label2
                End If
            Else
                MsgBox "�����𒆒f���܂�", vbCritical
                Unload Me
                Exit Sub
            End If
            ImportWbook.Close False
        Else
            Flag = True
'             MsgBox twbS03.Cells(i, 5) & " �l�̃f�[�^�͓��͂���Ă���悤�ł��B", vbYesNo + vbQuestion, "�m�F"
        End If
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    twbS04.Select
    Unload Me
    
'   ���{�o�^
    twbS00.Range("A12").Interior.ColorIndex = 44
    twbS00.Range("A12").Font.ColorIndex = 1
    twbS00.Range("E12") = "" & Now & ""
    If Flag Then
        MsgBox "�ꕔ�o�^����Ă���悤�ł��B�㏑���ۑ����܂��B"
    Else
        MsgBox "�o�^�������܂����B�㏑���ۑ����܂��B"
    End If
    On Error Resume Next
    ActiveWorkbook.Save
    If Err.Number > 0 Then MsgBox "�ۑ�����܂���ł���"
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbCritical
End Sub
Private Sub CommandButton10_Click()
'�K�@�ꗗ���R���\��JA���ɍ쐬
Dim sr As Integer, LastRow As Integer, tmp_Addend As Integer, i As Integer
    Dim sr_Addend As Integer, SetRow_Next As Integer
    Dim myFileName As String
    Dim Flag As Boolean
    Dim ThisWbook As Workbook
    Dim twbS01 As Worksheet, iwbS00 As Worksheet, twbS00 As Worksheet
On Error GoTo myError
    Set twbS01 = Worksheets("3_�ꗗ���R���\")
    Set twbS00 = Worksheets("�t�H�[���ďo")
    Unload Me
    With twbS01
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        tmp_Addend = .Cells(.Rows.Count, 4).End(xlUp).Row
        Workbooks.Add
        twbS01.Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        ActiveWorkbook.Sheets("3_�ꗗ���R���\").Name = "�R���ꗗ�\"
        MsgBox "�����t�p�ꗗ�R���\��ۑ����܂��B�ۑ��ꏊ���w�肵�ĉ������B"
        myFileName = Application.GetSaveAsFilename(InitialFileName:="�R���ꗗ�\", FileFilter:="Excel�u�b�N,*.xlsx")
        If myFileName <> "False" Then
            ActiveWorkbook.SaveAs FileName:=myFileName
            ActiveWorkbook.Close
        Else
            ActiveWorkbook.Close
            Exit Sub
        End If
        
'        �t�@�C���̍쐬
        Workbooks.Add
        twbS01.Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        ActiveWorkbook.Sheets("3_�ꗗ���R���\").Name = "JA�{�n�ܘa"
        Set iwbS00 = ActiveWorkbook.Worksheets("JA�{�n�ܘa")
        iwbS00.Range("A3:N50").ClearContents
        tmp = 3
        
'        JA�{�n���蓖��
        For i = 3 To tmp_Addend
            If .Cells(i, 4) Like "*�{�n�k" Or .Cells(i, 4) Like "*�h��" Or .Cells(i, 4) Like "*������" Or .Cells(i, 4) Like "*����" Or _
                .Cells(i, 4) Like "*�{�n��" Or .Cells(i, 4) Like "*�����V��" Or .Cells(i, 4) Like "*������" Or .Cells(i, 4) Like "*�{���{" Or _
                .Cells(i, 4) Like "*�k�l��" Or .Cells(i, 4) Like "*�쌴��" Or .Cells(i, 4) Like "*�k����" Or .Cells(i, 4) Like "*�{���V�x" Or _
                .Cells(i, 4) Like "*������" Or .Cells(i, 4) Like "*�Ð쒬" Or .Cells(i, 4) Like "*�۔���" Or .Cells(i, 4) Like "*�{�����͓�" Or _
                .Cells(i, 4) Like "*�����V��" Or .Cells(i, 4) Like "*��V��" Or .Cells(i, 4) Like "*���˒�" Or .Cells(i, 4) Like "*��쌴" Or _
                .Cells(i, 4) Like "*���l��" Or .Cells(i, 4) Like "*���c��" Or .Cells(i, 4) Like "*�u�`" Or .Cells(i, 4) Like "*���" Or _
                .Cells(i, 4) Like "*��l��" Or .Cells(i, 4) Like "*����" Or .Cells(i, 4) Like "*���Y��" Or .Cells(i, 4) Like "*���" Or _
                .Cells(i, 4) Like "*�鉺��" Or .Cells(i, 4) Like "*��쒬" Or .Cells(i, 4) Like "*�T�꒬�T��" Or .Cells(i, 4) Like "*�S�r" Or _
                .Cells(i, 4) Like "*�D�V����" Or .Cells(i, 4) Like "*�R�̎蒬" Or .Cells(i, 4) Like "*�T�꒬�H��" Or .Cells(i, 4) Like "*��]" Or _
                .Cells(i, 4) Like "*�l�蒬" Or .Cells(i, 4) Like "*�쌴�V��" Or .Cells(i, 4) Like "*��Y��" Or .Cells(i, 4) Like "*����" Or _
                .Cells(i, 4) Like "*��������" Or .Cells(i, 4) Like "*�z�K��" Or .Cells(i, 4) Like "*�d�F�y��" Or .Cells(i, 4) Like "*����" Or _
                .Cells(i, 4) Like "*�`��" Or .Cells(i, 4) Like "*�쒬" Or .Cells(i, 4) Like "*�{�n�x��" Then
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
        
'        �t�@�C���̕ۑ�
        MsgBox "JA�{�n�ܘa��ۑ����܂��B�ۑ��ꏊ���w�肵�ĉ������B"
        myFileName = Application.GetSaveAsFilename(InitialFileName:="JA�{�n�ܘa", FileFilter:="Excel�u�b�N,*.xlsx")

 '        JA�{�n�ܘa���f�[�^���󂾂�����A�t�@�C���͍쐬���Ȃ��B
        If tmp = 3 Then
            myFileName = False
        End If
        If myFileName <> "False" Then
            ActiveWorkbook.SaveAs FileName:=myFileName
            ActiveWorkbook.Close
        Else
            ActiveWorkbook.Close
            MsgBox "JA�{�n�ܘa���f�[�^���󂾂����ׁA�t�@�C���͍폜���܂����B"
        End If
        
'        �t�@�C���̍쐬
        Workbooks.Add
        twbS01.Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        ActiveWorkbook.Sheets("3_�ꗗ���R���\").Name = "JA���܂���"
        Set iwbS00 = ActiveWorkbook.Worksheets("JA���܂���")
        iwbS00.Range("A3:N50").ClearContents
        tmp = 3
        
'        JA���܂������蓖��
        For i = 3 To tmp_Addend
            If .Cells(i, 4) Like "*���ɒÒ�" Or .Cells(i, 4) Like "*�n��" Or .Cells(i, 4) Like "*�͉Y*" Or .Cells(i, 4) Like "*��]�R�Y" Or _
                .Cells(i, 4) Like "*���[��" Or .Cells(i, 4) Like "*�Ð�" Or .Cells(i, 4) Like "*���x" Or .Cells(i, 4) Like "*���c��" Or _
                .Cells(i, 4) Like "*�v�ʒ�" Or .Cells(i, 4) Like "*���D��" Or .Cells(i, 4) Like "*���" Or .Cells(i, 4) Like "*���c�k" Or _
                .Cells(i, 4) Like "*���ђ�" Or .Cells(i, 4) Like "*�ԍ�" Or .Cells(i, 4) Like "*���x" Or .Cells(i, 4) Like "*���l��" Or _
                .Cells(i, 4) Like "*��Y�����Y" Or .Cells(i, 4) Like "*�{�q" Or .Cells(i, 4) Like "*���" Or .Cells(i, 4) Like "*���l�k" Or _
                .Cells(i, 4) Like "*��Y���T�Y" Or .Cells(i, 4) Like "*��Y" Or .Cells(i, 4) Like "*����" Or .Cells(i, 4) Like "*���A��" Or _
                .Cells(i, 4) Like "*�[�C��" Or .Cells(i, 4) Like "*���" Or .Cells(i, 4) Like "*�{��͓�" Or .Cells(i, 4) Like "*���l�k" Or _
                .Cells(i, 4) Like "*�Y" Or .Cells(i, 4) Like "*��ÉY" Or .Cells(i, 4) Like "*�H��" Or .Cells(i, 4) Like "*��]��" Or _
                .Cells(i, 4) Like "*�{�c" Or .Cells(i, 4) Like "*���ÉY" Or .Cells(i, 4) Like "*�v��" Or .Cells(i, 4) Like "*���{�n" Or _
                .Cells(i, 4) Like "*�I��" Or .Cells(i, 4) Like "*�����q" Or .Cells(i, 4) Like "*���؉͓�" Or .Cells(i, 4) Like "*�命��" Or _
                .Cells(i, 4) Like "*�œc" Or .Cells(i, 4) Like "*�哇�q" Or .Cells(i, 4) Like "*�V��" Or .Cells(i, 4) Like "*��{�n" Or _
                .Cells(i, 4) Like "*�͓�" Or .Cells(i, 4) Like "*���c" Or .Cells(i, 4) Like "*��]" Or .Cells(i, 4) Like "*����" Or _
                .Cells(i, 4) Like "*�V�a��" Or .Cells(i, 4) Like "*�q��" Or .Cells(i, 4) Like "*���Y" Or .Cells(i, 4) Like "*�䏊�Y��" Or .Cells(i, 4) Like "*���c" Then
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

'        �t�@�C���̕ۑ�
        MsgBox "JA���܂�����ۑ����܂��B�ۑ��ꏊ���w�肵�ĉ������B"
        myFileName = Application.GetSaveAsFilename(InitialFileName:="JA���܂���", FileFilter:="Excel�u�b�N,*.xlsx")
        
'        JA���܂������f�[�^���󂾂�����A�t�@�C���͍쐬���Ȃ��B
        If tmp = 3 Then
            myFileName = False
        End If
        If myFileName <> "False" Then
            ActiveWorkbook.SaveAs FileName:=myFileName
            ActiveWorkbook.Close
        Else
            ActiveWorkbook.Close
            MsgBox "JA���܂������f�[�^���󂾂����ׁA�t�@�C���͍폜���܂����B"
        End If
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End With
    
'   ���{�o�^
        twbS00.Range("A13").Interior.ColorIndex = 44
        twbS00.Range("A13").Font.ColorIndex = 1
        twbS00.Range("E13") = "" & Now & ""
        MsgBox "�t�@�C���̕������������܂����B"
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
End Sub

Private Sub CommandButton11_Click()
'�L�@�F��؏����t��������i�p�Q�j
Dim sr As Integer, LastRow As Integer, Value As Integer
    Dim ThisWbook As Workbook
    Dim twbS08 As Worksheet, twbS03 As Worksheet, twbS00 As Worksheet
    Dim InputValue As Integer, i As Integer, sRow As Integer, tmpValue As Integer
    Dim saiNintei As Integer, henkou As Integer, shinki As Integer, zitai As Integer
    Dim SetMsg As String
    Dim tmp As VbMsgBoxResult
    Set twbS03 = Worksheets("2_�\���E�F���")
    Set twbS08 = Worksheets("�F��񍐈��")
    Set twbS00 = Worksheets("�t�H�[���ďo")
On Error GoTo myError
    MsgBox "�������O�Ƀv�����^�[�ݒ�" & vbLf & "�i������s�������v�����^�[��ʏ�g�p����v�����^�[�Ɂj" & vbLf & "���Ă����ĉ������B", vbExclamation
    tmp = MsgBox("���������́uC��v�Ƀ`�F�b�N�����Ă����ĉ������B" & vbLf & "�v���r���[��A1�l������́~�Ŋm�F���Ă��������B", vbYesNo)
    Unload Me
    With twbS03
    If tmp = vbNo Then Exit Sub
        Application.ScreenUpdating = False
        
'        �����`�F�b�N
        Value = WorksheetFunction.CountA(.Range("C:C")) - 1
        If Value = 0 Then
            MsgBox "�������ꍇ�AC��Ƀ`�F�b�N�����Ď��s���ĉ������B", vbExclamation
            Exit Sub
        End If
        LastRow = .Cells(.Rows.Count, 5).End(xlUp).Row
        
'        �敪�̐��l�`�F�b�N
        For sr = 3 To LastRow
            If .Cells(sr, 2).Value = "�ĔF��" Then
            saiNintei = saiNintei + 1
            End If
            If .Cells(sr, 2).Value = "�ύX" Then
            henkou = henkou + 1
            End If
            If .Cells(sr, 2).Value = "�V�K" Then
            shinki = shinki + 1
            End If
            If .Cells(sr, 2).Value = "����" Then
            zitai = zitai + 1
            End If
        Next
        tmpValue = .Cells(.Rows.Count, 11).End(xlUp).Row
        
'        ����v���r���[
        For sr = 3 To LastRow
            If .Cells(sr, 2).Value <> "����" And .Cells(sr, 3).Value = "��" Then
                If Value = 0 Or twbS03.Cells(sr, 5) = "" Or twbS03.Cells(sr, 11) = "" Or twbS03.Cells(sr, 12) = "" Then
                    MsgBox "�����͂̉ӏ�������܂��B�m�F���čĎ��s���ĉ������B", vbExclamation
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

'        �敪�̐��l����
        For i = 3 To 14
            If twbS00.Cells(i, 7) = Month(.Cells(tmpValue, 11)) Then
                sRow = twbS00.Cells(i, 7).Row
            End If
        Next i
        twbS00.Range("J" & sRow) = saiNintei
        twbS00.Range("K" & sRow) = shinki
        twbS00.Range("L" & sRow) = zitai
        twbS00.Range("M" & sRow) = henkou
        
'        �v���r���[�̈��
        Application.ScreenUpdating = True
        tmp = MsgBox("�v���r���[���������܂����B�����Ĉ�����܂��B��낵���ł����H", vbYesNo)
        If tmp = vbNo Then Exit Sub
            Application.ScreenUpdating = False
            LastRow = .Cells(.Rows.Count, 5).End(xlUp).Row
        For sr = 3 To LastRow
            If .Cells(sr, 2).Value <> "����" And .Cells(sr, 3).Value = "��" Then
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
    
'   ���{�o�^
    twbS00.Range("A15").Interior.ColorIndex = 44
    twbS00.Range("A15").Font.ColorIndex = 1
    twbS00.Range("E15") = "" & Now & ""
    MsgBox "������������܂����B"
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
End Sub

Private Sub CommandButton12_Click()
'�M�@�\�����t�@�C�����l�[��
Dim tmp As Integer, sr As Integer, LastRow As Integer, SetRow As Integer, SetDate As Integer
    Dim i As Integer, SearchValue As Integer, LastRow2 As Integer, k As Integer
    Dim OpenFileName As String, FileName As String, Path As String, SetFile As String, ThisWbookPass As String, SetFileName As String
    Dim SetName As String, ReturnValue As String, myFileName As String, TmpChar As String, SaveDir As String
    Dim FoundValue As Object
    Dim Flag As Boolean
    Dim ThisWbook As Workbook, ImportWbook2 As Workbook
    Dim twbS00 As Worksheet, twbS03 As Worksheet, twbS04 As Worksheet, iwbS11 As Worksheet
    Set ThisWbook = ActiveWorkbook
    Set twbS03 = Worksheets("2_�\���E�F���")
    Set twbS04 = Worksheets("3_�ꗗ���R���\")
    Set twbS00 = Worksheets("�t�H�[���ďo")
On Error GoTo myError
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    With twbS03
    LastRow = .Cells(.Rows.Count, 5).End(xlUp).Row
    UserForm6.Show
    
'    �����t�H���_�̍쐬
    tmp = MsgBox("�\������ۑ�����t�H���_�𓯂��K�w�ɍ쐬���܂��B", vbYesNo + vbQuestion, "�m�F")
        If tmp = vbYes Then
            Flag = True
            SaveDir = ActiveWorkbook.Path & "\00�F��_�Ǝ҃f�[�^�i" & index & "�������j����"
            If Dir(SaveDir, vbDirectory) = "" Then
                MkDir SaveDir
                MsgBox "���̃t�@�C���Ɠ����K�w�Ɂu00�F��_�Ǝ҃f�[�^�i" & index & "�������j�����v�Ƃ����ۑ��t�H���_���쐬���܂����B"
            End If
        Else
            Flag = False
            MsgBox "�����I�ɍ쐬���ꂽ�t�H���_���ɕۑ�����܂��̂ŁA�t�H���_�͍쐬���ĉ������B"
        End If
        
'    �V�K�ƍĔF��̂݃t�@�C�����̍X�V���s��
    For i = 3 To LastRow
        If .Cells(i, 2).Value = "�V�K" Or .Cells(i, 2).Value = "�ĔF��" Or .Cells(i, 2).Value = "�ύX" Then
Label1:
            tmp = MsgBox(.Cells(i, 5).Value & " �l�̃t�@�C����I�����Ă��������B", vbYesNo + vbQuestion, "�m�F")
            If tmp = vbYes Then
                OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")
                If OpenFileName <> "False" Then
                    SetFile = OpenFileName
                Else
                    MsgBox "�L�����Z������܂���", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "�L�����Z������܂���", vbCritical
                Unload Me
                Exit Sub
            End If
            Workbooks.Open FileName:=SetFile, ReadOnly:=True, UpdateLinks:=0
            Set ImportWbook2 = Workbooks.Open(Path & SetFile)
            Windows(ImportWbook2.Name).Visible = False
            Set iwbS11 = ImportWbook2.Worksheets("���̓V�[�g")
            
            If iwbS11.Range("D5") = "" Then
                SetName = iwbS11.Range("M5")
            Else
                SetName = iwbS11.Range("D5")
            End If
            If .Cells(i, 5).Value <> SetName Then
                MsgBox "�Z�b�g�����t�@�C��" & SetName & "�l�ł͂���܂���B�ēx�I�����ĉ������B"
                ImportWbook2.Close
                GoTo Label1
            End If
            
            ImportWbook2.Activate
            
'            MsgBox "�t�@�C�����i�F��ԍ��j���ԈႢ�Ȃ����m�F���ۑ����ĉ������B"
            myFileName = Format(twbS00.Range("G1"), "[DBNum3][$]e") & "-" & .Cells(i, 1).Value & "�F��\�����i" & SetName & " �j" & ".xlsm"
            ImportWbook2.SaveAs FileName:=SaveDir & "\" & myFileName
            Unload Me
            Windows(ImportWbook2.Name).Visible = True
            ImportWbook2.Close
        End If
    Next
        Application.DisplayAlerts = True
        MsgBox "�\�������l�[�����������܂����B"
        
 '   ���{�o�^
        twbS00.Range("A16").Interior.ColorIndex = 44
        twbS00.Range("A16").Font.ColorIndex = 1
        twbS00.Range("E16") = "" & Now & ""
    End With
    Exit Sub
myError:
    MsgBox "�I�������t�@�C�����Ⴂ�܂��I�������I�����܂��B", vbExclamation
End Sub

Private Sub CommandButton13_Click()
'�N�@�F��_�Ǝ҃f�[�^�x�[�X�o�^�E�폜
UserForm3.Show
End Sub

Private Sub CommandButton14_Click()
'�O�@�F��_�Ǝ҃f�[�^�x�[�X�ҏW�i�l�E�@�l�j
UserForm4.Show
End Sub

Private Sub CommandButton15_Click()
'�P�@������
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
    Set twbS00 = Worksheets("�t�H�[���ďo")
    Set twbS01 = Worksheets("�J�����_�[")
    Set twbS02 = Worksheets("0_�����\")
    Set twbS03 = Worksheets("2_�\���E�F���")
    Set twbS04 = Worksheets("3_�ꗗ���R���\")
    Set twbS05 = Worksheets("data")
    Set twbS06 = Worksheets("���ʒ��o")
    Set twbS10 = Worksheets("�ڎ�-�F���")
    Set twbS11 = Worksheets("�ڎ�-���ގ�")
On Error GoTo myError
    With twbS10
        Application.DisplayAlerts = False
        TmpChar = .Cells(.Rows.Count, 4).End(xlUp)
        Set StringObject = twbS03.Range("A2", "A200")
        
        If TmpChar = "�������͈���" Then
            MsgBox "�{�N�x����o�^�ł��B�ڎ��V�[�g�ɓo�^���܂��B"
            Flag = False
        Else
            MsgBox "�{�N�x2��ڈȍ~�̓o�^�ł��B�ڎ��V�[�g�ɓo�^���܂��B"
            Flag = True
        End If
        
        Value = twbS03.Range("B" & Rows.Count).End(xlUp).Row
        sr = .Range("D" & Rows.Count).End(xlUp).Row + 1
        sr0 = twbS11.Range("D" & Rows.Count).End(xlUp).Row + 1
        For i = 2 To Value
            If WorksheetFunction.CountIf(.Range("D3:D122"), twbS03.Cells(i, 5)) > 0 Then
                MsgBox "���ꖼ�����݂��܂��B" & vbCrLf & "�d���o�^�ł͂���܂��񂩁H�ڎ��V�[�g�ƃ`�F�b�N���ĉ������B" & vbCrLf & "��U�������I�����܂��B"
                Unload Me
                Exit Sub
            End If
        Next i
        For i = 3 To Value
            If twbS03.Cells(i, 2).Value Like "[�V�K,�ĔF��,�ύX]*" Then
                .Cells(sr, 1) = twbS03.Cells(i, 1)
                .Cells(sr, 3) = twbS03.Cells(i, 12)
                .Cells(sr, 4) = twbS03.Cells(i, 5)
                .Cells(sr, 5) = "�_�ƌo�c���P�v��F��\�����i" & twbS03.Cells(i, 2) & ")"
                sr = sr + 1
            ElseIf twbS03.Cells(i, 2).Value Like "����" Then
                twbS11.Cells(sr0, 3) = twbS03.Cells(i, 12)
                twbS11.Cells(sr0, 4) = twbS03.Cells(i, 5)
                twbS11.Cells(sr0, 5) = "�_�ƌo�c���P�v��F��\�����i" & twbS03.Cells(i, 2) & ")"
                sr0 = sr0 + 1
            Else
            End If
        Next i
        
        MsgBox "�{�Ɩ��̌������͂����肢���܂��B"
        UserForm6.Show
        
'   ���{�o�^
        twbS00.Range("A19").Interior.ColorIndex = 44
        twbS00.Range("A19").Font.ColorIndex = 1
        twbS00.Range("E19") = "" & Now & ""
    End With
    
    Unload UserForm5
    SaveDir = ThisWbook.Path
    myFileName = "10�\���҈ꗗ�\�i" & Format(twbS00.Range("G1"), "[DBNum3][$]ggge") & "�N�x" & index & "�����j����"
    ActiveWorkbook.SaveAs FileName:=SaveDir & "\" & myFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    MsgBox myFileName & "�@���ŕۑ����܂����B" & vbCrLf & "�ڎ��̔F��ԍ��̃Y�����������A�e�V�[�g���m�F���ĉ������B"

'    �����t�@�C�����쐬����i�N�x���͍쐬���Ȃ��j
    If index <> 3 Then
        MsgBox "�����t�@�C�����쐬���܂��B"
        If index = 12 Then
            nextIndex = 1
            nextYear = Format(twbS00.Range("G1"), "[$-ja-JP]e") + 1
        Else
            nextIndex = index + 1
            nextYear = Format(twbS00.Range("G1"), "[$-ja-JP]e")
        End If
        
        
    
        ParentDirName = Left(ThisWbook.Path, InStrRev(ThisWbook.Path, "\") - 1)

        SaveDir = ParentDirName & "\R" & nextYear & "." & nextIndex
        
        myFileName = "10�\���҈ꗗ�\�i" & Format(twbS00.Range("G1"), "[DBNum3][$]ggge") & "�N�x" & nextIndex & "�����j"
        If Dir(SaveDir, vbDirectory) = "" Then
            MkDir SaveDir
            MsgBox "���̊K�w��R" & nextYear & "." & nextIndex & "�Ƃ����ۑ��t�H���_���쐬���A���̒��� " & myFileName & "�t�@�C�����쐬���܂��B"
        End If
    
        ActiveWorkbook.SaveAs FileName:=SaveDir & "\" & myFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        
        '������
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
        
        '�F��ԍ��ݒ�
        tmp = twbS10.Cells(twbS10.Rows.Count, 3).End(xlUp).Row + 1
        
        twbS03.Range("A3") = twbS10.Cells(tmp, 1)
        ActiveWorkbook.Save
        Application.DisplayAlerts = True
        MsgBox "�ۑ����������܂����B�����Ɏg�p���ĉ������B"
    Else
        MsgBox "�F��_�ƎҍX�V�葱������ꂳ�܂ł����B�{�N�x�͏I���ƂȂ�܂��B"
    End If
    Exit Sub
myError:
    MsgBox "�\�����ʃG���[���������܂����I�������I�����܂��B", vbExclamation
End Sub


