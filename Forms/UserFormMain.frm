VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMain 
   Caption         =   "Product Tracker v01beta"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15045
   OleObjectBlob   =   "UserFormMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rangeCellItemCode As Variant
Public rangexCellItemCode As Variant
Public lRow As Long, lxRow As Long
Public icode As Integer
Public itemCode As String, desk As String, rpsStart As String, rpsEnd As String
Public wf As WorksheetFunction
Public cell As Range
Public Fini As Boolean
Public Lg As Single
Public Ht As Single
Public oDictionary As Object
Public Thisrow As Long

'API functions
Private Declare Function GetWindowLong Lib "user32" _
                         Alias "GetWindowLongA" _
                         (ByVal hWnd As Long, _
                          ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
                         Alias "SetWindowLongA" _
                         (ByVal hWnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" _
                         (ByVal hWnd As Long, _
                          ByVal hWndInsertAfter As Long, _
                          ByVal x As Long, _
                          ByVal Y As Long, _
                          ByVal cx As Long, _
                          ByVal cy As Long, _
                          ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" _
                         Alias "FindWindowA" _
                         (ByVal lpClassName As String, _
                          ByVal lpWindowName As String) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" _
                         () As Long
Private Declare Function SendMessage Lib "user32" _
                         Alias "SendMessageA" _
                         (ByVal hWnd As Long, _
                          ByVal wMsg As Long, _
                          ByVal wParam As Long, _
                          lParam As Any) As Long
Private Declare Function DrawMenuBar Lib "user32" _
                         (ByVal hWnd As Long) As Long


'Constants
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const WS_EX_APPWINDOW = &H40000
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&

Private Sub AppTasklist(myForm)
'Add this userform into the Task bar
    Dim WStyle As Long
    Dim Result As Long
    Dim hWnd As Long

    hWnd = FindWindow(vbNullString, myForm.Caption)
    WStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    WStyle = WStyle Or WS_EX_APPWINDOW
    Result = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, _
                          SWP_NOMOVE Or _
                          SWP_NOSIZE Or _
                          SWP_NOACTIVATE Or _
                          SWP_HIDEWINDOW)
    Result = SetWindowLong(hWnd, GWL_EXSTYLE, WStyle)
    Result = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, _
                          SWP_NOMOVE Or _
                          SWP_NOSIZE Or _
                          SWP_NOACTIVATE Or _
                          SWP_SHOWWINDOW)
End Sub

Private Sub AddIcon()
'Add an icon on the titlebar
    Dim hWnd As Long
    Dim lngRet As Long
    Dim hIcon As Long
    hIcon = Sheet16.Image1.Picture.Handle
    hWnd = FindWindow(vbNullString, Me.Caption)
    lngRet = SendMessage(hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
    lngRet = SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    lngRet = DrawMenuBar(hWnd)
End Sub

Private Sub Label1_Click()
    Libraries_Modules.GotoWeb
End Sub

Private Sub LabelBreakdown_Click()
    If Me.ListBoxBreakdown.ListCount <= 0 Then
    'do nothing listbox kosong
    Else
        If IsNull(ListBoxBreakdown) Then
            Run "SortListBox", ListBoxBreakdown, 2, 1, 1 'order breakdown, alfabet, ascending
    'label breakdown tambah panah
            With LabelBreakdown
                .Caption = ChrW(9660) & " Breakdown"
                .BackColor = &H8000000D          'highlight
            '.BackColor = &H8000000F  'button face
            End With
    'label time tanpa panah
            With Me.LabelBreakdownTime
                .Caption = "Time"
                .BackColor = &H8000000F          'button face
            End With
    'label kode tanpa panah
            With Me.LabelKodeBreakdown
                .Caption = "Kode"
                .BackColor = &H8000000F          'button face
            End With
        Else
            unselectListboxBreakdown
            Run "SortListBox", ListBoxBreakdown, 2, 1, 1 'order breakdown, alfabet, ascending
    'label breakdown tambah panah
            With LabelBreakdown
                .Caption = ChrW(9660) & " Breakdown"
                .BackColor = &H8000000D          'highlight
            '.BackColor = &H8000000F  'button face
            End With
    'label time tanpa panah
            With Me.LabelBreakdownTime
                .Caption = "Time"
                .BackColor = &H8000000F          'button face
            End With
    'label kode tanpa panah
            With Me.LabelKodeBreakdown
                .Caption = "Kode"
                .BackColor = &H8000000F          'button face
            End With
        End If
    End If
End Sub

Private Sub LabelBreakdownTime_Click()
    If Me.ListBoxBreakdown.ListCount <= 0 Then
    'do nothing listbox kosong
    Else
        If IsNull(ListBoxBreakdown) Then
            Run "SortListBox", ListBoxBreakdown, 0, 2, 2 'order time, number, desc
    'label breakdown tambah panah
            With LabelBreakdown
                .Caption = "Breakdown"
                .BackColor = &H8000000F          'button face
            End With
    'label time tanpa panah
            With Me.LabelBreakdownTime
                .Caption = ChrW(9660) & " Time"
                .BackColor = &H8000000D          'highlight
            End With
    'label kode tanpa panah
            With Me.LabelKodeBreakdown
                .Caption = "Kode"
                .BackColor = &H8000000F          'button face
            End With
        Else
            unselectListboxBreakdown
            Run "SortListBox", ListBoxBreakdown, 0, 2, 2 'order time, number, desc
    'label breakdown tambah panah
            With LabelBreakdown
                .Caption = "Breakdown"
                .BackColor = &H8000000F          'button face
            End With
    'label time tanpa panah
            With Me.LabelBreakdownTime
                .Caption = ChrW(9660) & " Time"
                .BackColor = &H8000000D          'highlight
            End With
    'label kode tanpa panah
            With Me.LabelKodeBreakdown
                .Caption = "Kode"
                .BackColor = &H8000000F          'button face
            End With
        End If
    End If
End Sub

Private Sub sortByBreakdownTime()
    If Me.ListBoxBreakdown.ListCount <= 0 Then
    'do nothing listbox kosong
    Else
        Run "SortListBox", ListBoxBreakdown, 0, 2, 2 'order time, number, desc
    'label breakdown tambah panah
        With LabelBreakdown
            .Caption = "Breakdown"
            .BackColor = &H8000000F              'button face
        End With
    'label time tanpa panah
        With Me.LabelBreakdownTime
            .Caption = ChrW(9660) & " Time"
            .BackColor = &H8000000D              'highlight
        End With
    'label kode tanpa panah
        With Me.LabelKodeBreakdown
            .Caption = "Kode"
            .BackColor = &H8000000F              'button face
        End With
    End If
End Sub

Private Sub LabelBreakdownTimeUniq_Click()
    sortByTimeUniq
End Sub

Private Sub sortByTimeUniq()
    If Me.ListBoxBreakdownUniq.ListCount <= 0 Then
    'do nothing listbox kosong
    Else
        If IsNull(ListBoxBreakdownUniq) Then
            Run "SortListBox", ListBoxBreakdownUniq, 1, 2, 2 'order time, number, desc
    'label breakdown tambah panah
            With LabelBreakdownUniq
                .Caption = "Breakdown"
                .BackColor = &H8000000F          'button face
            End With
    'label time tanpa panah
            With Me.LabelBreakdownTimeUniq
                .Caption = ChrW(9660) & " Time"
                .BackColor = &H8000000D          'highlight
            End With
    'label kode tanpa panah
            With Me.LabelKodeBreakdownUniq
                .Caption = "Kode"
                .BackColor = &H8000000F          'button face
            End With
        Else
            If IsNull(ListBoxBreakdownUniq) Then
        'MsgBox "Nothing Selected"
            Else
                unselectListboxBreakdownUniq
            End If
            Run "SortListBox", ListBoxBreakdownUniq, 1, 2, 2 'order time, number, desc
    'label breakdown tambah panah
            With LabelBreakdownUniq
                .Caption = "Breakdown"
                .BackColor = &H8000000F          'button face
            End With
    'label time tanpa panah
            With Me.LabelBreakdownTimeUniq
                .Caption = ChrW(9660) & " Time"
                .BackColor = &H8000000D          'highlight
            End With
    'label kode tanpa panah
            With Me.LabelKodeBreakdownUniq
                .Caption = "Kode"
                .BackColor = &H8000000F          'button face
            End With
        End If
    End If
End Sub

Private Sub LabelKodeBreakdown_Click()
    If Me.ListBoxBreakdown.ListCount <= 0 Then
    'do nothing listbox kosong
    Else
        If IsNull(ListBoxBreakdown) Then
            Run "SortListBox", ListBoxBreakdown, 1, 1, 1 'order kode, alfabet, asc
    'label breakdown tambah panah
            With LabelBreakdown
                .Caption = "Breakdown"
                .BackColor = &H8000000F          'button face
            End With
    'label time tanpa panah
            With Me.LabelBreakdownTime
                .Caption = "Time"
                .BackColor = &H8000000F          'button face
            End With
    'label kode tanpa panah
            With Me.LabelKodeBreakdown
                .Caption = ChrW(9660) & " Kode"
                .BackColor = &H8000000D          'highlight
            End With
        Else
            unselectListboxBreakdown
            Run "SortListBox", ListBoxBreakdown, 1, 1, 1 'order kode, alfabet, asc
    'label breakdown tambah panah
            With LabelBreakdown
                .Caption = "Breakdown"
                .BackColor = &H8000000F          'button face
            End With
    'label time tanpa panah
            With Me.LabelBreakdownTime
                .Caption = "Time"
                .BackColor = &H8000000F          'button face
            End With
    'label kode tanpa panah
            With Me.LabelKodeBreakdown
                .Caption = ChrW(9660) & " Kode"
                .BackColor = &H8000000D          'highlight
            End With
        End If
    End If
End Sub

Private Sub ListBoxBreakdown_Click()
    Dim i As Long
    With Me.ListBoxBreakdown
        i = .ListIndex
        MsgBox .List(i, 2)
    End With
End Sub

Private Sub ListBoxBreakdown_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If IsNull(ListBoxBreakdown) Then
        'MsgBox "Nothing Selected"
    Else
        unselectListboxBreakdown
    End If
End Sub

Private Sub ListBoxBreakdownUniq_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If IsNull(ListBoxBreakdownUniq) Then
        'MsgBox "Nothing Selected"
    Else
        unselectListboxBreakdownUniq
    End If
End Sub

Private Sub unselectListboxBreakdown()
On Error Resume Next
    ListBoxBreakdown.Selected(ListBoxBreakdown.ListIndex) = False
End Sub

Private Sub unselectListboxBreakdownUniq()
On Error Resume Next
    ListBoxBreakdownUniq.Selected(ListBoxBreakdownUniq.ListIndex) = False
End Sub

Private Sub ListBoxCounter_Click()
    With SheetInput
        .Activate
        .Range("$A$3:$AJ$9999").AutoFilter Field:=3, Operator:= _
                                                          xlFilterValues, Criteria2:=Array(2, ListBoxCounter.Value)
        .Range("$A$3:$AJ$9999").AutoFilter Field:=10, Criteria1:=TextBoxItemCode.Value
    End With
    fillListBoxBD
    sortByBreakdownTime
    ModuleMain.CreateUniqueListKodeDowntime
    ModuleMain.CreateUniqueListDowntime
'    Application.Wait (Now + TimeValue("00:00:03"))
    fillSheetDayCode
    CreateUniqueListSumIfsTime
    fillListBoxBDUniq
End Sub

Private Sub fillSheetDayCode()
    Dim i As Long
    With Me.ListBoxCounter
        i = .ListIndex
        SheetSumBD.Range("E2").Value = .List(i, 1)
    End With
'        SheetSumBD.Range("E2").Value = ListBoxCounter.List(i, 2)
End Sub

Private Sub TextBoxItemCode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TextBoxItemCode.Value = "" Then
            TextBoxItemCode.SetFocus
            TextBoxItemCode.SelStart = 0
        End If
    End If
End Sub

Private Sub UserForm_Activate()
    AddIcon
    AppTasklist Me
End Sub

Private Sub UserForm_Resize()
    Dim RtL As Single, RtH As Single
    If Me.Width < 300 Or Me.Height < 200 Or Fini Then Exit Sub
    RtL = Me.Width / Lg
    RtH = Me.Height / Ht
    Me.Zoom = IIf(RtL < RtH, RtL, RtH) * 100
End Sub

Private Sub UserForm_Terminate()
    Fini = True
End Sub

Private Sub showChart()
    Dim Fname As String
    Call SaveChart
    Fname = ThisWorkbook.Path & "\temp1.bmp"
    Me.Image1.Picture = LoadPicture(Fname)
End Sub

Private Sub SaveChart()
    Dim MyChart As Chart
    Dim Fname As String

    Set MyChart = Sheets("Daily").ChartObjects(1).Chart
    Fname = ThisWorkbook.Path & "\temp1.bmp"
    MyChart.Export Filename:=Fname, FilterName:="BMP"
End Sub

Private Sub TextBoxItemCode_Change()
    On Error Resume Next
    TextBoxItemCode = UCase(TextBoxItemCode)
    On Error GoTo 0
End Sub

Private Sub TextBoxItemCode_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                      ByVal x As Single, ByVal Y As Single)
    With TextBoxItemCode
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TextBoxItemCode_Enter()
Application.ScreenUpdating = False
    With SheetInput
        .Activate
        .Range("$A$3:$AJ$9999").AutoFilter Field:=3
        .Range("$A$3:$AJ$1000").AutoFilter Field:=10
    End With
Application.ScreenUpdating = True
End Sub

Private Sub TextBoxItemCode_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim produk As Variant
    If Me.TextBoxItemCode.Value = Me.LabelItemCodeResult.Caption Then
    '
        Exit Sub
    ElseIf TextBoxItemCode.Value = vbNullString Then
'        icode = MsgBox("Item Code tidak boleh kosong", vbRetryCancel + vbExclamation, "Item Code")
'        If icode = vbRetry Then
'        '== memaksa user untuk tidak mengosongkan kolom ==
'            TextBoxItemCode.SelStart = 0         'jika ada text maka kursor otomatis ke ujung kiri text
'            TextBoxItemCode.SelLength = TextBoxItemCode.TextLength 'semua text otomatis terblok
'            Exit Sub
'            Cancel = True
'        ElseIf icode = vbCancel Then
'            Cancel = False
'        End If
        Cancel = False
        TextBoxItemCode.SetFocus
        TextBoxItemCode.SelStart = 0
        Exit Sub
    Else
        '== handling error ketika item yang dimasukan tidak termasuk dalam list ==
        produk = Application.VLookup(TextBoxItemCode, Range("J:K"), 2, False)
        If IsError(produk) Then
            MsgBox "Item Code tidak ada/belum terdaftar", vbCritical
            With LabelProduk
                .BackColor = &H8080FF
                .Caption = "#N/A"
            End With
            With TextBoxItemCode
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            Exit Sub
            Cancel = True
        Else
        'On Error GoTo na
            LabelItemCodeResult.Caption = TextBoxItemCode.Value
            If LabelProduk.Caption = "Unknown Product" Then
                LabelProduk.BackColor = &H8080FF
                LabelItemCodeResult.BackColor = &H8080FF
            Else
                LabelProduk.BackColor = &H8000000A
                LabelItemCodeResult.BackColor = &H8000000A
            End If
            LabelProduk.Caption = produk
            ModuleMain.sumIfsFormula
            fillListBox
            fillListBoxWeekly
            showChart
            aveAR
            avePR
            aveQR
            aveOee
            ListBoxBreakdown.Clear
        End If
    End If
'na:     LabelProduk.Caption = "#N/A"
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer, L As Integer, TB
    InitMaxMin Me.Caption
    Ht = Me.Height
    Lg = Me.Width
    With SheetInput
        .Activate
        .Range("$A$3:$AJ$9999").AutoFilter Field:=3
        .Range("$A$3:$AJ$1000").AutoFilter Field:=10
    End With
    Me.Bar.Width = 0
End Sub

Sub fillListBox()
    itemCode = "DAY CODE"
    Set rangeCellItemCode = Cells.Find(What:=itemCode, LookIn:=xlFormulas, LookAt _
                                       :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                                       True, SearchFormat:=False)
    lRow = Cells(Rows.Count, rangeCellItemCode.Column).End(xlUp).Row
    ListBoxCounter.Clear
    With ListBoxCounter
        .ColumnWidths = "0;20;60;40;40;40;40"
        .ColumnCount = 7
        For Each cell In Range("C6:C" & lRow).SpecialCells(xlCellTypeVisible)
            .AddItem CStr(cell.Value)
            .List(.ListCount - 1, 1) = cell.Offset(0, -2).Value 'RPS
            .List(.ListCount - 1, 2) = cell.Offset(0, 0).Value 'cell.Offset(0, 0).Value 'produk
            .List(.ListCount - 1, 3) = Format(cell.Offset(0, 24).Value, "00.00%") 'ar
            .List(.ListCount - 1, 4) = Format(cell.Offset(0, 26).Value, "00.00%") 'pr
            .List(.ListCount - 1, 5) = Format(cell.Offset(0, 28).Value, "00.00%") 'qr
            .List(.ListCount - 1, 6) = Format(cell.Offset(0, 30).Value, "00.00%") 'oee
        Next cell
    End With
End Sub

Sub fillListBoxWeekly()
    ListBoxCounterWeekly.Clear
    With ListBoxCounterWeekly
        .ColumnWidths = "0;30;50;40;40;40;40"
        .ColumnCount = 7
        For Each cell In Worksheets("Weekly").Range("C5:C9").SpecialCells(xlCellTypeVisible)
            .AddItem CStr(cell.Value)
            .List(.ListCount - 1, 1) = cell.Offset(0, -1).Value 'Month
            .List(.ListCount - 1, 2) = cell.Offset(0, 0).Value 'cell.Offset(0, 0).Value 'produk
            .List(.ListCount - 1, 3) = Format(cell.Offset(0, 29).Value, "00.00%") 'ar
            .List(.ListCount - 1, 4) = Format(cell.Offset(0, 31).Value, "00.00%") 'pr
            .List(.ListCount - 1, 5) = Format(cell.Offset(0, 33).Value, "00.00%") 'qr
            .List(.ListCount - 1, 6) = Format(cell.Offset(0, 35).Value, "00.00%") 'oee
        Next cell
    End With
End Sub

Sub fillListBoxBDUniq_BACKUP()
    Set oDictionary = CreateObject("Scripting.Dictionary")
    desk = "DESKRIPSI KEGIATAN"
    Set rangexCellItemCode = SheetInput.Cells.Find(What:=desk, LookIn:=xlFormulas, LookAt _
                                                            :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                                                            True, SearchFormat:=False)
    lxRow = Cells(SheetInput.Rows.Count, rangexCellItemCode.Column).End(xlUp).Row
    Me.ListBoxBreakdownUniq.Clear
    With ListBoxBreakdownUniq
        .ColumnWidths = "50;40;150"
        .ColumnCount = 3
        For Each cell In SheetInput.Range("O6:O" & lxRow).SpecialCells(xlCellTypeVisible)
            Thisrow = cell.Row
            If Not cell.Rows.Hidden And Thisrow <> lxRow Then
                If oDictionary.exists(cell.Value) Then
                'Do nothing
                Else
                    oDictionary.Add cell.Value, 0
                    .AddItem CStr(cell.Value)
                    .List(.ListCount - 1, 1) = cell.Offset(0, -7).Value 'RPS
                    .List(.ListCount - 1, 2) = cell.Offset(0, 1).Value
                    .List(.ListCount - 1, 3) = cell.Offset(0, 3).Value 'RPS
                End If
            End If
            lxRow = Thisrow
        Next cell
    End With
'    Run "SortListBox", ListBoxBreakdownUniq, 1, 2, 2
' 'Sort by the 1st column in the ListBox Alphabetically in Ascending Order
'Run "SortListBox", Me.lbxSheet_Data, 0, 1, 1
 
' 'Sort by the 1st column in the ListBox Alphabetically in Descending Order
'Run "SortListBox", ListBox1, 0, 1, 2
'
' 'Sort by the 2nd column in the ListBox Numerically in Ascending Order
'Run "SortListBox", ListBox1, 1, 2, 1
'
' 'Sort by the 2nd column in the ListBox Numerically in Descending Order
'Run "SortListBox", ListBox1, 1, 2, 2
End Sub

Sub fillListBoxBD()
    desk = "DESKRIPSI KEGIATAN"
    Set rangexCellItemCode = SheetInput.Cells.Find(What:=desk, LookIn:=xlFormulas, LookAt _
                                                            :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                                                            True, SearchFormat:=False)
    lxRow = Cells(SheetInput.Rows.Count, rangexCellItemCode.Column).End(xlUp).Row
    Me.ListBoxBreakdown.Clear
    With ListBoxBreakdown
        .ColumnWidths = "40;50;150"
        .ColumnCount = 3
        For Each cell In SheetInput.Range("H6:H" & lxRow).SpecialCells(xlCellTypeVisible)
            .AddItem CStr(cell.Value)
            .List(.ListCount - 1, 1) = cell.Offset(0, 7).Value 'RPS
            .List(.ListCount - 1, 2) = cell.Offset(0, 8).Value
            .List(.ListCount - 1, 3) = cell.Offset(0, 9).Value 'RPS
        Next cell
    End With
    Run "SortListBox", ListBoxBreakdown, 0, 2, 2
End Sub

Sub fillListBoxBDUniq()
    Dim kodebd As String
    Dim lzRow As Long
    Dim cellz As Range
    lzRow = SheetSumBD.Cells(Rows.Count, 1).End(xlUp).Row
    Me.ListBoxBreakdownUniq.Clear
    With ListBoxBreakdownUniq
        .ColumnWidths = "50;40;150"
        .ColumnCount = 3
        For Each cellz In SheetSumBD.Range("A2:A" & lzRow)
            .AddItem CStr(cellz.Value)
            .List(.ListCount - 1, 1) = cellz.Offset(0, 1).Value 'RPS
            .List(.ListCount - 1, 2) = cellz.Offset(0, 2).Value
            .List(.ListCount - 1, 3) = cellz.Offset(0, 3).Value 'RPS
        Next cellz
    End With
    sortByTimeUniq
End Sub

Sub aveAR()
    Dim Ws As Worksheet
    Dim rng As Range
    Dim visibleTotal As Double

    Set Ws = ActiveSheet
    Set rng = Ws.Range("AA:AA")
    On Error Resume Next
    visibleTotal = Application.WorksheetFunction.Average(rng.SpecialCells(xlCellTypeVisible))
    ' print to the immediate window
    LabelSumAr.Caption = Format(visibleTotal, "00.00%")
End Sub

Sub avePR()
    Dim Ws As Worksheet
    Dim rng As Range
    Dim visibleTotal As Double

    Set Ws = ActiveSheet
    Set rng = Ws.Range("AC:AC")

    visibleTotal = Application.WorksheetFunction.Average(rng.SpecialCells(xlCellTypeVisible))
    ' print to the immediate window
    If visibleTotal > 10 Then
        LabelSumPr.Caption = ">1000%"
    Else
        LabelSumPr.Caption = Format(visibleTotal, "00.00%")
    End If
End Sub

Sub aveQR()
    Dim Ws As Worksheet
    Dim rng As Range
    Dim visibleTotal As Double

    Set Ws = ActiveSheet
    Set rng = Ws.Range("AE:AE")

    visibleTotal = Application.WorksheetFunction.Average(rng.SpecialCells(xlCellTypeVisible))
    ' print to the immediate window
    LabelSumQr.Caption = Format(visibleTotal, "00.00%")
End Sub

Sub aveOee()
    Dim Ws As Worksheet
    Dim rng As Range
    Dim visibleTotal As Double

    Set Ws = ActiveSheet
    Set rng = Ws.Range("AG:AG")

    visibleTotal = Application.WorksheetFunction.Average(rng.SpecialCells(xlCellTypeVisible))
    ' print to the immediate window
    LabelSumOee.Caption = Format(visibleTotal, "00.00%")
End Sub

'Attribute VB_Name = "systemButton"
'created by nandur93 12/07/2017 https://nandur93.com/VBA
'update 18/04/2019
'+fix code
'+add simple tutorial
'update 01/05/2019
'+fix indent
'+fix simple tutorial
'library ini untuk tutup form menggunakan X merah pojok kiri atas
'this library is to ask user some action if RED X button clicked
'copy and paste this code to end of your UserForm module
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    'when X button clicked
    If CloseMode = 0 Then                        'ketika X di klik
        'defined as integer
        Dim xClose As Integer                    'dim variablenya sebagai integer
        'ask user two question before quit between YES or NO
        xClose = MsgBox("Tutup Form?", vbYesNo + vbQuestion, "Keluar") 'tanyakan mau keluar apa tidak
        'if YES button clicked
        If xClose = vbYes Then                   'jika klik IYA
            '//put your logic here
            '...
            'here for example, my logic is Unload Me (close the userform)
            Unload Me                            'keluar dari aplikasi
            'else or if NO button clicked
        Else
            '//put your logic here
            '...
            'here for example, my logic is CANCEL the action (don't close the userform)
            Cancel = True
        End If
    End If
End Sub

