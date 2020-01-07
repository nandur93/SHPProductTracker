Attribute VB_Name = "ModuleMain"
Option Explicit
Public actWb As Workbook
Public shtInput As Worksheet
Public shtDaily As Worksheet
Public app As Application
Public wFun As WorksheetFunction

'--Libraries--
Public Declare Function FindWindowA& Lib "user32" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function GetWindowLongA& Lib "user32" (ByVal hWnd&, ByVal nIndex&)
Public Declare Function SetWindowLongA& Lib "user32" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)
 
' Déclaration des constantes
Public Const GWL_STYLE As Long = -16
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_FULLSIZING = &H70000

'Attention, envoyer après changement du caption de l'UF
Public Sub InitMaxMin(mCaption As String, Optional Max As Boolean = True, Optional Min As Boolean = True _
        , Optional Sizing As Boolean = True)
Dim hWnd As Long
    hWnd = FindWindowA(vbNullString, mCaption)
    If Min Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MINIMIZEBOX
    If Max Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MAXIMIZEBOX
    If Sizing Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_FULLSIZING
End Sub
'--Libraries--

Sub sumIfsFormula()
Application.ScreenUpdating = False

    Set app = Application
    Set wFun = app.WorksheetFunction
    Set actWb = ActiveWorkbook
    Set shtInput = actWb.Sheets("Input")
    Set shtDaily = actWb.Sheets("Daily")
    Dim InputH As Range, InputS As Range
    Dim InputO As Range
    Dim PDs As String, BDs As String, SAs As String
    Dim OTs As String, ODs As String, NMSs As String
    Dim OT As String
    Dim inputC As Range, InputJ As Range, inputT As Range, inputW As Range
    Dim dailyC As Range, dbBJ As Range
    Dim totalItem As Integer
    Dim itemCodeInput As String
    Dim pctCompl As Single
    
    'source sheets
    Set InputH = shtInput.Range("H:H")
    Set InputS = shtInput.Range("S:S")
    Set InputO = shtInput.Range("O:O")
    Set inputC = shtInput.Range("C:C")
    Set InputJ = shtInput.Range("J:J")
    Set inputT = shtInput.Range("T:T")
    Set inputW = shtInput.Range("W:W")
    Set dailyC = actWb.Sheets("Daily").Range("C6")
    Set dbBJ = actWb.Sheets("DB").Range("B:J")
    
    'sumifs parameter
    PDs = "PD*"
    BDs = "BD*"
    SAs = "SA*"
    OTs = "OT*"
    ODs = "OD*"
    NMSs = "NMS*"
    OT = "OT"
    itemCodeInput = UserFormMain.TextBoxItemCode.Value
    
    
    'Focus to sheet daily
    shtDaily.Activate
    
    'Find lastRow untuk mendapatkan row terakhir secara dinamis
    Dim lRow As Long
    lRow = Cells(Rows.Count, 3).End(xlUp).Row
        ActiveSheet.Range("$A$5:$AJ$9999").AutoFilter Field:=14
    
    'Loop sampai row terakhir
    Dim sRow As Integer
    For sRow = 6 To lRow - 1
        'breakdown time Q
        Cells(sRow, 17).Value = _
                              wFun.SumIfs(InputH, InputO, BDs, inputC, Cells(sRow, 3)) / 60
        'setup adjustment time S
        Cells(sRow, 19).Value = _
                              wFun.SumIfs(InputH, InputO, SAs, inputC, Cells(sRow, 3)) / 60
        'others downtime U
        Cells(sRow, 21).Value = _
                              wFun.SumIfs(InputH, InputO, ODs, inputC, Cells(sRow, 3)) / 60
        'nms time W
        Cells(sRow, 23).Value = _
                              wFun.SumIfs(InputH, InputO, NMSs, inputC, Cells(sRow, 3)) / 60
        'total downtime loss I
        Cells(sRow, 9).Value = _
                             app.Sum(Cells(sRow, 17).Value, Cells(sRow, 19).Value, _
                                     Cells(sRow, 21).Value, Cells(sRow, 23).Value)
        'operating time H
        Cells(sRow, 8).Value = _
                             wFun.SumIfs(InputH, InputO, OTs, inputC, Cells(sRow, 3)) / 60
        'plan downtime F
        Cells(sRow, 6).Value = _
                             wFun.SumIfs(InputH, InputO, PDs, inputC, Cells(sRow, 3)) / 60
        'loading time E
        Cells(sRow, 5).Value = _
                             app.Sum(Cells(sRow, 8), Cells(sRow, 9))
        'working hour D
        Cells(sRow, 4).Value = _
                             app.Sum(Cells(sRow, 5), Cells(sRow, 6))
    '=============================== Looping speed loss L
    'clear input untuk reset filter
        SheetInput.Range("$A$3:$AJ$1000").AutoFilter Field:=10
        SheetInput.Range("$A$3:$AJ$9999").AutoFilter Field:=3
        'filter dengan kriteria tanggal
        SheetInput.Range("$A$3:$AJ$9999").AutoFilter Field:=3, Criteria1:="=" & sRow - 5 & "-Jan" _
                                                                                    , Operator:=xlAnd
                                                                                    
        SheetInput.Range("$A$3:$AJ$9999").AutoFilter Field:=10, Criteria1:="=" & itemCodeInput _
                                                                                    , Operator:=xlAnd
'        SheetInput.Range("$A$3:$AJ$1000").AutoFilter Field:=10, criteria1:= _
'                                                  "<>libur", Operator:=xlAnd
        CreateUniqueList
        'temukan dulu item code yang related pada tanggal
        On Error GoTo NoBlanks
        totalItem = Worksheets("variable").Range("B2:B11").Cells.SpecialCells(xlCellTypeConstants).Count
NoBlanks:
        Resume Next
        If totalItem <= 1 Then
        '=(SUMIFS(Input!W:W,Input!J:J,Input!J8,Input!C:C,Daily!C6)/IFERROR(VLOOKUP(Input!J8,DB!B:J,9,0),0))
        Cells(sRow, 12).Formula = _
                                "=IFERROR((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0))),0)-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))/60,0)"
        Cells(sRow, 15).Formula = _
                                "=IFERROR((SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)),0)"
        Cells(sRow, 11).Formula = _
                                "=IFERROR(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)/60"
        ElseIf totalItem = 2 Then
        Cells(sRow, 12).Formula = _
                                "=IFERROR(((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)))/60),0)"
        Cells(sRow, 15).Formula = _
                                "=IFERROR((SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)),0)"
        Cells(sRow, 11).Formula = _
                                "=IFERROR((SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0)),0)/60"
        ElseIf totalItem = 3 Then
        Cells(sRow, 12).Formula = _
                                "=IFERROR(((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0)))/60),0)"
        Cells(sRow, 15).Formula = _
                                "=IFERROR((SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0)),0)"
        Cells(sRow, 11).Formula = _
                                "=IFERROR((SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0)),0)/60"
        ElseIf totalItem = 4 Then
        Cells(sRow, 12).Formula = _
                                "=IFERROR(((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B5").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B5").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B5").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B5").Value & """" & ",DB!B:J,9,0),0)))/60),0)"
        Cells(sRow, 15).Formula = _
                                "=IFERROR((SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B5").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B5").Value & """" & ",DB!B:J,9,0),0)),0)"
        Cells(sRow, 11).Formula = _
                                "=IFERROR((SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B5").Value & """" & ",DB!B:J,9,0)),0)/60"
        ElseIf totalItem = 5 Then
        Cells(sRow, 12).Formula = _
                                "=IFERROR(((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B5").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B5").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B5").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B5").Value & """" & ",DB!B:J,9,0),0)))/60) + " & _
                                "((((SUMIFS(Input!H:H,Input!O:O,""OT"",Input!J:J," & """" & Sheets("variable").Range("B6").Value & """" & ",Input!C:C,Daily!C" & sRow & ")*(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B6").Value & """" & ",DB!B:J,9,0),0)))-(SUMIFS(Input!T:T,Input!J:J," & """" & Sheets("variable").Range("B6").Value & """" & ",Input!C:C,Daily!C" & sRow & ")))/(IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B6").Value & """" & ",DB!B:J,9,0),0)))/60),0)"
        Cells(sRow, 15).Formula = _
                                "=IFERROR((SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B2").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B3").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B4").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B5").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B5").Value & """" & ",DB!B:J,9,0),0))+" & _
                                "(SUMIFS(Input!W:W,Input!J:J," & """" & Sheets("variable").Range("B6").Value & """" & ",Input!C:C,Daily!C" & sRow & ")/IFERROR(VLOOKUP(" & """" & Sheets("variable").Range("B6").Value & """" & ",DB!B:J,9,0),0)),0)"
        Cells(sRow, 11).Formula = _
                                "=IFERROR((SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B2").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B3").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B4").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B5").Value & """" & ",DB!B:J,9,0))+" & _
                                "(SUMIFS(Input!S:S,Input!O:O,""OT"",Input!C:C,Daily!C" & sRow & ")/VLOOKUP(" & """" & Sheets("variable").Range("B6").Value & """" & ",DB!B:J,9,0)),0)/60"
        End If
    '=============================== End Looping speed loss L
    
    
    '=IF($E6=0,"",S6/$E6)
        With Cells(sRow, 20)                     'T
            .Formula = "=IF($E" & sRow & "=0,0,S" & sRow & "/$E" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
    '=IF($E6=0,"",U6/$E6)
        With Cells(sRow, 22)                     'V
            .Formula = "=IF($E" & sRow & "=0,0,U" & sRow & "/$E" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
    '=IF($E6=0,"",W6/$E6)
        With Cells(sRow, 24)                     'X
            .Formula = "=IF($E" & sRow & "=0,0,W" & sRow & "/$E" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
    '=IF($E6=0,"",I6/$E6)
        With Cells(sRow, 10)                     'G
            .Formula = "=IF($E" & sRow & "=0,0,I" & sRow & "/$E" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
    '=IF($D6=0,"",F6/$D6)
        With Cells(sRow, 7)                      'G
            .Formula = "=IF($D" & sRow & "=0,0,F" & sRow & "/$D" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
    '=IF($E6=0,"",L6/$E6)
        With Cells(sRow, 13)                     'M
            .Formula = "=IF($E" & sRow & "=0,0,L" & sRow & "/$E" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
    '=K6-L6-O6
        Cells(sRow, 14).Formula = "=K" & sRow & "-L" & sRow & "-O" & sRow 'N"
    '=IFERROR(L6,0)
        Cells(sRow, 25).Formula = "=IFERROR(L" & sRow & ",0)" 'Y
    
        With Cells(sRow, 26)                     'Z
            .Formula = "=IF($E" & sRow & "=0,0,Y" & sRow & "/$E" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
    '=IF($E6=0,"",O6/$E6)
        With Cells(sRow, 16)                     'P
            .Formula = "=IF($E" & sRow & "=0,0,O" & sRow & "/$E" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
    '=IF($E6=0,"",Q6/$E6)
        With Cells(sRow, 18)                     'R
            .Formula = "=IF($E" & sRow & "=0,0,Q" & sRow & "/$E" & sRow & ")"
            .NumberFormat = "0.00%"
        End With
        
'    '=IFERROR(H6/E6,0)
'    With Cells(sRow, 27) 'AA
'        .Formula = "=IFERROR(H" & sRow & "/E" & sRow & ",0)"
'        .NumberFormat = "0.00%"
'    End With
'
'    With Cells(sRow, 28) 'AB
'        .Value = Cells(sRow, 27).Value
'        .NumberFormat = "0.00%"
'    End With
'    With Cells(sRow, 29) 'AC
'        .Value = Cells(sRow, 11).Value / Cells(sRow, 8).Value
'        .NumberFormat = "0.00%"
'    End With
'    With Cells(sRow, 30) 'AD
'        .Value = Cells(sRow, 29).Value
'        .NumberFormat = "0.00%"
'    End With
'    On Error GoTo eh
'    With Cells(sRow, 31) 'AE
'        .Value = Cells(sRow, 14).Value / Cells(sRow, 11).Value
'        .NumberFormat = "0.00%"
'    End With
'eh:                                             Cells(sRow, 31).Value = 0
'    On Error GoTo eh2
'    With Cells(sRow, 32) 'AF
'        .Value = Cells(sRow, 31).Value
'        .NumberFormat = "0.00%"
'    End With
'eh2:                                             Cells(sRow, 32).Value = 0
'        On Error GoTo NoBlankss
'Cells(sRow, 33).Value = Cells(sRow, 27).Value * Cells(sRow, 29).Value * Cells(sRow, 31).Value
'NoBlankss:
'        Resume Next
'    Cells(sRow, 34).Value = Cells(sRow, 28).Value * Cells(sRow, 30).Value * Cells(sRow, 32).Value
'    With Cells(sRow, 35)
'        .NumberFormat = "0.00%"
'        .Value = Cells(sRow, 5).Value / Cells(sRow, 4).Value
'    End With

        pctCompl = sRow
        progress pctCompl
    Next sRow
    'reset semua filter
    With UserFormMain
        .Text.Caption = "Done"
        .LabelDone.Caption = ChrW(&H221A)
    End With
    With SheetInput
        .Range("$A$3:$AJ$9999").AutoFilter Field:=3
        .Range("$A$3:$AJ$1000").AutoFilter Field:=10
    End With
    ActiveSheet.Range("$A$5:$AJ$99999").AutoFilter Field:=14, Operator:= _
        xlFilterNoFill
Application.ScreenUpdating = False
End Sub

Sub progress(pctCompl As Single)
    With UserFormMain
        .Frame1.Width = pctCompl * 2
        .Text.Caption = "Loading... "
        .LabelDone.Caption = (pctCompl - 6) + 1
        .Bar.Width = pctCompl * 2
    End With
    DoEvents
End Sub

Sub Range_End_Method()
    'Finds the last non-blank cell in a single row or column

    Dim lRow As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    MsgBox "Last Row: " & lRow
  
End Sub

Sub CreateUniqueList()
    Dim dict As Object, lastRow As Long, lLRow As Long, Champ, c
    Set dict = CreateObject("Scripting.dictionary")
    With SheetInput
        lastRow = .Cells(.Rows.Count, "J").End(xlUp).Row
        For Each c In .Range("J3:J" & lastRow).SpecialCells(xlCellTypeVisible)
            dict(c.Text) = 0
        Next
    End With
    Champ = dict.keys
  ' Now you have the "variables". To create the new sheet:
    With Sheets("variable")
        lLRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("B2:B" & lLRow).ClearContents
        .Range("B1").Resize(dict.Count).Value = Application.Transpose(dict.keys)
        .Range("B1").Value = "varValue"
    End With
End Sub

Sub CreateUniqueListKodeDowntime()
    Dim dict As Object, lastRow As Long, lLRow As Long, Champ, c
    Set dict = CreateObject("Scripting.dictionary")
    With SheetInput
        lastRow = .Cells(.Rows.Count, "O").End(xlUp).Row
        For Each c In .Range("O4:O" & lastRow).SpecialCells(xlCellTypeVisible)
            dict(c.Text) = 0
        Next
    End With
    Champ = dict.keys
  ' Now you have the "variables". To create the new sheet:
    With SheetSumBD
        lLRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("A2:A" & lLRow).ClearContents
        .Range("A2").Resize(dict.Count).Value = Application.Transpose(dict.keys)
        .Range("A1").Value = "kode bd"
    End With
End Sub

Sub CreateUniqueListDowntime()
    Dim dict As Object, lastRow As Long, lLRow As Long, Champ, c
    Set dict = CreateObject("Scripting.dictionary")
    With SheetInput
        lastRow = .Cells(.Rows.Count, "P").End(xlUp).Row
        For Each c In .Range("P4:P" & lastRow).SpecialCells(xlCellTypeVisible)
            dict(c.Text) = 0
        Next
    End With
    Champ = dict.keys
  ' Now you have the "variables". To create the new sheet:
    With SheetSumBD
        lLRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        .Range("C2:C" & lLRow).ClearContents
        .Range("C2").Resize(dict.Count).Value = Application.Transpose(dict.keys)
        .Range("C1").Value = "breakdown"
    End With
End Sub

Sub CreateUniqueListSumIfsTime()
    Dim ThisWB As Workbook: Set ThisWB = ThisWorkbook
    Dim Ws As Worksheet: Set Ws = ThisWB.Sheets("Input")
    Dim i As Integer

    Dim InputH As Range
    Dim InputO As Range
    Dim InputJ As Range
    Dim InputD As Range
    Dim lRow As Long
    Dim itemCode As String
    

    Set InputH = Ws.Range("H:H")
    Set InputO = Ws.Range("O:O")
    Set InputJ = Ws.Range("J:J")
    Set InputD = Ws.Range("D:D")
    
    itemCode = UserFormMain.TextBoxItemCode.Value
    
    lRow = SheetSumBD.Cells(Rows.Count, 1).End(xlUp).Row
    SheetSumBD.Range("B2:B9999").ClearContents
'=IF(A3="","",SUMIFS(Input!H:H,Input!O:O,sum_bd!A3,Input!J:J,sum_bd!$D$2,Input!D:D,sum_bd!$E$2))
    For i = 2 To lRow
        Dim x As Long
        With UserFormMain.ListBoxCounter
            x = .ListIndex
            SheetSumBD.Cells(i, 2) _
        = Application.WorksheetFunction.SumIfs(InputH, InputO, SheetSumBD.Cells(i, 1).Value, InputJ, itemCode, InputD, .List(x, 1))
        End With
    Next
End Sub

Sub CreateUniqueListDate()
    Dim dict As Object, lastRow As Long, lLRow As Long, Champ, c
    Set dict = CreateObject("Scripting.dictionary")
    With SheetInput
        lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        For Each c In .Range("C3:C" & lastRow).SpecialCells(xlCellTypeVisible)
            dict(c.Text) = 0
        Next
    End With
    Champ = dict.keys
  ' Now you have the "variables". To create the new sheet:
    With Sheets("variable")
        lLRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        .Range("C2:C" & lLRow).ClearContents
        .Range("C1").Resize(dict.Count).Value = Application.Transpose(dict.keys)
        .Range("C1").Value = "dateValue"
    End With
End Sub

Sub SortListBox(oLb As MSForms.ListBox, sCol As Integer, sType As Integer, sDir As Integer)
    Dim vaItems As Variant
    Dim i As Long, j As Long
    Dim c As Integer
    Dim vTemp As Variant
 
 'Put the items in a variant array
    vaItems = oLb.List
 
 'Sort the Array Alphabetically(1)
    If sType = 1 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
 'Sort Ascending (1)
                If sDir = 1 Then
                    If vaItems(i, sCol) > vaItems(j, sCol) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
 'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If vaItems(i, sCol) < vaItems(j, sCol) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If
 
            Next j
        Next i
 'Sort the Array Numerically(2)
 '(Substitute CInt with another conversion type (CLng, CDec, etc.) depending on type of numbers in the column)
    ElseIf sType = 2 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
 'Sort Ascending (1)
                If sDir = 1 Then
                    If CInt(vaItems(i, sCol)) > CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
 'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If CInt(vaItems(i, sCol)) < CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If
 
            Next j
        Next i
    End If
 
 'Set the list to the array
    oLb.List = vaItems
End Sub

