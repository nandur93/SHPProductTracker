Attribute VB_Name = "Libraries_Modules"
Option Explicit

Sub insertListBox()
'Untuk meng-insert data dari Listbox VBA ke Excel Sheet Cells
    Dim x
    'UserFormMain.ListBoxBreakdownUniq adalah nama dari userform dan lisbox, jika prosedur berada dalam module form, maka gunakan Me.ListBoxBreakdownUniq
    x = UserFormMain.ListBoxBreakdownUniq.List
    With SheetSumBD
        .Columns("A:A").ClearContents
        .Range("A1").Resize(UBound(x) + 1, 1).Value = x
    End With
End Sub

Sub Launcher(control As IRibbonControl)
'Untuk memasukkan .xlsm atau Add-In .xlam ke ribbon excel 2007 dan 2010
    UserFormMain.Show vbModeless
End Sub

Sub CountBlankVar()
    Dim n As Integer
    n = Worksheets("variable").Range("B2:B11").Cells.SpecialCells(xlCellTypeConstants).Count
    MsgBox n
End Sub

Sub GotoWeb()
    Dim IE As Object

    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    On Error GoTo NoCanDo
    'Go to this Web Page!
    IE.navigate "https://nandur93.com/"
    'Unload UserFormMain
    Exit Sub
NoCanDo:
    MsgBox "Cannot open " & IE
End Sub
