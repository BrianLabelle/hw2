Attribute VB_Name = "Module3"
'2019-06-16: VBA Script for Rice University | Bootcamp | #2. Unit 2 | Assignment - The VBA of Wall Street
'Submitted By Brian Labelle

'HARD
    ' Your solution will include everything from the moderate challenge.
    ' Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

Sub hard()
Attribute hard.VB_ProcData.VB_Invoke_Func = "q\n14"

Dim TickerName As String
Dim OpenDateRange As Range

Dim TickerNameRow As Integer
Dim TickerOpenValue As Integer
Dim TotalStockVolume As Double

Dim firstRow As Long
Dim lastRow As Long

Dim YearlyOpenPrice As Double
Dim OpenPrice As Double
Dim YearlyClosePrice As Double
Dim YearlyChange As Double

TickerNameRow = 2
TickerOpenValue = 3
TotalStockVolume = 0

'EXCEL GENERATED MACRO: Cleans up previously generated content
    
    Columns("I:T").Select
    Selection.Delete Shift:=xlToLeft
    Range("I1").Select
    ActiveWindow.SmallScroll Down:=-9
    
' --------------------------------------------------------------------------------------------------------

firstRow = Cells(Rows.Count, 1).End(xlDown).Row
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
YearlyOpenPrice = 0
YearlyClosePrice = 0
YearlyChange = 0


' Sets header row for all titles needed.
Cells(1, 9).Value = "Ticker"
Cells(1, 9).Font.Bold = True
Cells(1, 10).Value = "Yearly Change"
Cells(1, 10).Font.Bold = True
Cells(1, 11).Value = "Percent Change"
Cells(1, 11).Font.Bold = True
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 12).Font.Bold = True
Cells(2, 14).Value = "Greatest % Increase"
Cells(2, 14).Font.Bold = True
Cells(3, 14).Value = "Greatest % Decrease"
Cells(3, 14).Font.Bold = True
Cells(4, 14).Value = "Greatest Total Volume"
Cells(4, 14).Font.Bold = True
Cells(1, 15).Value = "Ticker"
Cells(1, 15).Font.Bold = True
Cells(1, 16).Value = "Value"
Cells(1, 16).Font.Bold = True

Cells(1, 18).Value = "CLOSE VALUE"
Cells(1, 18).Font.Bold = True
Cells(1, 19).Value = "OPEN VALUE"
Cells(1, 19).Font.Bold = True

' EXCEL GENERATED SORT MACRO: Sorts Column A ( By Ticker ) then by column b ( date ) smallest to largest
' This helps ensures the data is prepared

'Sorting Commented out for multi-year spreasheet, this assumes sheets are sorted but did account incase they were not.

'    Cells.Select
 '   ActiveWorkbook.Worksheets("A").SORT.SortFields.Clear
  '  ActiveWorkbook.Worksheets("A").SORT.SortFields.Add2 Key:=Range("A2:A70926"), _
   '     SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  '  ActiveWorkbook.Worksheets("A").SORT.SortFields.Add2 Key:=Range("B2:B70926"), _
  '      SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  '  With ActiveWorkbook.Worksheets("A").SORT
  '      .SetRange Range("A1:G70926")
  '      .Header = xlYes
  '      .MatchCase = False
  '      .Orientation = xlTopToBottom
  '      .SortMethod = xlPinYin
   '     .Apply
  ''  End With

'Code Cheat
Cells(2, 19).Value = Cells(2, 3).Value

For r = 2 To lastRow
         
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
                     
        YearlyOpenPrice = Cells(r + 1, 3).Value
        YearlyClosePrice = Cells(r, 6).Value
        TickerName = Cells(r, 1).Value
        TotalStockVolume = TotalStockVolume + Cells(r, 7).Value
        
        Range("I" & TickerNameRow).Value = TickerName
        Range("L" & TickerNameRow).Value = TotalStockVolume
        Range("R" & TickerNameRow).Value = YearlyClosePrice
        Range("S" & TickerOpenValue).Value = YearlyOpenPrice
        
        '2019-06-22: not elegant, but it gets it done.
        Range("J" & TickerNameRow).Value = Cells(TickerNameRow, 18) - Cells(TickerOpenValue - 1, 19)
        Range("K" & TickerNameRow).Value = Range("J" & TickerNameRow).Value / Cells(TickerOpenValue - 1, 19)
                
        ' sets conditional formating
        If Range("K" & TickerNameRow).Value < 0 Then Range("J" & TickerNameRow).Interior.ColorIndex = 3
        If Range("K" & TickerNameRow).Value < 0 Then Range("J" & TickerNameRow).Font.ColorIndex = 6
        If Range("K" & TickerNameRow).Value > 0 Then Range("J" & TickerNameRow).Interior.ColorIndex = 4
        If Range("K" & TickerNameRow).Value > 0 Then Range("J" & TickerNameRow).Font.ColorIndex = 1
        
        ' increments row counters
        TickerNameRow = TickerNameRow + 1
        TickerOpenValue = TickerOpenValue + 1
        TotalStockVolume = 0
        YearlyChange = 0

    Else
    
        ' rolls up total stock volume for range.
        TotalStockVolume = TotalStockVolume + Cells(r, 7).Value
        
    End If

Next r

' unclean code, multiple hacks to get this to work.
Dim ws As Worksheet
Dim currentName As String
currentName = ActiveSheet.Name
Set ws = ThisWorkbook.Sheets(currentName)

Dim MyTicMax As String
Dim MyTicMin As String

Dim MyTicVol As String

Dim TickerMax As Range
Dim TickerMin As Range
Dim TickerVol As Range
Dim FindString As String

Dim wb As Workbook
Dim ws2 As Worksheet
Dim FoundCell As Range
Set wb = ActiveWorkbook
Set ws2 = ActiveSheet

MyTicMax = Application.WorksheetFunction.Max(ws.Range("K:K"))
ws.Range("P2") = MyTicMax
FindString = Cells(2, 16).Value
Set FoundCell = ws2.Range("K:K").find(What:=FindString)
Range("O2") = Cells(FoundCell.Row, 9).Value

MyTicMin = Application.WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("P3") = MyTicMin
FindString = Cells(3, 16).Value
Set FoundCell = ws2.Range("K:K").find(What:=FindString)
Range("O3") = Cells(FoundCell.Row, 9).Value

MyTicVol = Application.WorksheetFunction.Max(ws.Range("L:L"))
ws.Range("P4") = MyTicVol
FindString = (Cells(4, 16))
Set FoundCell = ws2.Range("L:L").find(What:=FindString)
Range("O4") = Cells(FoundCell.Row, 9).Value


' EXCEL GENERATED MACRO: AutoFits Column I through T to Autofit for easier legibility
' This helps ensures the data is prepared
    Columns("I:T").Select
    Columns("I:T").EntireColumn.AutoFit


'DELETES TEMP COLUMNS TO CALCULATE OPEN & CLOSE
    
    Columns("R:S").Select
    Selection.ClearContents

    Range("P2:P3").Select
    Selection.NumberFormat = "0.00%"
    Range("P4").Select
    Selection.NumberFormat = "General"
    Range("K:K").Select
    Selection.NumberFormat = "0.00%"

    Range("A1").Select
End Sub


