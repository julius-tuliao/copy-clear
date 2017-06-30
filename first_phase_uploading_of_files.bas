Attribute VB_Name = "first_phase_uploading_of_files"
Option Explicit

Sub Upload_Click()
Call updatedata
End Sub
Sub Upload_selectives_Click()
Call updateselectives
End Sub
Sub updatedata()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Call deleteworksheet
Dim lastrow As Long
Dim cell As Long
Dim i As Long
Dim ws As Worksheet, strFile As String
Dim r1 As Range, r2 As Range, r3 As Range
Dim bank As Range, bank2 As Range
Dim agent As Range, agent2 As Range
Dim ob As Range, ob2 As Range
Dim level As Range, level2 As Range
Set ws = Sheets.Add
ws.Name = "Accounts Data"
'ws.Name = WorksheetFunction.Text(Now(), "m-d-yyyy h_mm_ss am/pm")
ws.UsedRange.Clear
strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select text file...")
With ws.QueryTables.Add(Connection:="TEXT;" & strFile, _
Destination:=ws.Range("A1"))
    .TextFileParseType = xlDelimited
    .TextFileCommaDelimiter = True
    .Refresh
End With


With ThisWorkbook.Worksheets("DATA")
    lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row - 1
End With

ThisWorkbook.Worksheets("SELECTIVES1").Rows("2:" & Rows.Count).ClearContents


 Worksheets("Accounts Data").Columns("K:K").Select
    Selection.Replace What:="4", Replacement:="L4 NCR", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 Worksheets("Accounts Data").Columns("K:K").Select
    Selection.Replace What:="1", Replacement:="L1 PL", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 Worksheets("Accounts Data").Columns("K:K").Select
    Selection.Replace What:="2", Replacement:="L2 GEO", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False



    Set r1 = Sheets("Accounts Data").Range("B2:B" & lastrow)
    Set r2 = Sheets("DATA").Range("A2")
    Set r3 = Sheets("DATA").Range("E2")
    r1.Copy r2
    r1.Copy r3
    
    
    Set bank = Sheets("Accounts Data").Range("R2:R" & lastrow)
    Set bank2 = Sheets("DATA").Range("B2")
    bank.Copy bank2
    
    
    
    Set agent = Sheets("Accounts Data").Range("P2:P" & lastrow)
    Set agent2 = Sheets("DATA").Range("C2")
    agent.Copy agent2
    
    Set ob = Sheets("Accounts Data").Range("G2:G" & lastrow)
    Set ob2 = Sheets("DATA").Range("I2")
    ob.Copy ob2
    
     Set level = Sheets("Accounts Data").Range("K2:K" & lastrow)
    Set level2 = Sheets("DATA").Range("H2")
    level.Copy level2
    
    'find and replace
    
    
    
    Worksheets("DATA").Range("D2:D" & lastrow).Formula = "=VLOOKUP(C2,INFO!A:B,2,FALSE)"
    Worksheets("DATA").Range("F2:F" & lastrow).Formula = "=VLOOKUP(C2,INFO!A:C,3,FALSE)"
    Worksheets("DATA").Range("G2:G" & lastrow).Formula = "=VLOOKUP(C2,INFO!A:D,4,FALSE)"
    Worksheets("DATA").Range("J2:J" & lastrow).Formula = "=SUMIFS(SELECTIVES1!C:C,SELECTIVES1!F:F,DATA!E2)"
    Worksheets("DATA").Range("K2:K" & lastrow).Formula = "=I2-J2"
    Worksheets("DATA").Range("L2:L" & lastrow).Formula = "=IFERROR(VLOOKUP(E2,EPA!C:I,7,FALSE),""-"")"
    Worksheets("DATA").Range("M2:M" & lastrow).Formula = "=CONCATENATE(C2,H2)"
    Worksheets("DATA").Range("N2:N" & lastrow).Formula = "=CONCATENATE(H2,D2,G2)"
    Worksheets("DATA").Range("R2:R" & lastrow).Formula = "=VLOOKUP(H2,'SUMMARY '!$I$1:$J$4,2,FALSE)"
    Worksheets("DATA").Range("S2:S" & lastrow).Formula = "=IFERROR(R2*L2,0)"
    
    'convert all to values
    With Worksheets("DATA").UsedRange
    .Value = .Value
End With

   Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub deleteworksheet()
Dim ws As Worksheet
For Each ws In Worksheets
    If ws.Name = "Accounts Data" Then
        Application.DisplayAlerts = False
        Sheets("Accounts Data").Delete
        Application.DisplayAlerts = True
        End
    End If
Next


End Sub

Sub deleteworksheetselectives()
Dim ws As Worksheet
For Each ws In Worksheets
    If ws.Name = "Selectives Data" Then
        Application.DisplayAlerts = False
        Sheets("Selectives Data").Delete
        Application.DisplayAlerts = True
        End
    End If
Next


End Sub
Sub updateselectives()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Call deleteworksheetselectives
Dim lastrow As Long, thisrow As Long, updaterow As Long
Dim cell As Long
Dim i As Long
Dim ws As Worksheet, strFile As String
Dim r1 As Range, r2 As Range, r3 As Range
Dim bank As Range, bank2 As Range
Dim agent As Range, agent2 As Range
Dim ob As Range, ob2 As Range
Dim level As Range, level2 As Range
Dim answer As String
Set ws = Sheets.Add
ws.Name = "Selectives Data"
'ws.Name = WorksheetFunction.Text(Now(), "m-d-yyyy h_mm_ss am/pm")
ws.UsedRange.Clear
strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select text file...")
With ws.QueryTables.Add(Connection:="TEXT;" & strFile, _
Destination:=ws.Range("A1"))
    .TextFileParseType = xlDelimited
    .TextFileCommaDelimiter = True
    .Refresh
End With

answer = MsgBox("Is this data be append on selectives existing data?", vbQuestion + vbYesNo, "???")

With ThisWorkbook.Worksheets("Selectives Data")
    lastrow = .Cells(.Rows.Count, "B").End(xlUp).Row
End With


If answer = vbNo Then


ThisWorkbook.Worksheets("SELECTIVES1").Rows("2:" & Rows.Count).ClearContents


 Set r1 = Sheets("Selectives Data").Range("B2:B" & lastrow)
    Set r2 = Sheets("SELECTIVES1").Range("A2")
    Set r3 = Sheets("SELECTIVES1").Range("F2")
    r1.Copy r2
    r1.Copy r3
    
 Set bank = Sheets("Selectives Data").Range("I2:I" & lastrow)
    Set bank2 = Sheets("SELECTIVES1").Range("B2")
    bank.Copy bank2
    
 Set ob = Sheets("Selectives Data").Range("M2:M" & lastrow)
    Set ob2 = Sheets("SELECTIVES1").Range("C2")
    ob.Copy ob2
    
 Set level = Sheets("Selectives Data").Range("O2:O" & lastrow)
    Set level2 = Sheets("SELECTIVES1").Range("D2")
    level.Copy level2
    
 Set agent = Sheets("Selectives Data").Range("D2:D" & lastrow)
    Set agent2 = Sheets("SELECTIVES1").Range("E2")
    agent.Copy agent2
    
Worksheets("SELECTIVES1").Range("G2:G" & lastrow).Formula = "=IF(C2>0,C2,0)"

 'convert all to values
    With Worksheets("SELECTIVES1").UsedRange
    .Value = .Value
End With


Else


With ThisWorkbook.Worksheets("SELECTIVES1")
    thisrow = .Cells(.Rows.Count, "B").End(xlUp).Row + 1
End With


 Set r1 = Sheets("Selectives Data").Range("B2:B" & lastrow)
    Set r2 = Sheets("SELECTIVES1").Range("A" & thisrow)
    Set r3 = Sheets("SELECTIVES1").Range("F" & thisrow)
    r1.Copy r2
    r1.Copy r3
    
 Set bank = Sheets("Selectives Data").Range("I2:I" & lastrow)
    Set bank2 = Sheets("SELECTIVES1").Range("B" & thisrow)
    bank.Copy bank2
    
 Set ob = Sheets("Selectives Data").Range("M2:M" & lastrow)
    Set ob2 = Sheets("SELECTIVES1").Range("C" & thisrow)
    ob.Copy ob2
    
 Set level = Sheets("Selectives Data").Range("O2:O" & lastrow)
    Set level2 = Sheets("SELECTIVES1").Range("D" & thisrow)
    level.Copy level2
    
 Set agent = Sheets("Selectives Data").Range("D2:D" & lastrow)
    Set agent2 = Sheets("SELECTIVES1").Range("E" & thisrow)
    agent.Copy agent2
    
With ThisWorkbook.Worksheets("SELECTIVES1")
    updaterow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With
    
Worksheets("SELECTIVES1").Range("G" & thisrow & ":G" & updaterow).Formula = "=IF(C2>0,C2,0)"

 'convert all to values
    With Worksheets("SELECTIVES1").UsedRange
    .Value = .Value
End With

End If
   Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
