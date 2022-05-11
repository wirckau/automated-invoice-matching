Option Explicit
'Application.ScreenUpdating = False

Dim strDirectory   : strDirectory = ":\data"
Dim docName
getExcel

Dim tempSheet   : tempSheet = "temp"
Dim outputSheet   : outputSheet = "output"
Dim FDSheet   : FDSheet = "Sheet1"

Dim dateNow : dateNow =     MyDateFormat(Date)
dIM letzteMonth : letzteMonth = DatePart("m",        DateAdd("m",-1,     FormatDateTime(Date,2)   )     )
dim letzteInYR : letzteInYR = DatePart("yyyy",        DateAdd("m",-1,     FormatDateTime(Date,2)   )     )
dim monthEarlier : monthEarlier = DatePart("m",  FormatDateTime(Date,2)      )
WScript.Echo "First day of date now is: "  & dateNow

Dim objConnection
Const adOpenStatic = 3
Const adLockOptimistic = 3

'creat obj/ an Excel instance:
Dim objExcel   : Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Visible = True
'get obj:
Dim objWorkbook   : Set objWorkbook = objExcel.Workbooks.Open(strDirectory &"\"& docName)
'>>>>>> set temp sheet:
Dim ws   : Set ws = objWorkbook.Sheets.Add(, objWorkbook.Sheets(objWorkbook.Sheets.Count))
ws.Name = tempSheet

announcerDataDB

'calculation>>>>>>>>>
Dim usedRowCnt   : usedRowCnt = ws.UsedRange.Rows.Count
Dim usedColmnCnt   : usedColmnCnt = ws.UsedRange.Columns.Count
Dim inputRange   : Set inputRange = ws.Range(ws.Cells(1, 1), ws.Cells(usedRowCnt, usedColmnCnt))

Dim i, uniqueKey
    Dim dict : Set dict = CreateObject("Scripting.Dictionary")
    Dim arr   : arr       =inputRange.Value

    For i = LBound(arr, 1) To UBound(arr, 1)
        uniqueKey = arr(i, 1) & "," & arr(i, 2)
        dict(uniqueKey) = dict(uniqueKey) + arr(i, 3) 'sum
    Next

    Dim key, tempArr
    Dim rowCounter   : rowCounter = inputRange.Offset(0, 0).Row 
    '>>> add output sheet>>
    Set ws = objWorkbook.Sheets.Add(, objWorkbook.Sheets(objWorkbook.Sheets.Count))
    ws.Name = outputSheet
    objWorkbook.Worksheets(outputSheet).Activate

    With ws
        For Each key In dict.keys
            .Cells(rowCounter, 1) = dict(key)
            tempArr = Split(key, ",")
            .Cells(rowCounter, 2).Resize(1, UBound(tempArr) + 1) = tempArr
            rowCounter = rowCounter + 1
        Next

    '>> rows used in output>>
    Dim rowsUsed : rowsUsed = .UsedRange.Rows.Count
    '>>move price col to end >>>
    Dim Cols, TCols
    Set Cols = .Range("A1","A"&rowsUsed)
    Set TCols = .Range("D1","D"&rowsUsed)
    Cols.Cut
    TCols.Insert
    
    End With

'lookup on FD sheet>>>>>
With objWorkbook.Worksheets(FDSheet)
    .Activate
    .UsedRange.Cells(1).Find("Sum").Next.Select
    objExcel.selection.EntireColumn.Insert 'New col next to "Sum"..
    objExcel.selection.EntireColumn.Insert
    Dim LastRow  : LastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
    objExcel.selection.Range(.Cells(1, 1), .Cells(LastRow, 1)) = "'F = person code col" 'info on lookup val field
    objExcel.selection.Range(.Cells(2, 1), .Cells(LastRow, 1)) = "'=VLOOKUP(LOOKUP(9,99999999999999E+307;SEARCH(output!$B$1:$B$"&rowsUsed&";F2);output!$B$1:$B$"&rowsUsed&");output!$B$1:$C$"&rowsUsed&";2;FALSE)"
    objExcel.selection.Cells(1, 2) = "'If FALSE, then dismatch" 'info on col head
    objExcel.selection.Range(.Cells(2, 2), .Cells(LastRow, 2)) = "=G2=H2"

    'add mapping for FD:
    .UsedRange.Cells(1).Find("Kl.Pk").Next.Select
    objExcel.selection.EntireColumn.Insert
    .Cells(1, 7) = "'Reg.piez.(ligums)"
    Dim objFSO   : Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile   : Set objFile = objFSO.OpenTextFile(strDirectory &"\mappingSubjects.csv", 1)

    Dim arrFields   : arrFields = Split(objFile.ReadAll, vbCrLf)
    objFile.Close

    Dim j
    Dim ij : ij = 1
    Do While ij <= LastRow
        For j = LBound(arrFields) To UBound(arrFields)
            If Split(arrFields(j),",")(1) = .Cells(ij,6) Then
                .Cells(ij, 7) = Split(arrFields(j),",")(2)
            End If
        Next
        ij = ij + 1
    Loop

End With
    

'objWorkbook.SaveAS strDirectory&"\test.xlsx" 'Save as..
Set objExcel = Nothing
Set objWorkbook = Nothing
Set dict = Nothing
Set objFSO = Nothing

'functions>>
Function MyDateFormat(TheDate)
    dim m, d, y
    If Not IsDate(TheDate) Then
        MyDateFormat = TheDate '- if input isn't a date, don't do anything
    Else
        m = Right(100 + Month(TheDate),2) '- pad month with a zero if needed
        d = Right(100 + Day(TheDate),2) '- ditto for the day
        y = Right(Year(TheDate),4)  'trim 4 or 2 if last dig of YR
        'MyDateFormat = y & "/" & m & "/" & d
        MyDateFormat = y & "/" & m & "/01" ' with firstDayOfMonth
    End If
End Function

Sub announcerDataDB
Dim msSQLstr, slctAnn, slctDIT, slctMnaAnn, objRs
Set objConnection = CreateObject("ADODB.Connection")
msSQLstr =  "Provider=MSOLEDBSQL;" &_
            "Server=10.10.100.220\xxx;" &_
            "Database=FAB;" &_
            "Trusted_Connection=yes;"
objConnection.ConnectionString = msSQLstr
objConnection.Open
WScript.Echo "activ Con: " & objConnection.State
'sql:
slctAnn = "Declare @firstDayOfMonth datetime = CONVERT(datetime, '"&dateNow&" 00:00:00', 120)" _
    & "  , @firstDayOfPrevMonth datetime  = CONVERT(datetime, '"&letzteInYR&"/"&letzteMonth&"/01 00:00:00', 120)" _
    & "SELECT * ....;"
slctDIT = "Declare @firstDayOfMonth datetime = CONVERT(datetime, '"&dateNow&" 00:00:00', 120) " _
    & " , @firstDayOfPrevMonth datetime  = CONVERT(datetime, '"&letzteInYR&"/"&letzteMonth&"/01 00:00:00', 120) " _
    & "SELECT count...;"
slctMnaAnn = "Declare @firstDayOfMonth datetime = CONVERT(datetime, '"&dateNow&" 00:00:00', 120) " _
    & " , @firstDayOfPrevMonth datetime  = CONVERT(datetime, '"&letzteInYR&"/"&letzteMonth&"/01 00:00:00', 120)" _
    & " SELECT sub..;"
'recrdset:
Set objRs = CreateObject("ADODB.RecordSet")
objRs.Open slctAnn, objConnection, adOpenStatic, adLockOptimistic

WScript.Echo "cnt is: "  & objRs.RecordCount
Dim rowCounterTemp : rowCounterTemp = 1
'loop:
Do Until objRs.eof
'WScript.Echo objRs.fields("name")
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,1) = objRs.fields("name")
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,2) = objRs.fields("sn")
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,3) = objRs.fields("summa")
rowCounterTemp = rowCounterTemp + 1
objRs.Movenext
Loop

objRs.Close
WScript.Echo "close state ann: " & objConnection.State & " next RS> dit >>>"
objRs.Open slctDIT, objConnection, adOpenStatic, adLockOptimistic

Do Until objRs.EOF
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,1) = objRs.fields("name")
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,2) = objRs.fields("sn")
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,3) = objRs.fields("price")
rowCounterTemp = rowCounterTemp + 1
objRs.Movenext
Loop

objRs.Close
WScript.Echo "close state dit: " & objConnection.State & " next RS > mna >>"
objRs.Open slctMnaAnn, objConnection, adOpenStatic, adLockOptimistic
Do Until objRs.EOF
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,1) = objRs.fields("Username")
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,2) = objRs.fields("sn")
objWorkbook.Worksheets(tempSheet).Cells(rowCounterTemp,3) = objRs.fields("price")
rowCounterTemp = rowCounterTemp + 1
objRs.Movenext
Loop

'clean up
objRs.Close
Set objRs = Nothing
objConnection.Close
WScript.Echo "closed: " & objConnection.State
Set objConnection = Nothing
End Sub

Sub getExcel
 Dim oLstExl : Set oLstExl = Nothing
  Dim oFile
  Dim goFS    : Set goFS    = CreateObject("Scripting.FileSystemObject")
  For Each oFile In goFS.GetFolder(strDirectory).Files
      If "xls" = LCase(goFS.GetExtensionName(oFile.Name)) Then
         If oLstExl Is Nothing Then 
            Set oLstExl = oFile ' the first could be the last
         Else
            If oLstExl.DateLastModified < oFile.DateLastModified Then
               Set oLstExl = oFile
            End If
         End If
      End If
  Next
  If oLstExl Is Nothing Then
     WScript.Echo "no .xls found"
  Else
     ' WScript.Echo "found", oLstExl.Name, oLstExl.DateLastModified 'debug
    docName = oLstExl.Name
  End If

  Set goFS  = Nothing
end Sub