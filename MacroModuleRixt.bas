Attribute VB_Name = "MyVBAModule"

'
' Nico Macro
'
' Easy debug
' MsgBox My Message
'

Sub Nico()
 
 ' Variable declaration
 
 ' Dim gives a global scope to the variable
 Dim excelSourceFileName As String
 excelSourceFileName = "training III_macro file.xlsm"
 sheetSourceName = "sourceSheet"
 cellToFindWhat = "K:"
 
 excelTargetFileName = "training III_macro file.xlsm"
 sheetTargetName = "targetSheet" 'the target sheet name
 rowTargetNumber = 2 'the first row where the data are copied
 
 Windows(excelSourceFileName).Activate
 
 Dim sourceSheet As Worksheet
 ' If want to save a reference to an object, you must use Set
 Set sourceSheet = Sheets(sheetSourceName)
 
 sourceSheet.Select
 
 'Find the cell
 Cells.Find(What:=cellToFindWhat, _
            After:=ActiveCell, _
            LookIn:=xlFormulas, _
            LookAt:= _
            xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False _
          , SearchFormat:=False).Activate
 
 
 'ActiveCell return a range object
 'Range Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.
 ActiveCell.Offset(1, 2).Select
  
 Dim counter As Double
 counter = 0
 While ActiveCell.Value <> vbNullString
 
    counter = counter + 1
    
    ' Source Address backup to come back
    sourceCellAddress = ActiveCell.Address
    
    ' Copy
    ' The range reference is relative
    ActiveCell.range("A1:J1").Select
    Selection.Copy
    
    'Paste
    Windows(excelTargetFileName).Activate
    Sheets(sheetTargetName).Activate
    
    'One over 2 is paste on the next row
    If (counter Mod 2) = 0 Then
        range("K" & rowTargetNumber).Select
        rowTargetNumber = rowTargetNumber + 1
    Else
        range("A" & rowTargetNumber).Select
    End If
    ActiveSheet.Paste
    
    'Go to the source and go one below
    Windows(excelSourceFileName).Activate
    Sheets(sheetSourceName).Activate
    range(sourceCellAddress).Select
    ActiveCell.Offset(1, 0).Select
    
 Wend
 
 End Sub
