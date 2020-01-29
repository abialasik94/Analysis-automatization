'Before run a main AllThree() function you have to specify only Source1, which is the path to folder with your xls files.
'By running AllThree() function 3 subfunctions descripted below will be started:
'1) Names() - names of all workbooks in specified directory will be added to each non-empty row, after last column (n+1)
'2) GetSheets() - first sheets from all .xls files in specified directory will be copied to new sheet in excel file
' from which the script was launched
'3) CombineDataFromAllSheets() - move the content of every sheet in excel file from which the script was launched to
' sheet Import created earlier

Sub AllThree()
    Dim Source1 As String
    'do not forget last backslash in source directory.
    Source1 = "C:\Users\abialasik\Desktop\Grupa_Muszkieterow\vba\testowy\"
    Call Names(Source1)
    Call GetSheets(Source1)
    Call CombineDataFromAllSheets

End Sub

'names of all workbooks in specified directory will be added to each non-empty row, after last column (n+1)

Sub Names(Source1)
    Dim Source As String
    Dim StrFile As String
    Dim wb As Workbook
    Dim rng As Range, r As Range
    Dim last_row As Long

    Source = Source1
    StrFile = Dir(Source)
    Do While Len(StrFile) > 0
        Set wb = Workbooks.Open(Source & StrFile)
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        wb.ActiveSheet.Range("AM6", "AM" & last_row).Value = wb.Name
        StrFile = Dir()
        wb.Close (True)
    Loop

End Sub

'first sheets from all .xls files in specified directory will be copied to new sheet in excel file from which the script was launched

Sub GetSheets(Source1)
    Dim LastRow As Long
    'Update ExcelJunction.com
    Path = Source1
    Filename = Dir(Path & "*.xls")
    ThisWorkbook.Sheets.Add.Name = "Import"
    Do While Filename <> ""
        Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
        ActiveWorkbook.Sheets(1).Copy After:=ThisWorkbook.Sheets(1)
        Workbooks(Filename).Close
        Filename = Dir()
    Loop
End Sub

'move the content of every sheet in excel file from which the script was launched to sheet Import created earlier

Public Sub CombineDataFromAllSheets()

    Dim wksSrc As Worksheet, wksDst As Worksheet
    Dim rngSrc As Range, rngDst As Range
    Dim lngLastCol As Long, lngSrcLastRow As Long, lngDstLastRow As Long
    
    'Notes: "Src" is short for "Source", "Dst" is short for "Destination"
    
    'Set references up-front
    Set wksDst = ThisWorkbook.Worksheets("Import")
    lngDstLastRow = LastOccupiedRowNum(wksDst) '<~ defined below (and in Toolbelt)!
    lngLastCol = LastOccupiedColNum(wksDst) '<~ defined below (and in Toolbelt)!
    
    'Set the initial destination range
    Set rngDst = wksDst.Cells(lngDstLastRow + 1, 1)
    
    'Loop through all sheets
    For Each wksSrc In ThisWorkbook.Worksheets
    
        'Make sure we skip the "Import" destination sheet!
        If wksSrc.Name <> "Import" Then
            
            'Identify the last occupied row on this sheet
            lngSrcLastRow = LastOccupiedRowNum(wksSrc)
            
            'Store the source data then copy it to the destination range
            With wksSrc
                Set rngSrc = .Range(.Cells(6, 1), .Cells(lngSrcLastRow, 39))
                rngSrc.Copy Destination:=rngDst
            End With
            
            'Redefine the destination range now that new data has been added
            lngDstLastRow = LastOccupiedRowNum(wksDst)
            Set rngDst = wksDst.Cells(lngDstLastRow + 1, 1)
            
        End If
    
    Next wksSrc

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Sheet, the worksheet we'll search to find the last row
'OUTPUT      : Long, the last occupied row
'SPECIAL CASE: if Sheet is empty, return 1
Public Function LastOccupiedRowNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              Lookat:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Row
        End With
    Else
        lng = 1
    End If
    LastOccupiedRowNum = lng
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Sheet, the worksheet we'll search to find the last column
'OUTPUT      : Long, the last occupied column
'SPECIAL CASE: if Sheet is empty, return 1
Public Function LastOccupiedColNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              Lookat:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Column
        End With
    Else
        lng = 1
    End If
    LastOccupiedColNum = lng
End Function

