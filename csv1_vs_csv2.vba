Sub OneClickForFileCompare()
Call TestDownloadCSVs
Call TestCompareWorksheets
Call TestCompare
End Sub
Sub TestDownloadCSVs()

Dim fPath   As String
Dim fCSV    As String
Dim wbCSV   As Workbook
Dim wbMST   As Workbook
Dim ws      As Worksheet
Dim x As Integer

MsgBox ("CSV File import will start once you clicked ok!")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    '''''''Deleting Old files
    For Each ws In Application.ActiveWorkbook.Worksheets
        If Left(Trim(ws.Name), 4) = "PMI_" Then
            ws.Delete
        End If
    Next

Set wbMST = ThisWorkbook
fPath = ThisWorkbook.Path & "\Files\"     'path to CSV files, include the final \
'Application.ScreenUpdating = False  'speed up macro
Application.DisplayAlerts = False   'no error messages, take default answers

If Dir(fPath & "*.csv") = "" Then
        MsgBox "Could not find the file in local machine, Please execute the shell first"
       Exit Sub
    End If
    
    ' Code will only reach here if file exists
fCSV = Dir(fPath & "*.csv")         'start the CSV file listing

    On Error Resume Next
    Do While Len(fCSV) > 0
        Set wbCSV = Workbooks.Open(fPath & fCSV)                    'open a CSV file
        wbMST.Sheets(ActiveSheet.Name).Delete                       'delete sheet if it exists
        ActiveSheet.Move After:=wbMST.Sheets(wbMST.Sheets.Count)    'move new sheet into Mstr
        Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", TrailingMinusNumbers:=True
    ActiveWindow.SmallScroll Down:=-36
    Columns.AutoFit             'clean up display
        fCSV = Dir              'ready next CSV
    Loop
 
Application.ScreenUpdating = True
Set wbCSV = Nothing

End Sub


Sub TestCompareWorksheets()
    
Dim shtName1 As String
Dim shtName2 As String
    
shtName1 = ActiveWorkbook.Worksheets(2).Name
shtName2 = ActiveWorkbook.Worksheets(3).Name

    ' compare two different worksheets in the active workbook
    CompareWorksheets Worksheets(shtName1), Worksheets(shtName2)
    ' compare two different worksheets in two different workbooks
    'CompareWorksheets ActiveWorkbook.Worksheets("Sheet1"), _
     '   Workbooks("WorkBookName.xls").Worksheets("Sheet2")
End Sub

Sub CompareWorksheets(ws1 As Worksheet, ws2 As Worksheet)

Dim r As Long, c As Integer
Dim lr1 As Long, lr2 As Long, lc1 As Integer, lc2 As Integer
Dim maxR As Long, maxC As Integer, cf1 As String, cf2 As String
Dim rptWB As Workbook, DiffCount As Long
MsgBox ("CSV Files Comparison will start once you clicked ok!")
    Application.ScreenUpdating = False
    Application.StatusBar = "Creating the report..."
    Set rptWB = Workbooks.Add
    Application.DisplayAlerts = False
    While Worksheets.Count > 1
        Worksheets(2).Delete
    Wend
    Application.DisplayAlerts = True
    With ws1.UsedRange
        lr1 = .Rows.Count
        lc1 = .Columns.Count
    End With
    With ws2.UsedRange
        lr2 = .Rows.Count
        lc2 = .Columns.Count
    End With
    maxR = lr1
    maxC = lc1
    If maxR < lr2 Then maxR = lr2
    If maxC < lc2 Then maxC = lc2
    DiffCount = 0
    For c = 1 To maxC
        Application.StatusBar = "Comparing cells " & Format(c / maxC, "0 %") & "..."
        For r = 1 To maxR
            cf1 = ""
            cf2 = ""
            On Error Resume Next
            cf1 = ws1.Cells(r, c).FormulaLocal
            cf2 = ws2.Cells(r, c).FormulaLocal
            On Error GoTo 0
            If cf1 <> cf2 Then
                DiffCount = DiffCount + 1
                Cells(r, c).Formula = "'" & cf1 & " <> " & cf2
            End If
        Next r
    Next c
    Application.StatusBar = "Formatting the report..."
    ws1.Range("A1:ALL1").Copy Destination:=Range("A1")
    With Range(Cells(1, 1), Cells(maxR, maxC))
        .Interior.ColorIndex = 19
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        On Error Resume Next
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        On Error GoTo 0
    End With
    Columns("A:IV").ColumnWidth = 20
    rptWB.Saved = True
    If DiffCount = 0 Then
    Set rptWB = Nothing
    'rptWB.Close False
    End If
    Set rptWB = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True

ActiveSheet.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
ActiveSheet.Name = "A_minus_B"
ActiveSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
ActiveSheet.Name = "A_minus_B-Column"
    MsgBox DiffCount & " cells contain different formulas!", vbInformation, _
        "Compare " & ws1.Name & " with " & ws2.Name
    
End Sub

Sub TestCompare()
    
    Range("A1").Select
    ActiveCell.EntireRow.Insert Shift:=xlShiftDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=COUNTBLANK(R[2]C:R[1048575]C)-1048574"
    
i = 1
j = 1
    Range("A2").Select
    Selection.End(xlToRight).Select
l = ActiveCell.Column

Range(Cells(i, j), Cells(i, l)).Select
Selection.FillRight


    Dim rng As Range

    For Each rng In ActiveSheet.UsedRange

        If rng.HasFormula Then

            rng.Formula = rng.Value

        End If

    Next rng

Range("A3:XFD1048576").EntireRow.Delete
Range("A4").Value = "Mismatch Count"
Range("B4").Value = "Attributes"
Range(Cells(i, j), Cells(2, l)).Copy
Range("A5").PasteSpecial Transpose:=True

Range("A4:B4").Select
Application.CutCopyMode = False
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .PatternTintAndShade = 0
End With
With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
End With

    Cells.EntireColumn.AutoFit

MsgBox ("All done your result is ready!")

End Sub
