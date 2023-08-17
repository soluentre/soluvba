Function GetColNumber(ws As Worksheet, colName As String) As Long
Dim rng As Range
Set rng = ws.Rows(1).Find(colName, LookAt:=xlWhole)

If Not rng Is Nothing Then
    GetColNumber = rng.Column
Else
    MsgBox "Column " & colName & " is NOT DETECTED in " & ws.Name & _
            " sheet 1st row, please fix this issue and run the project again. " & _
            "Terminating running now ...", vbExclamation
    End
End If
End Function
Function GetColNumberSafe(ws As Worksheet, colName As String) As Long
Dim rng As Range
Set rng = ws.Rows(1).Find(colName, LookAt:=xlWhole)

If Not rng Is Nothing Then
    GetColNumberSafe = rng.Column
Else
    GetColNumberSafe = 0 ' Just return 0 and proceed without stopping
End If
End Function
Function GetRangeByIdAndSeries(ws As Worksheet, colName As String, idValue As Long, Optional seriesValue As Long = -1) As Range
Dim colNumber As Long
Dim idColNumber As Long, seriesColNumber As Long
Dim i As Long, lastRow As Long
Dim startRow As Long, endRow As Long
Dim useSeries As Boolean

' Retrieve column numbers based on column names
colNumber = GetColNumber(ws, colName)
idColNumber = GetColNumber(ws, "id")

' Determine if seriesValue is used
If seriesValue <> -1 Then
    useSeries = True
    seriesColNumber = GetColNumber(ws, "series")
Else
    useSeries = False
End If

' Get the last row with data in the id column (to limit the loop)
lastRow = ws.Cells(ws.Rows.Count, idColNumber).End(xlUp).Row

' Loop through rows to find the matching id and series (if useSeries is True)
For i = 1 To lastRow
    If ws.Cells(i, idColNumber).Value = idValue Then
        If Not useSeries Then
            ' If match found and startRow is not yet initialized, set startRow
            If startRow = 0 Then startRow = i
            ' Update the endRow each time a match is found
            endRow = i
        ElseIf useSeries And ws.Cells(i, seriesColNumber).Value = seriesValue Then
            ' If match found and startRow is not yet initialized, set startRow
            If startRow = 0 Then startRow = i
            ' Update the endRow each time a match is found
            endRow = i
        End If
    ElseIf startRow > 0 Then ' Exit the loop if startRow is set and current row doesn't match
        Exit For
    End If
Next i

' Check if startRow and endRow are initialized, if yes, set the result range
If startRow > 0 And endRow >= startRow Then
    Set GetRangeByIdAndSeries = ws.Cells(startRow, colNumber).Resize(endRow - startRow + 1, 1)
Else
    Set GetRangeByIdAndSeries = Nothing
End If
End Function
Function ChartHasShape(chrt As Chart, shapeName As String) As Boolean
Dim shp As Shape
On Error Resume Next
Set shp = chrt.Shapes(shapeName)
On Error GoTo 0
ChartHasShape = Not shp Is Nothing
End Function
Function GetIndices(wsInfo As Worksheet, prefix As String) As Collection
Dim c As Range
Dim colLast As Long
Dim idx As Long
Dim indices As New Collection

colLast = wsInfo.Cells(1, wsInfo.Columns.Count).End(xlToLeft).Column

' Loop through each cell in the first row
For Each c In wsInfo.Range(wsInfo.Cells(1, 1), wsInfo.Cells(1, colLast))
    If Left(c.Value, Len(prefix)) = prefix Then
        ' Extract number after prefix and add to collection
        idx = CLng(Mid(c.Value, Len(prefix) + 1))
        indices.Add idx
    End If
Next c

Set GetIndices = indices
End Function
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
Dim element As Variant
On Error GoTo ErrorHandler
For Each element In arr
    If element = valToBeFound Then
        IsInArray = True
        Exit Function
    End If
Next element
ExitFunction:
Exit Function
ErrorHandler:
IsInArray = False
Resume ExitFunction
End Function
Function MaxSeriesForID(ws As Worksheet, targetID As Long) As Long
Dim lastRow As Long
Dim i As Long
Dim maxSeries As Long

'Find the last row with data in the id column (assuming id is in column F)
lastRow = ws.Cells(ws.Rows.Count, GetColNumber(ws, "id")).End(xlUp).Row

'Initialize maxSeries to a very small number
maxSeries = -99999

'Loop through the rows and find the max series for the target ID
For i = 1 To lastRow
    If ws.Cells(i, GetColNumber(ws, "id")).Value = targetID Then
        If ws.Cells(i, GetColNumber(ws, "series")).Value > maxSeries Then
            maxSeries = ws.Cells(i, GetColNumber(ws, "series")).Value
        End If
    End If
Next i

MaxSeriesForID = maxSeries
End Function
Sub ReplaceTextInIndices(targetChart As ChartObject, wsInfo As Worksheet, chartType As String, indices As Collection, prefix As String, classtype As String, rowNum As Long)
Dim tOrg As String
Dim tNew As String
Dim originalText As String
Dim newText As String
Dim TextboxName As String

For Each idx In indices
    tOrg = wsInfo.Cells(rowNum, GetColNumber(wsInfo, prefix & "org" & idx)).Value
    tNew = wsInfo.Cells(rowNum, GetColNumber(wsInfo, prefix & "new" & idx)).Value
    TextboxName = chartType & classtype
    
    With targetChart.Chart
        If ChartHasShape(targetChart.Chart, TextboxName) Then
            With .Shapes(TextboxName).OLEFormat.Object
                originalText = .text
                If InStr(originalText, tOrg) > 0 Then
                    newText = Replace(originalText, tOrg, tNew)
                    .text = newText
                End If
            End With
        End If
    End With
Next idx
End Sub
Sub ApplySpecialFormatting(indices As Collection, prefix As String, boxtype As String, chartType As String, wsInfo As Worksheet, sampleChart As ChartObject, targetChart As ChartObject, rowNum As Long)
Dim formatStart As String
Dim formatEnd As String
Dim sStartPos As Integer
Dim sEndPos As Integer
Dim tStartPos As Integer
Dim tEndPos As Integer
Dim origFont As Font
Dim sampleText As String
Dim targetText As String
Dim TextboxName As String

' Assuming the textbox names for the sample and target charts are known
TextboxName = chartType & boxtype ' adjust as needed

For Each idx In indices
    formatStart = wsInfo.Cells(rowNum, GetColNumber(wsInfo, prefix & "fstart" & idx)).Value
    formatEnd = wsInfo.Cells(rowNum, GetColNumber(wsInfo, prefix & "fend" & idx)).Value

    ' Extract formatting details from the sample chart
    With sampleChart.Chart
        If ChartHasShape(sampleChart.Chart, TextboxName) Then
            With .Shapes(TextboxName).OLEFormat.Object
                sampleText = .text

                ' Find starting and ending positions in the sample text
                sStartPos = InStr(1, sampleText, formatStart) + Len(formatStart)
                sEndPos = InStr(sStartPos, sampleText, formatEnd) - 1

                Set origFont = .Characters(sStartPos, sEndPos - sStartPos + 1).Font
            End With
        End If
    End With
    
    ' Apply the captured format to the target chart
    With targetChart.Chart
        If ChartHasShape(targetChart.Chart, TextboxName) Then
            With .Shapes(TextboxName).OLEFormat.Object
                targetText = .text

                ' Find starting and ending positions in the target text
                tStartPos = InStr(1, targetText, formatStart) + Len(formatStart)
                tEndPos = InStr(tStartPos, targetText, formatEnd) - 1

                With .Characters(tStartPos, tEndPos - tStartPos + 1).Font
                    .Name = origFont.Name
                    .Size = origFont.Size
                    .Color = origFont.Color
                    .Bold = origFont.Bold
                    ' ... [any other attributes you want to capture and apply]
                End With
            End With
        End If
    End With
Next idx
End Sub
Sub CreateChartsFromTable()

Dim wsInfo As Worksheet
Dim wsSample As Worksheet
Dim wsData As Worksheet
Dim wsArrangement As Worksheet
Dim lastRow As Long
Dim i As Long
Dim chartType As String
Dim targetChart As ChartObject
Dim targetRange As Range
Dim StartCell As Range
Dim GapColumns As Long, GapRows As Long, NumberPerRow As Long, CurrentRow As Long, CurrentCol As Long
Dim originalText As String
Dim newText As String
Dim noteOrg As String
Dim noteNew As String
Dim noteIndices As Collection
Dim boxTitleIndices As Collection
Dim notefIndices As Collection
Dim boxTitlefIndices As Collection
Dim idx As Variant
Dim sampleChart As ChartObject
Dim origMainTitleFont As Font, origSubTitleFont As Font
Dim splitPoint As Integer
Dim SpecialChartTypes() As Variant

dim1 = Array("bar", "barh", "pie") ' Add your specific chart type names here
dim2 = Array("stackbar")

' Define the worksheets
Set wsInfo = ActiveWorkbook.Sheets("info")
Set wsSample = ActiveWorkbook.Sheets("sample")
Set wsData = ActiveWorkbook.Sheets("data")
Set wsArrangement = ActiveWorkbook.Sheets("arrangement")

' Delete all existing charts from the "data" sheet
For Each chObj In wsData.ChartObjects
    chObj.Delete
Next chObj

' Get arrangement settings
Set StartCell = wsData.Range(wsArrangement.Range("B2").Value)
GapColumns = wsArrangement.Range("B3").Value
GapRows = wsArrangement.Range("B4").Value
NumberPerRow = wsArrangement.Range("B5").Value
TitleSize = wsArrangement.Range("B8").Value
CurrentRow = StartCell.Row
CurrentCol = StartCell.Column

' Get the last row of the "info" sheet
lastRow = wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row

' Loop through each row in the "info" sheet starting from the second row (assuming the first row is a header)
For i = 2 To lastRow

    wsData.Cells(CurrentRow - 1, CurrentCol).Value = wsInfo.Cells(i, GetColNumber(wsInfo, "dashboardtitle"))
    wsData.Cells(CurrentRow - 1, CurrentCol).Font.Bold = True
    wsData.Cells(CurrentRow - 1, CurrentCol).Font.Size = TitleSize
    
    ' Get the chart type from the "type" column and ...
    ' ... copy the chart
    ' ###################################################################
        
        chartType = wsInfo.Cells(i, GetColNumber(wsInfo, "type")).Value
        Application.Wait Now + TimeValue("00:00:01") * 0.5 'Waits for 1 second

        wsSample.ChartObjects(chartType).Copy
        Set sampleChart = wsSample.ChartObjects(chartType) ' Assuming chartType matches the name of the chart
    
    ' Define the target range in the "data" sheet and ...
    ' ... paste the chart
    ' ###################################################################
        
        Set targetRange = wsData.Cells(CurrentRow, CurrentCol)
        Application.Wait Now + TimeValue("00:00:01") * 0.5 'Waits for 1 second

        wsData.Paste targetRange
        Set targetChart = wsData.ChartObjects(wsData.ChartObjects.Count)

        If Not sampleChart.Chart.HasTitle Then
            On Error Resume Next ' To handle error if chart doesn't have a title
            On Error GoTo 0
        End If

        If sampleChart.Chart.HasTitle Then
            ' Find the position of the newline character in the sample chart's title
            splitPoint = InStr(sampleChart.Chart.ChartTitle.text, Chr(13))
            
            ' Capture original formatting
            Set origMainTitleFont = sampleChart.Chart.ChartTitle.Characters(1, splitPoint - 1).Font
            Set origSubTitleFont = sampleChart.Chart.ChartTitle.Characters(splitPoint + 1, Len(sampleChart.Chart.ChartTitle.text) - splitPoint).Font
            
            targetChart.Chart.ChartTitle.text = Trim(wsInfo.Cells(i, GetColNumber(wsInfo, "maintitle")).Value) & _
                                                Chr(10) & Trim(wsInfo.Cells(i, GetColNumber(wsInfo, "subtitle")).Value)
            targetChart.Height = wsInfo.Cells(i, GetColNumber(wsInfo, "height")).Value * 72 ' Assuming height in inches
            targetChart.Width = wsInfo.Cells(i, GetColNumber(wsInfo, "width")).Value * 72 ' Assuming width in inches
            
        
            ' Apply captured formatting to the new chart's title
            With targetChart.Chart.ChartTitle.Characters(1, InStr(1, targetChart.Chart.ChartTitle.text, Chr(10)) - 1).Font
                .Name = origMainTitleFont.Name
                .Size = origMainTitleFont.Size
                .Color = origMainTitleFont.Color
                .Bold = origMainTitleFont.Bold
                ' ... [any other attributes you want to capture and apply]
            End With
            
            With targetChart.Chart.ChartTitle.Characters(InStr(1, targetChart.Chart.ChartTitle.text, Chr(10)) + 1, _
            Len(targetChart.Chart.ChartTitle.text) - InStr(1, targetChart.Chart.ChartTitle.text, Chr(10))).Font
                .Name = origSubTitleFont.Name
                .Size = origSubTitleFont.Size
                .Color = origSubTitleFont.Color
                .Bold = origSubTitleFont.Bold
                ' ... [any other attributes you want to capture and apply]
            End With

        End If

    ' Modify the chart attributes based on the columns from the "info" sheet
    ' ###################################################################
        
        ' Rename the Chart
        targetChart.Name = wsInfo.Cells(i, GetColNumber(wsInfo, "chartname")).Value

        ' Get ID to filter series and category data
        Dim chartId As Long
        chartId = wsInfo.Cells(i, GetColNumber(wsInfo, "id")).Value

        If IsInArray(chartType, dim1) Then
            ' Check if chart has series
            If targetChart.Chart.SeriesCollection.Count > 0 Then
                ' Update existing series
                With targetChart.Chart.SeriesCollection(1)
                    .Values = GetRangeByIdAndSeries(wsData, "value", chartId)
                    .XValues = GetRangeByIdAndSeries(wsData, "catname", chartId)
                    .Name = "=" & GetRangeByIdAndSeries(wsData, "seriesname", chartId).Address(External:=True)
                End With
            Else
                ' If no series exists, create a new one
                With targetChart.Chart.SeriesCollection.NewSeries
                    .Values = GetRangeByIdAndSeries(wsData, "value", chartId)
                    .XValues = GetRangeByIdAndSeries(wsData, "catname", chartId)
                    .Name = "=" & GetRangeByIdAndSeries(wsData, "seriesname", chartId).Address(External:=True)
                End With
            End If
        End If
        
        If IsInArray(chartType, dim2) Then
            SeriesCount = targetChart.Chart.SeriesCollection.Count
            
            Dim m As Long
            For m = 1 To MaxSeriesForID(wsData, chartId)
                If i <= SeriesCount Then
                    targetChart.SeriesCollection(i).Values = GetRangeByIdAndSeries(wsData, "value", chartId, m)
                Else
                    Set CurrentSeries = targetChart.Chart.SeriesCollection.NewSeries
                    CurrentSeries.Values = GetRangeByIdAndSeries(wsData, "value", chartId, m)
                End If
                
                ' Update the X-axis labels and series names
                targetChart.Chart.SeriesCollection(m).XValues = GetRangeByIdAndSeries(wsData, "catname", chartId, m)
                targetChart.Chart.SeriesCollection(m).Name = GetRangeByIdAndSeries(wsData, "seriesname", chartId, m).Resize(1, 1)
                
            Next m
            
            Do While targetChart.Chart.SeriesCollection.Count > MaxSeriesForID(wsData, chartId)
                targetChart.Chart.SeriesCollection(targetChart.Chart.SeriesCollection.Count).Delete
            Loop
        End If
        
    ' Textbox on chart modification section
    ' ###################################################################

        ' Fetching the values for original and new note
        Set noteIndices = GetIndices(wsInfo, "noteorg")
        Set boxTitleIndices = GetIndices(wsInfo, "boxtitleorg")
        Set notefIndices = GetIndices(wsInfo, "notefstart")
        Set boxTitlefIndices = GetIndices(wsInfo, "boxtitlefstart")
        
        ReplaceTextInIndices targetChart, wsInfo, chartType, noteIndices, "note", "text", i
        ReplaceTextInIndices targetChart, wsInfo, chartType, boxTitleIndices, "boxtitle", "title", i
        
        ApplySpecialFormatting notefIndices, "note", "text", chartType, wsInfo, sampleChart, targetChart, i
        
        
    ' Move to the next column or row for the next chart
    ' ###################################################################
    
        If (i - 1) Mod NumberPerRow = 0 And i > 2 Then ' If reached the max number of charts per row
            CurrentRow = CurrentRow + GapRows
            CurrentCol = StartCell.Column ' Reset to initial column
        Else
            CurrentCol = CurrentCol + GapColumns
        End If
        
Next i

End Sub