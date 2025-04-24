Sub HighlightMaxBarDynamic()
    Dim ws As Worksheet
    Dim cht As ChartObject
    Dim s As Series
    Dim maxVal As Double
    Dim i As Integer, maxIndex As Integer

    ' Set worksheet
    Set ws = ActiveSheet

    ' Loop through all charts on the active sheet
    For Each cht In ws.ChartObjects
        ' Check if the chart has at least one series
        If cht.Chart.SeriesCollection.Count > 0 Then
            Set s = cht.Chart.SeriesCollection(1) 
           
            ' Handle case where no data is visible after filtering
            On Error Resume Next
            maxVal = Application.WorksheetFunction.Max(s.Values)
            On Error GoTo 0
            If maxVal = 0 Then Exit Sub ' Exit if no data available
           
            ' Find the index of the max value
            For i = 1 To s.Points.Count
                If s.Values(i) = maxVal Then
                    maxIndex = i
                    Exit For
                End If
            Next i
           
            ' Highlight the max bar with custom colour (#FFB300)
            For i = 1 To s.Points.Count
                If i = maxIndex Then
                    s.Points(i).Format.Fill.ForeColor.RGB = RGB(255, 179, 0) 'RGB for the Gold Highlight
                End If
            Next i
        End If
    Next cht
End Sub
