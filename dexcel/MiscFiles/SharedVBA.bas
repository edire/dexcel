
Sub RefreshWorkbook()

    Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.RefreshAll
	
    Application.CalculateUntilAsyncQueriesDone
    If Not Application.CalculationState = xlDone Then
        DoEvents
    End If
	
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
End Sub


Sub SavePDF(pdf_sheet As String, pdf_range As String, pdf_path As String):

	Dim rng As Range
	Set rng = Sheets(pdf_sheet).Range(pdf_range)
	Application.DisplayAlerts = False
	pdf_range.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdf_path
	Application.DisplayAlerts = True

End Sub


Sub RangetoImage(pic_sheet As String, pic_range As String, pic_path As Variant)

	Dim rng As Range
	Set rng = Sheets(pic_sheet).Range(pic_range)

    Dim i  As Integer
    Dim intCount As Integer
    Dim objPic As Shape
    Dim objChart As Chart
    Dim CurrSht As Worksheet
    Dim shtTemp As Worksheet
    
    Set CurrSht = ActiveSheet
    Set shtTemp = Worksheets.Add
    shtTemp.Activate
    Set objPic = ActiveSheet.Shapes.AddChart
    objPic.Select
    Set objChart = ActiveChart
    objPic.Line.Visible = msoFalse
    objPic.Width = rng.Width
    objPic.Height = rng.Height
    rng.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
    objChart.Paste
    objChart.Export (pic_path)
    
    Application.DisplayAlerts = False
    shtTemp.Delete
    Application.DisplayAlerts = True

End Sub