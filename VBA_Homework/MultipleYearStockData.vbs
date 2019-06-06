Sub StockCounter()
	Dim ws As Worksheet
	Dim starting_ws As Worksheet
	Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

	For Each ws In ThisWorkbook.Worksheets
		ws.Activate 'activate current worksheet
    
    
		Dim sht As Worksheet
		Dim NumRows As Long     'stores number of rows in the worksheet
		Dim holdingarr(1 To 5000, 1 To 5) As Variant    'array to hold each row in the Excel sheet
		Dim cnt As Long     'ticker counter
		Dim volumeSum As Double 'sum of volume per ticker
    
		Dim firstrow As Double 'counter to get first row per ticker
    
    
		Set sht = ActiveSheet
    
		NumRows = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row  'Last row on the worksheet
    
		cnt = 1
		volumeSum = 0
    
		firstrow = 0
    
		For i = 2 To NumRows
			If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
				holdingarr(cnt, 5) = volumeSum + Cells(i, 7).Value 'volume
				holdingarr(cnt, 4) = Cells(i, 6).Value 'closing price for last day of the year
				cnt = cnt + 1 'increment to next ticker
				volumeSum = 0 'reset volumeSum
				firstrow = 0
        
			Else
				holdingarr(cnt, 1) = Cells(i, 1).Value ' ticker
				volumeSum = volumeSum + Cells(i, 7).Value ' volume
         
				If firstrow = 0 Then
					holdingarr(cnt, 3) = Cells(i, 3).Value ' opening price for the year per ticker
				End If
         
				firstrow = firstrow + 1
        
        
			End If
            
            
		Next i
    
		'cnt = cnt - 1
    
		For j = 2 To cnt 'This loop populates the cells with the aggregated data
			Cells(j, 10).Value = holdingarr(j - 1, 1)
			Cells(j, 11).Value = holdingarr(j - 1, 4) - holdingarr(j - 1, 3)
			If (holdingarr(j - 1, 3) <> 0) Then
				Cells(j, 12).Value = (holdingarr(j - 1, 4) - holdingarr(j - 1, 3)) / holdingarr(j - 1, 3)
			Else
				Cells(j, 12).Value = 0
			End If
			Cells(j, 13).Value = holdingarr(j - 1, 5)
    
		Next j
	Next

	starting_ws.Activate 'activate the worksheet that was originally active
End Sub
