Attribute VB_Name = "Module11"
Sub stockmarket():


'Loop through each worksheet
For Each ws In worksheets

    'creating and setting variables
    Dim totalStockVolume As Double
    Dim inforow As Long
    Dim opening As Double
    Dim greatestStockVolume As Double
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim volumeTicker As String
    Dim increaseTicker As String
    Dim decreaseTicker As String
    

    greatestStockVolume = 0
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    totalStockVolume = 0
    inforow = 2
    opening = 0

    'Setting variable to hold ticket name and last row
    'I don't think this is needed but I'm not removing it
    Dim worksheetname As String

    'Also don't think this is needed but not removing it
    'Grabbing worksheet
    worksheetname = ws.Name

    'set column values to hold names
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"

    'set varibales to capture the entire row
    lastrow = ws.Cells(Rows.Count, 2).End(xlUp).Row

    'For loop to iterate through entire row
    For rownum = 2 To lastrow
        
        'Skip row if opening is 0
        If ws.Cells(rownum, 3).Value = 0 Then
            'Do nothing
        Else
        

            'Iteratively adding to total stock volume
            totalStockVolume = totalStockVolume + ws.Cells(rownum, 7).Value
            '
            If ws.Cells(rownum, 2).Value = ws.Cells(2, 2) Then
                opening = ws.Cells(rownum, 3).Value
            End If
    
            'Checks if the ticket has changed. If so, paste saved ticket values into appropriate fields
            If ws.Cells(rownum + 1, 1).Value <> ws.Cells(rownum, 1).Value Then
                
                ' Ticker
                ws.Cells(inforow, 9).Value = ws.Cells(rownum, 1).Value
                ws.Cells(inforow, 12).Value = totalStockVolume
    
                'BONUS - Set value for greatest total stock volume
                If greatestStockVolume < totalStockVolume Then
                    greatestStockVolume = totalStockVolume
                    volumeTicker = ws.Cells(rownum, 1).Value
                End If
                
                ' yearly change
                ws.Cells(inforow, 10).Value = (ws.Cells(rownum, 3).Value - opening)
                
                ' Percent Change
                ws.Cells(inforow, 11).Value = (ws.Cells(rownum, 3).Value - opening) / opening
                ws.Cells(inforow, 11).Value = ws.Cells(inforow, 11).Value * 1
                ws.Cells(inforow, 11).NumberFormat = "0.00%"
    
                ' Bonus Percentage Increase
                If greatestPercentIncrease < ws.Cells(inforow, 11).Value Then
                    greatestPercentIncrease = ws.Cells(inforow, 11).Value
                    increaseTicker = ws.Cells(rownum, 1).Value
                End If
                
                ' Bonus Percentage Decrease
                If greatestPercentDecrease > ws.Cells(inforow, 11).Value Then
                    greatestPercentDecrease = ws.Cells(inforow, 11).Value
                    decreaseTicker = ws.Cells(rownum, 1).Value
                End If
                
                ' Color Formatting
                If ws.Cells(inforow, 10).Value < 0 Then
                    ws.Cells(inforow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(inforow, 10).Interior.ColorIndex = 4
                End If
                totalStockVolume = 0
                inforow = inforow + 1
            End If
        End If
    Next rownum



    'Set Greatest Increase
    ws.Cells(2, 15).Value = increaseTicker
    ws.Cells(2, 16).Value = greatestPercentIncrease
    ws.Cells(2, 16).NumberFormat = "0.00%"

    'Set Greatest Decrease
    ws.Cells(3, 15).Value = decreaseTicker
    ws.Cells(3, 16).Value = greatestPercentDecrease
    ws.Cells(3, 16).NumberFormat = "0.00%"

    'Set Greatest Stock
    ws.Cells(4, 15).Value = volumeTicker
    ws.Cells(4, 16).Value = greatestStockVolume

Next ws


End Sub

