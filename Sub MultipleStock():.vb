Sub MultipleStock():

    Dim ws as Worksheet


    ' Loop through all sheets
    For Each ws in Worksheets
        ws.Activate

    ' Create a variable to hold file name
    Dim WorksheetName As String

    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Set up tables
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"

    'Declare the variables
    Dim VolumeTotal As Double
    Dim Ticker As String
    Dim YearChange As Long
    Dim YearPercentChange As Double
    Dim YearOpen as Long
    Dim YearClose As Long
    Dim MaximumChange As Double
    Dim MinimumChange As Double
    Dim MaximumVolume As Double
    Dim MaximumTicker As String
    Dim MinimumTicker As String
    Dim MaximumVolumeTicker As String
  

    ' Keep track of the location for each ticker in the summary table
    Dim SummaryTableRow as Long
    SummaryTableRow = 2

    ' Initialize varibles..

    YearChange = 0
    YearOpen = 0
    YearPercentChange = 0

    'Loop through all stock
    For i = 2 To LastRow

        
    
        'Check if we are still within the same stock, if it is not..
        If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the Ticker
            Ticker = ws.Cells(i, 1).Value

            ' Add to the Volume Total
            VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value

            ' Print the Ticker in the Summary Table
            Range("K" & SummaryTableRow).Value = Ticker

            ' Print the Volume Total to the Summary Table
            Range("N" & SummaryTableRow).Value = VolumeTotal

            
            ' End of year closing price
            YearClose = ws.Cells(i, 6).Value

            'Calculate Yearly Change
            YearChange = YearClose - YearOpen

            ' Print Yearly Change to Summary Table
            ws.Cells(SummaryTableRow, 12).Value = YearChange

            ' Conditionally Format Yearly Change

            If ws.Cells(SummaryTableRow, 12).Value > 0 Then
                ws.Cells(SummaryTableRow, 12).Interior.ColorIndex = 4

            Else
                ws.Cells(SummaryTableRow, 12).Interior.ColorIndex = 3

            End if 

            ' Calculate Year Percent Change
            If YearOpen= 0 Then
                YearPercentChange = 0
            Else
                YearPercentChange = (YearChange / YearOpen)
            End If

            ' Print Year Percent Change
            ws.Cells(SummaryTableRow, 13).Value = YearPercentChange

            ' Format Year Percent Change as Percent
            ws.Range("M2:M" & SummaryTableRow).NumberFormat = "#0.00%"

            'Reset Volume Total
            VolumeTotal = 0  

            'Reset Year Open
            YearOpen = 0

            ' Add one to the summary table row
            SummaryTableRow = SummaryTableRow + 1


        Else 

            VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value

            If YearOpen = 0 Then 

                YearOpen = ws.Cells(i, 3).Value  

            End if


       
        End If
     
    Next i


    ' Challenges 

    ' Initialize Variables

    MaximumChange = 0
    MinimumChange = 0
    MaximumVolume = 0
    MaximumTicker = None
    MinimumTicker = None
    MaximumVolumeTicker = None


    For i = 2 to SummaryTableRow

        ' Ticker with greatest percent increase

        If ws.Cells(i, 13) > MaximumChange Then
            MaximumChange = ws.Cells(i, 13).Value
            MaximumTicker = ws.Cells(i, 11).Value
        End If

        ' Ticker with the greatest percent decrease

        If ws.Cells(i, 13).Value < MaximumChange Then
            MinimumChange = ws.Cells(i, 13).Value
            MinimumTicker = ws.Cells(i, 11).Value
        End If

        ' Ticker with greatest stock volume

        If ws.Cells(i, 14).Value > MaximumVolume Then
            MaximumVolume = ws.Cells(i, 14).Value
            MaximumVolumeTicker = ws.Cells(i, 11).Value
        End If

    Next i

    ' Print tickers into table

    ws.Range("Q2").Value = MaximumTicker
    Ws.Range("Q3").Value = MinimumTicker
    ws.Range("Q4").Value = MaximumVolumeTicker

    ' Print values into table
    
    ws.Range("R2").Value = MaximumChange
    ws.Range("R3").Value = MinimumChange
    ws.Range("R4").Value = MaximumVolume

    ' Formate values as percentages

    ws.Range("R2:R" & SummaryTableRow).NumberFormat = "#0.00%"
    ws.Range("R3:R" & SummaryTableRow).NumberFormat = "#0.00%"
    ws.Range("R4:R" & SummaryTableRow).NumberFormat = "#0.00%"

    Next ws

End Sub
