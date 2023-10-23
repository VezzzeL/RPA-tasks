Sub ImportExchangeRatesFromNBU()
    Dim IE As Object
    Dim HTMLDoc As Object
    Dim URL As String
    Dim exchangeDate As String
    Dim rowCounter As Long
    Dim cell As Object

    ' Get the URL + check if date is valid
    exchangeDate = Format(Worksheets("Exchange rates").Range("H1").Value, "dd.mm.yyyy")
    
    If IsDate(exchangeDate) And exchangeDate <= Date Then
        URL = "https://bank.gov.ua/ua/markets/exchangerates?date=" & exchangeDate & "&period=daily"
    Else
        MsgBox "Invalid date: " & exchangeDate
        Exit Sub
    End If
    
    
    ' Create IE application
    Set IE = CreateObject("InternetExplorer.Application")

    ' Open IE and navigate to NBU
    IE.Visible = True
    IE.Navigate URL

    ' Wait IE loaded
    Do While IE.Busy Or IE.readyState <> 4
        DoEvents
    Loop

    ' Get the HTML
    Set HTMLDoc = IE.Document

    ' Set the exchange date
    'Dim dateField As Object
    'Set dateField = HTMLDoc.getElementById("date")
    
    ' Wait for load
    Application.Wait Now + TimeValue("00:00:02")

    ' Row counter init
    rowCounter = 2

    ' Find exchange table
    Dim exchangeTable As Object
    Set exchangeTable = HTMLDoc.getElementById("exchangeRates")

    'Get data from exchange table
    If Not exchangeTable Is Nothing Then
        Dim tbody As Object
        Set tbody = exchangeTable.getElementsByTagName("tbody")(0)

        ' Loop through rows
        For Each cell In tbody.getElementsByTagName("tr")
            'Get Currency
            Worksheets("Exchange rates").Cells(rowCounter, 2).Value = cell.getElementsByTagName("td")(1).innerText
            
            'Get rate and convert cell value to text
            Dim exchangeRateCell As Range
            Set exchangeRateCell = Worksheets("Exchange rates").Cells(rowCounter, 3)
            exchangeRateCell.NumberFormat = "@"
            exchangeRateCell.Value = cell.getElementsByTagName("td")(4).innerText
            
            'Get per
            Worksheets("Exchange rates").Cells(rowCounter, 4).Value = cell.getElementsByTagName("td")(2).innerText
            
            'Set exchange date
            Worksheets("Exchange rates").Cells(rowCounter, 5).Value = exchangeDate
            rowCounter = rowCounter + 1
        Next cell
    Else
        MsgBox "Exchange rates table not found."
    End If

    ' Clean up and close
    IE.Quit
    Set IE = Nothing
    Set HTMLDoc = Nothing
    MsgBox "Done."
End Sub