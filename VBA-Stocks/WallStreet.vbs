Attribute VB_Name = "Module1"
' The VBA of Wall Street
' Aaron Wollman

Option Explicit

' Constants for the data table
Const colTicker = 1
Const colDate = 2
Const colOpen = 3
Const colHigh = 4
Const colLow = 5
Const colClose = 6
Const colVolume = 7

' Constants for the summary table
Const colSummaryTicker = 9
Const colSummaryYearly = 10
Const colSummaryPercent = 11
Const colSummaryVolume = 12

' Constants for the greatest value table
Const colGreatLabels = 14
Const colGreatTicker = 15
Const colGreatValue = 16
Const rowGreatPercentInc = 2
Const rowGreatPercentDec = 3
Const rowGreatVolume = 4

' Color constants
Const colorRed = 3
Const colorGreen = 4
Const colorClear = 0

' Other constants
Const headerRow = 1
Const minRow = headerRow + 1

' This subroutine is the main subroutine.
' Run this subroutine when evaluating
Sub WallStreet()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Call SingleWallStreet(ws)
    Next ws
End Sub

' This subroutine does work for a single worksheet
' @Param ws - The current worksheet.
'             If not passed, will use current active worksheet.
Sub SingleWallStreet(Optional ws As Worksheet)
    If (ws Is Nothing) Then
        Set ws = Application.ActiveSheet
    End If
    
    Call GenerateSummaryTable(ws)
    Call GenerateGreatestTable(ws)
    
    ws.Columns("A:Z").AutoFit
End Sub

' This subroutine generates a summary table for the worksheet
' @Param ws - the current worksheet
Sub GenerateSummaryTable(ws As Worksheet)
    Dim maxRow As Long
    Dim row As Long
    Dim currentSummaryRow As Integer
    Dim currentTicker As String
    Dim nextTicker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As LongLong
    
    Call AddSummaryLabels(ws)
    
    currentSummaryRow = minRow
    
    ' Get the first open price
    openPrice = ws.Cells(minRow, colOpen)
    
    ' Loop over data table
    maxRow = GetMaxRow(ws, colTicker)
    For row = minRow To maxRow
        currentTicker = ws.Cells(row, colTicker).Value
        nextTicker = ws.Cells(row + 1, colTicker).Value
        
        totalVolume = totalVolume + ws.Cells(row, colVolume)
        If (currentTicker <> nextTicker) Then
            closePrice = ws.Cells(row, colClose).Value
            
            ' Populate Summary Table
            Call AddSummaryRow(ws, currentSummaryRow, currentTicker, openPrice, closePrice, totalVolume)
            
            ' Setup Next Ticker
            currentSummaryRow = currentSummaryRow + 1
            totalVolume = 0
            openPrice = ws.Cells(row + 1, colOpen)
        End If
    Next row
End Sub

' This subroutine adds labels for the summary table
' @Param ws - the current worksheet
Sub AddSummaryLabels(ws As Worksheet)
    ws.Cells(headerRow, colSummaryTicker).Value = "Ticker"
    ws.Cells(headerRow, colSummaryYearly).Value = "Yearly Change"
    ws.Cells(headerRow, colSummaryPercent).Value = "Percent Change"
    ws.Cells(headerRow, colSummaryVolume).Value = "Total Stock Volume"
End Sub

' Gets maximum row number for a spreadsheet
' @Param ws - the current worksheet
' @Param column - the column to find the max row of
' @Returns the max row number
Function GetMaxRow(ws As Worksheet, column As Integer)
    GetMaxRow = ws.Cells(Rows.Count, column).End(xlUp).row
End Function

' Adds a row to the Summary Table
' @Param ws - the current worksheet
' @Param row - the summary row to add the data to
' @Param ticker - the ticker to add
' @Param openPrice - the opening price of the stock
' @Param closePrice - the closing price of the stock
' @Param volume - the total volume of the stock
Sub AddSummaryRow(ws As Worksheet, row As Integer, ticker As String, openPrice As Double, closePrice As Double, volume As LongLong)
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    ws.Cells(row, colSummaryTicker).Value = ticker
            
    yearlyChange = closePrice - openPrice
    ws.Cells(row, colSummaryYearly).Value = yearlyChange
    Call FormatYearlyChange(ws, row, yearlyChange)
    
    If (openPrice <> 0) Then
        percentChange = yearlyChange / openPrice
    Else
        percentChange = 0
    End If
    
    ws.Cells(row, colSummaryPercent).Value = percentChange
    Call FormatPercentChange(ws, row)
    
    ws.Cells(row, colSummaryVolume).Value = volume
End Sub

' Formats a yearly change cell
' @Param ws - the current worksheet
' @Param row - the row of the cell to format
' @Param yearlyChange - the value of the yearly change
Sub FormatYearlyChange(ws As Worksheet, row As Integer, yearlyChange As Double)
    Call ColorYearlyChange(ws, row, yearlyChange)
End Sub

' Colors a yearly change cell depending on the value of the yearly change
' @Param ws - the current worksheet
' @Param row - the row of the cell to format
' @Param yearlyChange - the value of the yearly change
Sub ColorYearlyChange(ws As Worksheet, row As Integer, yearlyChange As Double)
    Dim color As Integer
    If (yearlyChange > 0) Then
        color = colorGreen
    ElseIf (yearlyChange < 0) Then
        color = colorRed
    Else
        color = colorClear
    End If
    ws.Cells(row, colSummaryYearly).Interior.ColorIndex = color
End Sub

' Formats a percent change cell
' @Param ws - the current worksheet
' @Param row - the row of the cell to format
Sub FormatPercentChange(ws As Worksheet, row As Integer)
    ws.Cells(row, colSummaryPercent).NumberFormat = "0.00%"
End Sub

' Creates a greates values table for the current worksheet
' @Param ws - the current worksheet
Sub GenerateGreatestTable(ws As Worksheet)
    Dim row As Integer
    Dim maxRow As Integer
    Dim maxIncRow As Integer
    Dim maxDecRow As Integer
    Dim maxVolumeRow As Integer
    Dim maxIncVal As Double
    Dim maxDecVal As Double
    Dim maxVolumeVal As LongLong
    Dim change As Double
    Dim volume As LongLong
    
    Call AddGreatestLabels(ws)
    ' Loop over summary table
    maxRow = GetMaxRow(ws, colSummaryTicker)
    For row = minRow To maxRow
    
        change = ws.Cells(row, colSummaryPercent)
        ' Is this the greatest increase?
        If (change > maxIncVal) Then
            maxIncRow = row
            maxIncVal = change
        End If
        
        ' Is this the greatest decrease?
        If (change < maxDecVal) Then
            maxDecRow = row
            maxDecVal = change
        End If
        
        ' Is this the greatest volume?
        volume = ws.Cells(row, colSummaryVolume)
        If (volume > maxVolumeVal) Then
            maxVolumeRow = row
            maxVolumeVal = volume
        End If
    Next row
    
    ' Populate Greatest % Increase
    ws.Cells(rowGreatPercentInc, colGreatTicker).Value = ws.Cells(maxIncRow, colSummaryTicker).Value
    ws.Cells(rowGreatPercentInc, colGreatValue).Value = maxIncVal
    ws.Cells(rowGreatPercentInc, colGreatValue).NumberFormat = "0.00%"
    
    ' Populate Greatest % Decrease
    ws.Cells(rowGreatPercentDec, colGreatTicker).Value = ws.Cells(maxDecRow, colSummaryTicker).Value
    ws.Cells(rowGreatPercentDec, colGreatValue).Value = maxDecVal
    ws.Cells(rowGreatPercentDec, colGreatValue).NumberFormat = "0.00%"
    
    ' Populate Greatest Total Volume
    ws.Cells(rowGreatVolume, colGreatTicker).Value = ws.Cells(maxVolumeRow, colSummaryTicker).Value
    ws.Cells(rowGreatVolume, colGreatValue).Value = maxVolumeVal
    
End Sub

' This subroutine adds labels for the greatest value table
' @Param ws - the current worksheet
Sub AddGreatestLabels(ws As Worksheet)
    ' Column Labels
    ws.Cells(headerRow, colGreatTicker).Value = "Ticker"
    ws.Cells(headerRow, colGreatValue).Value = "Value"
    ' Row Labels
    ws.Cells(rowGreatPercentInc, colGreatLabels).Value = "Greatest % Increase"
    ws.Cells(rowGreatPercentDec, colGreatLabels).Value = "Greatest % Decrease"
    ws.Cells(rowGreatVolume, colGreatLabels).Value = "Greatest Total Volume"
End Sub



