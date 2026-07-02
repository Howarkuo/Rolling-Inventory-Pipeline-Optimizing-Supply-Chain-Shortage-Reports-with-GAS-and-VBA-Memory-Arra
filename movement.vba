Option Explicit

Sub LogAllMovements()

    Dim mainSheet As Worksheet, logSheet As Worksheet
    Set mainSheet = ActiveSheet

    On Error Resume Next
    Set logSheet = ActiveWorkbook.Worksheets("Movements")
    On Error GoTo 0

    If logSheet Is Nothing Then
        MsgBox "'Movements' tab not found! Please create it first.", vbExclamation
        Exit Sub
    End If

    Const startRow As Long = 5
    Const startCol As Long = 17 ' Q
    Const endCol As Long = 42   ' AP

    Dim lastRow As Long
    lastRow = mainSheet.Cells(mainSheet.Rows.Count, 1).End(xlUp).Row

    If lastRow < startRow Then
        MsgBox "No data rows found.", vbExclamation
        Exit Sub
    End If

    Dim numRows As Long, numCols As Long
    numRows = lastRow - startRow + 1
    numCols = endCol - startCol + 1

    Dim headers As Variant, itemIds As Variant, colFValues As Variant
    Dim stockValues As Variant, timelineValues As Variant

    headers = mainSheet.Range(mainSheet.Cells(4, startCol), mainSheet.Cells(4, endCol)).Value
    itemIds = mainSheet.Range(mainSheet.Cells(startRow, 1), mainSheet.Cells(lastRow, 1)).Value
    colFValues = mainSheet.Range(mainSheet.Cells(startRow, 6), mainSheet.Cells(lastRow, 6)).Value
    stockValues = mainSheet.Range(mainSheet.Cells(startRow, 12), mainSheet.Cells(lastRow, 14)).Value
    timelineValues = mainSheet.Range(mainSheet.Cells(startRow, startCol), mainSheet.Cells(lastRow, endCol)).Value

    Dim output() As Variant
    ReDim output(1 To numRows * numCols, 1 To 3)

    Dim outCount As Long
    outCount = 0

    Dim r As Long, c As Long
    Dim itemId As Variant, dateVal As Variant
    Dim staticStockSum As Double
    Dim currentCellVal As Double
    Dim additionalQty As Double
    Dim adjustedCurrentCellVal As Double
    Dim previousCellVal As Double
    Dim foundPrevious As Boolean
    Dim xValue As Double

    For r = 1 To numRows

        itemId = itemIds(r, 1)

        If Len(Trim(CStr(itemId))) > 0 Then

            staticStockSum = ToNumber(stockValues(r, 1)) + _
                             ToNumber(stockValues(r, 2)) + _
                             ToNumber(stockValues(r, 3))

            previousCellVal = 0
            foundPrevious = False

            For c = 1 To numCols

                dateVal = headers(1, c)

                If Len(Trim(CStr(dateVal))) > 0 Then

                    currentCellVal = ToNumber(timelineValues(r, c))

                    additionalQty = GetQtyFromColumnFForTargetColumn( _
                        colFValues(r, 1), _
                        dateVal _
                    )

                    adjustedCurrentCellVal = currentCellVal + additionalQty

                    If c = 1 Then
                        If adjustedCurrentCellVal = 0 Then
                            xValue = 0
                        Else
                            xValue = staticStockSum + adjustedCurrentCellVal
                            previousCellVal = adjustedCurrentCellVal
                            foundPrevious = True
                        End If
                    Else
                        If adjustedCurrentCellVal = 0 Then
                            xValue = 0
                        Else
                            If foundPrevious Then
                                xValue = adjustedCurrentCellVal - previousCellVal
                            Else
                                xValue = adjustedCurrentCellVal - staticStockSum
                            End If

                            previousCellVal = adjustedCurrentCellVal
                            foundPrevious = True
                        End If
                    End If

                    If xValue <> 0 Then
                        outCount = outCount + 1
                        output(outCount, 1) = dateVal
                        output(outCount, 2) = itemId
                        output(outCount, 3) = xValue
                    End If

                End If

            Next c

        End If

    Next r

    If outCount > 0 Then
        Dim nextLogRow As Long
        nextLogRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1

        logSheet.Range( _
            logSheet.Cells(nextLogRow, 1), _
            logSheet.Cells(nextLogRow + outCount - 1, 3) _
        ).Value = output
    End If

    MsgBox "Batch completed. Logged " & outCount & " movement rows.", vbInformation

End Sub


Sub LogMovement()

    Dim mainSheet As Worksheet, logSheet As Worksheet
    Set mainSheet = ActiveSheet

    On Error Resume Next
    Set logSheet = ActiveWorkbook.Worksheets("Movements")
    On Error GoTo 0

    If logSheet Is Nothing Then
        MsgBox "'Movements' tab not found! Please create it first.", vbExclamation
        Exit Sub
    End If

    Dim rowNum As Long, colNum As Long
    rowNum = ActiveCell.Row
    colNum = ActiveCell.Column

    If colNum < 17 Or rowNum < 5 Then
        MsgBox "Please select a cell inside your timeline grid, Column Q onwards and Row 5 onwards.", vbExclamation
        Exit Sub
    End If

    Dim dateVal As Variant, itemId As Variant
    dateVal = mainSheet.Cells(4, colNum).Value
    itemId = mainSheet.Cells(rowNum, 1).Value

    If Len(Trim(CStr(dateVal))) = 0 Then
        MsgBox "Could not find a date/range header in Row 4.", vbExclamation
        Exit Sub
    End If

    Dim staticStockSum As Double
    staticStockSum = ToNumber(mainSheet.Cells(rowNum, 12).Value) + _
                     ToNumber(mainSheet.Cells(rowNum, 13).Value) + _
                     ToNumber(mainSheet.Cells(rowNum, 14).Value)

    Dim currentCellVal As Double
    currentCellVal = ToNumber(mainSheet.Cells(rowNum, colNum).Value)

    Dim colFValue As Variant
    colFValue = mainSheet.Cells(rowNum, 6).Value

    Dim additionalQty As Double
    additionalQty = GetQtyFromColumnFForTargetColumn(colFValue, dateVal)

    Dim adjustedCurrentCellVal As Double
    adjustedCurrentCellVal = currentCellVal + additionalQty

    Dim xValue As Double
    Dim previousCellVal As Double
    Dim foundPrevious As Boolean

    If colNum = 17 Then

        If adjustedCurrentCellVal = 0 Then
            xValue = 0
        Else
            xValue = staticStockSum + adjustedCurrentCellVal
        End If

    Else

        If adjustedCurrentCellVal = 0 Then
            xValue = 0
        Else

            Dim i As Long
            For i = colNum - 1 To 17 Step -1

                Dim previousRawVal As Double
                Dim previousAdditionalQty As Double
                Dim previousAdjustedVal As Double

                previousRawVal = ToNumber(mainSheet.Cells(rowNum, i).Value)

                previousAdditionalQty = GetQtyFromColumnFForTargetColumn( _
                    colFValue, _
                    mainSheet.Cells(4, i).Value _
                )

                previousAdjustedVal = previousRawVal + previousAdditionalQty

                If previousAdjustedVal <> 0 Then
                    previousCellVal = previousAdjustedVal
                    foundPrevious = True
                    Exit For
                End If

            Next i

            If foundPrevious Then
                xValue = adjustedCurrentCellVal - previousCellVal
            Else
                xValue = adjustedCurrentCellVal - staticStockSum
            End If

        End If

    End If

    Dim nextLogRow As Long
    nextLogRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1

    logSheet.Cells(nextLogRow, 1).Value = dateVal
    logSheet.Cells(nextLogRow, 2).Value = itemId
    logSheet.Cells(nextLogRow, 3).Value = xValue

    MsgBox "Successfully calculated and logged " & xValue & _
           " for Item """ & itemId & """." & vbCrLf & _
           "Column F quantity before subtraction: " & additionalQty, vbInformation

End Sub


Function GetQtyFromColumnFForTargetColumn(ByVal colFValue As Variant, ByVal targetHeader As Variant) As Double

    If Len(Trim(CStr(colFValue))) = 0 Then
        GetQtyFromColumnFForTargetColumn = 0
        Exit Function
    End If

    Dim entries As Variant
    entries = Split(CStr(colFValue), ";")

    Dim total As Double
    total = 0

    Dim i As Long
    For i = LBound(entries) To UBound(entries)

        Dim parts As Variant
        parts = Split(Trim(CStr(entries(i))), "*")

        If UBound(parts) = 1 Then

            Dim entryDate As Date
            Dim qty As Double

            If TryParseFullDate(Trim(CStr(parts(0))), entryDate) Then
                qty = ToNumber(parts(1))

                If qty <> 0 Then
                    If DateBelongsToHeader(entryDate, targetHeader) Then
                        total = total + qty
                    End If
                End If
            End If

        End If

    Next i

    GetQtyFromColumnFForTargetColumn = total

End Function


Function DateBelongsToHeader(ByVal entryDate As Date, ByVal header As Variant) As Boolean

    DateBelongsToHeader = False

    If IsDate(header) Then
        DateBelongsToHeader = DateValue(entryDate) = DateValue(CDate(header))
        Exit Function
    End If

    Dim headerText As String
    headerText = Trim(CStr(header))

    If InStr(headerText, "~") > 0 Then

        Dim rangeParts As Variant
        rangeParts = Split(headerText, "~")

        If UBound(rangeParts) = 1 Then
            Dim startDate As Date, endDate As Date

            If TryParseMonthDay(rangeParts(0), startDate) And _
               TryParseMonthDay(rangeParts(1), endDate) Then

                DateBelongsToHeader = DateValue(entryDate) >= DateValue(startDate) And _
                                      DateValue(entryDate) <= DateValue(endDate)
            End If
        End If

        Exit Function

    End If

    If Left(headerText, 4) = "Over" Then

        Dim cutoffDate As Date
        If TryParseMonthDay(Replace(headerText, "Over", ""), cutoffDate) Then
            DateBelongsToHeader = DateValue(entryDate) > DateValue(cutoffDate)
        End If

        Exit Function

    End If

    Dim fullDate As Date
    If TryParseFullDate(headerText, fullDate) Then
        DateBelongsToHeader = DateValue(entryDate) = DateValue(fullDate)
    End If

End Function


Function TryParseFullDate(ByVal text As String, ByRef resultDate As Date) As Boolean

    On Error GoTo Fail

    text = Trim(text)

    Dim parts As Variant
    parts = Split(text, "/")

    If UBound(parts) <> 2 Then GoTo Fail
    If Len(parts(0)) <> 4 Then GoTo Fail

    resultDate = DateSerial(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
    TryParseFullDate = True
    Exit Function

Fail:
    TryParseFullDate = False

End Function


Function TryParseMonthDay(ByVal text As String, ByRef resultDate As Date) As Boolean

    On Error GoTo Fail

    text = Trim(text)

    Dim parts As Variant
    parts = Split(text, "/")

    If UBound(parts) <> 1 Then GoTo Fail

    resultDate = DateSerial(2026, CLng(parts(0)), CLng(parts(1)))
    TryParseMonthDay = True
    Exit Function

Fail:
    TryParseMonthDay = False

End Function


Function ToNumber(ByVal value As Variant) As Double

    If IsError(value) Then
        ToNumber = 0
    ElseIf IsNumeric(value) Then
        ToNumber = CDbl(value)
    Else
        ToNumber = Val(CStr(value))
    End If

End Function
