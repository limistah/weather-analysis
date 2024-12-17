' Important for connection to Excel files
Imports System.Globalization
Imports System.Threading
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmRelativeHumidity
    ' Important for transfer of data between Excel and RelativeHumidity Application
    Dim xlTransfer As New Microsoft.Office.Interop.Excel.Application
    Dim objExlApp As Object
    Dim objWrkBk As Object
    Dim objWrkSheet1 As Object
    Dim objWrkSheet2 As Object

    Dim strFileName As String ' String to contain path to Excel file
    Dim intSheet As Integer ' Counter to manipulate various Excel sheets

    Private Sub FreqOfRHInTempRange(ByVal TheTable As DataGridView, ByVal TheProbabilityTable As DataGridView, ByVal TheEndRow As Integer)
        ' Declarations for manipulation of Excel workbook and worksheets
        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US")

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing
        Dim missing As Object = Type.Missing
        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        ' Declarations for counters used to loop through Excel worksheet horizontally and vertically
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim intCounter As Integer
        Dim intProbabilityCount As Integer
        Dim decProbability As Decimal

        Try
            ' Create new instance of Excel application and set necessary properties
            xlApp = New Excel.Application()
            xlApp.DisplayAlerts = False
            xlApp.UserControl = True
            xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)
            ' Change full path of Excel file here:
            ' Open workbook, supply all necessary parameters
            xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

            ' Obtain worksheet and range for temp from user textbox input
            intTempSheet = txtTempWorksheet.Text

            ' Obtain worksheet and range for RH from user textbox input
            intRHSheet = txtRHWorksheet.Text

            ' Get a sheet in the workbook
            xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
            xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

            ' Loop through corresponding Excel sheet rows
            For i = 5 To TheEndRow
                ' Loop through Excel sheet columns
                For j = 3 To 26
                    ' Declare variable for manipulating each Excel cell
                    Dim xlTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)
                    ' Verify temp data
                    If VerifiedData(xlTempCell.Value2) = True Then
                        ' Increment counter for probability calculation (counting number of data points)
                        intProbabilityCount += 1

                        ' Check for range of data
                        k = CheckTempRangeOfData(xlTempCell.Value2)

                        ' Get corresponding RH value
                        Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)
                        l = CheckRHRangeOfData(xlRHCell.Value2)

                        ' Increment counter
                        ' Place count value in appropriate column
                        If IsNothing(TheTable.Rows.Item(l).Cells(k).Value) = False Then
                            ' (Check if cell is presently empty)
                            ' If not empty, increment number in it. (CInt converts contents of cell to integer data type)
                            TheTable.Rows.Item(l).Cells(k).Value = CInt(TheTable.Rows.Item(l).Cells(k).Value) + 1
                        Else
                            intCounter += 1
                            TheTable.Rows.Item(l).Cells(k).Value = intCounter ' Place counter value in cell
                            intCounter = 0 ' Reset counter
                        End If
                    End If
                Next
            Next

            ' Find probability
            For i = 1 To TheTable.RowCount - 1
                For j = 1 To TheTable.ColumnCount - 1
                    If IsNothing(TheTable.Rows.Item(i).Cells(j).Value) = False Then
                        If TheTable.Rows.Item(i).Cells(j).Value <> 0 Then
                            TheTable.Rows.Item(i).Cells(j).Value = TheTable.Rows.Item(i).Cells(j).Value
                        End If

                        decProbability = CDec(TheTable.Rows.Item(i).Cells(j).Value) / intProbabilityCount

                        ' Round up or down
                        TheProbabilityTable.Rows.Item(i).Cells(j).Value = CStr(RoundingProcess(decProbability))
                    Else
                        TheTable.Rows.Item(i).Cells(j).Value = 0
                        TheProbabilityTable.Rows.Item(i).Cells(j).Value = 0
                    End If
                Next
            Next

            ' Close Excel workbook
            DirectCast(xlBook, Excel._Workbook).Close(True, missing, missing)

            ' Enable radio buttons
            rbtProbabilityPerMonth.Enabled = True
            rbtProbabilityPerYear.Enabled = True

        Catch ex As Exception
            MsgBox("Either the selected data path or worksheet is invalid, please input valid data path and worksheet.", MsgBoxStyle.Critical + vbOKOnly, "")
            Me.Cursor = Cursors.Default
            Exit Sub
        Finally
            xlApp.Quit() ' End Excel application
            releaseObject(xlApp) ' Release objects from computer memory
        End Try
    End Sub

    Private Function RoundingProcess(ByVal RoundMe As Decimal) As String
        On Error Resume Next

        Dim strCheck As String
        Dim intDecimalPosition As Integer
        Dim intLastDigit As Integer
        Dim strOutPut As String

        intDecimalPosition = Strings.InStr(RoundMe, ".")

        ' Return six places after decimal point
        strCheck = Strings.Left(CStr(RoundMe), intDecimalPosition + 7)

        intLastDigit = CInt(Strings.Right(strCheck, 1))
        strCheck = Strings.Left(CStr(RoundMe), intDecimalPosition + 6)

        ' Check if last digit is greater than 4
        If intLastDigit > 4 Then
            ' Round up
            strOutPut = CStr(CDec(strCheck + 0.000001))
            RoundingProcess = strOutPut
        Else
            ' Round down
            RoundingProcess = Strings.Left(CStr(RoundMe), intDecimalPosition + 6)
        End If
    End Function

    Private Sub FreqOfTempInTempRange()
        ' Make declarations for manipulating Excel workbook, worksheets
        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US")

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing
        Dim missing As Object = Type.Missing
        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        ' Declare variables for looping through rows and columns
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim intCounter As Integer
        Dim intDataPointCount As Integer

        Try
            xlApp = New Excel.Application()
            xlApp.DisplayAlerts = False
            xlApp.UserControl = True
            xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)
            ' Change full path of Excel file here:
            xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

            ' Worksheet and range for temp from user textbox input
            intTempSheet = txtTempWorksheet.Text

            ' Worksheet and range for RH from user textbox input
            intRHSheet = txtRHWorksheet.Text

            ' Get sheet in the workbook
            xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
            xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

            CreateTempProbabilityHeaders()

            intDataPointCount = 0

            ' Loop through corresponding Excel sheet columns
            For j = 3 To 26
                ' Loop through Excel sheet rows
                For i = 5 To 483
                    Dim xlTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                    ' Verify temp data
                    If VerifiedData(xlTempCell.Value2) = True Then
                        intDataPointCount += 1

                        ' Check if data falls within range of column k
                        k = CheckTempRangeOfData(xlTempCell.Value2)

                        ' Increment counter
                        ' Place count value in appropriate column
                        If IsNothing(dgvTempProbability.Rows.Item(0).Cells(k).Value) = False Then
                            ' (Check if presently empty)
                            ' If not empty, increment number in it
                            dgvTempProbability.Rows.Item(0).Cells(k).Value = CInt(dgvTempProbability.Rows.Item(0).Cells(k).Value) + 1
                        Else
                            intCounter += 1
                            dgvTempProbability.Rows.Item(0).Cells(k).Value = intCounter
                            intCounter = 0 ' Reset counter
                        End If
                    End If
                Next
            Next

            ' Loop through columns' ranges in vb grid (table)
            For k = 1 To dgvTempProbability.ColumnCount - 1
                If IsNothing(dgvTempProbability.Rows.Item(0).Cells(k).Value) = False Then
                    dgvTempProbability.Rows.Item(1).Cells(k).Value = Format(Math.Round(dgvTempProbability.Rows.Item(0).Cells(k).Value / intDataPointCount, 6), "0.######")
                End If
            Next

            DirectCast(xlBook, Excel._Workbook).Close(True, missing, missing)

        Catch ex As Exception
            MsgBox("Either the selected database path or worksheets are invalid, therefore input a valid database path and worksheets.", MsgBoxStyle.Critical + vbOKOnly, "")
            Me.Cursor = Cursors.Default
            Exit Sub
        Finally
            xlApp.Quit()
            releaseObject(xlApp)
        End Try
    End Sub

    Private Sub FreqOfRHInRHRange()
        ' Make declarations for manipulating Excel workbook, worksheets
        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US")

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing
        Dim missing As Object = Type.Missing
        Dim intRHSheet As Integer

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim intCounter As Integer
        Dim intDataPointCount As Integer

        Try
            xlApp = New Excel.Application() ' Create new Excel application
            xlApp.DisplayAlerts = False
            xlApp.UserControl = True
            xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)
            ' Change full path of Excel file here:
            xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

            ' Worksheet and range for RH from user textbox input
            intRHSheet = txtRHWorksheet.Text

            ' Get sheet in the workbook from user textbox input
            xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

            CreateRHProbabilityHeaders()

            intDataPointCount = 0

            ' Loop through corresponding Excel sheet columns
            For j = 3 To 26
                ' Loop through Excel sheet rows
                For i = 5 To 483
                    Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)

                    ' Verify temp data
                    If VerifiedData(xlRHCell.Value2) = True Then
                        intDataPointCount += 1

                        ' Check if data falls within range of column k
                        k = CheckRHRangeOfData(xlRHCell.Value2)

                        ' Increment counter
                        ' Place count value in appropriate column
                        If IsNothing(dgvRHProbability.Rows.Item(0).Cells(k).Value) = False Then
                            ' (Check if presently empty)
                            ' If not empty, increment number in it
                            dgvRHProbability.Rows.Item(0).Cells(k).Value = CInt(dgvRHProbability.Rows.Item(0).Cells(k).Value) + 1
                        Else
                            intCounter += 1
                            dgvRHProbability.Rows.Item(0).Cells(k).Value = intCounter
                            intCounter = 0 ' Reset counter
                        End If
                    End If
                Next
            Next

            ' Loop through columns' ranges in vb grid
            For k = 1 To dgvRHProbability.ColumnCount - 1
                If IsNothing(dgvRHProbability.Rows.Item(0).Cells(k).Value) = False Then
                    dgvRHProbability.Rows.Item(1).Cells(k).Value = Format(Math.Round(dgvRHProbability.Rows.Item(0).Cells(k).Value / intDataPointCount, 6), "0.######")
                End If
            Next

            DirectCast(xlBook, Excel._Workbook).Close(True, missing, missing)

        Catch ex As Exception
            MsgBox("Either the selected database path or worksheets are invalid, therefore input a valid database path and worksheets.", MsgBoxStyle.Critical + vbOKOnly, "")
            Exit Sub
            Me.Cursor = Cursors.Default
        Finally
            xlApp.Quit()
            releaseObject(xlApp)
        End Try
    End Sub

    Private Sub CalculateHDDActualDh()
        On Error Resume Next

        Dim i As Integer
        Dim j As Integer

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing

        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing

        Dim missing As Object = Type.Missing

        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        Dim sngDhSubtraction As Single
        Dim sngDhSum As Single

        xlApp = New Excel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

        ' Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

        ' Worksheet and range for temp
        intTempSheet = txtTempWorksheet.Text

        ' Worksheet and range for RH
        intRHSheet = txtRHWorksheet.Text

        ' Get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

        ' Loop thru Temp worksheet 
        For i = 5 To 483 ' rows
            ' Loop thru excel sheet columns
            For j = 3 To 26 ' last

                ' Get each temp value
                Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                ' Verify that data is a temp value
                If VerifiedData(xlSecondTempCell.Value2) Then

                    ' Calculate Dh total
                    sngDhSubtraction = txtBaseTempActualHDD.Text - xlSecondTempCell.Value2
                    If sngDhSubtraction > 0 Then
                        sngDhSum = sngDhSum + sngDhSubtraction
                    End If
                End If
            Next
        Next

        txtActualTotalHDD.Text = Math.Round(sngDhSum / 24, 1)
        txtActualAvgMonthlyHDD.Text = Math.Round((sngDhSum / 24) / txtNumofYearsHDD.Text, 1)

    End Sub

    Private Sub CalculateCDDActualDc()
        On Error Resume Next

        Dim i As Integer
        Dim j As Integer

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing

        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing

        Dim missing As Object = Type.Missing

        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        Dim sngDcSubtraction As Single
        Dim sngDcSum As Single

        xlApp = New Excel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

        ' Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

        ' Worksheet and range for temp
        intTempSheet = txtTempWorksheet.Text

        ' Worksheet and range for RH
        intRHSheet = txtRHWorksheet.Text

        ' Get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

        ' Loop thru Temp worksheet 
        For i = 5 To 483 ' rows
            ' Loop thru excel sheet columns
            For j = 3 To 26 ' last

                ' Get each temp value
                Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                ' Verify that data is a temp value
                If VerifiedData(xlSecondTempCell.Value2) Then

                    ' Calculate Dc total
                    sngDcSubtraction = xlSecondTempCell.Value2 - CSng(txtBaseTempActualCDD.Text)
                    If sngDcSubtraction > 0 Then
                        sngDcSum = sngDcSum + sngDcSubtraction
                    End If

                End If
            Next
        Next

        txtActualTotalCDD.Text = Math.Round(sngDcSum / 24, 1)
        txtActualAvgMonthlyCDD.Text = Math.Round((sngDcSum / 24) / txtNumofYearsCDD.Text, 1)
    End Sub

    Private Function VerifiedData(ByVal TheData As String) As Boolean
        On Error Resume Next

        ' Ensure not empty
        If IsNothing(TheData) = False Then
            ' Ensure it is numeric
            If IsNumeric(TheData) = True Then
                ' Ensure it's >= 1 
                If TheData >= 1 Then
                    VerifiedData = True
                End If
            End If
        End If

    End Function

    Private Sub AverageRH(ByVal TheTemp As Single, ByVal TheRH As Single)
        On Error Resume Next

        Dim boolPresent As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim intNextRow As Integer

        With dgvCoincidentRH
            If .ColumnCount < 1 Then
                ' Place temp in column header (topmost row)
                .Columns.Add("temp", CStr(TheTemp))
                .Columns(0).Width = 44
            End If

            ' Loop thru columns
            For i = 0 To .ColumnCount - 1
                ' Check if temp is already present in column header (topmost row)
                If .Columns.Item(i).HeaderCell.Value = CStr(TheTemp) Then
                    ' Increment row 
                    .Rows.Add()

                    ' Find next empty row cell under given column
                    For j = 0 To .RowCount - 1
                        If IsNothing(.Rows.Item(j).Cells(i).Value) = True Then
                            intNextRow = j
                            Exit For
                        Else
                            intNextRow = intNextRow + 1
                        End If
                    Next

                    ' Place coincident RH in next row
                    .Rows.Item(intNextRow).Cells(i).Value = TheRH
                    boolPresent = True
                End If

                ' Adjust width of columns
                If .Columns(0).Width <> 44 Then .Columns(i).Width = 44
            Next

            ' Check if temp is new (not already existing in topmost row)
            If boolPresent = False Then
                ' Place temp in column header (topmost row)
                .Columns.Add("temp", CStr(TheTemp))
                .Columns(i).Width = 44

                ' Place coincident RH in first row
                .Rows.Item(0).Cells(i).Value = TheRH
            End If
        End With

    End Sub

    Private Function CalculateHourlyTemp(ByVal TheAvTemp As Single, ByVal TheMaxTemp As Single, ByVal TheMinTemp As Single, ByVal Tasterix As Single) As Single
        On Error Resume Next

        CalculateHourlyTemp = TheAvTemp + (TheMaxTemp - TheMinTemp) * (0.4535 * Math.Cos(Tasterix - 3.7522) + 0.1207 * Math.Cos(2 * Tasterix - 0.3895) + 0.0146 * Math.Cos(3 * Tasterix - 0.8927) + 0.0212 * Math.Cos(4 * Tasterix - 0.2674))

    End Function

    Private Function CalculateFourierHourlyTemp(ByVal TheAvTemp As Single, ByVal TheRange As Single, ByVal Tasterix As Single) As Single
        On Error Resume Next

        CalculateFourierHourlyTemp = TheAvTemp + (TheRange) * (0.4535 * Math.Cos(Tasterix - 3.7522) + 0.1207 * Math.Cos(2 * Tasterix - 0.3895) + 0.0146 * Math.Cos(3 * Tasterix - 0.8927) + 0.0212 * Math.Cos(4 * Tasterix - 0.2674))

    End Function

    Private Function CalculateHourlyRH(ByVal TheAvRH As Single, ByVal TheMaxRH As Single, ByVal TheMinRH As Single, ByVal Tasterix As Single) As Single
        On Error Resume Next

        CalculateHourlyRH = TheAvRH + (TheMaxRH - TheMinRH) * (0.4602 * Math.Cos(Tasterix - 0.6038) + 0.1255 * Math.Cos(2 * Tasterix - 3.5427) + 0.0212 * Math.Cos(3 * Tasterix - 4.2635) + 0.0255 * Math.Cos(4 * Tasterix - 0.3833))

    End Function

    Private Function CalculateFourierHourlyRH(ByVal TheAvRH As Single, ByVal TheRange As Single, ByVal Tasterix As Single) As Single
        On Error Resume Next

        CalculateFourierHourlyRH = TheAvRH + (TheRange) * (0.4606 * Math.Cos(Tasterix - 0.6038) + 0.1255 * Math.Cos(2 * Tasterix - 3.5247) + 0.0212 * Math.Cos(3 * Tasterix - 4.2635) + 0.0255 * Math.Cos(4 * Tasterix - 0.3833))

    End Function

    Private Sub ComputeFourier()
        On Error Resume Next

        Dim sngCalcTempAv As Single
        Dim sngCalcRHAv As Single
        Dim sngMaxAv As Single
        Dim sngMinAv As Single
        Dim sngTav As Single
        Dim sngSTD As Single
        Dim i As Integer
        Dim j As Integer
        Dim intCountTemp As Integer
        Dim intCountRH As Integer
        Dim intValidHours As Integer
        Dim intAvOfAvCount As Integer

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing

        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing

        Dim missing As Object = Type.Missing

        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        Dim sngTasterix As Single

        xlApp = New Excel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

        ' Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

        ' Worksheet for temp and RH from user textbox input
        intTempSheet = txtTempWorksheet.Text
        intRHSheet = txtRHWorksheet.Text

        ' Get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

        CreateFourierHeaders()

        For j = 3 To 26 ' loop thru excel sheet columns

            ' Reset counters
            sngCalcTempAv = 0
            intCountTemp = 0
            sngCalcRHAv = 0
            intCountRH = 0

            ' Loop thru excel sheet rows
            For i = 5 To 483 ' last

                ' Get each temp value
                Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                ' Verify that data is a temperature value
                If VerifiedData(xlSecondTempCell.Value2) Then

                    ' Count columns with valid data
                    intValidHours = j

                    ' Get corresponding RH
                    Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)

                    sngCalcTempAv = sngCalcTempAv + xlSecondTempCell.Value2
                    intCountTemp = intCountTemp + 1

                    sngCalcRHAv = sngCalcRHAv + xlRHCell.Value2
                    intCountRH = intCountRH + 1
                End If

            Next

            If sngCalcTempAv <> 0 Then
                dgvFourierTh.Rows.Item(0).Cells(j - 2).Value = Math.Round(sngCalcTempAv / intCountTemp, 1)
                dgvFourierRh.Rows.Item(0).Cells(j - 2).Value = Math.Round(sngCalcRHAv / intCountRH, 1)
            Else
                dgvFourierTh.Rows.Item(0).Cells(j - 2).Value = ""
                dgvFourierRh.Rows.Item(0).Cells(j - 2).Value = ""
            End If
        Next

        sngCalcTempAv = 0
        sngMinAv = 10000
        sngMaxAv = 0

        intAvOfAvCount = 0

        For j = 1 To 24
            ' Find max average (ensure cell is not empty)
            If dgvFourierTh.Rows.Item(0).Cells(j).Value <> "" Then
                If dgvFourierTh.Rows.Item(0).Cells(j).Value > sngMaxAv Then
                    sngMaxAv = dgvFourierTh.Rows.Item(0).Cells(j).Value
                End If

                ' Find min average
                If dgvFourierTh.Rows.Item(0).Cells(j).Value < sngMinAv Then
                    sngMinAv = dgvFourierTh.Rows.Item(0).Cells(j).Value
                End If

                intAvOfAvCount = intAvOfAvCount + 1

                ' Sum the averages
                sngCalcTempAv = sngCalcTempAv + dgvFourierTh.Rows.Item(0).Cells(j).Value
            End If
        Next

        ' Get average of averages
        sngCalcTempAv = Math.Round(sngCalcTempAv / intAvOfAvCount, 1)

        For j = 1 To 24 ' columns

            ' Ensure cells are not empty
            If IsNothing(dgvFourierTh.Rows.Item(0).Cells(j).Value) = False Then

                ' Calculate fourier hourly temp
                sngTasterix = (2 * Math.PI * (j - 1)) / 24
                dgvFourierTh.Rows.Item(1).Cells(j).Value = Math.Round(CalculateFourierHourlyTemp(CSng(txtNewFourierTav.Text), CSng(txtNewFourierTrange.Text), sngTasterix), 1)

            End If
        Next

        intValidHours = 0
        sngSTD = 0
        For j = 3 To 26 ' loop thru excel sheet columns
            For i = 5 To 483 ' last

                ' Get each temp value
                Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                ' Verify that data is a temperature value
                If VerifiedData(xlSecondTempCell.Value2) Then

                    ' Count columns with valid data
                    intValidHours = intValidHours + 1

                    sngSTD = sngSTD + Math.Pow(xlSecondTempCell.Value2 - CSng(txtNewFourierTav.Text), 2)
                End If

            Next
        Next

        '***********************************************RH

        sngCalcRHAv = 0
        sngMinAv = 10000
        sngMaxAv = 0

        intAvOfAvCount = 0

        For j = 1 To 24
            If dgvFourierRh.Rows.Item(0).Cells(j).Value <> "" Then
                ' Find max average
                If dgvFourierRh.Rows.Item(0).Cells(j).Value > sngMaxAv Then
                    sngMaxAv = dgvFourierRh.Rows.Item(0).Cells(j).Value
                End If

                ' Find min average
                If dgvFourierRh.Rows.Item(0).Cells(j).Value < sngMinAv Then
                    sngMinAv = dgvFourierRh.Rows.Item(0).Cells(j).Value
                End If

                intAvOfAvCount = intAvOfAvCount + 1

                ' Sum averages
                sngCalcRHAv = sngCalcRHAv + dgvFourierRh.Rows.Item(0).Cells(j).Value
            End If
        Next

        sngCalcRHAv = Math.Round(sngCalcRHAv / intAvOfAvCount, 2)

        For j = 1 To 24 ' columns

            ' Ensure cell is not empty
            If IsNothing(dgvFourierRh.Rows.Item(0).Cells(j).Value) = False Then

                ' Calculate fourier Rh 
                sngTasterix = (2 * Math.PI * (j - 1)) / 24
                dgvFourierRh.Rows.Item(1).Cells(j).Value = Math.Round(CalculateFourierHourlyRH(CSng(txtNewFourierRHav.Text), CSng(txtNewFourierRHrange.Text), sngTasterix), 1)
            End If
        Next

    End Sub

    Private Sub ComputeAverages()
        On Error Resume Next

        Dim sngCalcTempAv As Single
        Dim sngCalcRHAv As Single
        Dim sngMaxAv As Single
        Dim sngMinAv As Single
        Dim sngTav As Single
        Dim sngSTD As Single
        Dim i As Integer
        Dim j As Integer
        Dim intCountTemp As Integer
        Dim intCountRH As Integer
        Dim intValidHours As Integer
        Dim intAvOfAvCount As Integer

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing

        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing

        Dim missing As Object = Type.Missing

        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        Dim sngTasterix As Single

        xlApp = New Excel.Application()
    Private Sub CalculateHDDActualDh()
        On Error Resume Next

        Dim i As Integer
        Dim j As Integer

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing

        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing

        Dim missing As Object = Type.Missing

        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        Dim sngDhSubtraction As Single
        Dim sngDhSum As Single

        xlApp = New Excel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

        ' Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

        ' Worksheet and range for temp
        intTempSheet = txtTempWorksheet.Text

        ' Worksheet and range for RH
        intRHSheet = txtRHWorksheet.Text

        ' Get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

        ' Loop thru Temp worksheet 
        For i = 5 To 483 ' rows
            ' Loop thru excel sheet columns
            For j = 3 To 26 ' last

                ' Get each temp value
                Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                ' Verify that data is a temp value
                If VerifiedData(xlSecondTempCell.Value2) Then

                    ' Calculate Dh total
                    sngDhSubtraction = txtBaseTempActualHDD.Text - xlSecondTempCell.Value2
                    If sngDhSubtraction > 0 Then
                        sngDhSum = sngDhSum + sngDhSubtraction
                    End If
                End If
            Next
        Next

        txtActualTotalHDD.Text = Math.Round(sngDhSum / 24, 1)
        txtActualAvgMonthlyHDD.Text = Math.Round((sngDhSum / 24) / txtNumofYearsHDD.Text, 1)

    End Sub

    Private Sub CalculateCDDActualDc()
        On Error Resume Next

        Dim i As Integer
        Dim j As Integer

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing

        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing

        Dim missing As Object = Type.Missing

        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        Dim sngDcSubtraction As Single
        Dim sngDcSum As Single

        xlApp = New Excel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

        ' Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

        ' Worksheet and range for temp
        intTempSheet = txtTempWorksheet.Text

        ' Worksheet and range for RH
        intRHSheet = txtRHWorksheet.Text

        ' Get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

        ' Loop thru Temp worksheet 
        For i = 5 To 483 ' rows
            ' Loop thru excel sheet columns
            For j = 3 To 26 ' last

                ' Get each temp value
                Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                ' Verify that data is a temp value
                If VerifiedData(xlSecondTempCell.Value2) Then

                    ' Calculate Dc total
                    sngDcSubtraction = xlSecondTempCell.Value2 - CSng(txtBaseTempActualCDD.Text)
                    If sngDcSubtraction > 0 Then
                        sngDcSum = sngDcSum + sngDcSubtraction
                    End If

                End If
            Next
        Next

        txtActualTotalCDD.Text = Math.Round(sngDcSum / 24, 1)
        txtActualAvgMonthlyCDD.Text = Math.Round((sngDcSum / 24) / txtNumofYearsCDD.Text, 1)
    End Sub

    Private Function VerifiedData(ByVal TheData As String) As Boolean
        On Error Resume Next

        ' Ensure not empty
        If IsNothing(TheData) = False Then
            ' Ensure it is numeric
            If IsNumeric(TheData) = True Then
                ' Ensure it's >= 1 
                If TheData >= 1 Then
                    VerifiedData = True
                End If
            End If
        End If

    End Function

    Private Sub AverageRH(ByVal TheTemp As Single, ByVal TheRH As Single)
        On Error Resume Next

        Dim boolPresent As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim intNextRow As Integer

        With dgvCoincidentRH
            If .ColumnCount < 1 Then
                ' Place temp in column header (topmost row)
                .Columns.Add("temp", CStr(TheTemp))
                .Columns(0).Width = 44
            End If

            ' Loop thru columns
            For i = 0 To .ColumnCount - 1
                ' Check if temp is already present in column header (topmost row)
                If .Columns.Item(i).HeaderCell.Value = CStr(TheTemp) Then
                    ' Increment row 
                    .Rows.Add()

                    ' Find next empty row cell under given column
                    For j = 0 To .RowCount - 1
                        If IsNothing(.Rows.Item(j).Cells(i).Value) = True Then
                            intNextRow = j
                            Exit For
                        Else
                            intNextRow = intNextRow + 1
                        End If
                    Next

                    ' Place coincident RH in next row
                    .Rows.Item(intNextRow).Cells(i).Value = TheRH
                    boolPresent = True
                End If

                ' Adjust width of columns
                If .Columns(0).Width <> 44 Then .Columns(i).Width = 44
            Next

            ' Check if temp is new (not already existing in topmost row)
            If boolPresent = False Then
                ' Place temp in column header (topmost row)
                .Columns.Add("temp", CStr(TheTemp))
                .Columns(i).Width = 44

                ' Place coincident RH in first row
                .Rows.Item(0).Cells(i).Value = TheRH
            End If
        End With

    End Sub

    Private Function CalculateHourlyTemp(ByVal TheAvTemp As Single, ByVal TheMaxTemp As Single, ByVal TheMinTemp As Single, ByVal Tasterix As Single) As Single
        On Error Resume Next

        CalculateHourlyTemp = TheAvTemp + (TheMaxTemp - TheMinTemp) * (0.4535 * Math.Cos(Tasterix - 3.7522) + 0.1207 * Math.Cos(2 * Tasterix - 0.3895) + 0.0146 * Math.Cos(3 * Tasterix - 0.8927) + 0.0212 * Math.Cos(4 * Tasterix - 0.2674))

    End Function

    Private Function CalculateFourierHourlyTemp(ByVal TheAvTemp As Single, ByVal TheRange As Single, ByVal Tasterix As Single) As Single
        On Error Resume Next

        CalculateFourierHourlyTemp = TheAvTemp + (TheRange) * (0.4535 * Math.Cos(Tasterix - 3.7522) + 0.1207 * Math.Cos(2 * Tasterix - 0.3895) + 0.0146 * Math.Cos(3 * Tasterix - 0.8927) + 0.0212 * Math.Cos(4 * Tasterix - 0.2674))

    End Function

    Private Function CalculateHourlyRH(ByVal TheAvRH As Single, ByVal TheMaxRH As Single, ByVal TheMinRH As Single, ByVal Tasterix As Single) As Single
        On Error Resume Next

        CalculateHourlyRH = TheAvRH + (TheMaxRH - TheMinRH) * (0.4602 * Math.Cos(Tasterix - 0.6038) + 0.1255 * Math.Cos(2 * Tasterix - 3.5427) + 0.0212 * Math.Cos(3 * Tasterix - 4.2635) + 0.0255 * Math.Cos(4 * Tasterix - 0.3833))

    End Function

    Private Function CalculateFourierHourlyRH(ByVal TheAvRH As Single, ByVal TheRange As Single, ByVal Tasterix As Single) As Single
        On Error Resume Next

        CalculateFourierHourlyRH = TheAvRH + (TheRange) * (0.4606 * Math.Cos(Tasterix - 0.6038) + 0.1255 * Math.Cos(2 * Tasterix - 3.5247) + 0.0212 * Math.Cos(3 * Tasterix - 4.2635) + 0.0255 * Math.Cos(4 * Tasterix - 0.3833))

    End Function

    Private Sub ComputeFourier()
        On Error Resume Next

        Dim sngCalcTempAv As Single
        Dim sngCalcRHAv As Single
        Dim sngMaxAv As Single
        Dim sngMinAv As Single
        Dim sngTav As Single
        Dim sngSTD As Single
        Dim i As Integer
        Dim j As Integer
        Dim intCountTemp As Integer
        Dim intCountRH As Integer
        Dim intValidHours As Integer
        Dim intAvOfAvCount As Integer

        Dim xlApp As Excel.Application = Nothing
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing

        Dim xlTempSheet As Excel.Worksheet = Nothing
        Dim xlRHSheet As Excel.Worksheet = Nothing

        Dim missing As Object = Type.Missing

        Dim intTempSheet As Integer
        Dim intRHSheet As Integer

        Dim sngTasterix As Single

        xlApp = New Excel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

        ' Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

        ' Worksheet for temp and RH from user textbox input
        intTempSheet = txtTempWorksheet.Text
        intRHSheet = txtRHWorksheet.Text

        ' Get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

        CreateFourierHeaders()

        For j = 3 To 26 ' loop thru excel sheet columns

            ' Reset counters
            sngCalcTempAv = 0
            intCountTemp = 0
            sngCalcRHAv = 0
            intCountRH = 0

            ' Loop thru excel sheet rows
            For i = 5 To 483 ' last

                ' Get each temp value
                Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                ' Verify that data is a temperature value
                If VerifiedData(xlSecondTempCell.Value2) Then

                    ' Count columns with valid data
                    intValidHours = j

                    ' Get corresponding RH
                    Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)

                    sngCalcTempAv = sngCalcTempAv + xlSecondTempCell.Value2
                    intCountTemp = intCountTemp + 1

                    sngCalcRHAv = sngCalcRHAv + xlRHCell.Value2
                    intCountRH = intCountRH + 1
                End If

            Next

            If sngCalcTempAv <> 0 Then
                dgvFourierTh.Rows.Item(0).Cells(j - 2).Value = Math.Round(sngCalcTempAv / intCountTemp, 1)
                dgvFourierRh.Rows.Item(0).Cells(j - 2).Value = Math.Round(sngCalcRHAv / intCountRH, 1)
            Else
                dgvFourierTh.Rows.Item(0).Cells(j - 2).Value = ""
                dgvFourierRh.Rows.Item(0).Cells(j - 2).Value = ""
            End If
        Next

        sngCalcTempAv = 0
        sngMinAv = 10000
        sngMaxAv = 0

        intAvOfAvCount = 0

        For j = 1 To 24
            ' Find max average (ensure cell is not empty)
            If dgvFourierTh.Rows.Item(0).Cells(j).Value <> "" Then
                If dgvFourierTh.Rows.Item(0).Cells(j).Value > sngMaxAv Then
                    sngMaxAv = dgvFourierTh.Rows.Item(0).Cells(j).Value
                End If

                ' Find min average
                If dgvFourierTh.Rows.Item(0).Cells(j).Value < sngMinAv Then
                    sngMinAv = dgvFourierTh.Rows.Item(0).Cells(j).Value
                End If

                intAvOfAvCount = intAvOfAvCount + 1

                ' Sum the averages
                sngCalcTempAv = sngCalcTempAv + dgvFourierTh.Rows.Item(0).Cells(j).Value
            End If
        Next

        ' Get average of averages
        sngCalcTempAv = Math.Round(sngCalcTempAv / intAvOfAvCount, 1)

        For j = 1 To 24 ' columns

            ' Ensure cells are not empty
            If IsNothing(dgvFourierTh.Rows.Item(0).Cells(j).Value) = False Then

                ' Calculate fourier hourly temp
                sngTasterix = (2 * Math.PI * (j - 1)) / 24
                dgvFourierTh.Rows.Item(1).Cells(j).Value = Math.Round(CalculateFourierHourlyTemp(CSng(txtNewFourierTav.Text), CSng(txtNewFourierTrange.Text), sngTasterix), 1)

            End If
        Next

        intValidHours = 0
        sngSTD = 0
        For j = 3 To 26 ' loop thru excel sheet columns
            For i = 5 To 483 ' last

                ' Get each temp value
                Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

                ' Verify that data is a temperature value
                If VerifiedData(xlSecondTempCell.Value2) Then

                    ' Count columns with valid data
                    intValidHours = intValidHours + 1

                    sngSTD = sngSTD + Math.Pow(xlSecondTempCell.Value2 - CSng(txtNewFourierTav.Text), 2)
                End If

            Next
        Next

        '***********************************************RH

        sngCalcRHAv = 0
        sngMinAv = 10000
        sngMaxAv = 0

        intAvOfAvCount = 0

        For j = 1 To 24
            If dgvFourierRh.Rows.Item(0).Cells(j).Value <> "" Then
                ' Find max average
                If dgvFourierRh.Rows.Item(0).Cells(j).Value > sngMaxAv Then
                    sngMaxAv = dgvFourierRh.Rows.Item(0).Cells(j).Value
                End If

                ' Find min average
                If dgvFourierRh.Rows.Item(0).Cells(j).Value < sngMinAv Then
                    sngMinAv = dgvFourierRh.Rows.Item(0).Cells(j).Value
                End If

                intAvOfAvCount = intAvOfAvCount + 1

                ' Sum averages
                sngCalcRHAv = sngCalcRHAv + dgvFourierRh.Rows.Item(0).Cells(j).Value
            End If
        Next

        sngCalcRHAv = Math.Round(sngCalcRHAv / intAvOfAvCount, 2)

        For j = 1 To 24 ' columns

            ' Ensure cell is not empty
            If IsNothing(dgvFourierRh.Rows.Item(0).Cells(j).Value) = False Then

                ' Calculate fourier Rh 
                sngTasterix = (2 * Math.PI * (j - 1)) / 24
                dgvFourierRh.Rows.Item(1).Cells(j).Value = Math.Round(CalculateFourierHourlyRH(CSng(txtNewFourierRHav.Text), CSng(txtNewFourierRHrange.Text), sngTasterix), 1)
            End If
        Next

    End Sub

    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
            obj.Workbooks.Close()
            obj.Quit()
            obj = Nothing
        Catch ex As System.Exception
            System.Diagnostics.Debug.Print(ex.ToString())
            obj.Workbooks.Close()
            obj.Quit()
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Function CheckTempRangeOfData(ByVal TheTempData As Single) As Single
        On Error Resume Next

        If Strings.InStr(TheTempData, ".") > 0 Then 'ensure decimal point exists
            'check for range
            'return whole number part of figure, cut off decimal point and decimal part
            CheckTempRangeOfData = CSng(Strings.Left(TheTempData, Strings.InStr(TheTempData, ".") - 1))
        Else
            'check if data is without decimal, return whole data
            CheckTempRangeOfData = TheTempData
        End If
    End Function

    Private Function CheckRHRangeOfData(ByVal TheRHData As Single) As Integer
        On Error Resume Next

        'check for range within which data falls
        'divide by 5 to get vertical position, round up if with fraction
        'if not fraction, add 1 to result (to effect, for instance, values from 60-64.9 being counted instead of 60-65. 
        '65 is then included in range 65-70)
        If Strings.InStr((TheRHData / 5), ".") > 0 Then 'ensure decimal point exists
            If Strings.Right((TheRHData / 5), Strings.InStr((TheRHData / 5), ".") + 1) > 0 Then 'check if fraction exists
                CheckRHRangeOfData = Strings.Left((TheRHData / 5), Strings.InStr((TheRHData / 5), ".") - 1) + 1 'round up integer part
            Else
                CheckRHRangeOfData = (TheRHData / 5) + 1 'no decimal, i.e. boarder line figure, thus add 1 to upgrade to next range
            End If
        Else
            CheckRHRangeOfData = (TheRHData / 5) + 1 'no decimal, i.e. boarder line figure, thus add 1 to upgrade to next range
        End If
    End Function

    Private Sub CreateCoincidentTempRHProbabilityHeaders()
        On Error Resume Next

        Dim i As Integer
        Dim intColumns As Integer
        Dim intRH As Integer

        ClearProbabilityTable()

        If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth" Then
            With dgvRangeProbability
                .Columns.Add("Temp", "Temp")
                .Width = 1240
                .Height = 486
                .Top = 80

                .Columns(0).Width = 41
                .RowHeadersWidth = 30

                For i = 1 To 50 'range starts at 1-2, ends at 49-50
                    .Columns.Add("Therange", i & "-" & i + 1)
                    intColumns += 1
                    .Columns(intColumns).Width = 64
                Next

                .Rows.Add()
                .Rows.Item(0).Cells(0).Value = "RH"
                For i = 1 To 21
                    .Rows.Add()
                    .Rows.Item(i).Cells(0).Value = intRH & "-" & intRH + 5 'range starts at 0-5, ends at 100-105
                    intRH += 5
                Next
            End With
        Else
            With dgvRangeProbabilityPerYear
                .Columns.Add("Temp", "Temp")
                .Width = 1240
                .Height = 486
                .Top = 80

                .Columns(0).Width = 41
                .RowHeadersWidth = 30

                For i = 1 To 50 'range starts at 1-2, ends at 49-50
                    .Columns.Add("Therange", i & "-" & i + 1)
                    intColumns += 1
                    .Columns(intColumns).Width = 64
                Next

                .Rows.Add()
                .Rows.Item(0).Cells(0).Value = "RH"
                For i = 1 To 21
                    .Rows.Add()
                    .Rows.Item(i).Cells(0).Value = intRH & "-" & intRH + 5 'range starts at 0-5, ends at 100-105
                    intRH += 5
                Next
            End With
        End If
    End Sub

    Private Sub CreateCoincidentTempRHHeaders()
        On Error Resume Next

        Dim i As Integer
        Dim intColumns As Integer
        Dim intRH As Integer

        ClearCountTable()

        If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth" Then
            With dgvRangeCategories
                .Columns.Add("Temp", "Temp")
                .Width = 1210
                .Height = 500
                .Top = 80

                .Columns(0).Width = 41
                .RowHeadersWidth = 30

                For i = 1 To 50 'range starts at 12-13, ends at 39-40
                    .Columns.Add("Therange", i & "-" & i + 1)
                    intColumns += 1
                    .Columns(intColumns).Width = 40
                Next

                .Rows.Add()
                .Rows.Item(0).Cells(0).Value = "RH"
                For i = 1 To 21
                    .Rows.Add()
                    .Rows.Item(i).Cells(0).Value = intRH & "-" & intRH + 5 'range starts at 0-5, ends at 100-105
                    intRH += 5
                Next
            End With
        Else
            With dgvRangeCategoriesPerYear
                .Columns.Add("Temp", "Temp")
                .Width = 1210
                .Height = 500
                .Top = 80

                .Columns(0).Width = 41
                .RowHeadersWidth = 30

                For i = 1 To 50 'range starts at 12-13, ends at 39-40
                    .Columns.Add("Therange", i & "-" & i + 1)
                    intColumns += 1
                    .Columns(intColumns).Width = 40
                Next

                .Rows.Add()
                .Rows.Item(0).Cells(0).Value = "RH"
                For i = 1 To 21
                    .Rows.Add()
                    .Rows.Item(i).Cells(0).Value = intRH & "-" & intRH + 5 'range starts at 0-5, ends at 100-105
                    intRH += 5
                Next
            End With
        End If
    End Sub

    Private Sub CreateTempProbabilityHeaders()
        On Error Resume Next

        Dim intColumns As Integer

        ClearTempProbabilityTable()

        With dgvTempProbability
            .Columns.Add("Temp", "Temp")
            .Width = 1240
            .Height = 100

            .Columns(0).Width = 72

            For i = 1 To 50 'range starts at 1-2, ends at 49-50
                .Columns.Add("Therange", i & "-" & i + 1)
                intColumns += 1
                .Columns(intColumns).Width = 64
            Next

            .Columns.Item(0).HeaderText = "Temp. Bin Data"

            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Freq."

            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Probability"
        End With
    End Sub

    Private Sub CreateRHProbabilityHeaders()
        On Error Resume Next

        Dim i As Integer
        Dim intColumns As Integer
        Dim intRH As Integer

        ClearRHProbabilityTable()

        With dgvRHProbability
            .Columns.Add("RH", "RH")
            .Width = 1248
            .Height = 90

            .Columns(0).Width = 94

            intColumns = 0 'reset counter

            For i = 1 To 20
                .Columns.Add("Therange", intRH + 1 & "-" & intRH + 5)
                intRH += 5
                intColumns += 1
                .Columns(intColumns).Width = 55
            Next

            .Columns.Item(0).HeaderText = "RH Bin Data"

            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Freq."

            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Probability"
        End With
    End Sub

    Private Sub ClearCoincidentTable()
        On Error Resume Next

        With dgvCoincidentRH
            .RowCount = 1
            .ColumnCount = 0
        End With
    End Sub

    Private Sub ClearFourierTables()
        On Error Resume Next

        With dgvFourierTh
            .RowCount = 1
            .ColumnCount = 0
        End With

        With dgvFourierRh
            .RowCount = 1
            .ColumnCount = 0
        End With
    End Sub

    Private Sub ClearAveragesTables()
        On Error Resume Next

        With dgvTempAverages
            .RowCount = 1
            .ColumnCount = 0
        End With

        With dgvRHAverages
            .RowCount = 1
            .ColumnCount = 0
        End With
    End Sub

    Private Sub ClearTempProbabilityTable()
        On Error Resume Next

        With dgvTempProbability
            .RowCount = 1
            .ColumnCount = 0
        End With
    End Sub

    Private Sub ClearRHProbabilityTable()
        On Error Resume Next

        With dgvRHProbability
            .RowCount = 1
            .ColumnCount = 0
        End With
    End Sub

    Private Sub ClearProbabilityTable()
        On Error Resume Next

        If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth" Then
            With dgvRangeProbability
                .RowCount = 1
                .ColumnCount = 0
            End With
        Else
            With dgvRangeProbabilityPerYear
                .RowCount = 1
                .ColumnCount = 0
            End With
        End If
    End Sub

    Private Sub ClearCountTable()
        On Error Resume Next

        If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth" Then
            With dgvRangeCategories
                .RowCount = 1
                .ColumnCount = 0
            End With
        Else
            With dgvRangeCategoriesPerYear
                .RowCount = 1
                .ColumnCount = 0
            End With
        End If
    End Sub

    Private Sub ClearCountTable()
        On Error Resume Next

        If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth" Then
            With dgvRangeCategories
                .RowCount = 1
                .ColumnCount = 0
            End With
        Else
            With dgvRangeCategoriesPerYear
                .RowCount = 1
                .ColumnCount = 0
            End With
        End If
    End Sub

    Private Sub CreateFourierHeaders()
        On Error Resume Next

        Dim i As Integer

        'create 23 columns for table
        With dgvFourierTh
            .Width = 1224
            .Height = 112
            .RowCount = 0
            .ColumnCount = 0

            For i = -1 To 23
                If i = -1 Then
                    .Columns.Add("Hour", "")
                    .Columns(i + 1).Width = 100
                Else
                    .Columns.Add("Hour", i)
                    .Columns(i + 1).Width = 45
                End If
            Next

            'create rows for table
            .Columns.Item(0).HeaderText = "Hour"
            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Hourly Th"
            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Fourier Temp"
        End With

        With dgvFourierRh
            .Width = 1224
            .Height = 112
            .RowCount = 0
            .ColumnCount = 0

            'create 23 columns for table
            For i = -1 To 23
                If i = -1 Then
                    .Columns.Add("Hour", "")
                    .Columns(i + 1).Width = 115
                Else
                    .Columns.Add("Hour", i)
                    .Columns(i + 1).Width = 45
                End If
            Next

            'create rows for table
            .Columns.Item(0).HeaderText = "Hour"
            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Hourly Rh"
            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Fourier RH"
        End With
    End Sub

    Private Sub CreateAveragesHeaders()
        On Error Resume Next

        Dim i As Integer

        ClearAveragesTables()

        'create 23 columns for table
        With dgvTempAverages
            .Width = 1224
            .Height = 112

            For i = -1 To 23
                If i = -1 Then
                    .Columns.Add("Hour", "")
                    .Columns(i + 1).Width = 100
                Else
                    .Columns.Add("Hour", i)
                    .Columns(i + 1).Width = 45
                End If
            Next

            'create rows for table
            .Columns.Item(0).HeaderText = "Hour"
            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Th"
            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Hourly Temp"
        End With

        With dgvRHAverages
            .Width = 1249
            .Height = 112

            'create 23 columns for table
            For i = -1 To 23
                If i = -1 Then
                    .Columns.Add("Hour", "")
                    .Columns(i + 1).Width = 115
                Else
                    .Columns.Add("Hour", i)
                    .Columns(i + 1).Width = 45
                End If
            Next

            'create rows for table
            .Columns.Item(0).HeaderText = "Hour"
            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "RHh"
            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Hourly RH"
        End With
    End Sub

    Private Sub ComputeHeatingDegreeDays()
        On Error Resume Next

        Dim sngHeatingHb As Single

        sngHeatingHb = (CSng(txtBaseTempModelHDD.Text) - CSng(txtTavHDD.Text)) / (CSng(txtSIGMAmModelHDD.Text) * Math.Sqrt(CSng(cboNumofDaysInMonthHDD.Text)))

        txtModelAvgMonthlyHDD.Text = Math.Round((CSng(txtSIGMAmModelHDD.Text) * Math.Pow(CSng(cboNumofDaysInMonthHDD.Text), (3 / 2))) * (0.072196 + (sngHeatingHb / 2) + (1 / 9.6) * Math.Log(Math.Cosh(4.8 * sngHeatingHb))), 2)
    End Sub

    Private Sub ComputeCoolingDegreeDays()
        On Error Resume Next

        Dim sngCoolingHb As Single

        sngCoolingHb = (CSng(txtTavCDD.Text) - CSng(txtBaseTempModelCDD.Text)) / (CSng(txtSIGMAmModelCDD.Text) * Math.Sqrt(CSng(cboNumofDaysInMonthCDD.Text)))

        txtModelAvgMonthlyCDD.Text = MathPrivate Sub ClearCountTable()
        On Error Resume Next

        If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth" Then
            With dgvRangeCategories
                .RowCount = 1
                .ColumnCount = 0
            End With
        Else
            With dgvRangeCategoriesPerYear
                .RowCount = 1
                .ColumnCount = 0
            End With
        End If
    End Sub

    Private Sub CreateFourierHeaders()
        On Error Resume Next

        Dim i As Integer

        'create 23 columns for table
        With dgvFourierTh
            .Width = 1224
            .Height = 112
            .RowCount = 0
            .ColumnCount = 0

            For i = -1 To 23
                If i = -1 Then
                    .Columns.Add("Hour", "")
                    .Columns(i + 1).Width = 100
                Else
                    .Columns.Add("Hour", i)
                    .Columns(i + 1).Width = 45
                End If
            Next

            'create rows for table
            .Columns.Item(0).HeaderText = "Hour"
            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Hourly Th"
            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Fourier Temp"
        End With

        With dgvFourierRh
            .Width = 1224
            .Height = 112
            .RowCount = 0
            .ColumnCount = 0

            'create 23 columns for table
            For i = -1 To 23
                If i = -1 Then
                    .Columns.Add("Hour", "")
                    .Columns(i + 1).Width = 115
                Else
                    .Columns.Add("Hour", i)
                    .Columns(i + 1).Width = 45
                End If
            Next

            'create rows for table
            .Columns.Item(0).HeaderText = "Hour"
            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Hourly Rh"
            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Fourier RH"
        End With
    End Sub

    Private Sub CreateAveragesHeaders()
        On Error Resume Next

        Dim i As Integer

        ClearAveragesTables()

        'create 23 columns for table
        With dgvTempAverages
            .Width = 1224
            .Height = 112

            For i = -1 To 23
                If i = -1 Then
                    .Columns.Add("Hour", "")
                    .Columns(i + 1).Width = 100
                Else
                    .Columns.Add("Hour", i)
                    .Columns(i + 1).Width = 45
                End If
            Next

            'create rows for table
            .Columns.Item(0).HeaderText = "Hour"
            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Th"
            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Hourly Temp"
        End With

        With dgvRHAverages
            .Width = 1249
            .Height = 112

            'create 23 columns for table
            For i = -1 To 23
                If i = -1 Then
                    .Columns.Add("Hour", "")
                    .Columns(i + 1).Width = 115
                Else
                    .Columns.Add("Hour", i)
                    .Columns(i + 1).Width = 45
                End If
            Next

            'create rows for table
            .Columns.Item(0).HeaderText = "Hour"
            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "RHh"
            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Hourly RH"
        End With
    End Sub

    Private Sub ComputeHeatingDegreeDays()
        On Error Resume Next

        Dim sngHeatingHb As Single

        sngHeatingHb = (CSng(txtBaseTempModelHDD.Text) - CSng(txtTavHDD.Text)) / (CSng(txtSIGMAmModelHDD.Text) * Math.Sqrt(CSng(cboNumofDaysInMonthHDD.Text)))

        txtModelAvgMonthlyHDD.Text = Math.Round((CSng(txtSIGMAmModelHDD.Text) * Math.Pow(CSng(cboNumofDaysInMonthHDD.Text), (3 / 2))) * (0.072196 + (sngHeatingHb / 2) + (1 / 9.6) * Math.Log(Math.Cosh(4.8 * sngHeatingHb))), 2)
    End Sub

    Private Sub ComputeCoolingDegreeDays()
        On Error Resume Next

        Dim sngCoolingHb As Single

        sngCoolingHb = (CSng(txtTavCDD.Text) - CSng(txtBaseTempModelCDD.Text)) / (CSng(txtSIGMAmModelCDD.Text) * Math.Sqrt(CSng(cboNumofDaysInMonthCDD.Text)))

        txtModelAvgMonthlyCDD.Text = Math.Round((CSng(txtSIGMAmModelCDD.Text) * Math.Pow(CSng(cboNumofDaysInMonthCDD.Text), (3 / 2))) * (0.072196 + (sngCoolingHb / 2) + (1 / 9.6) * Math.Log(Math.Cosh(4.8 * sngCoolingHb))), 1)
    End Sub

    Private Function VerifyInputs() As Boolean
        On Error Resume Next

        ' Make sure inputs are valid
        If txtTotalHDD.SelectedTab.Name = "tpgHeatingDegreeDays" Then
            If txtBaseTempModelHDD.Text <> "" And txtSIGMAmModelHDD.Text <> "" And cboNumofDaysInMonthHDD.Text <> "" And txtNumofYearsHDD.Text <> "" And txtBaseTempActualHDD.Text <> "" Then
                If IsNumeric(txtBaseTempModelHDD.Text) = True And IsNumeric(txtSIGMAmModelHDD.Text) = True And IsNumeric(txtNumofYearsHDD.Text) = True And IsNumeric(txtBaseTempActualHDD.Text) = True Then
                    VerifyInputs = True
                Else
                    VerifyInputs = False
                End If
            Else
                VerifyInputs = False
            End If
        ElseIf txtTotalHDD.SelectedTab.Name = "tpgCoolingDegreeDays" Then
            If txtBaseTempModelCDD.Text <> "" And txtSIGMAmModelCDD.Text <> "" And cboNumofDaysInMonthCDD.Text <> "" And txtNumofYearsCDD.Text <> "" And txtBaseTempActualCDD.Text <> "" Then
                If IsNumeric(txtBaseTempModelCDD.Text) = True And IsNumeric(txtSIGMAmModelCDD.Text) = True And IsNumeric(txtNumofYearsCDD.Text) = True And IsNumeric(txtBaseTempActualCDD.Text) = True Then
                    VerifyInputs = True
                Else
                    VerifyInputs = False
                End If
            Else
                VerifyInputs = False
            End If
        End If
    End Function

    Private Function VerifyFourierInputs() As Boolean
        On Error Resume Next

        ' Make sure inputs are valid
        If txtNewFourierTav.Text <> "" And txtNewFourierTrange.Text <> "" And txtNewFourierRHav.Text <> "" And txtNewFourierRHrange.Text <> "" Then
            If InStr(txtNewFourierTav.Text, " ") = 0 And InStr(txtNewFourierTrange.Text, " ") = 0 And InStr(txtNewFourierRHav.Text, " ") = 0 And InStr(txtNewFourierRHrange.Text, " ") = 0 Then
                If IsNumeric(CSng(txtNewFourierTav.Text)) = True And IsNumeric(CSng(txtNewFourierTrange.Text)) = True And IsNumeric(CSng(txtNewFourierRHav.Text)) = True And IsNumeric(CSng(txtNewFourierRHrange.Text)) = True Then
                    VerifyFourierInputs = True
                Else
                    VerifyFourierInputs = False
                End If
            Else
                VerifyFourierInputs = False
            End If
        Else
            VerifyFourierInputs = False
        End If
    End Function

    Private Sub ClearFourierBoxes()
        On Error Resume Next
        ' Add code to clear Fourier input boxes if needed
    End Sub

    Private Sub ClearHeatingDegreeBoxes()
        On Error Resume Next

        txtModelAvgMonthlyHDD.Text = ""
        txtActualAvgMonthlyHDD.Text = ""
        txtActualTotalHDD.Text = ""
    End Sub

    Private Sub ClearCoolingDegreeBoxes()
        On Error Resume Next

        txtModelAvgMonthlyCDD.Text = ""
        txtActualTotalCDD.Text = ""
        txtActualAvgMonthlyCDD.Text = ""
    End Sub

    Private Sub ClearBoxes()
        On Error Resume Next

        ' Clear user input boxes for new entry
        txtTempRange.Text = ""
        txtTStandardDeviation.Text = ""
        txtMaxTempAverage.Text = ""
        txtMinTempAverage.Text = ""
        txtTav.Text = ""

        txtRHRange.Text = ""
        txtRHStandardDeviation.Text = ""
        txtMaxRHAverage.Text = ""
        txtMinRHAverage.Text = ""
        txtRHav.Text = ""
    End Sub

    Private Sub frmRelativeHumidity_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        On Error Resume Next

        Me.Cursor = Cursors.WaitCursor

        CreateAveragesHeaders()
        CreateTempProbabilityHeaders()
        CreateRHProbabilityHeaders()

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnRangeCategoriesPerMonth_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRangeCategoriesPerMonth.Click
        On Error Resume Next

        Me.Cursor = Cursors.WaitCursor

        If VerifyWorksheets() = False Then
            MsgBox("Input data is invalid", vbOKOnly + MsgBoxStyle.Exclamation, "")
            Me.Cursor = Cursors.Default
            GoTo exit_sub
        End If

        CreateCoincidentTempRHHeaders()
        CreateCoincidentTempRHProbabilityHeaders()

        If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth" Then
            FreqOfRHInTempRange(dgvRangeCategories, dgvRangeProbability, 483)
        ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear" Then
            FreqOfRHInTempRange(dgvRangeCategoriesPerYear, dgvRangeProbabilityPerYear, 5800)
        End If

        TransferToWorkSheet()

        Me.Cursor = Cursors.Default

    exit_sub:
        Exit Sub
    End Sub

    Private Sub txtTempWorksheet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTempWorksheet.TextChanged
        On Error Resume Next

        If VerifyWorksheets() = True And txtExcelFile.Text <> "" Then
            ClearCountTable()
            ClearCoincidentTable()
            txtTotalHDD.Enabled = True
        Else
            txtTotalHDD.Enabled = False
        End If
    End Sub

    Private Sub txtExcelFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExcelFile.TextChanged
        On Error Resume Next

        ClearCountTable()
        ClearCoincidentTable()

        If txtExcelFile.Text <> "" And VerifyWorksheets() = True Then
            txtTotalHDD.Enabled = True

            ' Simulate selection of basic data radio button
            rbtBasicDataPerMonth.Checked = True
            rbtBasicDataPerYear.Checked = True
        Else
            txtTotalHDD.Enabled = False
        End If
    End Sub

    Private Sub txtRHWorksheet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRHWorksheet.TextChanged
        On Error Resume Next

        If VerifyWorksheets() = True And txtExcelFile.Text <> "" Then
            ClearCountTable()
            ClearCoincidentTable()
            txtTotalHDD.Enabled = True
        Else
            txtTotalHDD.Enabled = False
        End If
    End Sub

    Private Sub btnExcelFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcelFile.Click
        On Error Resume Next

        GetDatabasePath()
    End Sub


    Private Sub btnAverages_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAverages.Click
        On Error Resume Next

        Me.Cursor = Cursors.WaitCursor

        ClearBoxes()

        ' If VerifyInputs() = True Then
        ComputeAverages()
        TransferToWorkSheet()

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnTempProbability_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProbability.Click
        On Error Resume Next

        Me.Cursor = Cursors.WaitCursor

        If VerifyWorksheets() = False Then
            MsgBox("Input data is invalid", vbOKOnly + MsgBoxStyle.Exclamation, "")
            Me.Cursor = Cursors.Default
            GoTo exit_sub
        End If

        FreqOfTempInTempRange()
        FreqOfRHInRHRange()
        TransferToWorkSheet()

        Me.Cursor = Cursors.Default

    exit_sub:
        Exit Sub
    End Sub

    Private Sub tpgRangeCategories_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tpgRangeCategoriesPerMonth.Click
        On Error Resume Next

        ' Simulate selection of basic data radio button
        rbtBasicData_Click(sender, e)
    End Sub

    Private Sub rbtBasicData_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtBasicDataPerMonth.Click
        On Error Resume Next

        If rbtBasicDataPerMonth.Checked = True Then
            dgvRangeCategories.Visible = True
            dgvRangeProbability.Visible = False
        End If
    End Sub

    Private Sub rbtProbability_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtProbabilityPerMonth.Click
        On Error Resume Next

        If rbtProbabilityPerMonth.Checked = True Then
            dgvRangeProbability.Visible = True
            dgvRangeCategories.Visible = False
        End If
    End Sub

    Private Sub btnRangeCategoriesPerYear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRangeCategoriesPerYear.Click
        On Error Resume Next

        btnRangeCategoriesPerMonth_Click(sender, e)
    End Sub

    Private Sub rbtBasicDataPerYear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtBasicDataPerYear.Click
        On Error Resume Next

        If rbtBasicDataPerYear.Checked = True Then
            dgvRangeCategoriesPerYear.Visible = True
            dgvRangeProbabilityPerYear.Visible = False
        End If
    End Sub

    Private Sub rbtProbabilityPerYear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtProbabilityPerYear.Click
        On Error Resume Next

        If rbtProbabilityPerYear.Checked = True Then
            dgvRangeProbabilityPerYear.Visible = True
            dgvRangeCategoriesPerYear.Visible = False
        End If
    End Sub

    Private Sub btnFourier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFourier.Click
        On Error Resume Next

        Me.Cursor = Cursors.WaitCursor

        If VerifyFourierInputs() = True Then
            ComputeFourier()
            TransferToWorkSheet()
        Else
            MsgBox("Input data is invalid, please input valid data. Ensure there are no spaces between figures, or between figures and decimal point.", vbOKOnly + MsgBoxStyle.Exclamation, "")
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnHeatingDegreeDays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHeatingDegreeDays.Click
        On Error Resume Next

        Me.Cursor = Cursors.WaitCursor

        ClearHeatingDegreeBoxes()

        If VerifyInputs() = True Then
            ComputeHeatingDegreeDays()
            CalculateHDDActualDh()
            TransferToWorkSheet()
        Else
            MsgBox("Input data is invalid, please input valid data.", vbOKOnly + MsgBoxStyle.Exclamation, "")
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnCoolingDegreeDays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCoolingDegreeDays.Click
        On Error Resume Next

        Me.Cursor = Cursors.WaitCursor

        ClearCoolingDegreeBoxes()

        If VerifyInputs() = True Then
            ComputeCoolingDegreeDays()
            CalculateCDDActualDc()
            TransferToWorkSheet()
        Else
            MsgBox("Input data is invalid, please input valid data.", vbOKOnly + MsgBoxStyle.Exclamation, "")
        End If

        Me.Cursor = Cursors.Default
    End Sub
End Class