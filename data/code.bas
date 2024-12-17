'important for connection to excel files
Imports System.Globalization
Imports System.Threading
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel

PublicClassfrmRelativeHumidity
'important for transfer of data between excel and RelativeHumidity Application
Dim xlTransfer AsNew Microsoft.Office.Interop.Excel.Application
Dim objExlApp AsObject
Dim objWrkBk AsObject
Dim objWrkSheet1 AsObject
Dim objWrkSheet2 AsObject

Dim strFileName AsString'string to contain path to excel file

Dim intSheet AsInteger'counter to manipulate various excel sheets

PrivateSubFreqOfRHInTempRange(ByVal TheTable AsDataGridView, ByVal TheProbabilityTable AsDataGridView, ByVal TheEndRow AsInteger)
'declearations for manipulation of excel workbook and worksheets
Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US")

Dim xlApp As Excel.Application = Nothing
Dim xlBooks As Excel.Workbooks = Nothing
Dim xlBook As Excel.Workbook = Nothing

Dim xlTempSheet As Excel.Worksheet = Nothing
Dim xlRHSheet As Excel.Worksheet = Nothing

Dim missing AsObject = Type.Missing

Dim intTempSheet AsInteger
Dim intRHSheet AsInteger

'declearations for counters used to loop through excel worksheet horizontally and vertically
Dim i AsInteger
Dim j AsInteger
Dim k AsInteger
Dim l AsInteger
Dim intCounter AsInteger
Dim intProbabilityCount AsInteger
Dim decProbability AsDecimal

Try
'create new instance of excel application and set necessary properties
xlApp = NewExcel.Application()
xlApp.DisplayAlerts = False
xlApp.UserControl = True
xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)
'Change full path of Excel file here:
'open workbook, supply all necessary parameters
xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

'obtain worksheet and range for temp from user textbox input
intTempSheet = txtTempWorksheet.Text

'obtain worksheet and range for RH from user textbox input
intRHSheet = txtRHWorksheet.Text

'get a sheet in the workbook
xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

'loop thru corresponding excel sheet rows
For i = 5 To TheEndRow
    'loop thru excel sheet columns
    For j = 3 To 26
        'declear variable for manipulating each excel cell
        Dim xlTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)
        'verify temp data
        If VerifiedData(xlTempCell.Value2) = True Then

            'increament counter for probability calculation (counting number of data points)
            intProbabilityCount = intProbabilityCount + 1

            'check for range of data
            k = CheckTempRangeOfData(xlTempCell.Value2)

            'get corresponding RH value
            Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)
            l = CheckRHRangeOfData(xlRHCell.Value2)

            'increament counter
            'place count value in appropriate column
            If IsNothing(TheTable.Rows.Item(l).Cells(k).Value) = False Then
            '(check if cell is presently empty)
            'if not empty, increment number in it. (CInt converts contents of cell to integer data type)
                TheTable.Rows.Item(l).Cells(k).Value = CInt(TheTable.Rows.Item(l).Cells(k).Value) + 1
            Else
                intCounter = intCounter + 1
                TheTable.Rows.Item(l).Cells(k).Value = intCounter 'place counter value in cell
                intCounter = 0 'reset counter
            End If

        End If

      Next
Next

'find probability
For i = 1 To TheTable.RowCount - 1
For j = 1 To TheTable.ColumnCount - 1
IfIsNothing(TheTable.Rows.Item(i).Cells(j).Value) = FalseThen

IfTheTable.Rows.Item(i).Cells(j).Value <> 0 Then
TheTable.Rows.Item(i).Cells(j).Value = TheTable.Rows.Item(i).Cells(j).Value
EndIf

                        decProbability = CDec(TheTable.Rows.Item(i).Cells(j).Value) / (intProbabilityCount)

'round up or down
TheProbabilityTable.Rows.Item(i).Cells(j).Value = CStr(RoundingProcess(decProbability))
Else
TheTable.Rows.Item(i).Cells(j).Value = 0
TheProbabilityTable.Rows.Item(i).Cells(j).Value = 0
EndIf
Next
Next

'close excel workbook
DirectCast(xlBook, Excel._Workbook).Close(True, missing, missing) '

'enable radio buttons
            rbtProbabilityPerMonth.Enabled = True
            rbtProbabilityPerYear.Enabled = True

Catch ex AsException
MsgBox("Either the selected data path or worksheet is invalid, please input valid data path and worksheet.", MsgBoxStyle.Critical + vbOKOnly, "")

Me.Cursor = Cursors.Default
Exit Sub

Finally
xlApp.Quit() 'end excel application
releaseObject(xlApp) 'release objects from computer memory

EndTry
EndSub

PrivateFunctionRoundingProcess(ByVal RoundMe AsDecimal) AsString
OnErrorResumeNext

Dim strCheck AsString
Dim intDecimalPosition AsInteger
Dim intLastDigit AsInteger
Dim strOutPut AsString

        intDecimalPosition = Strings.InStr(RoundMe, ".")

'return six places after decimal point
        strCheck = Strings.Left(CStr(RoundMe), intDecimalPosition + 7)

        intLastDigit = CInt(Strings.Right(strCheck, 1))
        strCheck = Strings.Left(CStr(RoundMe), intDecimalPosition + 6)

'check if last digit is greater than 4
If intLastDigit > 4 Then
'round up
            strOutPut = CStr(CDec(strCheck + 0.000001))
            RoundingProcess = strOutPut
Else
'round down
            RoundingProcess = Strings.Left(CStr(RoundMe), intDecimalPosition + 6)
EndIf

EndFunction

PrivateSubFreqOfTempInTempRange()

'make declearations for manipulating excel workbook, worksheets
Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US") '<-- change culture on whatever you need

Dim xlApp As Excel.Application = Nothing
Dim xlBooks As Excel.Workbooks = Nothing
Dim xlBook As Excel.Workbook = Nothing

Dim xlTempSheet As Excel.Worksheet = Nothing

Dim xlRHSheet As Excel.Worksheet = Nothing

Dim missing AsObject = Type.Missing

Dim intTempSheet AsInteger
Dim intRHSheet AsInteger

'declear variables for looping through rows and columns
Dim i AsInteger
Dim j AsInteger
Dim k AsInteger
Dim intCounter AsInteger
Dim intDataPointCount AsInteger

Try

            xlApp = NewExcel.Application()

            xlApp.DisplayAlerts = False
            xlApp.UserControl = True
            xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)
'Change full path of Excel file here:
            xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

'worksheet and range for temp from user textbox input
            intTempSheet = txtTempWorksheet.Text

'worksheet and range for RH from user textbox input
            intRHSheet = txtRHWorksheet.Text

'get sheet in the workbook
            xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
            xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

CreateTempProbabilityHeaders()

            intDataPointCount = 0

'loop thru corresponding excel sheet columns
For j = 3 To 26

'loop thru excel sheet rows
For i = 5 To 483

Dim xlTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

'verify temp data
IfVerifiedData(xlTempCell.Value2) = TrueThen

                        intDataPointCount = intDataPointCount + 1

'check if data falls within range of column k
                        k = CheckTempRangeOfData(xlTempCell.Value2)

'increament counter
'place count value in appropriate column
IfIsNothing(dgvTempProbability.Rows.Item(0).Cells(k).Value) = FalseThen'(check if presently empty)
'if not empty, increment number in it
dgvTempProbability.Rows.Item(0).Cells(k).Value = CInt(dgvTempProbability.Rows.Item(0).Cells(k).Value) + 1
Else
                            intCounter = intCounter + 1
dgvTempProbability.Rows.Item(0).Cells(k).Value = intCounter

                            intCounter = 0 'reset counter
EndIf

EndIf

Next

Next

'loop thru columns' ranges in vb grid (table) 
For k = 1 To dgvTempProbability.ColumnCount - 1
IfIsNothing(dgvTempProbability.Rows.Item(0).Cells(k).Value) = FalseThen
dgvTempProbability.Rows.Item(1).Cells(k).Value = Format(Math.Round(dgvTempProbability.Rows.Item(0).Cells(k).Value / intDataPointCount, 6), "0.######")
EndIf
Next

DirectCast(xlBook, Excel._Workbook).Close(True, missing, missing) '

Catch ex AsException

MsgBox("Either the selected database path or worksheets are invalid, therefore input a valid database path and worksheets.", MsgBoxStyle.Critical + vbOKOnly, "")

Me.Cursor = Cursors.Default
Exit Sub

Finally
xlApp.Quit()
releaseObject(xlApp)
EndTry
EndSub

PrivateSubFreqOfRHInRHRange()
'make declearations for manipulating excel workbook, worksheets
Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US") '<-- change culture on whatever you need

Dim xlApp As Excel.Application = Nothing
Dim xlBooks As Excel.Workbooks = Nothing
Dim xlBook As Excel.Workbook = Nothing

Dim xlRHSheet As Excel.Worksheet = Nothing

Dim missing AsObject = Type.Missing

Dim intRHSheet AsInteger

Dim i AsInteger
Dim j AsInteger
Dim k AsInteger
Dim intCounter AsInteger
Dim intDataPointCount AsInteger

Try

            xlApp = NewExcel.Application() 'create new excel application

            xlApp.DisplayAlerts = False
            xlApp.UserControl = True
            xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)
'Change full path of Excel file here:
            xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

'worksheet and range for RH from user textbox input
            intRHSheet = txtRHWorksheet.Text

'get sheet in the workbook from user textbox input
            xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

CreateRHProbabilityHeaders()

            intDataPointCount = 0

'loop thru corresponding excel sheet columns
For j = 3 To 26

'loop thru excel sheet rows
For i = 5 To 483

Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)

'verify temp data
IfVerifiedData(xlRHCell.Value2) = TrueThen

                        intDataPointCount = intDataPointCount + 1

'check if data falls within range of column k
                        k = CheckRHRangeOfData(xlRHCell.Value2)

'increament counter
'place count value in appropriate column
IfIsNothing(dgvRHProbability.Rows.Item(0).Cells(k).Value) = FalseThen'(check if presently empty)
'if not empty, increment number in it
dgvRHProbability.Rows.Item(0).Cells(k).Value = CInt(dgvRHProbability.Rows.Item(0).Cells(k).Value) + 1
Else
                            intCounter = intCounter + 1
dgvRHProbability.Rows.Item(0).Cells(k).Value = intCounter

                            intCounter = 0 'reset counter
EndIf

EndIf

Next
Next

'loop thru columns' ranges in vb grid 
For k = 1 To dgvRHProbability.ColumnCount - 1
IfIsNothing(dgvRHProbability.Rows.Item(0).Cells(k).Value) = FalseThen
dgvRHProbability.Rows.Item(1).Cells(k).Value = Format(Math.Round(dgvRHProbability.Rows.Item(0).Cells(k).Value / intDataPointCount, 6), "0.######")
EndIf
Next

DirectCast(xlBook, Excel._Workbook).Close(True, missing, missing) '

Catch ex AsException

MsgBox("Either the selected database path or worksheets are invalid, therefore input a valid database path and worksheets.", MsgBoxStyle.Critical + vbOKOnly, "")
Exit Sub

Me.Cursor = Cursors.Default
Finally
xlApp.Quit()
releaseObject(xlApp)
EndTry
EndSub

PrivateSubCoincidentRHs()
OnErrorResumeNext

Dim intMaximumRow AsInteger
Dim intTotalRows AsInteger
Dim sngAverage AsSingle
Dim i AsInteger
Dim j AsInteger

DimarryAverages() AsSingle

Dim xlApp As Excel.Application = Nothing
Dim xlBooks As Excel.Workbooks = Nothing
Dim xlBook As Excel.Workbook = Nothing

Dim xlTempSheet As Excel.Worksheet = Nothing

Dim xlRHSheet As Excel.Worksheet = Nothing

Dim missing AsObject = Type.Missing

Dim intTempSheet AsInteger

Dim intRHSheet AsInteger

        xlApp = NewExcel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

'Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

'worksheet and range for temp from user textbox input
        intTempSheet = txtTempWorksheet.Text

'worksheet and range for RH from user textbox input
        intRHSheet = txtRHWorksheet.Text

'get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

ClearCoincidentTable()

'loop thru Temp worksheet 
'loop thru corresponding RH worksheet
For i = 5 To 483 'rows
'loop thru excel sheet columns
For j = 3 To 26 'last

'get each temp value
Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

'verify that data is a temp value
IfVerifiedData(xlSecondTempCell.Value2) Then
'get coincident RH value of temp value
Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)

AverageRH(xlSecondTempCell.Value2, xlRHCell.Value2)
EndIf
Next
Next

With dgvCoincidentRH
'loop thru columns
For j = 0 To .ColumnCount - 1
'find average and place at end of each column
For i = 0 To .RowCount - 1

IfIsNothing(.Rows.Item(i).Cells(j).Value) = FalseThen
                        sngAverage = sngAverage + CSng(.Rows.Item(i).Cells(j).Value)
Else
'maximum is reached
                        intTotalRows = i

Exit For
EndIf
Next

'place average in array
ReDimPreservearryAverages(j)
arryAverages(j) = sngAverage / (intTotalRows)

If intMaximumRow < intTotalRows Then
                    intMaximumRow = intTotalRows
EndIf

'reset 
                sngAverage = 0
Next

'cut off rows which are 1 row after maximum number of rows
For i = .RowCount - 1 To intMaximumRow + 2 Step -1
                .Rows.RemoveAt(i)
Next

'loop thru array and place averages in appropriate columns on last row
For i = 0 ToarryAverages.Length()
'place average in 1 row after last row
                .Rows.Item(intMaximumRow + 1).Cells(i).Value = Math.Round(arryAverages(i), 2)

Next
EndWith

EndSub

PrivateSubCalculateHDDActualDh()
OnErrorResumeNext

Dim i AsInteger
Dim j AsInteger

DimarryAverages() AsSingle

Dim xlApp As Excel.Application = Nothing
Dim xlBooks As Excel.Workbooks = Nothing
Dim xlBook As Excel.Workbook = Nothing

Dim xlTempSheet As Excel.Worksheet = Nothing

Dim xlRHSheet As Excel.Worksheet = Nothing

Dim missing AsObject = Type.Missing

Dim intTempSheet AsInteger

Dim intRHSheet AsInteger

Dim sngDhSubtraction AsSingle
Dim sngDcSubtraction AsSingle
Dim sngDhSum AsSingle

        xlApp = NewExcel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

'Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

'worksheet and range for temp
        intTempSheet = txtTempWorksheet.Text

'worksheet and range for RH
        intRHSheet = txtRHWorksheet.Text

'get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

'loop thru Temp worksheet 
For i = 5 To 483 'rows
'loop thru excel sheet columns
For j = 3 To 26 'last

'get each temp value
Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

'verify that data is a temp value
IfVerifiedData(xlSecondTempCell.Value2) Then

'calcualte Dh total
                    sngDhSubtraction = txtBaseTempActualHDD.Text - xlSecondTempCell.Value2
If sngDhSubtraction > 0 Then
                        sngDhSum = sngDhSum + sngDhSubtraction
EndIf
EndIf
Next
Next

        txtActualTotalHDD.Text = Math.Round(sngDhSum / 24, 1)

        txtActualAvgMonthlyHDD.Text = Math.Round((sngDhSum / 24) / txtNumofYearsHDD.Text, 1)

EndSub

PrivateSubCalculateCDDActualDc()
OnErrorResumeNext

Dim i AsInteger
Dim j AsInteger

DimarryAverages() AsSingle

Dim xlApp As Excel.Application = Nothing
Dim xlBooks As Excel.Workbooks = Nothing
Dim xlBook As Excel.Workbook = Nothing

Dim xlTempSheet As Excel.Worksheet = Nothing

Dim xlRHSheet As Excel.Worksheet = Nothing

Dim missing AsObject = Type.Missing

Dim intTempSheet AsInteger

Dim intRHSheet AsInteger

Dim sngDcSubtraction AsSingle
Dim sngDcSum AsSingle

        xlApp = NewExcel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

'Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

'worksheet and range for temp
        intTempSheet = txtTempWorksheet.Text

'worksheet and range for RH
        intRHSheet = txtRHWorksheet.Text

'get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

'loop thru Temp worksheet 
For i = 5 To 483 'rows
'loop thru excel sheet columns
For j = 3 To 26 'last

'get each temp value
Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

'verify that data is a temp value
IfVerifiedData(xlSecondTempCell.Value2) Then

'calculate Dc total
                    sngDcSubtraction = xlSecondTempCell.Value2 - CSng(txtBaseTempActualCDD.Text)
If sngDcSubtraction > 0 Then
                        sngDcSum = sngDcSum + sngDcSubtraction
EndIf

EndIf
Next
Next

        txtActualTotalCDD.Text = Math.Round(sngDcSum / 24, 1)

        txtActualAvgMonthlyCDD.Text = Math.Round((sngDcSum / 24) / txtNumofYearsCDD.Text, 1)
EndSub

PrivateFunctionVerifiedData(ByVal TheData AsString) AsBoolean
OnErrorResumeNext

'ensure not empty
IfIsNothing(TheData) = FalseThen
'ensure it is numeric
IfIsNumeric(TheData) = TrueThen

'ensure it's >= 1 
If TheData >= 1 Then
                    VerifiedData = True
EndIf

EndIf
EndIf

EndFunction

PrivateSubAverageRH(ByVal TheTemp AsSingle, ByVal TheRH AsSingle)
OnErrorResumeNext

Dim boolPresent AsBoolean
Dim i AsInteger
Dim j AsInteger
Dim intNextRow AsInteger

With dgvCoincidentRH
If .ColumnCount < 1 Then
'place temp in column header (topmost row)
                .Columns.Add("temp", CStr(TheTemp))

                .Columns(0).Width = 44
EndIf

'loop thru columns
For i = 0 To .ColumnCount - 1

'check if temp is already present in column header (topmost row)
If .Columns.Item(i).HeaderCell.Value = CStr(TheTemp) Then

'increment row 
                    .Rows.Add()

'find next empty row cell under given column
For j = 0 To .RowCount - 1
IfIsNothing(.Rows.Item(j).Cells(i).Value) = TrueThen
                            intNextRow = j
Exit For
Else
                            intNextRow = intNextRow + 1
EndIf
Next

'place coincident RH in next row
                    .Rows.Item(intNextRow).Cells(i).Value = TheRH

                    boolPresent = True
EndIf

'adjust width of columns
If .Columns(0).Width <> 44 Then .Columns(i).Width = 44
Next

'check if temp is new (not already existing in topmost row)
If boolPresent = FalseThen
'place temp in column header (topmost row)
                .Columns.Add("temp", CStr(TheTemp))

                .Columns(i).Width = 44

'place coincident RH in first row
                .Rows.Item(0).Cells(i).Value = TheRH
EndIf

EndWith

EndSub

PrivateFunction CalculateHourlyTemp(ByVal TheAvTemp AsSingle, ByVal TheMaxTemp AsSingle, ByVal TheMinTemp AsSingle, ByVal Tasterix AsSingle) AsSingle
OnErrorResumeNext

        CalculateHourlyTemp = TheAvTemp + (TheMaxTemp - TheMinTemp) * (0.4535 * Math.Cos(Tasterix - 3.7522) + 0.1207 * Math.Cos(2 * Tasterix - 0.3895) + 0.0146 * Math.Cos(3 * Tasterix - 0.8927) + 0.0212 * Math.Cos(4 * Tasterix - 0.2674))

EndFunction

PrivateFunctionCalculateFourierHourlyTemp(ByVal TheAvTemp AsSingle, ByVal TheRange AsSingle, ByVal Tasterix AsSingle) AsSingle
OnErrorResumeNext

        CalculateFourierHourlyTemp = TheAvTemp + (TheRange) * (0.4535 * Math.Cos(Tasterix - 3.7522) + 0.1207 * Math.Cos(2 * Tasterix - 0.3895) + 0.0146 * Math.Cos(3 * Tasterix - 0.8927) + 0.0212 * Math.Cos(4 * Tasterix - 0.2674))

EndFunction

PrivateFunction CalculateHourlyRH(ByVal TheAvRH AsSingle, ByVal TheMaxRH AsSingle, ByVal TheMinRH AsSingle, ByVal Tasterix AsSingle) AsSingle
OnErrorResumeNext

        CalculateHourlyRH = TheAvRH + (TheMaxRH - TheMinRH) * (0.4602 * Math.Cos(Tasterix - 0.6038) + 0.1255 * Math.Cos(2 * Tasterix - 3.5427) + 0.0212 * Math.Cos(3 * Tasterix - 4.2635) + 0.0255 * Math.Cos(4 * Tasterix - 0.3833))

EndFunction

PrivateFunctionCalculateFourierHourlyRH(ByVal TheAvRH AsSingle, ByVal TheRange AsSingle, ByVal Tasterix AsSingle) AsSingle
OnErrorResumeNext

        CalculateFourierHourlyRH = TheAvRH + (TheRange) * (0.4606 * Math.Cos(Tasterix - 0.6038) + 0.1255 * Math.Cos(2 * Tasterix - 3.5247) + 0.0212 * Math.Cos(3 * Tasterix - 4.2635) + 0.0255 * Math.Cos(4 * Tasterix - 0.3833))

EndFunction

PrivateSubComputeFourier()
OnErrorResumeNext

Dim sngCalcTempAv AsSingle
Dim sngCalcRHAv AsSingle
Dim sngMaxAv AsSingle
Dim sngMinAv AsSingle
Dim sngTav AsSingle
Dim sngSTD AsSingle
Dim i AsInteger
Dim j AsInteger
Dim intCountTemp AsInteger
Dim intCountRH AsInteger
Dim intValidHours AsInteger
Dim intAvOfAvCount AsInteger

Dim xlApp As Excel.Application = Nothing
Dim xlBooks As Excel.Workbooks = Nothing
Dim xlBook As Excel.Workbook = Nothing

Dim xlTempSheet As Excel.Worksheet = Nothing
Dim xlRHSheet As Excel.Worksheet = Nothing

Dim missing AsObject = Type.Missing

Dim intTempSheet AsInteger
Dim intRHSheet AsInteger

Dim sngTasterix AsSingle

        xlApp = NewExcel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

'Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

'worksheet for temp and RH from user textbox input
        intTempSheet = txtTempWorksheet.Text
        intRHSheet = txtRHWorksheet.Text

'get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

CreateFourierHeaders()

For j = 3 To 26 'loop thru excel sheet columns

'reset counters
            sngCalcTempAv = 0
            intCountTemp = 0
            sngCalcRHAv = 0
            intCountRH = 0

'loop thru excel sheet rows
For i = 5 To483  'last

'get each temp value
Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

'verify that data is a temperature value
IfVerifiedData(xlSecondTempCell.Value2) Then

'count columns with valid data
                    intValidHours = j

'get corresponding RH
Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)

                    sngCalcTempAv = sngCalcTempAv + xlSecondTempCell.Value2
                    intCountTemp = intCountTemp + 1

                    sngCalcRHAv = sngCalcRHAv + xlRHCell.Value2
                    intCountRH = intCountRH + 1
EndIf

Next

If sngCalcTempAv <> 0 Then
dgvFourierTh.Rows.Item(0).Cells(j - 2).Value = Math.Round(sngCalcTempAv / intCountTemp, 1)
dgvFourierRh.Rows.Item(0).Cells(j - 2).Value = Math.Round(sngCalcRHAv / intCountRH, 1)
Else
dgvFourierTh.Rows.Item(0).Cells(j - 2).Value = ""
dgvFourierRh.Rows.Item(0).Cells(j - 2).Value = ""
EndIf
Next


        sngCalcTempAv = 0
        sngMinAv = 10000
        sngMaxAv = 0

        intAvOfAvCount = 0

For j = 1 To 24
'find max average (ensure cell is not empty)
IfdgvFourierTh.Rows.Item(0).Cells(j).Value <>""Then
IfdgvFourierTh.Rows.Item(0).Cells(j).Value > sngMaxAv Then
                    sngMaxAv = dgvFourierTh.Rows.Item(0).Cells(j).Value
EndIf

'find min average
IfdgvFourierTh.Rows.Item(0).Cells(j).Value < sngMinAv Then
                    sngMinAv = dgvFourierTh.Rows.Item(0).Cells(j).Value
EndIf

                intAvOfAvCount = intAvOfAvCount + 1

'sum the averages
                sngCalcTempAv = sngCalcTempAv + dgvFourierTh.Rows.Item(0).Cells(j).Value
EndIf
Next

'get average of averages
        sngCalcTempAv = Math.Round(sngCalcTempAv / intAvOfAvCount, 1)

For j = 1 To 24 'columns

'ensure cells are not empty
IfIsNothing(dgvFourierTh.Rows.Item(0).Cells(j).Value) = FalseThen


'calculate fourier hourly temp
                sngTasterix = (2 * Math.PI * (j - 1)) / 24
dgvFourierTh.Rows.Item(1).Cells(j).Value = Math.Round(CalculateFourierHourlyTemp(CSng(txtNewFourierTav.Text), CSng(txtNewFourierTrange.Text), sngTasterix), 1)

EndIf
Next

        intValidHours = 0
        sngSTD = 0
For j = 3 To 26 'loop thru excel sheet columns
For i = 5 To483  'last

'get each temp value
Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

'verify that data is a temperature value
IfVerifiedData(xlSecondTempCell.Value2) Then

'count columns with valid data
                    intValidHours = intValidHours + 1

                    sngSTD = sngSTD + Math.Pow(xlSecondTempCell.Value2 - CSng(txtNewFourierTav.Text), 2)
EndIf

Next
Next

'***********************************************RH


        sngCalcRHAv = 0
        sngMinAv = 10000
        sngMaxAv = 0

        intAvOfAvCount = 0

For j = 1 To 24
IfdgvFourierRh.Rows.Item(0).Cells(j).Value <>""Then
'find max average
IfdgvFourierRh.Rows.Item(0).Cells(j).Value > sngMaxAv Then
                    sngMaxAv = dgvFourierRh.Rows.Item(0).Cells(j).Value
EndIf

'find min average
IfdgvFourierRh.Rows.Item(0).Cells(j).Value < sngMinAv Then
                    sngMinAv = dgvFourierRh.Rows.Item(0).Cells(j).Value
EndIf

                intAvOfAvCount = intAvOfAvCount + 1

'sum averages
                sngCalcRHAv = sngCalcRHAv + dgvFourierRh.Rows.Item(0).Cells(j).Value
EndIf
Next

        sngCalcRHAv = Math.Round(sngCalcRHAv / intAvOfAvCount, 2)

For j = 1 To 24 'columns

'ensure cell is not empty
IfIsNothing(dgvFourierRh.Rows.Item(0).Cells(j).Value) = FalseThen

'calculate fourier Rh 
                sngTasterix = (2 * Math.PI * (j - 1)) / 24
dgvFourierRh.Rows.Item(1).Cells(j).Value = Math.Round(CalculateFourierHourlyRH(CSng(txtNewFourierRHav.Text), CSng(txtNewFourierRHrange.Text), sngTasterix), 1)
EndIf
Next

EndSub

PrivateSubComputeAverages()
OnErrorResumeNext

Dim sngCalcTempAv AsSingle
Dim sngCalcRHAv AsSingle
Dim sngMaxAv AsSingle
Dim sngMinAv AsSingle
Dim sngTav AsSingle
Dim sngSTD AsSingle
Dim i AsInteger
Dim j AsInteger
Dim intCountTemp AsInteger
Dim intCountRH AsInteger
Dim intValidHours AsInteger
Dim intAvOfAvCount AsInteger

Dim xlApp As Excel.Application = Nothing
Dim xlBooks As Excel.Workbooks = Nothing
Dim xlBook As Excel.Workbook = Nothing

Dim xlTempSheet As Excel.Worksheet = Nothing
Dim xlRHSheet As Excel.Worksheet = Nothing

Dim missing AsObject = Type.Missing

Dim intTempSheet AsInteger
Dim intRHSheet AsInteger

Dim sngTasterix AsSingle

        xlApp = NewExcel.Application()
        xlApp.DisplayAlerts = False
        xlApp.UserControl = True
        xlBooks = DirectCast(xlApp.Workbooks, Excel.Workbooks)

'Change full path of Excel file here:
        xlBook = DirectCast(xlBooks.Open(strFileName, True, False, missing, "", missing, False, missing, missing, True, missing, missing, missing, missing, missing), Excel.Workbook)

'worksheet for temp and RH from user textbox input
        intTempSheet = txtTempWorksheet.Text
        intRHSheet = txtRHWorksheet.Text

'get sheet in the workbook
        xlTempSheet = DirectCast(xlBook.Worksheets.Item(intTempSheet), Excel.Worksheet)
        xlRHSheet = DirectCast(xlBook.Worksheets.Item(intRHSheet), Excel.Worksheet)

CreateAveragesHeaders()

For j = 3 To 26 'loop thru excel sheet columns

'reset counters
            sngCalcTempAv = 0
            intCountTemp = 0
            sngCalcRHAv = 0
            intCountRH = 0

'loop thru excel sheet rows
For i = 5 To483  'last

'get each temp value
Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

'verify that data is a temperature value
IfVerifiedData(xlSecondTempCell.Value2) Then

'count columns with valid data
                    intValidHours = j

'get corresponding RH
Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)

                    sngCalcTempAv = sngCalcTempAv + xlSecondTempCell.Value2
                    intCountTemp = intCountTemp + 1

                    sngCalcRHAv = sngCalcRHAv + xlRHCell.Value2
                    intCountRH = intCountRH + 1
EndIf

Next

If sngCalcTempAv <> 0 Then
dgvTempAverages.Rows.Item(0).Cells(j - 2).Value = Math.Round(sngCalcTempAv / intCountTemp, 1)
dgvRHAverages.Rows.Item(0).Cells(j - 2).Value = Math.Round(sngCalcRHAv / intCountRH, 1)
Else
dgvTempAverages.Rows.Item(0).Cells(j - 2).Value = ""
dgvRHAverages.Rows.Item(0).Cells(j - 2).Value = ""
EndIf
Next

        sngCalcTempAv = 0
        sngMinAv = 10000
        sngMaxAv = 0

        intAvOfAvCount = 0

For j = 1 To 24
'find max average (ensure cell is not empty)
IfdgvTempAverages.Rows.Item(0).Cells(j).Value <>""Then
IfdgvTempAverages.Rows.Item(0).Cells(j).Value > sngMaxAv Then
                    sngMaxAv = dgvTempAverages.Rows.Item(0).Cells(j).Value
EndIf

'find min average
IfdgvTempAverages.Rows.Item(0).Cells(j).Value < sngMinAv Then
                    sngMinAv = dgvTempAverages.Rows.Item(0).Cells(j).Value
EndIf

                intAvOfAvCount = intAvOfAvCount + 1

'sum the averages
                sngCalcTempAv = sngCalcTempAv + dgvTempAverages.Rows.Item(0).Cells(j).Value
EndIf
Next

'get average of averages
        sngCalcTempAv = Math.Round(sngCalcTempAv / intAvOfAvCount, 1)

        txtMaxTempAverage.Text = sngMaxAv
        txtMinTempAverage.Text = sngMinAv
        txtTempRange.Text = Math.Round(sngMaxAv - sngMinAv, 1)
        txtTav.Text = sngCalcTempAv

For j = 1 To 24 'columns

'ensure cells are not empty
IfIsNothing(dgvTempAverages.Rows.Item(0).Cells(j).Value) = FalseThen
'calculate hourly temp
                sngTasterix = (2 * Math.PI * (j - 1)) / 24
dgvTempAverages.Rows.Item(1).Cells(j).Value = Math.Round(CalculateHourlyTemp(sngCalcTempAv, sngMaxAv, sngMinAv, sngTasterix), 1)

EndIf
Next

        intValidHours = 0
        sngSTD = 0
For j = 3 To 26 'loop thru excel sheet columns
For i = 5 To483  'last

'get each temp value
Dim xlSecondTempCell As Excel.Range = DirectCast(xlTempSheet.Cells(i, j), Excel.Range)

'verify that data is a temperature value
IfVerifiedData(xlSecondTempCell.Value2) Then

'count columns with valid data
                    intValidHours = intValidHours + 1

                    sngSTD = sngSTD + Math.Pow(xlSecondTempCell.Value2 - CSng(txtTav.Text), 2)
EndIf

Next
Next

        txtTStandardDeviation.Text = Math.Round(Math.Sqrt(sngSTD / intValidHours), 1)

'***********************************************RH


        sngCalcRHAv = 0
        sngMinAv = 10000
        sngMaxAv = 0

        intAvOfAvCount = 0

For j = 1 To 24
IfdgvRHAverages.Rows.Item(0).Cells(j).Value <>""Then
'find max average
IfdgvRHAverages.Rows.Item(0).Cells(j).Value > sngMaxAv Then
                    sngMaxAv = dgvRHAverages.Rows.Item(0).Cells(j).Value
EndIf

'find min average
IfdgvRHAverages.Rows.Item(0).Cells(j).Value < sngMinAv Then
                    sngMinAv = dgvRHAverages.Rows.Item(0).Cells(j).Value
EndIf

                intAvOfAvCount = intAvOfAvCount + 1

'sum averages
                sngCalcRHAv = sngCalcRHAv + dgvRHAverages.Rows.Item(0).Cells(j).Value
EndIf
Next

        sngCalcRHAv = Math.Round(sngCalcRHAv / intAvOfAvCount, 2)

        txtMaxRHAverage.Text = sngMaxAv
        txtMinRHAverage.Text = sngMinAv
        txtRHRange.Text = Math.Round(sngMaxAv - sngMinAv, 1)
        txtRHav.Text = Math.Round(sngCalcRHAv, 1)

For j = 1 To 24 'columns

'ensure cell is not empty
IfIsNothing(dgvRHAverages.Rows.Item(0).Cells(j).Value) = FalseThen

'calculate hourly RH
                sngTasterix = (2 * Math.PI * (j - 1)) / 24
dgvRHAverages.Rows.Item(1).Cells(j).Value = Math.Round(CalculateHourlyRH(sngCalcRHAv, sngMaxAv, sngMinAv, sngTasterix), 1)
EndIf
Next

        intValidHours = 0
        sngSTD = 0
For j = 3 To 26 'loop thru excel sheet columns
For i = 5 To483  'last

'get each temp value
Dim xlRHCell As Excel.Range = DirectCast(xlRHSheet.Cells(i, j), Excel.Range)

'verify that data is an rh value
IfVerifiedData(xlRHCell.Value2) Then

'count columns with valid data
                    intValidHours = intValidHours + 1

                    sngSTD = sngSTD + Math.Pow(xlRHCell.Value2 - CSng(txtRHav.Text), 2)
EndIf

Next
Next

        txtRHStandardDeviation.Text = Math.Round(Math.Sqrt(sngSTD / intValidHours), 1)

EndSub

PublicSubreleaseObject(ByVal obj AsObject)
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
EndTry
EndSub

PrivateFunctionCheckTempRangeOfData(ByVal TheTempData AsSingle) AsSingle
OnErrorResumeNext

IfStrings.InStr(TheTempData, ".") > 0 Then'ensure decimal point exists
'check for range
'return whole number part of figure, cut off decimal point and decimal part
            CheckTempRangeOfData = CSng(Strings.Left(TheTempData, Strings.InStr(TheTempData, ".") - 1))
Else
'check if data is without decimal, return whole data
'If CheckTempRangeOfData = 0 Then
            CheckTempRangeOfData = TheTempData
EndIf

EndFunction

PrivateFunctionCheckRHRangeOfData(ByVal TheRHData AsSingle) AsInteger
OnErrorResumeNext

'check for range within which data falls
'divide by 5 to get vertical position, round up if with fraction
'if not fraction, add 1 to result (to effect, for instance, values from 60-64.9 being counted instead of 60-65. 
'65 is then included in range 65-70)
IfStrings.InStr((TheRHData / 5), ".") > 0 Then'ensure decimal point exists
IfStrings.Right((TheRHData / 5), Strings.InStr((TheRHData / 5), ".") + 1) > 0 Then'check if fraction exists
                CheckRHRangeOfData = Strings.Left((TheRHData / 5), Strings.InStr((TheRHData / 5), ".") - 1) + 1 'round up integer part
Else
                CheckRHRangeOfData = (TheRHData / 5) + 1 'no decimal, i.e. boarder line figure, thus add 1 to upgrade to next range
EndIf
Else
            CheckRHRangeOfData = (TheRHData / 5) + 1 'no decimal, i.e. boarder line figure, thus add 1 to upgrade to next range
EndIf

EndFunction

PrivateSubCreateCoincidentTempRHProbabilityHeaders()
OnErrorResumeNext

Dim i AsInteger
Dim intColumns AsInteger
Dim intRH AsInteger

ClearProbabilityTable()

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then
With dgvRangeProbability
                .Columns.Add("Temp", "Temp")
                .Width = 1240
                .Height = 486
                .Top = 80

                .Columns(0).Width = 41
                .RowHeadersWidth = 30

For i = 1 To 50 'range starts at 1-2, ends at 49-50
                    .Columns.Add("Therange", i &"-"& i + 1)

                    intColumns = intColumns + 1
                    .Columns(intColumns).Width = 64
Next


                .Rows.Add()
                .Rows.Item(0).Cells(0).Value = "RH"
For i = 1 To 21
                    .Rows.Add()
                    .Rows.Item(i).Cells(0).Value = intRH &"-"& intRH + 5 'range starts at 0-5, ends at 100-105
                    intRH = intRH + 5
Next

EndWith
Else
With dgvRangeProbabilityPerYear
                .Columns.Add("Temp", "Temp")
                .Width = 1240
                .Height = 486
                .Top = 80

                .Columns(0).Width = 41
                .RowHeadersWidth = 30

For i = 1 To 50 'range starts at 1-2, ends at 49-50
                    .Columns.Add("Therange", i &"-"& i + 1)

                    intColumns = intColumns + 1
                    .Columns(intColumns).Width = 64
Next


                .Rows.Add()
                .Rows.Item(0).Cells(0).Value = "RH"
For i = 1 To 21
                    .Rows.Add()
                    .Rows.Item(i).Cells(0).Value = intRH &"-"& intRH + 5 'range starts at 0-5, ends at 100-105
                    intRH = intRH + 5
Next

EndWith
EndIf

EndSub

PrivateSubCreateCoincidentTempRHHeaders()
OnErrorResumeNext

Dim i AsInteger
Dim intColumns AsInteger
Dim intRH AsInteger

ClearCountTable()

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then
With dgvRangeCategories
                .Columns.Add("Temp", "Temp")
                .Width = 1210
                .Height = 500
                .Top = 80

                .Columns(0).Width = 41
                .RowHeadersWidth = 30

For i = 1 To 50 'range starts at 12-13, ends at 39-40
                    .Columns.Add("Therange", i &"-"& i + 1)

                    intColumns = intColumns + 1
                    .Columns(intColumns).Width = 40
Next


                .Rows.Add()
                .Rows.Item(0).Cells(0).Value = "RH"
For i = 1 To 21
                    .Rows.Add()
                    .Rows.Item(i).Cells(0).Value = intRH &"-"& intRH + 5 'range starts at 0-5, ends at 100-105
                    intRH = intRH + 5
Next

EndWith
Else
With dgvRangeCategoriesPerYear
                .Columns.Add("Temp", "Temp")
                .Width = 1210
                .Height = 500
                .Top = 80

                .Columns(0).Width = 41
                .RowHeadersWidth = 30

For i = 1 To 50 'range starts at 12-13, ends at 39-40
                    .Columns.Add("Therange", i &"-"& i + 1)

                    intColumns = intColumns + 1
                    .Columns(intColumns).Width = 40
Next


                .Rows.Add()
                .Rows.Item(0).Cells(0).Value = "RH"
For i = 1 To 21
                    .Rows.Add()
                    .Rows.Item(i).Cells(0).Value = intRH &"-"& intRH + 5 'range starts at 0-5, ends at 100-105
                    intRH = intRH + 5
Next

EndWith
EndIf
EndSub

PrivateSubCreateTempProbabilityHeaders()
OnErrorResumeNext

Dim intColumns AsInteger

ClearTempProbabilityTable()

With dgvTempProbability

            .Columns.Add("Temp", "Temp")
            .Width = 1240
            .Height = 100

            .Columns(0).Width = 72

For i = 1 To 50 'range starts at 1-2, ends at 49-50
                .Columns.Add("Therange", i &"-"& i + 1)

                intColumns = intColumns + 1

                .Columns(intColumns).Width = 64
Next

            .Columns.Item(0).HeaderText = "Temp. Bin Data"

            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Freq."

            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Probability"

EndWith

EndSub

PrivateSubCreateRHProbabilityHeaders()
OnErrorResumeNext

Dim i AsInteger
Dim intColumns AsInteger
Dim intRH AsInteger

ClearRHProbabilityTable()

With dgvRHProbability
            .Columns.Add("RH", "RH")
            .Width = 1248
            .Height = 90

            .Columns(0).Width = 94

            intColumns = 0 'reset counter

For i = 1 To 20
                .Columns.Add("Therange", intRH + 1 &"-"& intRH + 5)
                intRH = intRH + 5
                intColumns = intColumns + 1

                .Columns(intColumns).Width = 55

Next

            .Columns.Item(0).HeaderText = "RH Bin Data"

            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Freq."

            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Probability"

EndWith
EndSub

PrivateSubClearCoincidentTable()
OnErrorResumeNext

With dgvCoincidentRH
            .RowCount = 1
            .ColumnCount = 0
EndWith
EndSub

PrivateSubClearFourierTables()

OnErrorResumeNext

With dgvFourierTh
            .RowCount = 1
            .ColumnCount = 0
EndWith

With dgvFourierRh
            .RowCount = 1
            .ColumnCount = 0
EndWith

EndSub


PrivateSubClearAveragesTables()
OnErrorResumeNext

With dgvTempAverages
            .RowCount = 1
            .ColumnCount = 0
EndWith

With dgvRHAverages
            .RowCount = 1
            .ColumnCount = 0
EndWith
EndSub

PrivateSubClearTempProbabilityTable()
OnErrorResumeNext

With dgvTempProbability

            .RowCount = 1
            .ColumnCount = 0
EndWith

EndSub

PrivateSubClearRHProbabilityTable()
OnErrorResumeNext

Dim i AsInteger

With dgvRHProbability

            .RowCount = 1
            .ColumnCount = 0
EndWith
EndSub

PrivateSubClearProbabilityTable()
OnErrorResumeNext

Dim i AsInteger

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then
With dgvRangeProbability

                .RowCount = 1
                .ColumnCount = 0
EndWith
Else
With dgvRangeProbabilityPerYear

                .RowCount = 1
                .ColumnCount = 0
EndWith
EndIf
EndSub

PrivateSubClearCountTable()
OnErrorResumeNext

Dim i AsInteger

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then
With dgvRangeCategories

                .RowCount = 1
                .ColumnCount = 0
EndWith
Else
With dgvRangeCategoriesPerYear

                .RowCount = 1
                .ColumnCount = 0
EndWith
EndIf

EndSub

PrivateSubTransferToWorkSheet()
OnErrorResumeNext

Dim wkbTransfer AsNew Microsoft.Office.Interop.Excel.Workbook
Dim wksTransfer AsNew Microsoft.Office.Interop.Excel.Worksheet
Dim wksTransferProbability AsNew Microsoft.Office.Interop.Excel.Worksheet

Dim intCountColumns AsInteger
Dim intCountRows AsInteger
Dim intRow AsInteger
Dim intColumn AsInteger
Dim strDisplayFile AsString

        strDisplayFile = "C:\Users\Public\xxRelativeHumidity\RelativeHumidityData.xlsx"

        intSheet = intSheet + 1

'check if it's first creation
If intSheet = 1 Then
            wkbTransfer = xlTransfer.Workbooks.Open(strDisplayFile) ', , True
            objWrkBk = wkbTransfer

            wksTransfer = wkbTransfer.Sheets.Item("sheet"& intSheet)

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Or txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then
                intSheet = intSheet + 1
                wksTransferProbability = wkbTransfer.Sheets.Item("sheet"& intSheet)
EndIf

            objWrkSheet1 = wksTransfer
            objWrkSheet2 = wksTransferProbability
Else
            objWrkSheet1 = objWrkBk.Sheets.Item("sheet"& intSheet)

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Or txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then
                intSheet = intSheet + 1
                objWrkSheet2 = objWrkBk.Sheets.Item("sheet"& intSheet)
EndIf
EndIf

'get the number of columns involved
If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then
            intCountColumns = dgvRangeCategories.ColumnCount
            intCountRows = dgvRangeCategories.RowCount
ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then
            intCountColumns = dgvRangeCategoriesPerYear.ColumnCount
            intCountRows = dgvRangeCategoriesPerYear.RowCount
ElseIf txtTotalHDD.SelectedTab.Name = "tpgCoincidentRHs"Then
            intCountColumns = dgvCoincidentRH.ColumnCount
            intCountRows = dgvCoincidentRH.RowCount
ElseIf txtTotalHDD.SelectedTab.Name = "tpgAverages"Then
            intCountColumns = dgvTempAverages.ColumnCount
            intCountRows = dgvTempAverages.RowCount
ElseIf txtTotalHDD.SelectedTab.Name = "tpgFourierSeries"Then
            intCountColumns = dgvFourierTh.ColumnCount
            intCountRows = dgvFourierTh.RowCount
ElseIf txtTotalHDD.SelectedTab.Name = "tpgBinData"Then
            intCountColumns = dgvTempProbability.ColumnCount
            intCountRows = dgvTempProbability.RowCount
EndIf

With objWrkSheet1
'clear the worksheet
            .Range("1:"& intCountColumns, "1:"& .Rows.Count).Clear()

'copy each column header
For intColumn = 0 To intCountColumns

If txtTotalHDD.SelectedTab.Name = "tpgCoincidentRHs"Then
                    .Cells(1, intColumn + 1) = CStr(dgvCoincidentRH.Columns.Item(intColumn).HeaderCell.Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then
                    .Cells(1, intColumn + 1) = CStr(dgvRangeCategories.Columns.Item(intColumn).HeaderCell.Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then
                    .Cells(1, intColumn + 1) = CStr(dgvRangeCategoriesPerYear.Columns.Item(intColumn).HeaderCell.Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgAverages"Then
                    .Cells(1, intColumn + 1) = CStr(dgvTempAverages.Columns.Item(intColumn).HeaderCell.Value)
                    .Cells(13, intColumn + 1) = CStr(dgvRHAverages.Columns.Item(intColumn).HeaderCell.Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgFourierSeries"Then
                    .Cells(1, intColumn + 1) = CStr(dgvFourierTh.Columns.Item(intColumn).HeaderCell.Value)
                    .Cells(13, intColumn + 1) = CStr(dgvFourierRh.Columns.Item(intColumn).HeaderCell.Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgBinData"Then
                    .Cells(1, intColumn + 1) = CStr(dgvTempProbability.Columns.Item(intColumn).HeaderCell.Value)
                    .Cells(11, intColumn + 1) = CStr(dgvRHProbability.Columns.Item(intColumn).HeaderCell.Value)
EndIf

'copy rows
For intRow = 0 To intCountRows
If txtTotalHDD.SelectedTab.Name = "tpgCoincidentRHs"Then

                        .Cells(intRow + 2, intColumn + 1) = CStr(dgvCoincidentRH.Rows.Item(intRow).Cells(intColumn).Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then

                        .Cells(intRow + 2, intColumn + 1) = CStr(dgvRangeCategories.Rows.Item(intRow).Cells(intColumn).Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then

                        .Cells(intRow + 2, intColumn + 1) = CStr(dgvRangeCategoriesPerYear.Rows.Item(intRow).Cells(intColumn).Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgAverages"Then

                        .Cells(intRow + 2, intColumn + 1) = CStr(dgvTempAverages.Rows.Item(intRow).Cells(intColumn).Value)

                        .Cells(intRow + 14, intColumn + 1) = CStr(dgvRHAverages.Rows.Item(intRow).Cells(intColumn).Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgFourierSeries"Then

                        .Cells(intRow + 2, intColumn + 1) = CStr(dgvFourierTh.Rows.Item(intRow).Cells(intColumn).Value)

                        .Cells(intRow + 14, intColumn + 1) = CStr(dgvFourierRh.Rows.Item(intRow).Cells(intColumn).Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgBinData"Then

                        .Cells(intRow + 2, intColumn + 1) = CStr(dgvTempProbability.Rows.Item(intRow).Cells(intColumn).Value)

                        .Cells(intRow + 12, intColumn + 1) = CStr(dgvRHProbability.Rows.Item(intRow).Cells(intColumn).Value)
EndIf
Next

Next

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Or txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then
With objWrkSheet2
'clear the worksheet
                    .Range("1:"& intCountColumns, "1:"& .Rows.Count).Clear()

'copy each column header
For intColumn = 0 To intCountColumns

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then
                            .Cells(1, intColumn + 1) = CStr(dgvRangeProbability.Columns.Item(intColumn).HeaderCell.Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then
                            .Cells(1, intColumn + 1) = CStr(dgvRangeProbabilityPerYear.Columns.Item(intColumn).HeaderCell.Value)
EndIf

'copy rows
For intRow = 0 To intCountRows

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then

                                .Cells(intRow + 2, intColumn + 1) = CStr(dgvRangeProbability.Rows.Item(intRow).Cells(intColumn).Value)

ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then

                                .Cells(intRow + 2, intColumn + 1) = CStr(dgvRangeProbabilityPerYear.Rows.Item(intRow).Cells(intColumn).Value)

EndIf
Next

Next
EndWith
EndIf

If txtTotalHDD.SelectedTab.Name = "tpgAverages"Then
                .Cells(intRow + 2, 1) = lblTmax.Text
                .Cells(intRow + 3, 1) = lblTmin.Text
                .Cells(intRow + 4, 1) = lblRange.Text
                .Cells(intRow + 5, 1) = lblTav.Text
                .Cells(intRow + 6, 1) = lblTSTD.Text

                .Cells(intRow + 2, 2) = txtMaxTempAverage.Text
                .Cells(intRow + 3, 2) = txtMinTempAverage.Text
                .Cells(intRow + 4, 2) = txtTempRange.Text
                .Cells(intRow + 5, 2) = txtTav.Text
                .Cells(intRow + 6, 2) = txtTStandardDeviation.Text

'************************************

                .Cells(intRow + 14, 1) = lblRHmax.Text
                .Cells(intRow + 15, 1) = lblRHMin.Text
                .Cells(intRow + 16, 1) = lblRHRange.Text
                .Cells(intRow + 17, 1) = lblRHav.Text
                .Cells(intRow + 18, 1) = lblRHSTD.Text

                .Cells(intRow + 14, 2) = txtMaxRHAverage.Text
                .Cells(intRow + 15, 2) = txtMinRHAverage.Text
                .Cells(intRow + 16, 2) = txtRHRange.Text
                .Cells(intRow + 17, 2) = txtRHav.Text
                .Cells(intRow + 18, 2) = txtRHStandardDeviation.Text

                .Range("A:A").ColumnWidth = 15
ElseIf txtTotalHDD.SelectedTab.Name = "tpgFourierSeries"Then

                .Range("A:A").ColumnWidth = 15

ElseIf txtTotalHDD.SelectedTab.Name = "tpgBinData"Then

                .Range("A:A").ColumnWidth = 18

ElseIf txtTotalHDD.SelectedTab.Name = "tpgCoolingDegreeDays"Then

                .Cells(intRow + 2, 2) = txtModelAvgMonthlyCDD.Text

                .Cells(intRow + 2, 5) = txtActualTotalCDD.Text
                .Cells(intRow + 3, 5) = txtActualAvgMonthlyCDD.Text

                .Range("A:A").ColumnWidth = 18
                .Range("D:D").ColumnWidth = 18

                .Cells(intRow + 1, 1) = grpModelCDD.Text
                .Cells(intRow + 2, 1) = lblModelAvgMonthlyCDD.Text

                .Cells(intRow + 1, 4) = grpActualCDD.Text
                .Cells(intRow + 2, 4) = lblActualTotalCDD.Text
                .Cells(intRow + 3, 4) = lblActualAvgMonthlyCDD.Text

ElseIf txtTotalHDD.SelectedTab.Name = "tpgHeatingDegreeDays"Then

                .Cells(intRow + 2, 2) = txtModelAvgMonthlyHDD.Text

                .Cells(intRow + 2, 5) = txtActualTotalHDD.Text
                .Cells(intRow + 3, 5) = txtActualAvgMonthlyHDD.Text

                .Range("A:A").ColumnWidth = 18
                .Range("D:D").ColumnWidth = 18

                .Cells(intRow + 1, 1) = grpModelHDD.Text
                .Cells(intRow + 2, 1) = lblModelAvgMonthlyHDD.Text

                .Cells(intRow + 1, 4) = grpActualHDD.Text
                .Cells(intRow + 2, 4) = lblActualTotalHDD.Text
                .Cells(intRow + 3, 4) = lblActualAvgMonthlyHDD.Text

EndIf
EndWith

objWrkBk.SaveAs(strDisplayFile)

        xlTransfer.Visible = True

EndSub

PrivateFunctionVerifiedWorksheets() AsBoolean
OnErrorResumeNext

If txtTempWorksheet.Text <>""Then
If txtRHWorksheet.Text <>""Then
IfIsNumeric(txtTempWorksheet.Text) = TrueAnd IsNumeric(txtRHWorksheet.Text) = TrueThen
                    VerifiedWorksheets = True
EndIf
EndIf
EndIf

EndFunction

PublicSubGetDatabasePath()
OnErrorGoTo err_handler 'if error occurs, display error handler message

ofdDatabasePath.ShowDialog()
        txtExcelFile.Text = ofdDatabasePath.FileName
        strFileName = txtExcelFile.Text

Exit Sub

err_handler:
MsgBox("The selected data path is invalid, please input a valid data path.", MsgBoxStyle.Critical + vbOKOnly, "")

Me.Cursor = Cursors.Default
Exit Sub

EndSub

PrivateSubCreateFourierHeaders()
OnErrorResumeNext

Dim i AsInteger

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
EndIf

Next

'create rows for table
            .Columns.Item(0).HeaderText = "Hour"

            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Hourly Th"

'.Rows.Add()
'.Rows.Item(1).Cells(0).Value = "(Th-Tav)/Range"

            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Fourier Temp"

EndWith


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
EndIf

Next

'create rows for table
            .Columns.Item(0).HeaderText = "Hour"

            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Hourly Rh"

            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Fourier RH"

EndWith

EndSub

PrivateSubCreateAveragesHeaders()
OnErrorResumeNext

Dim i AsInteger

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
EndIf

Next

'create rows for table
            .Columns.Item(0).HeaderText = "Hour"

            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "Th"

            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Hourly Temp"
EndWith


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
EndIf

Next

'create rows for table
            .Columns.Item(0).HeaderText = "Hour"

            .Rows.Add()
            .Rows.Item(0).Cells(0).Value = "RHh"

            .Rows.Add()
            .Rows.Item(1).Cells(0).Value = "Hourly RH"
EndWith

EndSub

PrivateSubComputeHeatingDegreeDays()
OnErrorResumeNext

Dim sngHeatingHb AsSingle

        sngHeatingHb = (CSng(txtBaseTempModelHDD.Text) - CSng(txtTavHDD.Text)) / (CSng(txtSIGMAmModelHDD.Text) * Math.Sqrt(CSng(cboNumofDaysInMonthHDD.Text)))

        txtModelAvgMonthlyHDD.Text = Math.Round((CSng(txtSIGMAmModelHDD.Text) * Math.Pow(CSng(cboNumofDaysInMonthHDD.Text), (3 / 2))) * (0.072196 + (sngHeatingHb / 2) + (1 / 9.6) * Math.Log(Math.Cosh(4.8 * sngHeatingHb))), 2)
EndSub

PrivateSubComputeCoolingDegreeDays()
OnErrorResumeNext

Dim sngCoolingHb AsSingle

        sngCoolingHb = (CSng(txtTavCDD.Text) - CSng(txtBaseTempModelCDD.Text)) / (CSng(txtSIGMAmModelCDD.Text) * Math.Sqrt(CSng(cboNumofDaysInMonthCDD.Text)))

        txtModelAvgMonthlyCDD.Text = Math.Round((CSng(txtSIGMAmModelCDD.Text) * Math.Pow(CSng(cboNumofDaysInMonthCDD.Text), (3 / 2))) * (0.072196 + (sngCoolingHb / 2) + (1 / 9.6) * Math.Log(Math.Cosh(4.8 * sngCoolingHb))), 1)
EndSub

PrivateFunctionVerifyInputs() AsBoolean
OnErrorResumeNext

'make usre inputs are valid
If txtTotalHDD.SelectedTab.Name = "tpgHeatingDegreeDays"Then
If txtBaseTempModelHDD.Text <>""And txtSIGMAmModelHDD.Text <>""And cboNumofDaysInMonthHDD.Text <>""And txtNumofYearsHDD.Text <>""And txtBaseTempActualHDD.Text <>""Then
IfIsNumeric(txtBaseTempModelHDD.Text) = TrueAnd IsNumeric(txtSIGMAmModelHDD.Text) = TrueAnd IsNumeric(txtNumofYearsHDD.Text) = TrueAnd IsNumeric(txtBaseTempActualHDD.Text) = TrueThen
                    VerifyInputs = True
Else
                    VerifyInputs = False
EndIf
Else
                VerifyInputs = False
EndIf
ElseIf txtTotalHDD.SelectedTab.Name = "tpgCoolingDegreeDays"Then
If txtBaseTempModelCDD.Text <>""And txtSIGMAmModelCDD.Text <>""And cboNumofDaysInMonthCDD.Text <>""And txtNumofYearsCDD.Text <>""And txtBaseTempActualCDD.Text <>""Then
IfIsNumeric(txtBaseTempModelCDD.Text) = TrueAnd IsNumeric(txtSIGMAmModelCDD.Text) = TrueAnd IsNumeric(txtNumofYearsCDD.Text) = TrueAnd IsNumeric(txtBaseTempActualCDD.Text) = TrueThen
                    VerifyInputs = True
Else
                    VerifyInputs = False
EndIf
Else
                VerifyInputs = False
EndIf
EndIf

EndFunction

PrivateFunctionVerifyFourierInputs() AsBoolean
OnErrorResumeNext

'make usre inputs are valid
If txtNewFourierTav.Text <>""And txtNewFourierTrange.Text <>""And txtNewFourierRHav.Text <>""And txtNewFourierRHrange.Text <>""Then

IfInStr(txtNewFourierTav.Text, " ") = 0 And InStr(txtNewFourierTrange.Text, " ") = 0 And InStr(txtNewFourierRHav.Text, " ") = 0 And InStr(txtNewFourierRHrange.Text, " ") = 0 Then

If IsNumeric(CSng(txtNewFourierTav.Text)) = TrueAnd IsNumeric(CSng(txtNewFourierTrange.Text)) = TrueAnd IsNumeric(CSng(txtNewFourierRHav.Text)) = TrueAnd IsNumeric(CSng(txtNewFourierRHrange.Text)) = TrueThen
                    VerifyFourierInputs = True
Else
                    VerifyFourierInputs = False
EndIf
Else
                VerifyFourierInputs = False
EndIf
Else
            VerifyFourierInputs = False
EndIf
EndFunction

PrivateSubClearFourierBoxes()
OnErrorResumeNext

EndSub

PrivateSubClearHeatingDegreeBoxes()
OnErrorResumeNext

        txtModelAvgMonthlyHDD.Text = ""
        txtActualAvgMonthlyHDD.Text = ""
        txtActualTotalHDD.Text = ""
EndSub

PrivateSubClearCoolingDegreeBoxes()
OnErrorResumeNext

        txtModelAvgMonthlyCDD.Text = ""
        txtActualTotalCDD.Text = ""
        txtActualAvgMonthlyCDD.Text = ""
EndSub

PrivateSubClearBoxes()
OnErrorResumeNext

'clear user input boxes for new entry

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
EndSub

PrivateSub frmRelativeHumidity_Load(ByVal sender AsObject, ByVal e As System.EventArgs) HandlesMe.Load
OnErrorResumeNext

Me.Cursor = Cursors.WaitCursor

CreateAveragesHeaders()

CreateTempProbabilityHeaders()
CreateRHProbabilityHeaders()

Me.Cursor = Cursors.Default
EndSub

PrivateSub btnRangeCategoriesPerMonth_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles btnRangeCategoriesPerMonth.Click
OnErrorResumeNext

Me.Cursor = Cursors.WaitCursor

IfVerifiedWorksheets() = FalseThen
MsgBox("Input data is invalid", vbOKOnly + MsgBoxStyle.Exclamation, "")
Me.Cursor = Cursors.Default
GoTo exit_sub
EndIf

CreateCoincidentTempRHHeaders()
CreateCoincidentTempRHProbabilityHeaders()

If txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerMonth"Then
FreqOfRHInTempRange(dgvRangeCategories, dgvRangeProbability, 483)
ElseIf txtTotalHDD.SelectedTab.Name = "tpgRangeCategoriesPerYear"Then
FreqOfRHInTempRange(dgvRangeCategoriesPerYear, dgvRangeProbabilityPerYear, 5800)
EndIf

TransferToWorkSheet()

Me.Cursor = Cursors.Default

exit_sub:
Exit Sub
EndSub

PrivateSub txtTempWorksheet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTempWorksheet.TextChanged
OnErrorResumeNext

IfVerifiedWorksheets() = TrueAnd txtExcelFile.Text <>""Then
ClearCountTable()
ClearCoincidentTable()

            txtTotalHDD.Enabled = True
Else
            txtTotalHDD.Enabled = False
EndIf

EndSub

PrivateSub txtExcelFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExcelFile.TextChanged
OnErrorResumeNext

ClearCountTable()
ClearCoincidentTable()

If txtExcelFile.Text <>""AndVerifiedWorksheets() = TrueThen
            txtTotalHDD.Enabled = True

'simulate selection of basic data radio button
            rbtBasicDataPerMonth.Checked = True
            rbtBasicDataPerYear.Checked = True
Else
            txtTotalHDD.Enabled = False
EndIf
EndSub

PrivateSub txtRHWorksheet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRHWorksheet.TextChanged
OnErrorResumeNext

IfVerifiedWorksheets() = TrueAnd txtExcelFile.Text <>""Then
ClearCountTable()
ClearCoincidentTable()

            txtTotalHDD.Enabled = True
Else
            txtTotalHDD.Enabled = False
EndIf

EndSub

PrivateSub btnExcelFile_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles btnExcelFile.Click
OnErrorResumeNext

GetDatabasePath()
EndSub

PrivateSub btnCoincidentRh_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles btnCoincidentRh.Click
OnErrorResumeNext

Me.Cursor = Cursors.WaitCursor

        dgvCoincidentRH.Height = 486

CoincidentRHs()

TransferToWorkSheet()

Me.Cursor = Cursors.Default
EndSub

PrivateSub btnAverages_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles btnAverages.Click
OnErrorResumeNext

Me.Cursor = Cursors.WaitCursor

ClearBoxes()

'If VerifyInputs() = True Then

ComputeAverages()

TransferToWorkSheet()

Me.Cursor = Cursors.Default
EndSub

PrivateSub btnTempProbability_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles btnProbability.Click
OnErrorResumeNext

Me.Cursor = Cursors.WaitCursor

IfVerifiedWorksheets() = FalseThen
MsgBox("Input data is invalid", vbOKOnly + MsgBoxStyle.Exclamation, "")
Me.Cursor = Cursors.Default
GoTo exit_sub
EndIf

FreqOfTempInTempRange()
FreqOfRHInRHRange()

TransferToWorkSheet()

Me.Cursor = Cursors.Default

exit_sub:
Exit Sub
EndSub

PrivateSub tpgRangeCategories_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles tpgRangeCategoriesPerMonth.Click
OnErrorResumeNext

'simulate selection of basic data radio button
        rbtBasicData_Click(sender, e)
EndSub

PrivateSub rbtBasicData_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles rbtBasicDataPerMonth.Click
OnErrorResumeNext

If rbtBasicDataPerMonth.Checked = TrueThen
            dgvRangeCategories.Visible = True
            dgvRangeProbability.Visible = False
EndIf
EndSub

PrivateSub rbtProbability_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles rbtProbabilityPerMonth.Click
OnErrorResumeNext

If rbtProbabilityPerMonth.Checked = TrueThen
            dgvRangeProbability.Visible = True
            dgvRangeCategories.Visible = False
EndIf
EndSub

PrivateSub btnRangeCategoriesPerYear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRangeCategoriesPerYear.Click
OnErrorResumeNext

        btnRangeCategoriesPerMonth_Click(sender, e)
EndSub

PrivateSub rbtBasicDataPerYear_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles rbtBasicDataPerYear.Click
OnErrorResumeNext

If rbtBasicDataPerYear.Checked = TrueThen
            dgvRangeCategoriesPerYear.Visible = True
            dgvRangeProbabilityPerYear.Visible = False
EndIf
EndSub

PrivateSub rbtProbabilityPerYear_Click(ByVal sender AsObject, ByVal e As System.EventArgs) Handles rbtProbabilityPerYear.Click
OnErrorResumeNext

If rbtProbabilityPerYear.Checked = TrueThen
            dgvRangeProbabilityPerYear.Visible = True
            dgvRangeCategoriesPerYear.Visible = False
EndIf
EndSub

PrivateSub btnFourier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFourier.Click
OnErrorResumeNext

Me.Cursor = Cursors.WaitCursor


IfVerifyFourierInputs() = TrueThen

ComputeFourier()

TransferToWorkSheet()
Else
MsgBox("Input data is invalid, please input valid data. Ensure there are no spaces between figures, or between figures and decimal point.", vbOKOnly + MsgBoxStyle.Exclamation, "")
EndIf

Me.Cursor = Cursors.Default
EndSub

PrivateSub btnHeatingDegreeDays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHeatingDegreeDays.Click
OnErrorResumeNext

Me.Cursor = Cursors.WaitCursor

ClearHeatingDegreeBoxes()

IfVerifyInputs() = TrueThen

ComputeHeatingDegreeDays()
CalculateHDDActualDh()

TransferToWorkSheet()
Else
MsgBox("Input data is invalid, please input valid data.", vbOKOnly + MsgBoxStyle.Exclamation, "")
EndIf

Me.Cursor = Cursors.Default
EndSub

PrivateSub btnCoolingDegreeDays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCoolingDegreeDays.Click
OnErrorResumeNext

Me.Cursor = Cursors.WaitCursor

ClearCoolingDegreeBoxes()

IfVerifyInputs() = TrueThen

ComputeCoolingDegreeDays()
CalculateCDDActualDc()

TransferToWorkSheet()
Else
MsgBox("Input data is invalid, please input valid data.", vbOKOnly + MsgBoxStyle.Exclamation, "")
EndIf

Me.Cursor = Cursors.Default
EndSub
EndClass
