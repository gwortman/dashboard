'Refresh a given department's data and update dashboard
Private Sub Refresh(dept As String, fileName As String)
    'Disable nonessential processes
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    Dim todayDate As Date
    Dim thisShift As Integer
    Dim machineCount As Integer
    Dim dataCount As Integer
    Dim arrRow As Integer
    Dim equipData() As String
    'Number of DB columns - temporarily hardcoded
    dataCount = 22
    
    'Get the desired date and shift from the dashboard
    todayDate = Cells(1, 6)
    thisShift = Cells(1, 9)
    
    'Find the bounds of the desired department on the dashboard
    numRows = Cells(Rows.Count, "A").End(xlUp).Row 'Count total number of dashboard rows
    For thisRow = 1 To numRows
        If Cells(thisRow, 1) = dept Then 'found the correct department header
            firstRow = (thisRow + 1)     'data starts on the next row
            foundStart = True            'now look for the end
        ElseIf Cells(thisRow, 1) = "TOTALS" And foundStart = True Then 'found the end of the department
            lastRow = (thisRow - 1)      'data ended on previous row
            Exit For                     'aaaaand we're done
        End If
    Next thisRow
    
    'Calculate number of machines in the department and resize array
    machineCount = lastRow - firstRow + 1
    ReDim equipData(machineCount, dataCount)
    
    'Go get data from database
    equipData() = Copy_Data(fileName, todayDate, thisShift, machineCount, dataCount)
    
    'Print Current Data to Screen
    'arrRow = 0
    'For thisRow = firstRow To lastRow
    'Currently Running
    'Cells(thisRow, 4) = equipData(arrRow, 7)
    'Shift Plan
    'Cells(thisRow, 5) = equipData(arrRow, 8)
    'Shift Actual
        'If thisShift = 1 Then
            'Cells(thisRow, 6) = equipData(arrRow, 9)
        'ElseIf thisShift = 2 Then
            'Cells(thisRow, 7) = equipData(arrRow, 9)
        'ElseIf thisShift = 3 Then
            'Cells(thisRow, 8) = equipData(arrRow, 9)
        'End If
    'Downtime
    'Cells(thisRow, 12) = equipData(arrRow, 11)
    'Comments
    'Cells(thisRow, 14) = equipData(arrRow, 21)
    'arrRow = arrRow + 1
    'Next thisRow
    
    'Paste Part Number
    arrRow = 0
    For thisRow = firstRow To lastRow
        Cells(thisRow, 4) = equipData(arrRow, 7)
        arrRow = arrRow + 1
    Next thisRow
    
    'Paste Shift Plan
    arrRow = 0
    For thisRow = firstRow To lastRow
        Cells(thisRow, 5) = equipData(arrRow, 8)
        arrRow = arrRow + 1
    Next thisRow
    
    'Paste Actual Production
    arrRow = 0
    For thisRow = firstRow To lastRow
        If thisShift = 1 Then
            Cells(thisRow, 6) = equipData(arrRow, 9)
        ElseIf thisShift = 2 Then
            Cells(thisRow, 7) = equipData(arrRow, 9)
        ElseIf thisShift = 3 Then
            Cells(thisRow, 8) = equipData(arrRow, 9)
        End If
        arrRow = arrRow + 1
    Next thisRow
    
    'Paste Downtime
    arrRow = 0
    For thisRow = firstRow To lastRow
        Cells(thisRow, 12) = equipData(arrRow, 11)
        arrRow = arrRow + 1
    Next thisRow
    
    'Paste Comments
    arrRow = 0
    For thisRow = firstRow To lastRow
        Cells(thisRow, 14) = equipData(arrRow, 21)
        arrRow = arrRow + 1
    Next thisRow
    
    
    'Re-Enable screen updating
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
End Sub
'Pull data from database
Private Function Copy_Data(fileName As String, todayDate As Date, thisShift As Integer, machineCount As Integer, dataCount As Integer) As String()
    Dim pw As String
    Dim foundFirst As Boolean
    Dim noData As Boolean
    Dim equipData() As String
    Dim thisRow As Integer
    Dim arrayRow As Integer
    Dim arrayCol As Integer
    
    'Set variable starting values
    pw = "hanshotfirst"
    foundFirst = False
    thisRow = 2
    arrayRow = 0
    
    'Size array to match department
    ReDim equipData(machineCount, dataCount)
    
    'Open the data file and sort by date and then by shift
    Workbooks.Open fileName:=fileName, Password:=pw, IgnoreReadOnlyRecommended:=True
    Worksheets("Plan1").Columns("A:U").Sort key1:=Worksheets("Plan1").Columns("A:A"), order1:=xlDescending, key2:=Worksheets("Plan1").Columns("B:B"), order2:=xlDescending, Header:=xlYes
    
    'Seek to first line of desired data
    While Not (Worksheets("Plan1").Cells(thisRow, 1) = todayDate And Worksheets("Plan1").Cells(thisRow, 2) = thisShift)
        If Worksheets("Plan1").Cells(thisRow, 1) = "" Then
            noData = True
            GoTo Jump 'No data present so bypass the while loop's check
        End If
        thisRow = thisRow + 1
    Wend

    'Load array with data
    While (Worksheets("Plan1").Cells(thisRow, 1) = todayDate And Worksheets("Plan1").Cells(thisRow, 2) = thisShift)
Jump: 'If no data is present, jump directly here to load the top row of the array with "NO DATA"
        For arrayCol = 1 To dataCount
            If noData = False Then 'load the actual data
            equipData(arrayRow, arrayCol) = Worksheets("Plan1").Cells(thisRow, arrayCol)
            Else: equipData(arrayRow, arrayCol) = "NO DATA" 'no data, so load one row with the flag.  Loop won't repeat.
            End If
        Next arrayCol
        arrayRow = arrayRow + 1
        thisRow = thisRow + 1
    Wend
    
    'Close the database file
    ActiveWorkbook.Close SaveChanges:=False
    
    'Return
    Copy_Data = equipData()
    
End Function
'BUTTONS!
Private Sub CommandButton1_Click()
    Call Refresh("GMA", "W:\160216pd\GMA\GMA_Data.xlsx")
End Sub
Private Sub CommandButton2_Click()
    Call Refresh("Torsional", "W:\160216pd\TOR\TOR_Data.xlsx")
End Sub
Private Sub CommandButton3_Click()
    Call Refresh("MSBD", "W:\160216pd\MSBD\MSBD_Data.xlsx")
End Sub
Private Sub CommandButton4_Click()
    Call Refresh("LVNSK", "W:\160216pd\LVNSK\LVNSK_Data.xlsx")
End Sub
Private Sub CommandButton5_Click()
    Call Refresh("NSKN", "W:\160216pd\NSKN\NSKN_Data.xlsx")
End Sub
Private Sub CommandButton6_Click()
    Call Refresh("Housing", "W:\160216pd\HOUS\HOUS_Data.xlsx")
End Sub
Private Sub CommandButton7_Click()
    Call Refresh("High Volume", "W:\160216pd\HV\HV_Data.xlsx")
End Sub
Private Sub CommandButton8_Click()
    Call Refresh("NSKS", "W:\160216pd\NSKS\NSKS_Data.xlsx")
End Sub
Private Sub CommandButton9_Click()
    Call Refresh("Low Volume", "W:\160216pd\LV\LV_Data.xlsx")
End Sub
Private Sub CommandButton10_Click()
    Call Refresh("Steel Blanking/Sep Finishing", "W:\160216pd\SB\SB_Data.xlsx")
End Sub
Private Sub CommandButton11_Click()
    Call Refresh("Paper Blanking", "W:\160216pd\PB\PB_Data.xlsx")
End Sub
Private Sub CommandButton12_Click()
    Call Refresh("Acid Etch and Glue", "W:\160216pd\AEG\AEG_Data.xlsx")
End Sub
Private Sub CommandButton13_Click()
    Call Refresh("Saturation", "W:\160216pd\SAT\SAT_Data.xlsx")
End Sub
Private Sub CommandButton14_Click()
    Call CommandButton1_Click
    Call CommandButton2_Click
    Call CommandButton3_Click
    Call CommandButton4_Click
    Call CommandButton5_Click
    Call CommandButton6_Click
    Call CommandButton7_Click
    Call CommandButton8_Click
    Call CommandButton9_Click
    Call CommandButton10_Click
    Call CommandButton11_Click
    Call CommandButton12_Click
    Call CommandButton13_Click
End Sub

