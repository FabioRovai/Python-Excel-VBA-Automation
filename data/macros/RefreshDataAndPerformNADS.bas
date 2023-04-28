

Sub Run()

    StartRequestTable
    '
    Call togglebProcessClk
    Call ProcessRequestTable
    Call togglebProcessClk
    '
    Dim columnCount As Integer
    Dim myRange1 As Range
    Set myRange1 = Worksheets("tSpec").Range("A2:C30")
    Dim counter As Integer
    Dim tabName As String
    'Call BR3AK
    'Call DeleteTextInH3
    Call BR3AK
    columnCount = 3
    counter = 0
    Dim startTime As Date
    startTime = Now
    Dim endTime As Date
    endTime = startTime + TimeValue("00:10:00")
    'UPDATE BBG
    Call BBG_Update

    'PERFORM DS UPDATE FOR GIVEN CELLS
    For Each cell In myRange1

        If (counter = 0) Then
            tabName = cell.Value
            If (IsEmpty(cell.Value)) Then
                Debug.Print "cell value was empty"
            Else
                tabName = cell.Value
                Debug.Print "tblName = "
                Debug.Print tabName
            End If
        End If

        If (counter = 2) Then
            If (cell.Value = 1) Then

                Debug.Print "needsUpdate = "
                Debug.Print cell.Value
                Worksheets(tabName).Activate 'SET THE ACTIVE WORKSHEET TO THIS SO DS REFRESH WORKS
                Call DS_refresh1

            End If
        End If
        counter = counter + 1
        If (counter > 2) Then
            counter = 0
        End If

        If Now > endTime Then
            MsgBox "Time limit exceeded. Shutting down."
            Exit Sub
        End If

    Next cell
    Debug.Print "PERFORM #N/A FOR SELECTED TABS IF APPLICABLE"


   'PERFORM #N/A FOR SELECTED TABS IF APPLICABLE
    For Each cell In myRange1

        If (counter = 0) Then
            tabName = cell.Value

            If (IsEmpty(cell.Value)) Then
                Debug.Print "cell value was empty"
            Else
                Debug.Print "tblName = "
                Debug.Print cell.Value
            End If

        End If

        If (counter = 1) Then
            If (cell.Value = 1) Then
                'Worksheets(tabName).Activate 'SET THE ACTIVE WORKSHEET TO THIS SO DS REFRESH WORKS
                'Call DS_refresh1
                Debug.Print "n/a will be refreshed"
                Debug.Print cell.Value
                Call removeNotAvailable(tabName)
            End If
            'counter = 0
        End If

        counter = counter + 1
        If (counter > 2) Then
            counter = 0
        End If

            'CHECK IF TIME LIMIT HAS BEEN EXCEEDED
        If Now > endTime Then

            Exit Sub
        End If

    Next cell
    Application.Wait (Now + TimeValue("0:00:04"))
    call save_Sheet
    'call WriteTodayDate
End Sub
Private Sub DS_refresh1()

    Debug.Print "CALLING DS9"
    Application.COMAddIns("PowerlinkCOMAddIn.COMAddIn").Object.RefreshWorkbook
    Application.COMAddIns("PowerlinkCOMAddIn.COMAddIn").Object.RefreshSelection
    Application.COMAddIns("PowerlinkCOMAddIn.COMAddIn").Object.RefreshActiveSheet

End Sub


Private Sub removeNotAvailable(sth1 As String)

    Dim fnd As Variant
    Dim rplc As Variant

    Debug.Print "removing n/a from.. " + sth1
    Dim nSheet As Worksheet
    Set nSheet = ThisWorkbook.Sheets(sth1)

End Sub
