Sub run2()

    Dim columnCount As Integer
    Dim myRange1 As Range
    Set myRange1 = Worksheets("tSpec").Range("A2:C30")
    Dim counter As Integer
    Dim tabName As String
    Dim cell As Range
    Dim timer

    columnCount = 3
    counter = 0

    Application.DisplayAlerts = False
    'Call Enable_IC
    'Call BR3AK
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

        ' Check if 10 minutes have passed
        If Timer > (TimeValue("00:15:00") + timerStart) Then
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

        ' Check if 10 minutes have passed
        If Timer > (TimeValue("00:10:00") + timerStart) Then
            Exit Sub
        End If
    Next cell

    Application.Wait (Now + TimeValue("0:00:04"))

    ' Check if 10 minutes have passed
    If Timer > (TimeValue("00:10:00") + timerStart) Then
        Exit Sub
    End If

    'Application.DisplayAlerts = True
    'MsgBox "done deal!"
    Application.OnTime (Now + TimeValue("00:00:20")), "refresh_Formulas"
    'call save_Sheet
    'Application.OnTime (Now + TimeValue("00:00:10")), "close_Sheet"

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

    Debug.Print "refreshing.. " + sth1.Name
    Dim nSheet As Worksheet
    nSheet = ThisWorkbook.Sheets(sth1)

    fnd = "#N/A"
    rplc = ""
    nSheet.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

End Sub