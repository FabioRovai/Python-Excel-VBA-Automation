Attribute VB_Name = "basEvents"
Option Explicit
Private bProcessClick As Boolean
Private m_retryCount As Long
Private Const cnst_AutomationError = 429
Private Const cnst_AddInLoadDelay = 438
Public Sub togglebProcessClk()
    bProcessClick = Not bProcessClick
End Sub
Sub btnProcessTable_Click()
    bProcessClick = True
    Call ProcessRequestTable
    bProcessClick = False
End Sub
Sub btnAddToIndex_Click()
    On Error GoTo ErrProc:
    Application.EnableCancelKey = xlDisabled
    If Not IsNothing(gRTCOMInterface) Then
        Call gRTCOMInterface.AddToIndex
    End If
    Exit Sub
ErrProc:
    Call dsSetError("btnAddToIndex_Click", vbCritical)
End Sub
Sub btnRTHelp_Click()
    On Error GoTo ErrProc:
    Application.EnableCancelKey = xlDisabled
    If Not IsNothing(gRTCOMInterface) Then
        Call gRTCOMInterface.OpenRTHelp
    End If
    Exit Sub
ErrProc:
    Call dsSetError("btnRTHelp_Click", vbCritical)
End Sub
Sub lblFrequency_Click()
    Dim oWorksheet As Worksheet
    Dim oFreqList As ComboBox
    On Error GoTo ErrProc:
    
    Set oWorksheet = ThisWorkbook.Sheets(SHEET_NAME)
    Set oFreqList = ThisWorkbook.Worksheets(SHEET_NAME).cboFrequency
    Application.EnableCancelKey = xlDisabled
    If Not IsNothing(gRTCOMInterface) Then
        Call gRTCOMInterface.chkFrequencySelection(oFreqList.Text)
    End If
    Exit Sub
ErrProc:
    Call dsSetError("lblFreq_Click", vbCritical)
End Sub

Sub chkDisplayDetails_Click()
    Dim oDispDetails As CheckBox
    Dim bVal As Boolean
    On Error GoTo ErrProc:
    
    Set oDispDetails = ThisWorkbook.Sheets(SHEET_NAME).DrawingObjects("chkDisplayDetails")
    bVal = IIf(oDispDetails.Value = xlOn, True, False)
    
    If Not IsNothing(gRTCOMInterface) Then
        Call gRTCOMInterface.chkDisplayDetailsClick(bVal)
    End If
    Exit Sub
ErrProc:
    Call dsSetError("chkDisplayDetails_Click", vbCritical)
End Sub

Sub chkDispExcelDestination_Click()
    Dim oDispExcel As CheckBox
    Dim oR1C1Sytle As CheckBox
    Dim bDispExcel As Boolean
    Dim bR1C1Sytle As Boolean
    
    On Error GoTo ErrProc:
    
    Set oDispExcel = ThisWorkbook.Sheets(SHEET_NAME).DrawingObjects("chkDispExcelDestination")
    bDispExcel = IIf(oDispExcel.Value = xlOn, True, False)

    Set oR1C1Sytle = ThisWorkbook.Sheets(SHEET_NAME).DrawingObjects("chkR1C1Ref")
    bR1C1Sytle = IIf(oR1C1Sytle.Value = xlOn, True, False)
    
    If Not IsNothing(gRTCOMInterface) Then
        Call gRTCOMInterface.chkExcelFormulaClick(bDispExcel, bR1C1Sytle)
    End If
    Exit Sub
ErrProc:
    Call dsSetError("chkDispExcelDestination_Click", vbCritical)
End Sub

Sub chkNAWrite_Click()
    Dim oNAString As CheckBox
    Dim bNAString As Boolean
    
    Set oNAString = ThisWorkbook.Sheets(SHEET_NAME).DrawingObjects("chkNAWrite")
    bNAString = IIf(oNAString.Value = xlOn, True, False)
    ThisWorkbook.Sheets(SHEET_NAME).NAStringEnableDisable (bNAString)
    Exit Sub
ErrProc:
    Call dsSetError("chkDispExcelDestination_Click", vbCritical)
End Sub

Public Sub ProcessRequestTable()
Attribute ProcessRequestTable.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error GoTo ErrProc:
    Dim getRT As Boolean
    Application.EnableCancelKey = xlDisabled
    If Not IsNothing(gRTCOMInterface) Then
        getRT = gRTCOMInterface.RTShortcutMacro
        If Not getRT Then
            If Not bProcessClick Then Exit Sub
        Else
            If Not bProcessClick Then
                If Not IsRequestTable() Then
                    Exit Sub
                End If
            End If
        End If
        Call gRTCOMInterface.ProcessRequestTable
    End If
    Exit Sub
ErrProc:
    Call dsSetError("ProcessRequestTable", vbCritical)
End Sub
Private Function IsRequestTable() As Boolean
    Dim wb As Workbook
    Dim rtSheet As Worksheet
    On Error GoTo ErrProc:
    Set wb = Application.ActiveWorkbook
    Set rtSheet = wb.Worksheets(SHEET_NAME)
    If rtSheet Is Nothing Then
        IsRequestTable = False
    Else
        IsRequestTable = True
    End If
    Exit Function
ErrProc:
    IsRequestTable = False
End Function
Public Function IsProcessTblEnabled() As Boolean
    Dim oProcessTable As CheckBox
    Dim bProcessTable As Boolean
    On Error GoTo ErrProc:
    Set oProcessTable = ThisWorkbook.Sheets(SHEET_NAME).DrawingObjects("chkProcessTbl")
    bProcessTable = IIf(oProcessTable.Value = xlOn, True, False)
    IsProcessTblEnabled = bProcessTable
    Exit Function
ErrProc:
    Call dsSetError("IsProcessTblEnabled", vbCritical)
End Function
Public Sub StartRequestTable()
   On Error GoTo ErrProc:
     Application.EnableCancelKey = xlDisabled
           If Not IsNothing(gRTCOMInterface) Then
                 gRTCOMInterface.SaveRequestTableOnOpen
           Else
            If IsProcessTblEnabled Then
             m_retryCount = m_retryCount + 1
                If (m_retryCount <= 10) Then    ' only try for up to 50 seconds
                    ' try again in 5 seconds
                 Application.OnTime Now + TimeValue("00:00:05"), "StartRequestTable"
                End If
             Else
             Exit Sub
            End If
            
           End If
    Exit Sub
ErrProc:
    If Err.Number = cnst_AutomationError Then
        Err.Clear
    ElseIf Err.Number = cnst_AddInLoadDelay Then
        Err.Clear
    Else
        Call dsSetError("StartRequestTable : ", vbCritical)
    End If
End Sub

