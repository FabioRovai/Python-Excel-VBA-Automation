Attribute VB_Name = "basPublic"
Option Explicit
Public gRTCOMInterface As Object
Public Const SHEET_NAME = "REQUEST_TABLE"   ' Sheet name of table
' Freq options text
Private Const DAILY_FREQ = "Daily"
Private Const WEEKLY_FREQ = "Weekly"
Private Const MONTHLY_FREQ = "Monthly"
Private Const QUARTER_FREQ = "Quarterly"
Private Const YEARLY_FREQ = "Yearly"

Private Const DFO_ENVIRON_VAR = "ENV_DFOADDIN_SETUP_PATH"

Private Const AccessErrorNo = 245755
Private Const cnstERR_Access = vbObjectError - AccessErrorNo


Public Sub fill_freqlist()
    On Error GoTo ErrProc:
    Dim oFreqList As ComboBox
    Set oFreqList = ThisWorkbook.Worksheets(SHEET_NAME).cboFrequency
    oFreqList.Clear
    With oFreqList
        .AddItem DAILY_FREQ
        .AddItem WEEKLY_FREQ
        .AddItem MONTHLY_FREQ
        .AddItem QUARTER_FREQ
        .AddItem YEARLY_FREQ
    End With
    oFreqList.Value = DAILY_FREQ
    Exit Sub
ErrProc:
    Call dsSetError("fill_freqlist", vbCritical)
End Sub
Public Sub dsSetError(ErrorMessage As String, vbErrorType As Integer)
     
    Dim sMessage As String
    
    'On Error Resume Next
    
    Application.Cursor = xlDefault
    Application.StatusBar = "Ready"
    
    sMessage = "Sorry, an error has occurred within DFO.  " & _
                "Should you wish to report this, please" & Chr$(13) & Chr$(13) & _
                "   1.  Note down your last actions" & Chr$(13) & _
                "   2.  Note down the error message below." & Chr$(13) & _
                "   3.  Contact your local Datastream representative with log files." & Chr$(13) & Chr$(13) _
                & ErrorMessage
    
    MsgBox sMessage & Chr$(13) _
        & "ERROR:" & Err & ":" & Error & Chr$(13) & "SOURCE:" & Err.Source, vbErrorType, "Datastream for Office"
End Sub


Public Function IsNothing(obj As Object) As Boolean
     On Error GoTo ErrProc
    If Not (obj Is Nothing) Then
        IsNothing = False
    Else
        Dim COAddInObject As Object
        Set COAddInObject = Application.COMAddIns("PowerlinkCOMAddIn.COMAddIn").Object
        If Not (COAddInObject Is Nothing) Then
            If Not IsNull(COAddInObject) Then
                Set gRTCOMInterface = COAddInObject.GetRTComHelperInstance
                If Not (gRTCOMInterface Is Nothing) Then
                  If Not IsNull(gRTCOMInterface) Then
                    Call gRTCOMInterface.SetActiveWorkbook(ThisWorkbook)
                    Dim bRet As Boolean: bRet = gRTCOMInterface.SetSharedObject
                    IsNothing = Not (bRet)
                  Else
                   IsNothing = True
                  End If
                Else
                  IsNothing = True
                End If
            Else
              IsNothing = True
            End If
         Else
           IsNothing = True
         End If
     End If
     
    Exit Function
ErrProc:
    IsNothing = True
End Function
