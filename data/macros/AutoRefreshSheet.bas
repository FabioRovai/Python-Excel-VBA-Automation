Attribute VB_Name = "Module3"
Sub refresh_Formulas()
    Application.Calculate
    Application.OnTime (Now + TimeValue("00:00:20")), "save_Sheet"
End Sub




