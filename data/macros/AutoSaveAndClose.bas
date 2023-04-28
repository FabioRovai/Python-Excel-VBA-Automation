Attribute VB_Name = "Module2"
Sub save_Sheet()

    ThisWorkbook.Save


    Application.OnTime (Now + TimeValue("00:00:10")), "close_Sheet"
End Sub





