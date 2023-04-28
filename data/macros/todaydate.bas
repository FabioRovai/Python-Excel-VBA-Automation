Sub WriteTodayDate()
    Dim todayDate As Date
    todayDate = Date

    'Write today's date to cell A1
    Worksheets("tSpec").Range("H3").Value = todayDate
End Sub