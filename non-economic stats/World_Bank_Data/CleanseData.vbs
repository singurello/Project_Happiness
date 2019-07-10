Sub CleanseData()

  ' Setup a counter to track cell number
  Dim cellnumber As Integer
  Dim i, j As Integer

  ' Loop through each row of the board
  For i = 1 To 59

    ' Loop through each column of the board
    For j = 1 To 264

      ' If we are on a cell that is divisible by 2 then color it black
       Cells(j + 1 + (i - 1) * 264, 1).Value = Cells(j + 1, 4).Value
       Cells(j + 1 + (i - 1) * 264, 2).Value = Cells(1, 5 + i).Value
       Cells(j + 1 + (i - 1) * 264, 3).Value = Cells(j + 1, 5 + i).Value

    Next j

  Next i

End Sub
