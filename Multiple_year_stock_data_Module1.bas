Attribute VB_Name = "Module1"
'Reading stock data
Sub readStock()

Cells(1, 9).Value = "Ticker"
Cells(1, 9).Font.Bold = True

Cells(1, 10).Value = "Yearly Change"
Cells(1, 10).Font.Bold = True

Cells(1, 11).Value = "Percent Change"
Cells(1, 11).Font.Bold = True

Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 12).Font.Bold = True

Dim str As String

Dim cnt As Integer
cnt = 2

Dim openStk As Double
Dim closeStk As Double
Dim diffStk As Double
Dim totalStk As Double
totalStk = 0

Dim i As Long
For i = 2 To 1048576
    If Cells(i, 1).Value = "" Then Exit For
    If i = 2 Then
        str = Cells(i, 1).Value
        Cells(cnt, 9).Value = str
        'Save open stock
        openStk = Cells(i, 3).Value
        totalStk = Cells(i, 7).Value
    Else
        If str <> Cells(i, 1).Value Then
            cnt = cnt + 1
            str = Cells(i, 1).Value
            Cells(cnt, 9).Value = str
            'Save close
            closeStk = Cells(i - 1, 6).Value
            'print on the excel sheet
            diffStk = closeStk - openStk
            Cells(cnt - 1, 10).Value = diffStk
            ' conditional formatting
            If (diffStk < 0) Then
                Cells(cnt - 1, 10).Interior.ColorIndex = 3
            Else
                Cells(cnt - 1, 10).Interior.ColorIndex = 4
            End If
            
            ' if either closeStk or OpenStk is zero then this will fail
            If closeStk <> 0 And openStk <> 0 Then
                Cells(cnt - 1, 11).Value = (closeStk * 100 / openStk) - 100
            Else
                Cells(cnt - 1, 11).Value = 0
            End If
            
            Cells(cnt - 1, 12).Value = totalStk
            'Save open stock
            openStk = Cells(i, 3).Value
            'total stock reset and save new value
            totalStk = 0
            totalStk = Cells(i, 7).Value
        Else
            totalStk = totalStk + Cells(i, 7).Value
        End If
    End If
Next i
End Sub
