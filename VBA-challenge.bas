Attribute VB_Name = "Module1"
Sub stockmarket()

'Apply code through each sheet
Dim sheet As Worksheet
For Each sheet In ActiveWorkbook.Worksheets
sheet.Activate


'Declare Variables
Dim Ticker As String
Dim yopen, yclose, totalvol, yc, pc As Double
Dim summary As Integer
Dim lrow As Double
Dim count As Double
Dim j As Long


'set summary location and total stock volume
summary = 2
totalsvol = 0
lrow = Cells(Rows.count, 1).End(xlUp).Row


' Set Sheet Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


'Establish loop through ticker

For i = 2 To lrow

'Obtain total volume
totalsvol = totalsvol + Cells(i, 7).Value

'Obtain value for Year Open accounting for if open value is zero

If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
    yopen = Cells(i, 3).Value
      
    If yopen = 0 Then
            count = i + 1
        For j = count To lrow
            yopen = Cells(j, 3)
          If (yopen <> 0) Then
        Exit For
          End If
       Next j
    End If
       
End If
      
'Obtain closing value and make calculations
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           Ticker = Cells(i, 1).Value
           yclose = Cells(i, 6).Value
           yc = yclose - yopen
           
'account for if stock has zero value
        If yopen = 0 And yclose = 0 Then
            pc = 0
        Else
            pc = yc / yopen
        End If
    
'insert into summary table
        Range("I" & summary).Value = Ticker
        Range("L" & summary).Value = totalsvol
        Range("J" & summary).Value = yc
        Range("K" & summary).Value = pc

'Move to next row

summary = summary + 1

'Reset total volumne
totalsvol = 0
 

End If


Next i

'Establish Variable for last row of calculated values

Dim crow As Integer
    crow = Cells(Rows.count, 11).End(xlUp).Row

For k = 2 To crow
    If Cells(k, 10) >= 0 Then
        Cells(k, 10).Interior.ColorIndex = 4
    Else
        Cells(k, 10).Interior.ColorIndex = 3
     End If
Next k

' Bonus table

Dim hvols, lps, hps As String

Dim GP, LP, HV As Double
GP = Cells(2, 11).Value
HV = Cells(2, 12).Value
LP = Cells(2, 11).Value

For l = 2 To crow
   
   If Cells(l, 11).Value > GP Then
        GP = Cells(l, 11).Value
        hps = Cells(l, 9).Value
   End If
   
   If Cells(l, 12).Value > HV Then
        HV = Cells(l, 12).Value
        hvols = Cells(l, 9).Value
   End If
   
   If Cells(l, 11).Value < LP Then
       LP = Cells(l, 11).Value
       lps = Cells(l, 9).Value
   End If

Next l

Cells(2, 17).Value = GP
Cells(2, 16).Value = hps
Cells(3, 17).Value = LP
Cells(3, 16).Value = lps
Cells(4, 17).Value = HV
Cells(4, 16).Value = hvols

'Format Cells to proper format

Columns("I:Q").Select
    Columns("I:Q").EntireColumn.AutoFit
Columns("K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"

Next sheet

End Sub

