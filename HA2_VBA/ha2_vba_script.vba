Private Sub GenerateData_Click()

'0 = KITS01
'1 = KITS02
'2 = KITS03
'3 = KITS04

'4 = UKIS01
'5 = UKIS02
'6 = UKIS03
'7 = UKIS04

'8 = PLUS01
'9 = PLUS02
'10 = PLUS03

'11 = GTBS01

Dim max_entries As Integer
max_entries = 5000

Dim results(0 To 11, 0 To 23) As Integer
Dim data_entry As Integer
Dim entry_hour_index As Integer

Dim hour_position(0 To 23) As Integer

Dim i As Integer
Dim j As Integer

Dim filter_by_date As Boolean
Dim filter_date As String

If Not IsEmpty(Sheets("Part_Process").Cells(4, 6).Value) = True Then
    filter_by_date = True
    filter_date = Sheets("Part_Process").Cells(4, 6).Text
Else
    filter_by_date = False
End If
    
For data_entry = 7 To max_entries
    ' Check machine name and input into correct array

    ' Check if optional date filter has been applied
    ' Ignore all dates that do not work with this filter
    If filter_by_date = True Then
        If Not Left(Sheets("Part_Process").Cells(data_entry, 8).Value, 10) = filter_date Then
            GoTo End_of_loop
        End If
    End If

    ' Convert time to array index to match completed time
    If Not IsEmpty(Sheets("Part_Process").Cells(data_entry, 8)) = True Then
    
        data_hour_index = CInt(Left(Right(Sheets("Part_Process").Cells(data_entry, 8).Value, 12), 2))

        Select Case Sheets("Part_Process").Cells(data_entry, 11).Value
        Case "KITS01"
            results(0, data_hour_index) = results(0, data_hour_index) + 1
        Case "KITS02"
            results(1, data_hour_index) = results(1, data_hour_index) + 1
        Case "KITS03"
            results(2, data_hour_index) = results(2, data_hour_index) + 1
        Case "KITS04"
            results(3, data_hour_index) = results(3, data_hour_index) + 1
        Case "UKIS01"
            results(4, data_hour_index) = results(4, data_hour_index) + 1
        Case "UKIS02"
            results(5, data_hour_index) = results(5, data_hour_index) + 1
        Case "UKIS03"
            results(6, data_hour_index) = results(6, data_hour_index) + 1
        Case "UKIS04"
            results(7, data_hour_index) = results(7, data_hour_index) + 1
        Case "PLUS01"
            results(8, data_hour_index) = results(8, data_hour_index) + 1
        Case "PLUS02"
            results(9, data_hour_index) = results(9, data_hour_index) + 1
        Case "PLUS03"
            results(10, data_hour_index) = results(10, data_hour_index) + 1
        Case "GTBS01"
            results(11, data_hour_index) = results(11, data_hour_index) + 1
        End Select
    End If
    
End_of_loop:

Next data_entry
    
    
' Format Sheet(Hours)
For i = 0 To 23
    If i > 6 Then
        If i > 9 Then
            Cells(1, i - 5).Value = CStr(i) & ":00"
        Else
            Cells(1, i - 5).Value = "0" & CStr(i) & ":00"
        End If
    Else
        Cells(1, i + 19).Value = "0" & CStr(i) & ":00"
    End If
    Next i

' Format Sheet (Asset Names)
Cells(2, 1).Value = "KITS01"
Cells(3, 1).Value = "KITS02"
Cells(4, 1).Value = "KITS03"
Cells(5, 1).Value = "KITS04"
Cells(6, 1).Value = "UKIS01"
Cells(7, 1).Value = "UKIS02"
Cells(8, 1).Value = "UKIS03"
Cells(9, 1).Value = "UKIS04"
Cells(10, 1).Value = "PLUS01"
Cells(11, 1).Value = "PLUS02"
Cells(12, 1).Value = "PLUS03"
Cells(13, 1).Value = "GTBS01"

' Format cells to general
Range("B2", "Y13").NumberFormat = "General"

' Spit Data onto sheet
For i = 0 To 11
    For j = 0 To 23
        If j > 6 Then
            Cells(i + 2, j - 5).Value = results(i, j)
        Else
            Cells(i + 2, j + 19).Value = results(i, j)
        End If
        Next j
    Next i

End Sub
