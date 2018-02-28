ttribute VB_Name = "Module1"
' Constants
Const Start_Row = 9
Const End_Row = 17
Const Start_Column = 5
Const End_Column = 13
Const Row_Diff = 8
Const Col_Diff = 4

' Variables
Dim Count_Loner As Integer
Dim Count_Negatives As Integer
Dim Number_Array(10, 10, 10) As Integer
Dim All_Loners_Complete As Boolean
Dim All_Negatives_Complete As Boolean
Dim Show_Hint As Boolean
Private Sub Sudoku()
Attribute Sudoku.VB_ProcData.VB_Invoke_Func = "M\n14"

All_Loners_Complete = False
All_Negatives_Complete = False

' Identify Loners and Plot them
Do
    Call Scan_Memory_For_Loners
Loop Until (All_Loners_Complete = True)

If (Show_Hint = True) And (Count_Loner > 0) Then
   Exit Sub
End If

' Identify Negatives and Plot them
Do
    Call Scan_Memory_For_Negatives
Loop Until (All_Negatives_Complete = True)

End Sub
Private Sub Scan_Memory_For_Loners()
Attribute Scan_Memory_For_Loners.VB_Description = "Sudoku Loner"
Attribute Scan_Memory_For_Loners.VB_ProcData.VB_Invoke_Func = "M\n14"

Box_Row = Start_Row
Box_Column = Start_Column

' Initialize Memory

For I = 1 To 9
    For J = 1 To 9
        For K = 1 To 9
            Number_Array(I, J, K) = K
        Next
    Next
Next

' Box by Box Traverse
Count_Loner = 0

For I = 1 To 9
    For J = Box_Row To Box_Row + 2
        For K = Box_Column To Box_Column + 2
            L = Cells(J, K).Value
            If IsEmpty(L) = True Then
               Call Plot_Loner(J, K, I, Box_Row, Box_Column)
               If (Show_Hint = True) And (Count_Loner > 0) Then
                  All_Loners_Complete = True
                  Exit Sub
               End If
            End If
        Next
    Next
Box_Column = Box_Column + 3
If I Mod 3 = 0 Then
    Box_Row = Box_Row + 3
    Box_Column = Start_Column
End If
Next

Box_Row = Start_Row
Box_Column = Start_Column

If Count_Loner = 0 Then
   All_Loners_Complete = True
End If

End Sub
Private Sub Plot_Loner(ByVal Loner_Row As Integer, ByVal Loner_Column As Integer, ByVal Loner_Box As Integer, ByVal Loner_Box_Row As Integer, ByVal Loner_Box_Column As Integer)

Loner_Counter = 0

' Column Scan

For M = Start_Row To End_Row
    N = Cells(M, Loner_Column).Value
    If IsEmpty(N) = False Then
        Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, N) = Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, N) - Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, N)
    End If
Next

' Row Scan

For M = Start_Column To End_Column
    N = Cells(Loner_Row, M).Value
    If IsEmpty(N) = False Then
        Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, N) = Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, N) - Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, N)
    End If
Next

' Box Scan

    For M = Loner_Box_Row To Loner_Box_Row + 2
        For N = Loner_Box_Column To Loner_Box_Column + 2
            P = Cells(M, N).Value
            If IsEmpty(P) = False Then
               Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, P) = Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, P) - Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, P)
            End If
        Next
    Next

' Loner Validation

For M = 1 To 9
If Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, M) > 0 Then
   Loner_Counter = Loner_Counter + 1
End If
Next

' Loner Value Assign

If Loner_Counter = 1 Then
   For M = 1 To 9
   If Number_Array(Loner_Row - Row_Diff, Loner_Column - Col_Diff, M) > 0 Then
      Cells(Loner_Row, Loner_Column).Value = M
   End If
   Next
   Count_Loner = Count_Loner + 1
   If Show_Hint = True Then
         Exit Sub
   End If
End If

End Sub

Private Sub Scan_Memory_For_Negatives()

Box_Row = Start_Row
Box_Column = Start_Column
Count_Negatives = 0

' Box by Box Traverse

For I = 1 To 9
    For J = Box_Row To Box_Row + 2
        For K = Box_Column To Box_Column + 2
            L = Cells(J, K).Value
            If IsEmpty(L) = True Then
                Call Plot_Negatives(J, K, I, Box_Row, Box_Column)
                If (Show_Hint = True) And (Count_Negatives > 0) Then
                  All_Negatives_Complete = True
                  Exit Sub
                End If
            End If
        Next
    Next
Box_Column = Box_Column + 3
If I Mod 3 = 0 Then
    Box_Row = Box_Row + 3
    Box_Column = Start_Column
End If
Next

Box_Row = Start_Row
Box_Column = Start_Column

If Count_Negatives = 0 Then
   All_Negatives_Complete = True
End If

End Sub

Private Sub Plot_Negatives(ByVal Negative_Row As Integer, ByVal Negative_Column As Integer, ByVal Negative_Box As Integer, ByVal Negative_Box_Row As Integer, ByVal Negative_Box_Column As Integer)

' Column Scan
For L = 1 To 9
    Negative_Counter = 0
    If Number_Array(Negative_Row - Row_Diff, Negative_Column - Col_Diff, L) > 0 Then
        For M = Start_Row To End_Row
            N = Cells(M, Negative_Column).Value
            If IsEmpty(N) = False Then
               Negative_Counter = Negative_Counter + 1
            ElseIf Number_Array(M - Row_Diff, Negative_Column - Col_Diff, L) <= 0 Then
               Negative_Counter = Negative_Counter + 1
            End If
        Next
    End If
    If Negative_Counter = 8 Then
       Cells(Negative_Row, Negative_Column).Value = L
       Count_Negatives = Count_Negatives + 1
       'msgbox "done"
       Exit For
    End If
Next

If IsEmpty(Cells(Negative_Row, Negative_Column).Value) = False Then
    If Show_Hint = True Then
         Exit Sub
    End If
    All_Loners_Complete = False
    Do
        Call Scan_Memory_For_Loners
    Loop Until (All_Loners_Complete = True)
    Exit Sub
End If

' Row Scan

For L = 1 To 9
    Negative_Counter = 0
    If Number_Array(Negative_Row - Row_Diff, Negative_Column - Col_Diff, L) > 0 Then
        For M = Start_Column To End_Column
            N = Cells(Negative_Row, M).Value
            If IsEmpty(N) = False Then
               Negative_Counter = Negative_Counter + 1
            ElseIf Number_Array(Negative_Row - Row_Diff, M - Col_Diff, L) <= 0 Then
               Negative_Counter = Negative_Counter + 1
            End If
        Next
    End If
    If Negative_Counter = 8 Then
       Cells(Negative_Row, Negative_Column).Value = L
       Count_Negatives = Count_Negatives + 1
       'msgbox "done"
       Exit For
    End If
Next

If IsEmpty(Cells(Negative_Row, Negative_Column).Value) = False Then
   If Show_Hint = True Then
         Exit Sub
   End If
   All_Loners_Complete = False
   Do
        Call Scan_Memory_For_Loners
   Loop Until (All_Loners_Complete = True)
   Exit Sub
End If

' Box Scan
For L = 1 To 9
    Negative_Counter = 0
    If Number_Array(Negative_Row - Row_Diff, Negative_Column - Col_Diff, L) > 0 Then
    For M = Negative_Box_Row To Negative_Box_Row + 2
        For N = Negative_Box_Column To Negative_Box_Column + 2
            P = Cells(M, N).Value
            If IsEmpty(P) = False Then
               Negative_Counter = Negative_Counter + 1
            ElseIf Number_Array(M - Row_Diff, N - Col_Diff, L) <= 0 Then
               Negative_Counter = Negative_Counter + 1
            End If
        Next
    Next
    End If
If Negative_Counter = 8 Then
       Cells(Negative_Row, Negative_Column).Value = L
       Count_Negatives = Count_Negatives + 1
       'msgbox "done"
       Exit For
End If
Next

If IsEmpty(Cells(Negative_Row, Negative_Column).Value) = False Then
   If Show_Hint = True Then
         Exit Sub
   End If
   All_Loners_Complete = False
   Do
        Call Scan_Memory_For_Loners
   Loop Until (All_Loners_Complete = True)
   Exit Sub
End If

End Sub
Sub Hint()

Show_Hint = True

Call Sudoku

End Sub

Sub Solve()

Show_Hint = False

Call Sudoku

End Sub
