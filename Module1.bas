Attribute VB_Name = "Statistics"
Public fMainForm As frmMain
Public TestOpen As Boolean, ExitApp As Boolean, CancelUnload As Boolean
Public ErrorNumber As Integer
'General
Public LabMaxRowsDisplay As Integer, LabMaxColsDisplay As Integer
Sub Main()
Set fMainForm = New frmMain
fMainForm.Show
End Sub
Public Function IsStringANumber(NumString As String) As Boolean
If NumString = "" Then Exit Function
IsStringANumber = True
Dim I As Integer
For I = 1 To Len(NumString)
 Select Case Asc(Mid$(NumString, I, 1))
  Case Is > 57, Is < 45, 47: IsStringANumber = False: I = Len(NumString) + 1
 End Select
Next I
End Function
Public Function Average!(StatArray() As Single)
Dim I As Integer, AvgSum As Single, Total As Integer
For I = 1 To UBound(StatArray)
 AvgSum = AvgSum + Val(StatArray(I)): Total = Total + 1
Next I
Average = AvgSum / Total
End Function
Public Function Median(StatArray() As Single) As Single
Sort StatArray
If UBound(StatArray) / 2 = Int(UBound(StatArray) / 2) Then
 Median = (StatArray(UBound(StatArray) / 2) + StatArray(UBound(StatArray) / 2 + 1)) / 2
Else
 Median = StatArray(UBound(StatArray) / 2)
End If
End Function
Public Sub Swap(ByRef VarA As Variant, ByRef VarB As Variant)
Dim VarT As Variant
VarT = VarA
VarA = VarB
VarB = VarT
End Sub
Public Function SD(StatArray() As Single, Avg As Single) As Single
Dim I As Integer, Result As Single
For I = 1 To UBound(StatArray): Result = Result + (StatArray(I) - Avg) ^ 2: Next I
SD = Sqr(Result / (UBound(StatArray) - 1))
End Function

Public Function StandardDeviation(msfg As Control, Average As Single, Count As Integer) As Single
Dim Result As Single, I As Integer
For I = 1 To msfg.Rows - 1
 If IsStringANumber(msfg.TextMatrix(I, msfg.Col)) = True Then Result = Result + (Val(msfg.TextMatrix(I, msfg.Col)) - Average) ^ 2
Next I
If Count > 1 Then Result = Sqr(Result / (Count - 1)) 'Actual formula of SD is calculated using count & not count-1
StandardDeviation = Result
End Function
Public Function RSD(StdDeviation As Single, Average As Single) As Single
RSD = StdDeviation / Average * 100
End Function
Public Sub Sort(SortArray() As Single)
Dim I As Integer, J As Integer
For I = LBound(SortArray) To UBound(SortArray) - 1
 For J = I + 1 To UBound(SortArray)
  If SortArray(I) > SortArray(J) Then Swap SortArray(I), SortArray(J)
 Next J
Next I
End Sub
