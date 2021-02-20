VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private StatisticArray() As Single
Public Average!, Mode!, Quartile1!, Quartile2!, Quartile3!, IQR!, SIQR!, NIQR!
Public SD!, RSD!
Public Sub Sort()
Dim I As Integer, J As Integer
For I = LBound(StatisticArray) To UBound(StatisticArray) - 1
 For J = I + 1 To UBound(StatisticArray)
  If StatisticArray(I) > StatisticArray(J) Then Swap StatisticArray(I), StatisticArray(J)
 Next J
Next I
End Sub
Public Sub FillArray(StatArray() As Single)
ReDim StatisticArray(LBound(StatArray) To UBound(StatArray)), Zc(LBound(StatArray) To UBound(StatArray)), Zr(LBound(StatArray) To UBound(StatArray))
Dim I As Integer, AvgSum As Single
For I = LBound(StatArray) To UBound(StatArray)
 StatisticArray(I) = StatArray(I)
 AvgSum = AvgSum + StatisticArray(I)
Next I
Average = AvgSum / (UBound(StatisticArray) - LBound(StatisticArray) + 1)
For I = LBound(StatisticArray) To UBound(StatisticArray)
 SD = SD + (StatisticArray(I) - Average) ^ 2
Next I
SD = Sqr(SD / (UBound(StatisticArray) - 1)): RSD = SD / Average * 100
Dim TempArr() As Single
ReDim TempArr(LBound(StatisticArray) To UBound(StatisticArray))
Dim LMiddle As Integer, UMiddle As Integer
TempArr = StatisticArray
Statistics.Sort TempArr
Mode = 0 'Find mode
If UBound(TempArr) / 2 = Int(UBound(TempArr) / 2) Then 'Even
 Quartile2 = (TempArr(UBound(TempArr) / 2) + TempArr(UBound(TempArr) / 2 + 1)) / 2
 LMiddle = UBound(TempArr) / 2: UMiddle = LMiddle + 1
Else 'Odd
 Quartile2 = TempArr(UBound(TempArr) / 2 + 0.5)
 LMiddle = Int(UBound(TempArr) / 2): UMiddle = LMiddle + 2
End If
If LMiddle / 2 = Int(LMiddle / 2) Then   'Even
 Quartile1 = (TempArr(LMiddle / 2) + TempArr(LMiddle / 2 + 1)) / 2
Else 'odd
 Quartile1 = TempArr(LMiddle / 2 + 0.5)
End If
Dim Inx As Integer
Inx = UBound(TempArr) - UMiddle + 1
If Inx / 2 = Int(Inx / 2) Then   'Even
 Quartile3 = (TempArr(UMiddle + Inx / 2 - 1) + TempArr(UMiddle + Inx / 2)) / 2
Else 'odd
 Quartile3 = TempArr(UMiddle + Inx / 2 - 0.5)
End If
IQR = Quartile3 - Quartile1
SIQR = IQR / 2
NIQR = IQR / 1.3489
End Sub

Public Function Robust(StNum As Single) As Single
Robust = (StNum - Quartile2) / NIQR
End Function
Public Function Classical(StNum As Single) As Single
Classical = (StNum - Average) / SD
End Function
