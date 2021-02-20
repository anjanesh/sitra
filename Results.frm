VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmResult 
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   Icon            =   "Results.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10440
   Begin MSFlexGridLib.MSFlexGrid msfgResult 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9340
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Public ResultFileName As String, ResultName As String
Private Sub Form_Load()
If ResultFileName = "" Then NewResult Else OpenResult
Dim I%, J%, TotalChecked%, strFormulae$()
For I = 0 To frmResultDialog.chkFormula.Count - 1
 If frmResultDialog.chkFormula(I) = Checked Then
  TotalChecked = TotalChecked + 1
  ReDim Preserve strFormulae(TotalChecked)
  strFormulae(TotalChecked) = frmResultDialog.chkFormula(I).Caption
 End If
Next I
With msfgResult
  .TextArray(0) = "Lab / Sample"
  .TextMatrix(.Rows - 8, 0) = "Average"
  .TextMatrix(.Rows - 7, 0) = "Median"
  .TextMatrix(.Rows - 6, 0) = "SD"
  .TextMatrix(.Rows - 5, 0) = "RSD"
  .TextMatrix(.Rows - 4, 0) = "Quartile1"
  .TextMatrix(.Rows - 3, 0) = "Quartile3"
  .TextMatrix(.Rows - 2, 0) = "IQR"
  .TextMatrix(.Rows - 1, 0) = "NIQR"
  .Row = 0
  For I = 0 To msfgResult.Cols - 1
   .Col = I: .CellAlignment = 4: .CellFontBold = True
  Next I
  .Col = 0
  For I = 0 To msfgResult.Rows - 1
   .Row = I: .CellAlignment = 4: .CellFontBold = True
  Next I
  For I = 1 To .Rows - 1
   For J = 1 To .Cols - 1
    .Row = I: .Col = J: .CellAlignment = 4
   Next J
  Next I
 .ColWidth(0) = .ColWidth(0) * 1.5
End With
End Sub

Public Sub SaveResultAs(FileName As String)
Dim FileNum As Integer, I As Integer, J As Integer, NoOfRows As Integer, NoOfCols As Integer
FileNum = FreeFile
NoOfRows = msfgResult.Rows - 9: NoOfCols = msfgResult.Cols
Open FileName For Output As FileNum
Print #FileNum, "ZScore"
Write #FileNum, NoOfRows, NoOfCols
Write #FileNum, ResultName
For I = 1 To NoOfCols - 4 Step 3: Write #1, msfgResult.TextMatrix(0, I),: Next I
Write #1, msfgResult.TextMatrix(0, I)
For I = 1 To NoOfRows
 For J = 0 To NoOfCols - 2
  Write #FileNum, msfgResult.TextMatrix(I, J),
 Next J
 Write #FileNum, msfgResult.TextMatrix(I, J)
Next I
For I = NoOfRows + 1 To NoOfRows + 8
 For J = 1 To NoOfCols - 4 Step 3: Write #1, msfgResult.TextMatrix(I, J),: Next J
 Write #1, msfgResult.TextMatrix(I, J)
Next I
Close FileNum
End Sub
Public Sub NewResult()
Dim I As Integer, J As Integer, K As Integer, Flag As Boolean, SampleCol() As String
Dim Field As String, FileNum As Integer, NoOfRows As Integer, NoOfCols As Integer, RowCount As Integer
For I = 1 To frmResultDialog.TreeViewCheck.Nodes.Count
If frmResultDialog.TreeViewCheck.Nodes(I).Checked = True Then
 FileNum = FreeFile: RowCount = RowCount + 1
 Open frmResultDialog.TreeViewCheck.Nodes(I).Key For Input As FileNum
  Input #FileNum, NoOfRows, NoOfCols
  Input #FileNum, Field: msfgResult.TextMatrix(msfgResult.Rows - 1, 0) = Field
  ReDim SampleCol(NoOfCols - 1)
  For J = 1 To NoOfCols - 1
  Input #FileNum, SampleCol(J)
   Flag = False
   For K = 1 To msfgResult.Cols - 1 Step 3
    If msfgResult.TextMatrix(0, K) = SampleCol(J) Then Flag = True: K = msfgResult.Cols
   Next K
   If Flag = False Then
    msfgResult.TextMatrix(0, msfgResult.Cols - 1) = SampleCol(J)
    msfgResult.Cols = msfgResult.Cols + 3
    msfgResult.TextMatrix(0, msfgResult.Cols - 2) = "Zr"
    msfgResult.TextMatrix(0, msfgResult.Cols - 3) = "Zc"
   End If
  Next J
  For J = 1 To NoOfRows - 1: Line Input #FileNum, Field: Next J
  For J = 1 To NoOfCols - 1
  Input #FileNum, Field
   For K = 1 To msfgResult.Cols - 1 Step 3
    If msfgResult.TextMatrix(0, K) = SampleCol(J) Then msfgResult.TextMatrix(msfgResult.Rows - 1, K) = Field: K = msfgResult.Cols
   Next K
  Next J
 Close FileNum
msfgResult.Rows = msfgResult.Rows + 1
End If
Next I
msfgResult.Cols = msfgResult.Cols - 1: msfgResult.Rows = msfgResult.Rows + 7
Dim Zcr As New ZScore, SampAr() As Single, NumCount As Integer
For I = 1 To msfgResult.Cols - 1 Step 3
 NumCount = 0
 For J = 1 To RowCount
  If IsStringANumber(msfgResult.TextMatrix(J, I)) = True Then
   NumCount = NumCount + 1
   ReDim Preserve SampAr(NumCount) As Single
   SampAr(NumCount) = CSng(msfgResult.TextMatrix(J, I))
  End If
 Next J
 Zcr.FillArray SampAr
 With msfgResult
  .TextMatrix(.Rows - 8, I) = Format(Zcr.Average, "#0.00")
  .TextMatrix(.Rows - 7, I) = Format(Zcr.Quartile2, "#0.00")
  .TextMatrix(.Rows - 6, I) = Format(Zcr.SD, "#0.00")
  .TextMatrix(.Rows - 5, I) = Format(Zcr.RSD, "#0.00")
  .TextMatrix(.Rows - 4, I) = Format(Zcr.Quartile1, "#0.00")
  .TextMatrix(.Rows - 3, I) = Format(Zcr.Quartile3, "#0.00")
  .TextMatrix(.Rows - 2, I) = Format(Zcr.IQR, "#0.00")
  .TextMatrix(.Rows - 1, I) = Format(Zcr.NIQR, "#0.00")
End With
 For J = 1 To RowCount
  If IsStringANumber(msfgResult.TextMatrix(J, I)) = True Then
   msfgResult.TextMatrix(J, I + 1) = Format(Zcr.Robust(CSng(msfgResult.TextMatrix(J, I))), "#0.00")
   msfgResult.TextMatrix(J, I + 2) = Format(Zcr.Classical(CSng(msfgResult.TextMatrix(J, I))), "#0.00")
  End If
 Next J
Next I
ResultName = frmResultDialog.Text1
If frmResultDialog.CheckGraph.Value = 1 Then
 Unload frmResultDialog
 Load frmGraph
 frmGraph.Show
Else
 Unload frmResultDialog
End If
Me.Caption = ResultName & " - <Not Saved>"
End Sub

Public Sub OpenResult()
Dim FileNum As Integer, I As Integer, J As Integer, NoOfRows As Integer, NoOfCols As Integer, Field As String
FileNum = FreeFile
Open ResultFileName For Input As FileNum
Line Input #FileNum, Field
If Field <> "ZScore" Then MsgBox "Not a valid Result Sheet File.", vbExclamation, "Error Opening File !": Exit Sub
Input #FileNum, NoOfRows, NoOfCols, ResultName
msfgResult.Rows = NoOfRows + 9: msfgResult.Cols = NoOfCols
For I = 1 To NoOfCols - 1 Step 3
 Input #FileNum, Field
 msfgResult.TextMatrix(0, I) = Field
 msfgResult.TextMatrix(0, I + 1) = "Zc"
 msfgResult.TextMatrix(0, I + 2) = "Zr"
Next I
For I = 1 To NoOfRows
 For J = 0 To NoOfCols - 1
  Input #FileNum, Field
  msfgResult.TextMatrix(I, J) = Field
 Next J
Next I
For I = NoOfRows + 1 To NoOfRows + 8
 For J = 1 To NoOfCols - 1 Step 3
  Input #FileNum, Field
  msfgResult.TextMatrix(I, J) = Field
 Next J
Next I
Close FileNum
Me.Tag = "RsZ" & ResultName
Me.Caption = ResultName & " - " & Right$(ResultFileName, Len(ResultFileName) - InStrRev(ResultFileName, "\"))
frmTestExplorer.TestTree.Nodes(Me.Tag).Image = "Result"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ResultFileName <> "" Then CancelUnload = False: Exit Sub
Select Case MsgBox("This Z-Score Result Sheet has not been saved." & Chr(13) & Space$(30) & "Save File ?", vbYesNoCancel + vbQuestion, ResultName + " not saved !")
Case vbCancel: Cancel = True
Case vbYes: If ResultFileName <> "" Then SaveResultAs ResultFileName Else fMainForm.mnuFileSaveAs_Click: If ErrorNumber = cdlCancel Then Cancel = True
End Select
CancelUnload = Cancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
With frmTestExplorer.TestTree
If ResultFileName = "" Then
 If .Nodes(Me.Tag).FirstSibling.Key = .Nodes(Me.Tag).LastSibling.Key Then .Nodes("Result").Expanded = False
 .Nodes.Remove (Me.Tag)
Else
 .Nodes(Me.Tag).Image = "ClosedResult"
End If
End With
End Sub
