VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLab 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5055
   ClientLeft      =   495
   ClientTop       =   615
   ClientWidth     =   7545
   Icon            =   "MSFlexgridForm.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   7545
   Begin MSFlexGridLib.MSFlexGrid FgCaption 
      Height          =   915
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1614
      _Version        =   393216
      Rows            =   3
      FixedRows       =   0
      BackColor       =   16777215
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox fg2Picture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   1100
      ScaleHeight     =   915
      ScaleWidth      =   1170
      TabIndex        =   2
      Top             =   960
      Width           =   1170
      Begin MSFlexGridLib.MSFlexGrid Fg2 
         Height          =   915
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   1614
         _Version        =   393216
         Rows            =   3
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         GridLines       =   2
         ScrollBars      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4020
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   315
   End
   Begin MSFlexGridLib.MSFlexGrid Fg1 
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1614
      _Version        =   393216
      Rows            =   3
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      FillStyle       =   1
      ScrollBars      =   0
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
Attribute VB_Name = "frmLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public labfilename As String
Dim samplecount() As Integer

Private Sub Form_Load()
If labfilename <> "" Then OpenLab Else NewLab
Dim I%, Coldisp%, Rowdisp%, ColWidthInc%, RowWidthInc%
With Fg1
 txtEdit.Move 0, 0, .CellWidth, .CellHeight
 .Row = 0: .Col = 0: .CellFontItalic = True
 .Move 0, 0
 If .Cols >= LabMaxColsDisplay Then Coldisp = LabMaxColsDisplay: ColWidthInc = 250: .ScrollBars = .ScrollBars Or flexScrollBarHorizontal Else Coldisp = .Cols
 If .Rows >= LabMaxRowsDisplay Then Rowdisp = LabMaxRowsDisplay: RowWidthInc = 250: .ScrollBars = .ScrollBars Or flexScrollBarVertical Else Rowdisp = .Rows
 .Width = .ColWidth(0) * Coldisp + 100 + RowWidthInc
 .Height = .RowHeight(0) * Rowdisp + 100 + ColWidthInc
 For I = 1 To .Rows - 1: .TextMatrix(I, 0) = I: Next I
 .Row = 1: .Col = 1:
End With
fg2Picture.Move Fg2.ColWidth(1) + 30, Fg1.Height + 30, Fg2.ColWidth(0) * (Fg2.Cols - 1) + 30
Fg2.Left = Fg2.Left - Fg2.ColWidth(1) - 30
Fg2.Width = Fg2.ColWidth(0) * Fg2.Cols + 100
With FgCaption
 .TextMatrix(0, 0) = "Average"
 .TextMatrix(1, 0) = "SD"
 .TextMatrix(2, 0) = "RSD"
 .ColAlignment(0) = 4
 .Width = .ColWidth(0)
 .Top = fg2Picture.Top
End With
txtEdit = ""
Me.Move Me.Left, Me.Top, Fg1.Width + 100, Fg1.Height + Fg2.Height + 500
End Sub
Private Sub fg1_DblClick()
MSFlexGridEdit 32
End Sub
Private Sub fg1_GotFocus()
If txtEdit.Visible = False Then Exit Sub
Fg1 = txtEdit
txtEdit.Visible = False
End Sub

Private Sub Fg1_KeyPress(KeyAscii As Integer)
MSFlexGridEdit KeyAscii
End Sub

Sub MSFlexGridEdit(KeyAscii As Integer)
Select Case KeyAscii
 Case 0 To 32
  txtEdit = Fg1
  txtEdit.SelStart = 0
  txtEdit.SelLength = Len(txtEdit)
 Case Else
  txtEdit = Chr(KeyAscii)
  txtEdit.SelStart = 1
End Select
txtEdit.Move Fg1.CellLeft, Fg1.CellTop
txtEdit.Visible = True
txtEdit.SetFocus
End Sub

Private Sub Fg1_LeaveCell()
If txtEdit.Visible = True Then SetAvgSDRSD: txtEdit.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim FileNum%, Field As String, NoOfRows%, NoOfCols%
Dim I As Integer, J As Integer, NotSaved As Boolean
If labfilename <> "" Then
 FileNum = FreeFile
 Open labfilename For Input As FileNum
 Input #FileNum, NoOfRows, NoOfCols
 If NoOfRows = Fg1.Rows And NoOfCols = Fg1.Cols Then
  For I = 0 To Fg1.Cols - 1
   Input #FileNum, Field: If Field <> Fg1.TextMatrix(0, I) Then NotSaved = True: I = Fg1.Cols - 1
  Next I
  If NotSaved = False Then
   For I = 1 To Fg1.Rows - 1: For J = 1 To Fg1.Cols - 1
     Input #FileNum, Field: If Field <> Fg1.TextMatrix(I, J) Then NotSaved = True: J = Fg1.Cols - 1: I = Fg1.Rows - 1
   Next J: Next I
  End If
 End If
 Close FileNum
Else
 NotSaved = True
End If
If NotSaved = False Then CancelUnload = False: Exit Sub
Select Case MsgBox("This lab file data has not been saved since the last change." & Chr(13) & Space$(35) & "Save File ?", vbYesNoCancel + vbQuestion, Fg1.TextArray(0) + " not saved !")
Case vbCancel: Cancel = True
Case vbYes: If labfilename <> "" Then SaveLabAs labfilename Else fMainForm.mnuFileSaveAs_Click: If ErrorNumber = cdlCancel Then Cancel = True
End Select
CancelUnload = Cancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
With frmTestExplorer.TestTree
If labfilename = "" Then
 If .Nodes(Me.Tag).FirstSibling.Key = .Nodes(Me.Tag).LastSibling.Key Then .Nodes("Lab").Expanded = False
 .Nodes.Remove (Me.Tag)
Else
 .Nodes(Me.Tag).Image = "ClosedLab"
End If
End With
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode Fg1, txtEdit, KeyCode, Shift
End Sub
Public Sub SetAvgSDRSD()
If IsStringANumber(txtEdit) = False And IsStringANumber(Fg1) = False Then Fg1 = txtEdit: Exit Sub
  If IsStringANumber(txtEdit) = True Then
    If IsStringANumber(Fg1) = False Then
        samplecount(Fg1.Col) = samplecount(Fg1.Col) + 1
        Fg2.TextMatrix(0, Fg1.Col) = Format(((Val(Fg2.TextMatrix(0, Fg1.Col)) * (samplecount(Fg1.Col) - 1)) + txtEdit) / samplecount(Fg1.Col), "#0.00")
    Else
        Fg2.TextMatrix(0, Fg1.Col) = Val(Fg2.TextMatrix(0, Fg1.Col)) * samplecount(Fg1.Col) - Val(Fg1)
        Fg2.TextMatrix(0, Fg1.Col) = Format((Val(Fg2.TextMatrix(0, Fg1.Col)) + txtEdit) / samplecount(Fg1.Col), "#0.00")
    End If
  Else
    If IsStringANumber(Fg1) = True Then
        Fg2.TextMatrix(0, Fg1.Col) = Fg2.TextMatrix(0, Fg1.Col) * samplecount(Fg1.Col) - Fg1
        samplecount(Fg1.Col) = samplecount(Fg1.Col) - 1
        Fg2.TextMatrix(0, Fg1.Col) = Format(Fg2.TextMatrix(0, Fg1.Col) / samplecount(Fg1.Col), "#0.00")
    End If
  End If
  Fg1 = txtEdit
  Fg2.TextMatrix(1, Fg1.Col) = Format(StandardDeviation(Fg1, Val(Fg2.TextMatrix(0, Fg1.Col)), samplecount(Fg1.Col)), "#0.00")
  Fg2.TextMatrix(2, Fg1.Col) = Format(RSD(Val(Fg2.TextMatrix(1, Fg1.Col)), Val(Fg2.TextMatrix(0, Fg1.Col))), "#0.00")
End Sub
Sub EditKeyCode(msfg As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
 Case 27
  Edt.Visible = False
  msfg.SetFocus
 Case 13
  SetAvgSDRSD
  If Fg1.Row < Fg1.Rows - 1 Then
   Fg1.Row = Fg1.Row + 1
  Else
   Fg1.Rows = Fg1.Rows + 1
   Fg1.Row = Fg1.Row + 1
   Fg1.TextMatrix(Fg1.Row, 0) = Fg1.Rows - 1
   If Fg1.Rows = LabMaxRowsDisplay + 1 Then
      Fg1.ScrollBars = Fg1.ScrollBars Or flexScrollBarVertical
      Fg1.Width = Fg1.Width + 250
      Me.Move Me.Left, Me.Top, Fg1.Width + 100
    End If
   If Fg1.Rows < LabMaxRowsDisplay + 1 Then
      Fg1.Height = Fg1.Height + Fg1.CellHeight + 15
      fg2Picture.Top = fg2Picture.Top + txtEdit.Height + 15
      FgCaption.Top = fg2Picture.Top
      Me.Move Me.Left, Me.Top, Me.Width, Fg1.Height + Fg2.Height + 500
   End If
  End If
 Fg1.SetFocus
 Case 38
  SetAvgSDRSD
  If msfg.Row > msfg.FixedRows Then msfg.Row = msfg.Row - 1
  msfg.SetFocus
 Case 40
  SetAvgSDRSD
  If msfg.Row < msfg.Rows - 1 Then msfg.Row = msfg.Row + 1
  msfg.SetFocus
End Select
End Sub

Public Sub InsertSample(Optional samplename As String)
If samplename = "" Then samplename = InputBox("Sample name ?", "Sample name", "Sample " & Fg2.Cols)
If samplename = "" Then Exit Sub
With Fg1
 ReDim Preserve samplecount(1 To .Cols) As Integer
 .Cols = .Cols + 1
 .Row = 1: .Col = .Cols - 1
 .TextMatrix(0, .Col) = samplename
 .ColAlignment(.Col) = flexAlignCenterCenter
End With
With Fg2
 .Cols = .Cols + 1
 .ColAlignment(.Cols - 1) = flexAlignCenterCenter
End With
If Fg1.Cols = LabMaxColsDisplay + 1 Then
 Fg1.ScrollBars = Fg1.ScrollBars Or flexScrollBarHorizontal
 Fg1.Height = Fg1.Height + 275:
 Me.Height = Fg1.Height + Fg2.Height + 500
 fg2Picture.Top = fg2Picture.Top + 275
 FgCaption.Top = fg2Picture.Top
End If
If Fg1.Cols >= LabMaxColsDisplay + 1 Then Fg1.LeftCol = Fg1.LeftCol + 1
If Fg1.Cols < LabMaxColsDisplay + 1 Then
 Fg1.Width = Fg1.Width + Fg1.ColWidth(Fg1.Col)
 Fg2.Width = Fg1.Width
 fg2Picture.Width = Fg2.Width - 30
 Me.Move Me.Left, Me.Top, Fg1.Width + 100
End If
Fg1.Row = 1: Fg1.Col = Fg1.Cols - 1: Fg1.SetFocus
End Sub

Public Sub SaveLabAs(FileName As String)
Dim FileNum As Integer
FileNum = FreeFile
Open FileName For Output As FileNum
Write #FileNum, Fg1.Rows, Fg1.Cols
Dim I As Integer, J As Integer
For I = 0 To Fg1.Cols - 2: Write #FileNum, Fg1.TextMatrix(0, I),: Next I: Write #FileNum, Fg1.TextMatrix(0, I)
For I = 1 To Fg1.Rows - 1
 For J = 1 To Fg1.Cols - 2
  Write #FileNum, Fg1.TextMatrix(I, J),
 Next J
 Write #FileNum, Fg1.TextMatrix(I, J)
Next I
For I = 0 To Fg2.Rows - 1
 For J = 1 To Fg2.Cols - 2
  Write #FileNum, Fg2.TextMatrix(I, J),
 Next J
 Write #FileNum, Fg2.TextMatrix(I, J)
Next I
Close FileNum
End Sub
Private Sub NewLab()
ReDim samplecount(1 To 1) As Integer
Fg1.ColAlignment(0) = 4: Fg1.ColAlignment(1) = 4: Fg1.TextMatrix(0, 1) = "Sample 1": Fg2.ColAlignment(1) = 4
End Sub
Private Sub OpenLab()
Dim FileNum As Integer, NoOfRows As Integer, NoOfCols As Integer, I As Integer, J As Integer, Field As String
FileNum = FreeFile
Open labfilename For Input As FileNum
Input #FileNum, NoOfRows, NoOfCols
ReDim samplecount(1 To NoOfCols - 1) As Integer
Fg1.Cols = NoOfCols: Fg1.Rows = NoOfRows: Fg2.Cols = NoOfCols
For I = 0 To NoOfCols - 1: Input #FileNum, Field: Fg1.TextMatrix(0, I) = Field: Fg1.ColAlignment(I) = 4: Next I
For I = 1 To NoOfRows - 1
 For J = 1 To NoOfCols - 1
  Input #FileNum, Field
  Fg1.TextMatrix(I, J) = Field
  If IsStringANumber(Field) = True Then samplecount(J) = samplecount(J) + 1
 Next J
Next I
For I = 0 To Fg2.Rows - 1
 For J = 1 To Fg2.Cols - 1
  Input #FileNum, Field
  Fg2.TextMatrix(I, J) = Field
 Next J
Next I
For J = 1 To NoOfCols - 1: Fg1.ColAlignment(J) = 4: Fg2.ColAlignment(J) = 4: Next J
Close FileNum
Me.Tag = "Lab" & Me.Fg1.TextArray(0)
frmTestExplorer.TestTree.Nodes(Me.Tag).Image = "Lab"
End Sub
Public Sub RefreshValues()
Dim I As Integer, J As Integer, Avg As Single, AvgArray() As Single
ReDim AvgArray(1 To Fg1.Rows - 1)
For I = 1 To Fg1.Cols - 1
 samplecount(I) = 1
 For J = 1 To Fg1.Rows - 1
  If IsStringANumber(Fg1.TextMatrix(J, I)) = True Then AvgArray(samplecount(I)) = CSng(Fg1.TextMatrix(J, I)): samplecount(I) = samplecount(I) + 1
 Next J
 If samplecount(I) > 1 Then
  ReDim Preserve AvgArray(1 To samplecount(I) - 1) As Single
  Fg2.TextMatrix(0, I) = Format(Average(AvgArray), "#0.00")
  Fg2.TextMatrix(1, I) = Format(SD(AvgArray, CSng(Fg2.TextMatrix(0, I))), "#0.00")
  Fg2.TextMatrix(2, I) = Format(RSD(CSng(Fg2.TextMatrix(1, I)), CSng(Fg2.TextMatrix(0, I))), "#0.00")
 End If
Next I
End Sub

Private Sub Fg1_Scroll()
If Fg1.ScrollBars = flexScrollBarHorizontal Then Fg2.Left = -Fg1.LeftCol * Fg2.ColWidth(0): Fg2.Width = Fg2.Width + Fg2.ColWidth(0)
End Sub
Private Sub Fg1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu fMainForm.mnuEdit
End Sub

