VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGraph 
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "Graph.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   9600
   Begin VB.Frame Frame2 
      Caption         =   "Legend"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   8040
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
      Begin VB.Label Label2 
         Caption         =   "Zr"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Zc"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   120
         X2              =   480
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   120
         X2              =   480
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Markers"
      Height          =   1335
      Left            =   8040
      TabIndex        =   2
      Top             =   960
      Width           =   1335
      Begin VB.OptionButton OptionMarker 
         Caption         =   "&Both"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton OptionMarker 
         Caption         =   "&Lines"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton OptionMarker 
         Caption         =   "&Points"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin MSChart20Lib.MSChart ZScoreGraph 
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "Graph.frx":0442
      TabIndex        =   1
      Top             =   1080
      Width           =   7695
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8493
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Graphical Representation of Z-Score  Analysis"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Public GraphName As String, GraphFileName As String
Dim ChartArray() As Variant
Dim SampleNames() As String
Private Sub Form_Load()
Me.Move 0, 0, 9720, 5880
Line1.BorderColor = RGB(250, 20, 35): Line2.BorderColor = RGB(60, 200, 60)
OpenResultFile
GraphName = frmGraphDialog.Text1
ZScoreGraph = ChartArray
Dim I%
TabStrip1.Tabs(1).Caption = SampleNames(1)
For I = 2 To UBound(SampleNames)
 TabStrip1.Tabs.Add I, , SampleNames(I)
Next I

With ZScoreGraph.Plot.SeriesCollection
 For I = 1 To UBound(SampleNames) * 2 Step 2
  .Item(I).ShowLine = False: .Item(I + 1).ShowLine = False
  .Item(I).Pen.Width = 1: .Item(I + 1).Pen.Width = 1
  .Item(I).Pen.VtColor.Set 250, 20, 35
  .Item(I + 1).Pen.VtColor.Set 60, 200, 60
 Next I
.Item(1).ShowLine = True
.Item(2).ShowLine = True
End With

With ZScoreGraph.Plot.SeriesCollection(UBound(SampleNames) * 2 + 1)
 .ShowLine = False
 .StatLine.VtColor.Set 128, 128, 255
 .StatLine.Flag = VtChStatsMaximum
 .StatLine.Style(VtChStatsMaximum) = VtPenStyleSolid
 .StatLine.Width = 1
End With

With ZScoreGraph.Plot.SeriesCollection(UBound(SampleNames) * 2 + 2)
 .ShowLine = False
 .StatLine.VtColor.Set 128, 128, 255
 .StatLine.Flag = VtChStatsMaximum Or VtChStatsMinimum
 .StatLine.Style(VtChStatsMaximum) = VtPenStyleSolid
 .StatLine.Style(VtChStatsMinimum) = VtPenStyleSolid
 .StatLine.Width = 1
End With

Unload frmGraphDialog
Me.Caption = GraphName & " - <No Saved>"
End Sub

Public Sub OpenResultFile()
Dim FileNum As Integer, I%, J%, K%, NoOfRows%, NoOfCols%, Field As String
FileNum = FreeFile
Open frmGraphDialog.FileName For Input As FileNum
Line Input #FileNum, Field
If Field <> "ZScore" Then MsgBox "Not a valid Result Sheet File.", vbExclamation, "Error Opening File !": Exit Sub
Input #FileNum, NoOfRows, NoOfCols, Field
ReDim SampleNames((NoOfCols - 1) / 3)
ReDim ChartArray(NoOfRows, ((NoOfCols - 1) / 3) * 2 + 3)
For I = 1 To (NoOfCols - 1) / 3
 Input #FileNum, SampleNames(I)
Next I
For I = 1 To NoOfRows
 K = 1
 For J = 1 To NoOfCols
  Input #FileNum, Field
  If (J + 1) / 3 <> Int((J + 1) / 3) Then ChartArray(I, K) = Field: K = K + 1
 Next J
Next I
Close FileNum
Dim GridLines() As Integer
ReDim GridLines(2)
GridLines(1) = frmGraphDialog.TextLOutlier: GridLines(2) = frmGraphDialog.TextUOutlier
For I = 1 To NoOfRows
 For J = 1 To 2
  ChartArray(I, ((NoOfCols - 1) / 3) * 2 + 1 + J) = GridLines(J)
 Next J
Next I
ChartArray(1, ((NoOfCols - 1) / 3) * 2 + 3) = 0
End Sub

Private Sub OptionMarker_Click(Index As Integer)
ShowMarkers TabStrip1.SelectedItem.Index * 2 - 1
End Sub

Private Sub TabStrip1_Click()
Dim I%
For I = 1 To UBound(SampleNames) * 2
 ZScoreGraph.Plot.SeriesCollection(I).ShowLine = False
 ZScoreGraph.Plot.SeriesCollection(I).SeriesMarker.Show = False
Next I
ShowMarkers TabStrip1.SelectedItem.Index * 2 - 1
End Sub

Private Sub ShowMarkers(Index As Integer)
Dim I%, OpValue%
For I = 0 To 2
If OptionMarker(I).Value = True Then OpValue = I: I = 3
Next I
With ZScoreGraph.Plot.SeriesCollection
Select Case OpValue
Case 0: ' Points
.Item(Index).SeriesMarker.Show = True
.Item(Index + 1).SeriesMarker.Show = True
.Item(Index).ShowLine = False
.Item(Index + 1).ShowLine = False

'.Item(Index).DataPoints(-1).Marker.Visible = True
'.Item(Index).DataPoints(-1).Marker.Style = VtMarkerStyleFilledCircle
'.Item(Index).DataPoints(-1).Marker.Size = 1500
'.Item(Index).DataPoints(-1).Brush.Style = VtBrushStylePattern


Case 1: ' Lines
.Item(Index).SeriesMarker.Show = False
.Item(Index + 1).SeriesMarker.Show = False
.Item(Index).ShowLine = True
.Item(Index + 1).ShowLine = True
Case 2: 'Both
.Item(Index).SeriesMarker.Show = True
.Item(Index + 1).SeriesMarker.Show = True
.Item(Index).ShowLine = True
.Item(Index + 1).ShowLine = True
End Select
End With
End Sub
