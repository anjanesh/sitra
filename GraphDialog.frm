VERSION 5.00
Begin VB.Form frmGraphDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graph Generator"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "GraphDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Outliers"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   5175
      Begin VB.TextBox TextUOutlier 
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TextLOutlier 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Upper Bound :"
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Lower Bound :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   365
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox comboResults 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate Graph"
      Default         =   -1  'True
      Height          =   465
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Select a Result Sheet"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   615
   End
End
Attribute VB_Name = "frmGraphDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FileName As String
Private Sub Form_Load()
Text1 = "Graph " & fMainForm.GraphCount + 1
Dim I As Integer
For I = 5 To frmTestExplorer.TestTree.Nodes.Count
 If frmTestExplorer.TestTree.Nodes(I).Parent.Key = "Result" And frmTestExplorer.TestTree.Nodes(I).Tag <> "" Then
  comboResults.AddItem frmTestExplorer.TestTree.Nodes(I).Text
 End If
Next I
If comboResults.ListCount = 0 Then
 MsgBox "Only SAVED data can be used to generate a graph", vbInformation, "No Saved Reports available !"
 Unload Me
Else
 comboResults.ListIndex = 0
End If
End Sub
Private Sub CommandCancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Me.Visible = False
Dim I As Integer
With frmTestExplorer.TestTree
For I = 5 To .Nodes.Count
 If .Nodes(I).Parent.Key = "Result" And .Nodes(I).Text = comboResults.Text Then FileName = .Nodes(I).Tag: I = .Nodes.Count + 1
Next I
End With
Dim frmG As frmGraph
Set frmG = New frmGraph
frmG.Show
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0: Text1.SelLength = Len(Text1)
End Sub

Private Sub TextLOutlier_GotFocus()
TextLOutlier.SelStart = 0: TextLOutlier.SelLength = Len(TextLOutlier)
End Sub
Private Sub TextUOutlier_GotFocus()
TextUOutlier.SelStart = 0: TextUOutlier.SelLength = Len(TextUOutlier)
End Sub

Private Sub TextLOutlier_KeyPress(KeyAscii As Integer)
OutlierValidation KeyAscii
End Sub
Private Sub TextUOutlier_KeyPress(KeyAscii As Integer)
OutlierValidation KeyAscii
End Sub

Private Sub OutlierValidation(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
'If KeyAscii = 45 and textloutlier.
If KeyAscii < 45 Or KeyAscii > 57 Or KeyAscii = 47 Then KeyAscii = 0
End Sub
