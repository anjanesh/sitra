VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResultDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Result Sheet Generator"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "ResultDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFormulae 
      Caption         =   "Show"
      Height          =   2775
      Left            =   3480
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
      Begin VB.CheckBox chkFormula 
         Caption         =   "NIQR"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Normal Inter Quartile Range"
         Top             =   2400
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "SIQR"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Semi Inter Quartile Range"
         Top             =   2160
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "IQR"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Inter Quartile Range"
         Top             =   1920
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "Q2"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "3rdt Quartile"
         Top             =   1680
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "Q1"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "1st Quartile"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "RSD"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Relative Standard Deviation"
         Top             =   1200
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "SD"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Standard Deviation"
         Top             =   960
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "Mode"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "Median"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Median or 2nd Quartile"
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkFormula 
         Caption         =   "Average"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Mean"
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CheckBox CheckGraph 
      Caption         =   "&Generate Graph"
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton CommandGen 
      Caption         =   "&Generate Result"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   365
      Left            =   840
      TabIndex        =   3
      Top             =   75
      Width           =   2535
   End
   Begin VB.CheckBox chkSelectAll 
      Caption         =   "&Select All / None"
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin MSComctlLib.TreeView TreeViewCheck 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8281
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "TreeImages"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList TreeImages 
      Left            =   7560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResultDialog.frx":000C
            Key             =   "Lab"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3600
      Picture         =   "ResultDialog.frx":016C
      Top             =   120
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmResultDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1 = "Result " & fMainForm.ResultCount + 1
Dim I As Integer
For I = 5 To frmTestExplorer.TestTree.Nodes.Count
 If frmTestExplorer.TestTree.Nodes(I).Parent.Key = "Lab" And frmTestExplorer.TestTree.Nodes(I).Tag <> "" Then
  TreeViewCheck.Nodes.Add , tvwNext, frmTestExplorer.TestTree.Nodes(I).Tag, frmTestExplorer.TestTree.Nodes(I).Text, "Lab"
 End If
Next I
If TreeViewCheck.Nodes.Count = 0 Then
 MsgBox "Only SAVED data can be used to generate a report", vbInformation, "No Saved Labs available !"
 Unload Me
End If
End Sub

Private Sub chkSelectAll_Click()
Dim I As Integer
For I = 1 To TreeViewCheck.Nodes.Count: TreeViewCheck.Nodes(I).Checked = chkSelectAll.Value: Next I
End Sub

Private Sub CommandGen_Click()
If Text1 = "" Then MsgBox "Must Specify Name", vbInformation, "Name Required": Exit Sub
Dim I%, Total%
For I = 5 To TreeViewCheck.Nodes.Count
 If TreeViewCheck.Nodes(I).Checked = True Then Total = Total + 1
Next I
If Total < 3 Then MsgBox "Minimum 3 Labs required to generate a result.", vbInformation, "Not Enough Labs": Exit Sub
for i=
Me.Visible = False
fMainForm.LoadNewResult Text1
End Sub

Private Sub TreeViewCheck_NodeClick(ByVal Node As MSComctlLib.Node)
Node.Checked = Not (Node.Checked)
End Sub
