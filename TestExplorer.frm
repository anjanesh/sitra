VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestExplorer 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Test - "
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList TreeImages 
      Left            =   2160
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestExplorer.frx":0000
            Key             =   "CloseFolder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestExplorer.frx":0454
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestExplorer.frx":08A8
            Key             =   "Lab"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestExplorer.frx":0A08
            Key             =   "Graph"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestExplorer.frx":0E5C
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestExplorer.frx":12B0
            Key             =   "Result"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestExplorer.frx":1704
            Key             =   "ClosedLab"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestExplorer.frx":1864
            Key             =   "ClosedResult"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TestTree 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2566
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "TreeImages"
      Appearance      =   1
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
Attribute VB_Name = "frmTestExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Me.Caption = "Test - " & fMainForm.testname
 Me.Width = fMainForm.Width / 4.67: Me.Height = fMainForm.Height / 1.5
 Me.Left = fMainForm.Width - Me.Width - 175
 Me.Top = 0
 With TestTree.Nodes
 .Add , , "Test", fMainForm.testname + " (<Not Saved>)", "OpenFolder"
 .Add "Test", tvwChild, "Lab", "Labs", "OpenFolder"
 .Add "Test", tvwChild, "Graph", "Graphs", "CloseFolder"
 .Add "Test", tvwChild, "Result", "Results", "CloseFolder"
 .Add "Test", tvwChild, "Misc", "Miscellaneous", "CloseFolder"
 End With
 TestTree.Nodes("Lab").EnsureVisible
 TestTree.Nodes("Lab").Expanded = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormMDIForm Or ExitApp = True Or TestOpen = False Then Exit Sub
Cancel = True
Me.Visible = False
End Sub

Private Sub Form_Resize()
TestTree.Width = Me.Width - 125
TestTree.Height = Me.Height - 125 - 225
TestTree.Refresh
End Sub

Private Sub TestTree_Collapse(ByVal Node As MSComctlLib.Node)
 Node.Image = "CloseFolder"
End Sub

Public Sub TestTree_DblClick()
Dim I As Integer
Select Case TestTree.SelectedItem.Image
Case "Lab", "Result"
 For I = 1 To Forms.Count - 1
  If Forms(I).Tag = TestTree.SelectedItem.Key Then Forms(I).SetFocus: I = Forms.Count - 1
 Next I
Case "ClosedLab"
 fMainForm.LoadExistingDoc TestTree.SelectedItem.Tag
Case "ClosedResult"
 fMainForm.LoadExistingResult TestTree.SelectedItem.Tag
End Select
End Sub

Private Sub TestTree_Expand(ByVal Node As MSComctlLib.Node)
 Node.Image = "OpenFolder"
End Sub

Private Sub TestTree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Or TestTree.SelectedItem.Index < 6 Then Exit Sub
Select Case Left$(TestTree.SelectedItem.Key, 3)
Case "Lab"
 fMainForm.ExplorerLab(3).Caption = "&Copy Lab"
 fMainForm.ExplorerLab(3).Caption = "&Save Lab"
 fMainForm.ExplorerLab(4).Caption = "Save Lab &As..."
 fMainForm.ExplorerLab(5).Caption = "&Remove Lab"
Case "RsZ"
 fMainForm.ExplorerLab(3).Caption = "&Copy Result"
 fMainForm.ExplorerLab(3).Caption = "&Save Result"
 fMainForm.ExplorerLab(4).Caption = "Save Result &As..."
 fMainForm.ExplorerLab(5).Caption = "&Remove Result"
End Select
PopupMenu fMainForm.FloatMenus
End Sub
