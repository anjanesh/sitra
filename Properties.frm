VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties - "
   ClientHeight    =   6540
   ClientLeft      =   1335
   ClientTop       =   1020
   ClientWidth     =   9015
   Icon            =   "Properties.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   1
      Left            =   240
      ScaleHeight     =   5100
      ScaleWidth      =   8325
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   8325
      Begin VB.TextBox txtSummary 
         Height          =   5055
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "Properties.frx":000C
         Top             =   0
         Width           =   8295
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   5100
      ScaleWidth      =   8325
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   8325
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   5100
      ScaleWidth      =   8325
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   8325
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   0
      Left            =   240
      ScaleHeight     =   5100
      ScaleWidth      =   8325
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   8325
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtCollab 
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtCoord 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtOrganized 
         Height          =   975
         Left            =   4440
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   550
         Width           =   2295
      End
      Begin VB.TextBox txtTestDesp 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtTestName 
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Date :"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Collaborator :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Co-Ordinator :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Organized by :"
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Test Desciption :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Test Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   6015
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   6015
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   6015
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   5685
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10028
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            Object.ToolTipText     =   "Who, Where & When Specifications"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Summary"
            Key             =   "Summary"
            Object.ToolTipText     =   "Summary of the project"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Methods"
            Key             =   "Methods"
            Object.ToolTipText     =   "Methods adopted by laboratories"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Participants"
            Key             =   "Participants"
            Object.ToolTipText     =   "List of Participants who have sent the results"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
Select Case tbsOptions.SelectedItem.Key
 Case "General"
  If txtTestName <> "" Then
   fMainForm.Caption = "Proficiency Test - " & txtTestName & " - "
   frmTestExplorer.TestTree.Nodes(1).Text = txtTestName
   If fMainForm.testfilename = "" Then
    fMainForm.Caption = fMainForm.Caption + "<Not Saved>"
    frmTestExplorer.TestTree.Nodes(1).Text = frmTestExplorer.TestTree.Nodes(1).Text + " (<Not Saved>)"
   Else
    fMainForm.Caption = fMainForm.Caption + fMainForm.testfilename
    frmTestExplorer.TestTree.Nodes(1).Text = frmTestExplorer.TestTree.Nodes(1).Text + " (" + Right$(fMainForm.testfilename, Len(fMainForm.testfilename) - InStrRev(fMainForm.testfilename, "\")) + ")"
   End If
  End If
 Case "Summary"
  If txtSummary <> "Enter Summary Here" Then SaveSummary
End Select
End Sub
Public Sub SaveSummary()
If fMainForm.testfilename = "" Then MsgBox "Need to save test before creating a summary", vbInformation, "Cannot save test Summary:": Exit Sub
Dim FileNum%
FileNum = FreeFile
Open "Test Summary of " + fMainForm.testname + ".txt" For Output As FileNum
Print #FileNum, txtSummary
Close FileNum
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        I = tbsOptions.SelectedItem.Index
        If I = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(I + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 txtSummary.SelStart = 0
 txtSummary.SelLength = Len(txtSummary) - 1
End Sub

Private Sub tbsOptions_Click()
    Dim I As Integer
    For I = 0 To tbsOptions.Tabs.Count - 1
        If I = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(I).Visible = True
            If I = 1 Then txtSummary.SetFocus
        Else
            picOptions(I).Visible = False
        End If
    Next I
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ExitApp = True Or TestOpen = False Then Exit Sub
Cancel = True
Me.Visible = False
End Sub

