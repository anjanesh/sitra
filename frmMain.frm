VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   6105
   ClientLeft      =   210
   ClientTop       =   720
   ClientWidth     =   8880
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Add Lab"
            ImageKey        =   "NewLab"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ExistingLab"
                  Text            =   "Add Existing Lab"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Test"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveTest"
            Object.ToolTipText     =   "Save Test"
            ImageKey        =   "SaveAll"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Object.ToolTipText     =   "Close Test"
            ImageKey        =   "Close"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Explorer"
            Object.ToolTipText     =   "Test Explorer"
            ImageKey        =   "Explorer"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sample"
            Object.ToolTipText     =   "Insert Sample Column"
            ImageKey        =   "Sample"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Result"
            Object.ToolTipText     =   "Generate a Result Sheet"
            ImageKey        =   "Result"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ExistingResult"
                  Text            =   "Add Existing Result"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Graph"
            Object.ToolTipText     =   "Generate a Graph"
            ImageKey        =   "Graph"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Calculator"
            Object.ToolTipText     =   "Microsoft Calculator"
            ImageKey        =   "Calculator"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Website"
            Object.ToolTipText     =   "Sitra Online : www.sitraindia.org"
            ImageKey        =   "Website"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1320
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0890
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A2
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C76
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DD2
            Key             =   "Website"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1226
            Key             =   "Lab"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1386
            Key             =   "SaveAll"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E2
            Key             =   "Explorer"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":163E
            Key             =   "Graph"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17A2
            Key             =   "Result"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BF6
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D52
            Key             =   "NewLab"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22EE
            Key             =   "Calculator"
         EndProperty
      EndProperty
   End
   Begin VB.Menu FloatMenus 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu ExplorerLab 
         Caption         =   "&View"
         Index           =   0
      End
      Begin VB.Menu ExplorerLab 
         Caption         =   "P&roperties"
         Index           =   1
      End
      Begin VB.Menu ExplorerLab 
         Caption         =   "&Copy"
         Index           =   2
      End
      Begin VB.Menu ExplorerLab 
         Caption         =   "&Save"
         Index           =   3
      End
      Begin VB.Menu ExplorerLab 
         Caption         =   "Save &As"
         Index           =   4
      End
      Begin VB.Menu ExplorerLab 
         Caption         =   "&Remove"
         Index           =   5
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Test"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Test"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close Test"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveTest 
         Caption         =   "Save &Test"
      End
      Begin VB.Menu mnuFileBarSaveProject 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Test Propert&ies"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditCopyLab 
         Caption         =   "Cop&y Lab"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSample 
         Caption         =   "&Sample Name"
      End
      Begin VB.Menu mnuEditLab 
         Caption         =   "&Lab Name"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim lLabCount As Integer
Public ResultCount As Integer, GraphCount As Integer
Public testfilename As String, testname As String

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    Dim ScreenResolution As String
    ScreenResolution = (Screen.Width / Screen.TwipsPerPixelX) & "x" & (Screen.Height / Screen.TwipsPerPixelY)
    Select Case ScreenResolution
     Case "1024x768"
      LabMaxRowsDisplay = 20: LabMaxColsDisplay = 9
     Case "800x600"
      LabMaxRowsDisplay = 15: LabMaxColsDisplay = 7
     Case "640x480"
      LabMaxRowsDisplay = 10: LabMaxColsDisplay = 5
    End Select
    LoadNewTest
    End Sub
Private Sub LoadNewTest()
If TestOpen = True Then Unload Me
If CancelUnload = True Then Exit Sub
lLabCount = 0: ResultCount = 0
TestOpen = True
testname = "Untitled"
Me.Caption = "Proficiency Test - " & testname & " - <Not Saved>"
Load frmTestExplorer ' frmTestExplorer is the 2nd form that has to be loaded & 2nd last to be unloaded
Load frmProperties
LoadNewDoc
Dim I As Integer
For I = 1 To tbToolBar.Buttons.Count: tbToolBar.Buttons(I).Enabled = True: Next I
End Sub
Public Sub LoadExistingTest(FileName As String)
Dim FileNum As Integer, NoofLabs As Integer, I As Integer
Dim Field1 As String, Field2 As String, Field3 As String
FileNum = FreeFile
testfilename = FileName
Open FileName For Input As FileNum
Input #FileNum, testname
frmTestExplorer.TestTree.Nodes(1).Text = testname + " (" + Right$(testfilename, Len(testfilename) - InStrRev(testfilename, "\")) + ")"
While Not EOF(FileNum)
 Input #FileNum, Field1, Field2, Field3
 Select Case Field1
 Case "Lab"
  frmTestExplorer.TestTree.Nodes.Add "Lab", tvwChild, "Lab" & Field2, Field2 + " (" + Right$(Field3, Len(Field3) - InStrRev(Field3, "\")) + ")", "ClosedLab"
  frmTestExplorer.TestTree.Nodes("Lab" & Field2).Tag = Field3
 Case "ResultZScore"
  frmTestExplorer.TestTree.Nodes.Add "Result", tvwChild, "RsZ" & Field2, Field2 + " (" + Right$(Field3, Len(Field3) - InStrRev(Field3, "\")) + ")", "ClosedResult"
  frmTestExplorer.TestTree.Nodes("RsZ" & Field2).Tag = Field3
 End Select
Wend
Close FileNum
lLabCount = 0: ResultCount = 0
TestOpen = True
Me.Caption = "Proficiency Test - " & testname & " - " & testfilename
For I = 1 To tbToolBar.Buttons.Count: tbToolBar.Buttons(I).Enabled = True: Next I
frmTestExplorer.Show
Load frmProperties
End Sub
Private Sub LoadNewDoc()
    Dim frmD As frmLab
    lLabCount = lLabCount + 1
    Set frmD = New frmLab
    frmD.Fg1.TextArray(0) = "Lab " & lLabCount
    frmD.Caption = "<Not Saved>"
    If lLabCount = 1 Then frmD.Move 0, 0
    frmD.Tag = "Lab" & frmD.Fg1.TextArray(0) ' The key is Lab and its lab name for example : LabSitraLab211
    frmTestExplorer.TestTree.Nodes.Add "Lab", tvwChild, frmD.Tag, frmD.Fg1.TextArray(0) + " (" + frmD.Caption + ")", "Lab"
    frmD.Show
End Sub
Public Sub LoadExistingDoc(FileName As String)
 Dim frmD As frmLab
 Set frmD = New frmLab
 frmD.labfilename = FileName
 frmD.Caption = Right$(frmD.labfilename, Len(frmD.labfilename) - InStrRev(frmD.labfilename, "\"))
End Sub
Public Sub LoadExistingResult(FileName As String)
 Dim frmD As frmResult
 Set frmD = New frmResult
 frmD.ResultFileName = FileName
 frmD.Show
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If TestOpen = False Then Exit Sub
Dim FileNum As Integer, Field As String, NoofLabs As Integer
Dim I As Integer, J As Integer, NotSaved As Boolean
If testfilename <> "" Then
 FileNum = FreeFile
 Open testfilename For Input As FileNum
 Input #FileNum, Field, NoofLabs
 If Field <> testname Then NotSaved = True
 'For I = 5 To frmTestExplorer.TestTree.Nodes.Count
  'If Left$(frmTestExplorer.TestTree.Nodes(I).Key,3) = "Lab" Then
 'Next I
 Close FileNum
Else
 NotSaved = True
End If
If NotSaved = True Then
 Select Case MsgBox("Save Test ?", vbYesNoCancel Or vbQuestion, "Closing Test")
 Case vbCancel: Cancel = True
 Case vbYes: mnuFileSaveTest_Click: If ErrorNumber = cdlCancel Then Cancel = True
 End Select
End If
If Cancel = True Then Exit Sub
For I = 1 To Forms.Count - 3
Unload Forms(3)
If CancelUnload = True Then Exit For
Next I
If CancelUnload = True Then Cancel = True: Exit Sub
TestOpen = False
If UnloadMode = vbFormCode And ExitApp = False Then
 Unload frmTestExplorer: Unload frmProperties
 testname = "untitled"
 testfilename = ""
 Me.Caption = "Proficiency Test"
 Cancel = True
End If

For I = 1 To tbToolBar.Buttons.Count - 3: tbToolBar.Buttons(I).Enabled = False: Next I
tbToolBar.Buttons("Open").Enabled = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload frmProperties
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuEdit_Click()
If ActiveForm Is Nothing Then Exit Sub
Select Case ActiveForm.Name
 Case "frmLab"
  mnuEditSample.Enabled = True
  mnuEditLab.Enabled = True
  mnuEditCopyLab.Caption = "Cop&y Lab"
  mnuEditCopyLab.Enabled = True
 Case "frmResult"
  mnuEditSample.Enabled = False
  mnuEditLab.Enabled = False
  mnuEditCopyLab.Caption = "Cop&y Result"
  mnuEditCopyLab.Enabled = True
 Case Else
  mnuEditSample.Enabled = False
  mnuEditLab.Enabled = False
  mnuEditCopyLab.Enabled = False
End Select
End Sub
Private Sub mnuFile_Click()
If ActiveForm Is Nothing Then Exit Sub
Select Case ActiveForm.Name
 Case "frmLab"
  mnuFileSave.Caption = "&Save Lab"
  mnuFileSaveAs.Caption = "Save Lab &As.."
 Case "frmResult"
  mnuFileSave.Caption = "&Save Result"
  mnuFileSaveAs.Caption = "Save Result &As.."
 Case "frmTestExplorer"
  mnuFileSave.Caption = "&Save "
  mnuFileSaveAs.Caption = "Save &As.."
End Select
End Sub

Private Sub mnuEditCopyLab_Click()
If ActiveForm Is Nothing Then Exit Sub
Clipboard.Clear
Dim I As Integer, J As Integer, ClipString As String
Select Case ActiveForm.Name
 Case "frmLab"
  For I = 0 To ActiveForm.Fg1.Rows - 1
   For J = 0 To ActiveForm.Fg1.Cols - 2
    ClipString = ClipString + ActiveForm.Fg1.TextMatrix(I, J) + vbTab
   Next J
   ClipString = ClipString + ActiveForm.Fg1.TextMatrix(I, J) + vbCrLf
  Next I
  For I = 0 To ActiveForm.Fg2.Rows - 1
   For J = 0 To ActiveForm.Fg2.Cols - 2
    If J = 0 Then
     Select Case I
      Case 0: ClipString = ClipString + "Average"
      Case 1: ClipString = ClipString + "RD"
      Case 2: ClipString = ClipString + "RSD"
     End Select
     ClipString = ClipString + vbTab
    Else
     ClipString = ClipString + ActiveForm.Fg2.TextMatrix(I, J) + vbTab
    End If
   Next J
   ClipString = ClipString + ActiveForm.Fg2.TextMatrix(I, J) + vbCrLf
  Next I
 Case "frmResult"
  For I = 0 To ActiveForm.msfgResult.Rows - 1
   For J = 0 To ActiveForm.msfgResult.Cols - 2
    ClipString = ClipString + ActiveForm.msfgResult.TextMatrix(I, J) + vbTab
   Next J
   ClipString = ClipString + ActiveForm.msfgResult.TextMatrix(I, J) + vbCrLf
  Next I
End Select
Clipboard.SetText ClipString
End Sub

Private Sub mnuEditLab_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name <> "frmLab" Then Exit Sub
Dim labname As String
With ActiveForm
 labname = InputBox("Lab name ?", "Lab name", .Fg1.TextArray(0))
 If labname <> "" Then
  .Fg1.TextArray(0) = labname
  frmTestExplorer.TestTree.Nodes(.Tag).Text = labname & " (" & .Caption & ")"
  frmTestExplorer.TestTree.Nodes(.Tag).Key = "Lab" & labname
  .Tag = "Lab" & labname
 End If
End With
End Sub

Private Sub mnuEditSample_Click()
If ActiveForm.Name <> "frmLab" Then Exit Sub
Dim samplename As String
samplename = InputBox("Sample name ?", "Sample name", ActiveForm.Fg1.TextMatrix(0, ActiveForm.Fg1.Col))
If samplename <> "" Then ActiveForm.Fg1.TextMatrix(0, ActiveForm.Fg1.Col) = samplename
End Sub

Private Sub mnuFileSaveTest_Click()
On Error GoTo ErrHandler
If testfilename = "" Then
    With dlgCommonDialog
        .DialogTitle = "Save Test As..."
        .CancelError = True
        .Filter = "Test Files (*.tst)|*.tst"
        .Flags = cdlOFNOverwritePrompt
        .FileName = "test1"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        testfilename = .FileName
        Me.Caption = "Proficiency Test - " & testname & " - " & testfilename
        frmTestExplorer.TestTree.Nodes(1).Text = testname + " (" + .FileTitle + " )"
        frmTestExplorer.TestTree.Nodes(1).Tag = .FileName
    End With
End If
Dim FileNum As Integer, I As Integer, LName As String, FTitle As String
FileNum = FreeFile
Open testfilename For Output As FileNum
Write #FileNum, testname
For I = 3 To Forms.Count - 1
Select Case Forms(I).Name
 Case "frmLab"
  If Forms(I).labfilename = "" Then mnuFileSaveAs_Click Else Forms(I).SaveLabAs (Forms(I).labfilename)
End Select
Next I
If CancelUnload = True Then Exit Sub
For I = 5 To frmTestExplorer.TestTree.Nodes.Count
 Select Case Left$(frmTestExplorer.TestTree.Nodes(I).Key, 3)
    Case "Lab"
     LName = Left$(frmTestExplorer.TestTree.Nodes(I).Text, InStrRev(frmTestExplorer.TestTree.Nodes(I).Text, "(") - 2)
     Write #FileNum, "Lab", LName, frmTestExplorer.TestTree.Nodes(I).Tag
    Case "GrZ" ' ZScore Graph
    Case "RsZ" ' ZScore Result
     LName = Left$(frmTestExplorer.TestTree.Nodes(I).Text, InStrRev(frmTestExplorer.TestTree.Nodes(I).Text, "(") - 2)
     Write #FileNum, "ResultZScore", LName, frmTestExplorer.TestTree.Nodes(I).Tag
 End Select
Next I
Close FileNum
frmProperties.SaveSummary
Exit Sub
ErrHandler:
If Err.Number = cdlCancel Then ErrorNumber = cdlCancel Else Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim FormName As String
    FormName = ActiveForm.Name
    Select Case Button.Key
        Case "New": LoadNewDoc
        Case "Open": mnuFileOpen_Click
        Case "Save": mnuFileSave_Click
        Case "SaveTest": mnuFileSaveTest_Click
        Case "Close": Unload Me
        Case "Cut"
        Case "Copy": mnuEditCopy_Click
        Case "Paste"
        Case "Find"
        Case "Refresh": ActiveForm.RefreshValues
        Case "Properties": mnuFileProperties_Click
        Case "Help": MsgBox "Not yet ready"
        Case "Sample": ActiveForm.InsertSample
        Case "Explorer": If ActiveForm.Name <> "frmTestExplorer" Then frmTestExplorer.Show: frmTestExplorer.SetFocus
        Case "Result": frmResultDialog.Show 1
        Case "Calculator": Shell "C:\Windows\CALC.EXE"
        Case "Graph": frmGraphDialog.Show 1
        Case "Website": Shell "C:\Program Files\Internet Explorer\IEXPLORE www.sitraindia.org", vbMaximizedFocus
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuEditCopy_Click()
Select Case ActiveForm.Name
 Case "frmLab", "frmResult"
  Clipboard.SetText ActiveForm.ActiveControl.Clip
End Select
End Sub

Private Sub mnuFileExit_Click()
ExitApp = True: Unload Me
End Sub
Private Sub mnuFileProperties_Click()
frmProperties.Show 1
End Sub

Public Sub mnuFileSaveAs_Click()
    On Error GoTo ErrHandler
    If ActiveForm Is Nothing Then Exit Sub
    With dlgCommonDialog
    .CancelError = True
    .Flags = cdlOFNOverwritePrompt
    Select Case ActiveForm.Name
    Case "frmLab"
        .DialogTitle = "Save " & ActiveForm.Fg1.TextArray(0) & " As..."
        .Filter = "Lab Files (*.lab)|*.lab"
        .FileName = ActiveForm.Fg1.TextArray(0)
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        '
        ActiveForm.SaveLabAs .FileName
        ActiveForm.labfilename = .FileName
        ActiveForm.Caption = .FileTitle
        frmTestExplorer.TestTree.Nodes(ActiveForm.Tag).Text = ActiveForm.Fg1.TextArray(0) + " (" + .FileTitle + ")"
        frmTestExplorer.TestTree.Nodes(ActiveForm.Tag).Tag = .FileName
    Case "frmResult"
        .DialogTitle = "Save " & ActiveForm.ResultName & " As..."
        .Filter = "Result Sheet Files (*.rlt)|*.rlt"
        .FileName = ActiveForm.ResultName
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        ActiveForm.SaveResultAs .FileName
        ActiveForm.ResultFileName = .FileName
        ActiveForm.Caption = ActiveForm.ResultName + " - " + .FileTitle
        frmTestExplorer.TestTree.Nodes(ActiveForm.Tag).Text = ActiveForm.ResultName + " (" + .FileTitle + ")"
        frmTestExplorer.TestTree.Nodes(ActiveForm.Tag).Tag = .FileName
    End Select
    End With
    Exit Sub
ErrHandler:
If Err.Number = cdlCancel Then ErrorNumber = cdlCancel Else Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub mnuFileSave_Click()
 If ActiveForm Is Nothing Then Exit Sub
 Select Case ActiveForm.Name
  Case "frmLab": If ActiveForm.Caption = "<Not Saved>" Then mnuFileSaveAs_Click Else ActiveForm.SaveLabAs ActiveForm.labfilename
  Case "frmResult": If ActiveForm.ResultFileName = "" Then mnuFileSaveAs_Click Else ActiveForm.SaveResultAs ActiveForm.ResultFileName
 End Select
End Sub

Private Sub mnuFileClose_Click()
Unload Me
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo ErrorHandler
Dim FreeNum As Integer
FreeNum = FreeFile
If TestOpen = True Then Unload Me
If CancelUnload = True Or TestOpen = True Then Exit Sub
    With dlgCommonDialog
        .DialogTitle = "Open Test"
        .CancelError = True
        .Filter = "Test Files (*.tst)|*.tst"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        Open .FileName For Input As FreeNum: Close FreeNum
        LoadExistingTest .FileName
    End With
Exit Sub
ErrorHandler:
Select Case Err.Number
Case cdlCancel: ErrorNumber = cdlCancel
Case 53: MsgBox "Test File not found !", vbExclamation, "Proficiency Test"
Case Else: Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Select
End Sub
Private Sub mnuFileNew_Click()
LoadNewTest
End Sub

Private Sub tbToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ActiveForm Is Nothing Then Exit Sub
On Error GoTo ErrHandler
Dim FileNum As Integer, Field As String, Ascii0Pos As Integer, PthName As String, FlName As String
FileNum = FreeFile
Select Case ButtonMenu.Key
Case "ExistingLab"
    With dlgCommonDialog
        .DialogTitle = "Add an Existing Lab"
        .CancelError = True
        .Flags = cdlOFNExplorer Or cdlOFNAllowMultiselect Or cdlOFNHideReadOnly
        .Filter = "Lab Files (*.lab)|*.lab"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        Ascii0Pos = InStr(.FileName, Chr$(0))
        If Ascii0Pos <> 0 Then
           PthName = Left$(.FileName, Ascii0Pos - 1) + "\"
           Do
             If InStr(Ascii0Pos + 1, .FileName, Chr$(0)) <> 0 Then
              FlName = Mid$(.FileName, Ascii0Pos + 1, InStr(Ascii0Pos + 1, .FileName, Chr$(0)) - Ascii0Pos - 1)
             Else
              FlName = Mid$(.FileName, Ascii0Pos + 1)
             End If
             Open PthName + FlName For Input As FileNum: Line Input #FileNum, Field: Input #FileNum, Field: Close FileNum
             frmTestExplorer.TestTree.Nodes.Add "Lab", tvwChild, "Lab" & Field, Field + " (" + FlName + ")", "Lab"
             frmTestExplorer.TestTree.Nodes("Lab" & Field).Tag = PthName + FlName
             LoadExistingDoc PthName + FlName
             Ascii0Pos = InStr(Ascii0Pos + 1, .FileName, Chr$(0))
           Loop While Ascii0Pos <> 0
        Else
           Open .FileName For Input As FileNum: Line Input #FileNum, Field: Input #FileNum, Field: Close FileNum
           frmTestExplorer.TestTree.Nodes.Add "Lab", tvwChild, "Lab" & Field, Field + " (" + .FileTitle + ")", "Lab"
           frmTestExplorer.TestTree.Nodes("Lab" & Field).Tag = .FileName
           LoadExistingDoc .FileName
        End If
    End With
Case "ExistingResult"
    With dlgCommonDialog
        .DialogTitle = "Add an Existing Result Sheet"
        .CancelError = True
        .Filter = "Result Sheet Files (*.rlt)|*.rlt"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        Open .FileName For Input As FileNum
        Dim KindOfResult As String
        Line Input #FileNum, KindOfResult: Line Input #FileNum, Field: Input #FileNum, Field
        Close FileNum
        Select Case KindOfResult
         Case "ZScore"
          frmTestExplorer.TestTree.Nodes.Add "Result", tvwChild, "RsZ" & Field, Field + " (" + .FileTitle + ")", "Result"
          frmTestExplorer.TestTree.Nodes("RsZ" & Field).Tag = .FileName
          LoadExistingResult .FileName
        End Select
    End With
End Select
Exit Sub
ErrHandler:
Select Case Err.Number
Case cdlCancel: ErrorNumber = Err.Number
Case 53: MsgBox "File not found !", vbExclamation, "File not found"
Case Else: Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Select
End Sub
Private Sub ExplorerLab_Click(Index As Integer)
With frmTestExplorer.TestTree.SelectedItem
Dim I As Integer
Select Case Index
 Case 0: 'View
  frmTestExplorer.TestTree_DblClick
 Case 1: 'Properties
 Case 2: 'Copy
   Select Case .Image
    Case "Lab":
     For I = 1 To Forms.Count - 1
      If Forms(I).Name = "frmLab" And Forms(I).Tag = .Key Then Forms(I).SetFocus: mnuEditCopyLab_Click: I = Forms.Count
     Next I
    Case "ClosedLab": CopyClosedLab .Tag
    Case "Result"
     For I = 1 To Forms.Count - 1
      If Forms(I).Name = "frmResult" And Forms(I).Tag = .Key Then Forms(I).SetFocus: mnuEditCopyLab_Click: I = Forms.Count
     Next I
    Case "ClosedResult": CopyClosedResult .Tag
   End Select
 Case 3: 'Save
 Case 4: 'Save As
 Case 5: 'Remove
End Select
End With
End Sub

Private Sub CopyClosedLab(FileName As String)
Dim FileNum As Integer, NoOfRows As Integer, NoOfCols As Integer, I As Integer, J As Integer, Field As String, ClipString As String
FileNum = FreeFile
Open FileName For Input As FileNum
Input #FileNum, NoOfRows, NoOfCols
For I = 1 To NoOfCols - 1: Input #FileNum, Field: ClipString = ClipString + Field + vbTab: Next I
Input #FileNum, Field: ClipString = ClipString + Field + vbCrLf
For I = 1 To NoOfRows - 1
 ClipString = ClipString + CStr(I) + vbTab
 For J = 1 To NoOfCols - 2
  Input #FileNum, Field
  ClipString = ClipString + Field + vbTab
 Next J
 Input #FileNum, Field
 ClipString = ClipString + Field + vbCrLf
Next I
For I = 1 To 3
 Select Case I
  Case 1: ClipString = ClipString + "Average" + vbTab
  Case 2: ClipString = ClipString + "RD" + vbTab
  Case 3: ClipString = ClipString + "RSD" + vbTab
 End Select
 For J = 1 To NoOfCols - 2
  Input #FileNum, Field
  ClipString = ClipString + Field + vbTab
 Next J
 Input #FileNum, Field
 ClipString = ClipString + Field + vbCrLf
Next I
Close FileNum
Clipboard.Clear
Clipboard.SetText ClipString
End Sub
Public Sub CopyClosedResult(FileName As String)
Dim FileNum As Integer, I As Integer, J As Integer, NoOfRows As Integer, NoOfCols As Integer, Field As String, ClipString As String
FileNum = FreeFile
Open FileName For Input As FileNum
Line Input #FileNum, Field
If Field <> "ZScore" Then MsgBox "Not a valid Result Sheet File.", vbExclamation, "Error Opening File !": Exit Sub
Input #FileNum, NoOfRows, NoOfCols, Field
ClipString = Field + vbCrLf + "Lab / Sample" + vbTab
For I = 1 To NoOfCols - 4 Step 3
 Input #FileNum, Field
 ClipString = ClipString + Field + vbTab + "Zc" + vbTab + "Zr" + vbTab
Next I
Input #FileNum, Field
ClipString = ClipString + Field + vbTab + "Zc" + vbTab + "Zr" + vbCrLf
For I = 1 To NoOfRows
 For J = 1 To NoOfCols - 1
  Input #FileNum, Field: ClipString = ClipString + Field + vbTab
 Next J
 Input #FileNum, Field: ClipString = ClipString + Field + vbCrLf
Next I
For I = 1 To 8
 For J = 1 To NoOfCols - 4 Step 3
  Input #FileNum, Field: ClipString = ClipString + Field + vbTab + vbTab + vbTab
 Next J
 Input #FileNum, Field: ClipString = ClipString + Field + vbCrLf
Next I
Close FileNum
Clipboard.Clear
Clipboard.SetText ClipString
End Sub
Public Sub LoadNewResult(RName As String)
Dim frmR As frmResult
ResultCount = ResultCount + 1
Set frmR = New frmResult
frmR.ResultName = RName
frmR.Caption = RName + " - <Not Saved>"
frmR.Tag = "RsZ" & frmR.ResultName  ' The key is RsZ and its name for example : RsZSitraResultOfLab211 for ZScore
frmTestExplorer.TestTree.Nodes.Add "Result", tvwChild, frmR.Tag, frmR.ResultName + " (<Not Saved>)", "Result"
frmR.Show
End Sub
