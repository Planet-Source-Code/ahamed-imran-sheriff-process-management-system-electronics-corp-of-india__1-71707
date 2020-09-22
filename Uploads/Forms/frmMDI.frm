VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "ECI Corporation, India [Process Management System - Inspirat Edition]"
   ClientHeight    =   10290
   ClientLeft      =   180
   ClientTop       =   750
   ClientWidth     =   14970
   Icon            =   "frmMDI.frx":0000
   Picture         =   "frmMDI.frx":15162
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpCommand     =   1
      HelpContext     =   1
      HelpFile        =   "HELP.hlp"
      HelpKey         =   "F1"
   End
   Begin VB.PictureBox picRightNavigation 
      Align           =   3  'Align Left
      Height          =   9915
      Left            =   0
      Picture         =   "frmMDI.frx":44840
      ScaleHeight     =   9855
      ScaleWidth      =   3150
      TabIndex        =   1
      Top             =   0
      Width           =   3205
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Agreements Summary"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   6720
         Width           =   2775
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchases Summary"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   495
         Index           =   2
         Left            =   700
         TabIndex        =   5
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Agreement Details"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   495
         Index           =   1
         Left            =   370
         TabIndex        =   4
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Details"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   495
         Index           =   0
         Left            =   810
         TabIndex        =   3
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblCurrentUser 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
   End
   Begin MSComctlLib.StatusBar BottomStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9915
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
            Text            =   "ECI Corporation, India [Process Management System - Inspirat Edition]"
            TextSave        =   "ECI Corporation, India [Process Management System - Inspirat Edition]"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "AD000001"
            TextSave        =   "AD000001"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Administrator"
            TextSave        =   "Administrator"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCustomers 
      Caption         =   "&Customers"
      Begin VB.Menu mnuPurchases 
         Caption         =   "Purchases"
      End
      Begin VB.Menu mnuServiceAgreements 
         Caption         =   "Service Agreements"
      End
      Begin VB.Menu mnuViewPurchases 
         Caption         =   "View Purchases"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuPurchasesMasterReport 
         Caption         =   "Purchases Master Report"
      End
      Begin VB.Menu mnuServiceAgreementsMasterReport 
         Caption         =   "Service Agreements Master Report"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuMicrosoftMagnifier 
         Caption         =   "Microsoft Magnifier"
      End
      Begin VB.Menu mnuSeparator95 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMicrosoftNarrator 
         Caption         =   "Microsoft Narrator"
      End
      Begin VB.Menu mnuSeparator94 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemMediaPlayer 
         Caption         =   "System Media Player"
      End
      Begin VB.Menu mnuSeparator93 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalendar 
         Caption         =   "Calendar"
      End
      Begin VB.Menu mnuSeparator92 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemCalculator 
         Caption         =   "System Calculator"
      End
      Begin VB.Menu mnuSeparator91 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemNotepad 
         Caption         =   "System Notepad"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuSeparator88 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuSeparator87 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuSeparator86 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSeparator85 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "Credits"
      End
   End
   Begin VB.Menu mnuUserAccount 
      Caption         =   "&Account"
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuSeparator101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------
'Process Management System - Inspirat Edition
'Form Name: Menu Driven Interface (MDI)
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Salvin Ali Saleh
'Start Date: January 19, 2009.
'Date Of Last Modification: January 19, 2009.
'The Name Of The Database Being Accessed: None
'The Name/s Of The Database Table/s Being Accessed: None
'---------------------------------------------------------------------

Option Explicit
Dim iExitReply As Integer 'This variable will hold the user's choice, once he has been asked whether he wants to exit or not
Dim iLogOutReply As Integer 'This variable will hold the user's choice, once he has been asked whether he wants to log out or not


Private Sub MDIForm_Load()
    
    'In the following lines of code, I am checking the user access level
    'and appropriately disabling certain restricted functions
             
    'lblDateTime.Caption = "Today is " & FormatDateTime(Now, vbLongDate)
    
    'frmQuickLaunch.Show
    
    If accessLevel = "Operator" Then
        lblShortcut(2).Enabled = False
        lblShortcut(3).Enabled = False
        mnuReports.Enabled = False
        mnuViewPurchases.Enabled = False
    End If
    
    If accessLevel = "Administrator" Then
        lblShortcut(2).Enabled = True
        lblShortcut(3).Enabled = True
        mnuReports.Enabled = True
        mnuViewPurchases.Enabled = True
    End If
    
    BottomStatusBar.Panels(4).Text = userName
    BottomStatusBar.Panels(5).Text = userID
    
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'This event occurs when the user tries to quit the application by clicking the
    'standard red cross button, on the top left hand corner of the interface
    
    If UnloadMode = 0 Then
        iExitReply = MsgBox(userName & ", Are You Sure You Wish To Exit The Application?", vbYesNo + vbQuestion, "Exit Application?")
        If iExitReply = vbNo Then
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub


Private Sub mnuCalendar_Click()
    frmCalendar.Show
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCloseAll_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCredits_Click()
    frmCredits.Show
End Sub

Private Sub mnuLogOff_Click()

    iLogOutReply = MsgBox(userName & ", Are You Sure You Wish To Log Out Of Your Account?", vbYesNo + vbQuestion, "Log Out?")
    If iLogOutReply = vbYes Then
        frmLogin.Show
        Unload Me
    End If
    
End Sub

Private Sub mnuMicrosoftMagnifier_Click()   'Opens Up The Magnifier Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\magnify.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run Microsoft Magnifier on your computer", vbError, "Error Loading Microsoft Magnifier!"
    Resume Next
End Sub

Private Sub mnuMicrosoftNarrator_Click()    'Opens Up The Narrator Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\narrator.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run Microsoft Narrator on your computer", vbError, "Error Loading Microsoft Narrator!"
    Resume Next
End Sub

Private Sub mnuPurchases_Click()
    frmPurchaseDetails.Show
End Sub


Private Sub mnuPurchasesMasterReport_Click()
    rptPurchaseReport.Show
End Sub


Private Sub mnuServiceAgreements_Click()
    frmServicingDetails.Show
End Sub

Private Sub mnuServiceAgreementsMasterReport_Click()
    rptServiceAgreement.Show
End Sub

Private Sub mnuSystemCalculator_Click() 'Opens Up The Calculator Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\calc.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run the Calculator Utility on your computer", vbError, "Error Loading Calculator!"
    Resume Next
End Sub

Private Sub mnuSystemMediaPlayer_Click()    'Opens Up The System Media Player Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\dvdplay.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run System Media Player on your computer", vbError, "Error Loading System Media Player!"
    Resume Next
End Sub

Private Sub mnuSystemNotepad_Click()    'Opens Up The Notepad Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\notepad.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run the Notepad Utility on your computer", vbError, "Error Loading Notepad!"
    Resume Next
End Sub

Private Sub mnuTileHorizontally_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertically_Click()
    Me.Arrange vbTileVertical
End Sub




Private Sub mnuExit_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Quit The Application?", vbYesNo + vbQuestion, "Quit Application?") = vbYes Then
        End
    End If
    
End Sub


Private Sub lblShortcut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Here, I am creating Rollover Effects For Each Label On The Shortcut Panel

    Select Case (Index)
        Case 0: 'Purchase Details Label
            lblShortcut(0).ForeColor = &H80000003
            lblShortcut(1).ForeColor = &H80000004
            lblShortcut(2).ForeColor = &H80000004
            lblShortcut(3).ForeColor = &H80000004

            lblShortcut(0).FontUnderline = True
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False

            
        Case 1: 'Service Agreement Details Label
            lblShortcut(0).ForeColor = &H80000004
            lblShortcut(1).ForeColor = &H80000003
            lblShortcut(2).ForeColor = &H80000004
            lblShortcut(3).ForeColor = &H80000004

            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = True
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False

            
        Case 2: 'Purchases Summary Label
            lblShortcut(0).ForeColor = &H80000004
            lblShortcut(1).ForeColor = &H80000004
            lblShortcut(2).ForeColor = &H80000003
            lblShortcut(3).ForeColor = &H80000004

            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = True
            lblShortcut(3).FontUnderline = False

            
        Case 3: 'Service Agreements Summary
            lblShortcut(0).ForeColor = &H80000004
            lblShortcut(1).ForeColor = &H80000004
            lblShortcut(2).ForeColor = &H80000004
            lblShortcut(3).ForeColor = &H80000003

            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = True

            
    End Select
    
End Sub


Private Sub mnuViewPurchases_Click()
    
    frmViewPurchases.Show
    
End Sub

Private Sub picRightNavigation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'The Following Block Of Code Ensures That The Mouse Pointer
    'Returns To Normal When It Is Not Over A Button
    
            lblShortcut(0).ForeColor = &H80000004
            lblShortcut(1).ForeColor = &H80000004
            lblShortcut(2).ForeColor = &H80000004
            lblShortcut(3).ForeColor = &H80000004

            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
    
End Sub



Private Sub lblShortcut_Click(Index As Integer)

    'The following block of code illustrates which interfaces are displayed on click of
    'each respective label
    
    Select Case (Index)
        Case 0: 'Purchase Details Label
            frmPurchaseDetails.Show
            
        Case 1: 'Service Agreement Details Label
            frmServicingDetails.Show
            
        Case 2: 'Purchases Summary Label
            rptPurchaseReport.Show
            
        Case 3: 'Service Agreements Summary
            rptServiceAgreement.Show
    
    End Select
    
End Sub


