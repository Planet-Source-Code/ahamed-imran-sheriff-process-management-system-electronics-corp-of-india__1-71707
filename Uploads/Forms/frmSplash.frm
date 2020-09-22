VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   ClientHeight    =   4440
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4440
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   6480
      Top             =   600
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   840
      TabIndex        =   0
      Top             =   3780
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing Internal Assets....Please Standby..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   3480
      Width           =   5055
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
'    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Static ictr As Integer
    'Run the timer and check the condition
    'when the condition is false
    'stop the timer and disply the login form
    If ictr <= 100 Then
        ProgressBar1.Value = ictr
        ictr = ictr + 1
    Else
        frmLogin.Show
        'frmLogin.Show
        Unload Me
    End If
End Sub
