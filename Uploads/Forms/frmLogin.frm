VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   4485
   ClientLeft      =   -120
   ClientTop       =   -135
   ClientWidth     =   8235
   ForeColor       =   &H80000018&
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmLogin.frx":17780
      OLEDBString     =   $"frmLogin.frx":1780C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGo 
      DisabledPicture =   "frmLogin.frx":17898
      Height          =   496
      Left            =   7200
      Picture         =   "frmLogin.frx":17B38
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click To Access System"
      Top             =   2490
      Width           =   511
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   4545
      TabIndex        =   0
      Top             =   2040
      Width           =   2445
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4545
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2550
      Width           =   2445
   End
   Begin VB.CommandButton cmdExit 
      Height          =   496
      Left            =   6000
      Picture         =   "frmLogin.frx":1877C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   511
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   3360
      TabIndex        =   6
      Top             =   2055
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   3360
      TabIndex        =   5
      Top             =   2565
      Width           =   1320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Quit Application"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   3330
      Width           =   1935
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------
'Process Management System - Inspirat Edition
'Form Name: Login Interface
'Programmer: Ahamed Imran Sheriff
'Quality Assurance Engineer (Testing): Dinithi Vithanage
'Start Date: January 19, 2009.
'Date Of Last Modification: January 19, 2009.
'The Name Of The Database Being Accessed: dbECI
'The Name/s Of The Database Table/s Being Accessed: User_Account Table
'---------------------------------------------------------------------

Option Explicit
Dim rsLogin As ADODB.Recordset 'Creating a Recordset Variable
Dim iLoginFailure As Integer    'This variable will count the number of times the user's login is unsuccessful.


Private Sub Form_Initialize()
    
    Call Connection 'Calling the Connection Procedure.
    
    'Creating a New Recordset To Be Used For Login Purposes Only
    Set rsLogin = New ADODB.Recordset
    
    With rsLogin
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .CursorLocation = adUseClient
    End With

    iLoginFailure = 1    ' When a login attempt is unsuccessful, I decrement this variable's value.
    
End Sub


Private Sub txtUserName_keypress(KeyAscii As Integer)
    
    'This block of code prevents the user from using "Copy-Paste" (Ctrl+C, Ctrl+V) functions.
    
    If KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then   'This is for using the Enter key
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
    
End Sub

Private Sub txtPassword_Keypress(KeyAscii As Integer)

    'This block of code prevents the user from using "Copy-Paste" (Ctrl+C, Ctrl+V) functions.

    If KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Then
        KeyAscii = 0
        ElseIf KeyAscii = 13 Then   'This is for using the Enter key
        KeyAscii = 0
        SendKeys "{Tab}"
       Call cmdGO_Click
    End If
    
    
End Sub

Private Sub cmdGO_Click()

    On Error GoTo errorLogin

    If iLoginFailure <= 3 Then  'Checking If The User Is Still Allowed To Login
    
       'Selecting the Related Login Record from the User_Account Table.
        rsLogin.Open "select * from UserAccount where UserID='" & txtUserName.Text & "'", conn
                
        With rsLogin
                            
            If .RecordCount = 0 Then    'This Means That There Is No Matching Record
                
                iLoginFailure = iLoginFailure + 1   'Decrementing The Value Of i On Each Unsuccessful Login Attempt
                MsgBox "Sorry! Invalid User Name! Please Try Again!", vbCritical, "Invalid Login!"
                txtUserName.BackColor = &H80000018  'Highlighting The Textbox With The Error
                txtPassword.BackColor = &H80000005  'Highlighting The Textbox With The Error
                txtUserName.Text = ""
                txtUserName.SetFocus
                
            End If
        
            If .RecordCount <> 0 Then   'This Means That There Is A Matching Record
                
                If txtPassword.Text = .Fields(2).Value Then 'Checking Password
                
                    If .Fields(3) = "Administrator" Then    'Checking Designation
                    
                        'Passing Necessary Values To Global Variables
                        accessLevel = "Administrator"
                        userName = .Fields(1).Value
                        userID = .Fields(0).Value
                        frmMDI.lblCurrentUser.Caption = userName    'Displaying the current user's name on the MDI
                        frmMDI.Show
                        Unload Me
                        
                    ElseIf .Fields(3) = "Operator" Then    'Checking Designation
                        'Passing Necessary Values To Global Variables
                        accessLevel = "Operator"
                        userName = .Fields(1).Value
                        userID = .Fields(0).Value
                        frmMDI.lblCurrentUser.Caption = userName    'Displaying the current user's name on the MDI
                        frmMDI.Show
                        Unload Me
                    
                    End If
                        
                End If
                    
                Else
                
                    'Error Mesage For Invalid Password
                    iLoginFailure = iLoginFailure + 1   'Decrementing The Value Of i On Each Unsuccessful Login Attempt
                    MsgBox "Sorry! Invalid Password! Please Try Again!", vbCritical, "Invalid Login!"
                    txtPassword.BackColor = &H80000018  'Highlighting The Textbox With The Error
                    txtUserName.BackColor = &H80000005  'Highlighting The Textbox With The Error
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                    
                End If
            
            .Close  'Closing Recordset
                        
        End With
        
        Else
            'Error Message If User's Login Attempt Is Unsuccesful On Three
            'Consecutive Occasions
            MsgBox "Sorry! You Have To Login Within Three Tries! Unloading...", vbCritical, "Login Failure"
        End
        
    End If
    
    Exit Sub
    
errorLogin:
    MsgBox Err.Description & "" & Err.Number, vbCritical, "There was an error in your login!"
    
End Sub


Private Sub cmdExit_Click()

    'This block of code will be executed if the user decides to quit the application
    'from the Login page
    
    Dim ans As Variant
    ans = MsgBox("Are You Sure You Wish To Quit The Application?", vbYesNo + vbQuestion, "Quit Application?")
    
    If ans = vbYes Then
        End
    End If
    
End Sub

