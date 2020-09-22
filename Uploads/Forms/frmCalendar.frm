VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCalendar.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1695
   End
   Begin MSACAL.Calendar cldrSystemCalendar 
      Height          =   5055
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   6255
      _Version        =   524288
      _ExtentX        =   11033
      _ExtentY        =   8916
      _StockProps     =   1
      BackColor       =   -2147483632
      Year            =   2008
      Month           =   1
      Day             =   13
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()    'On click of the Close Button
    
    'Obtaining confirmation from the user
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    cldrSystemCalendar.Today
    
End Sub
