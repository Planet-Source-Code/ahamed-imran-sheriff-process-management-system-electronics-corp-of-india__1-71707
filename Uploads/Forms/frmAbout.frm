VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About This Software"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Close"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Click Here To Close This Interface"
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "We Hope You Enjoy Using This Software! Thank You!"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   6720
      Width           =   9615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "requests department on (Toll Free) 011 2598346."
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
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   6240
      Width           =   9615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "slogans, screen shots, copyrighted designs or other brand features, please contact the permission "
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
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   6000
      Width           =   9615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "If you are seeking permission to use ECI's trademarks, logos, service marks, trade dress,  "
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
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   5760
      Width           =   9615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "action against such person/s."
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
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   5400
      Width           =   9615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "our intellectual property rights have been otherwise violated, we would be forced to resort to legal "
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
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   5160
      Width           =   9615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "If we believe that our work has been copied in a way that constitutes copyright infringement, or that"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   4920
      Width           =   9615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "civil and criminal penalties, and will be prosecuted to the maximum extent possible under the law."
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
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   4560
      Width           =   9615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Unauthorized reproduction or distribution of this program, or any portion of it, may result in severe"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   4320
      Width           =   9615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ECI, India (Pvt) Ltd. have asserted their moral rights."
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
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   3960
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning: This computer program is protected by copyright law and international treaties."
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
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3720
      Width           =   8415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved! Copyright ECI, India (Pvt) Ltd., 2009."
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
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   3360
      Width           =   8895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ECI Corporation's Process Mangement System is wholly owned by ECI, India (Pvt) Ltd."
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
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   3120
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to PMS!"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblFrameTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "About This Software"
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
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   960
      X2              =   1200
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   3240
      X2              =   10920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   10920
      X2              =   10920
      Y1              =   2400
      Y2              =   7320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   960
      X2              =   10920
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   960
      X2              =   960
      Y1              =   2400
      Y2              =   7320
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
