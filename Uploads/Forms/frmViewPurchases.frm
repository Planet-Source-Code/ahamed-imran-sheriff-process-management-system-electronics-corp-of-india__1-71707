VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmViewPurchases 
   Caption         =   "View Purchases"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmViewPurchases.frx":0000
   ScaleHeight     =   10095
   ScaleWidth      =   12420
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dgrdPurchases 
      Height          =   4815
      Left            =   1800
      TabIndex        =   0
      Top             =   2160
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483629
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "View Purchases"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   5295
      Left            =   1560
      Top             =   1920
      Width           =   9255
   End
End
Attribute VB_Name = "frmViewPurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()

    'Obtaining confirmation from the user
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    
    Call ViewPurchaseDetails    'Calling the ViewPurchaseDetails Procedure to interact with the recordset
    
    Set dgrdPurchases.DataSource = rsViewPurchaseDetails 'Setting the DataSource of the DataGrid

End Sub
