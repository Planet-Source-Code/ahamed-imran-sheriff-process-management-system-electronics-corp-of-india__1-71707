VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmReportOutpatientRevenue 
   Caption         =   "Outpatient Revenue Report"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   LinkTopic       =   "Form3"
   ScaleHeight     =   4365
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame 
      Caption         =   "Please Choose the Dates of the Report Requested"
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtdate1 
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "M/D/YYYY"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtdate2 
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Text            =   "M/D/YYYY"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4800
         TabIndex        =   1
         Top             =   3120
         Width           =   1455
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   5953
         _StockProps     =   1
         BackColor       =   16761024
         Year            =   2008
         Month           =   4
         Day             =   1
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
      Begin VB.Label Label2 
         Caption         =   "Date To:"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date From:"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmReportOutpatientRevenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmDate As Date
Dim endDate As Date

Private Sub Calendar1_Click()
    txtdate1.Text = Calendar1.Value
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo e
        frmDate = txtdate1.Text
        endDate = txtdate2.Text
        
                DataEnvironment1.Commands("OutpatientRevenue").Parameters(0) = txtdate1
                DataEnvironment1.Commands("OutpatientRevenue").Parameters(1) = txtdate2
                With RptOutpatientRevenue
                'Set .DataSource = rsTemp
                '.DataMember = rsTemp.DataMember
                .Sections("Section2").Controls("lblDate1").Caption = txtdate1.Text
                .Sections("Section2").Controls("lblDate2").Caption = txtdate2.Text
                .Show
            End With
            DataEnvironment1.rsOutpatientRevenue.Close
        
        Unload Me
    Exit Sub
e:
    If Err.Number <> 3704 Then
        MsgBox Err.Description, vbCritical
    End If
End Sub


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtDate1_change()
    Dim todate As Date
    todate = txtdate1.Text
    
    txtdate2.Text = todate + 31
End Sub








