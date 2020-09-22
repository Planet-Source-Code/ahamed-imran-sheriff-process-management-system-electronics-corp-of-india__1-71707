VERSION 5.00
Begin VB.Form frmReportAging 
   Caption         =   "Form3"
   ClientHeight    =   1905
   ClientLeft      =   6915
   ClientTop       =   4695
   ClientWidth     =   4650
   LinkTopic       =   "Form3"
   ScaleHeight     =   1905
   ScaleWidth      =   4650
   Begin VB.ComboBox cbobillstat 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please Choose Which Bill Status to View in the Aging Report"
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Please Choose Bill Status"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmReportAging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Inpatient_Billing
    rsInpatientBilling.MoveFirst
    cbobillstat.AddItem ("PAID")
    cbobillstat.AddItem ("UNPAID")
End Sub

Private Sub cmdOK_Click()
    On Error GoTo e
        DataEnvironment1.Commands("AgingReport").Parameters(0) = cbobillstat
        With RptAging
            .Sections("Section2").Controls("lblbillstat").Caption = cbobillstat
            .Show
        End With
        DataEnvironment1.rsAgingReport.Close
        
        Unload Me
    Exit Sub
e:
    If Err.Number <> 3704 Then
        MsgBox Err.Description & "" & Err.Number, vbCritical
    End If
End Sub



