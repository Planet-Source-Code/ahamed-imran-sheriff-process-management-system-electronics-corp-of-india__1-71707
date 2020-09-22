VERSION 5.00
Begin VB.Form frmReportInpatientInvoice 
   Caption         =   "Inpatients Invoice"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4710
   LinkTopic       =   "Form3"
   ScaleHeight     =   2010
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Please Insert Patient ID to view the Invoice "
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtinpatientno 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "INP"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Please Insert Patient ID"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmReportInpatientInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo e
        
        DataEnvironment1.Commands("InpatientInvoice").Parameters(0) = ""
        DataEnvironment1.Commands("InpatientInvoice").Parameters(1) = txtinpatientno.Text
        RptInpatientInvoice.Show

        DataEnvironment1.rsInpatientInvoice.Close
        
        Unload Me
    Exit Sub
e:
    If Err.Number <> 3704 Then
        MsgBox Err.Description, vbCritical
    End If
End Sub


