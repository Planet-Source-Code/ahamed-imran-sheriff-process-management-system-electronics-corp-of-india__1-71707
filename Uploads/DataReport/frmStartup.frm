VERSION 5.00
Begin VB.Form frmStartup 
   Caption         =   "Startup"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form3"
   ScaleHeight     =   4560
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    frmMDI.Show
End Sub
