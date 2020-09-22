VERSION 5.00
Begin VB.Form frmStartup 
   Caption         =   "Form3"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form3"
   ScaleHeight     =   4965
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton frmStartup 
      Caption         =   "Command1"
      Height          =   975
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub frmStartup_Click()
    Unload Me
    frmMDI.Show
End Sub
