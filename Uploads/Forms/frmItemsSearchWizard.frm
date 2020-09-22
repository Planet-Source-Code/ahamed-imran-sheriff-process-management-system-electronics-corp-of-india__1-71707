VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmItemsSearchWizard 
   Caption         =   "Items Search Wizard"
   ClientHeight    =   7650
   ClientLeft      =   3945
   ClientTop       =   1980
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   7260
   Begin VB.ComboBox cboSearchType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmItemsSearchWizard.frx":0000
      Left            =   1320
      List            =   "frmItemsSearchWizard.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdApply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Apply"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dgrdItemsInfo 
      Height          =   3375
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483629
      HeadLines       =   1
      RowHeight       =   15
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
      Caption         =   "Items Information Table"
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   1335
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   360
      Top             =   1080
      Width           =   6495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label lblSearchText 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   3855
      Width           =   1215
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   3855
      Width           =   855
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Search Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   0
      Left            =   0
      Picture         =   "frmItemsSearchWizard.frx":0033
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   2535
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   360
      Top             =   2280
      Width           =   6495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2535
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "ECI Corporation, India."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   3
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Index           =   0
      Left            =   0
      Picture         =   "frmItemsSearchWizard.frx":00D5
      Stretch         =   -1  'True
      Top             =   -1200
      Width           =   9810
   End
End
Attribute VB_Name = "frmItemsSearchWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This variable will determine if the DataGrid has been clicked or not
Dim Flag As Boolean


Private Sub Form_Load() 'Form Load Procedure

    Flag = False    'The Flag variable is being initialized to False
    
    Call ItemDetails    'Calling the ItemDetails Procedure to interact with the recordset
        
    Set dgrdItemsInfo.DataSource = rsItemDetails 'Setting the DataSource of the DataGrid
    
End Sub



Private Sub cmdClose_Click()    'This procedure will close the Wizard

    Unload Me   'Unloading the Wizard
    
End Sub

Private Sub dgrdItemsInfo_Click()    'This procedure is executed if the user clicks the DataGrid
    
    'Setting the Flag variable to True, to indicate that the user
    'has clicked the DataGrid
    Flag = True
    
End Sub


Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    On Error GoTo errorGotcha
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsItemDetails
        
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[Item_ID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[Item_Type] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
                    .Filter = "[Manufacturer] Like '" & txtSearch.Text & "%" & "'"
            End Select
    
        End With
        
        Set dgrdItemsInfo.DataSource = rsItemDetails 'Setting the DataSource of the DataGrid
            
    Else
        
        Form_Load   'Calling the Form_Load Procedure
        
    End If
    
    Exit Sub
    
errorGotcha:
    MsgBox Err.Description & "" & Err.Number, vbCritical

    
End Sub


Private Sub cmdApply_Click()    'This code is executed when the user clicks the Apply Button
    
On Error GoTo errorHold

    'Here, I am checkin to see if the user has chosen a record
    If Flag = True And rsItemDetails.RecordCount > 0 Then
    
        With rsItemDetails
        
            'Reset the textfields with the selected record
            frmPurchaseDetails.txtItemID.Text = .Fields(0).Value
            frmPurchaseDetails.txtRetailPrice.Text = .Fields(4).Value
            
            Unload Me   'Unload the Wizard
            
        End With
    
    Else    'Displaying an error message, asking the user to choose a record
    
        MsgBox "Please Select a Record First!", vbExclamation, "No Record Selected!"
        Exit Sub
        
    End If
    
    Exit Sub

errorHold:
    MsgBox Err.Description & "" & Err.Number, vbCritical

    
End Sub





