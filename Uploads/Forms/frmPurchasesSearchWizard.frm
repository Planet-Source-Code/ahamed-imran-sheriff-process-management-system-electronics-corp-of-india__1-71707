VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPurchasesSearchWizard 
   Caption         =   "Purchases Search Wizard"
   ClientHeight    =   7515
   ClientLeft      =   3555
   ClientTop       =   2175
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   7215
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
      ItemData        =   "frmPurchasesSearchWizard.frx":0000
      Left            =   1320
      List            =   "frmPurchasesSearchWizard.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   9
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
      TabIndex        =   8
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dgrdPurchasesInfo 
      Height          =   3375
      Left            =   360
      TabIndex        =   0
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
      Caption         =   "Purchases Information Table"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   7200
      Width           =   2175
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchases Search Wizard"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   0
      Left            =   0
      Picture         =   "frmPurchasesSearchWizard.frx":0046
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgbg2 
      Height          =   7275
      Index           =   0
      Left            =   0
      Picture         =   "frmPurchasesSearchWizard.frx":00E8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   9810
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   120
      Picture         =   "frmPurchasesSearchWizard.frx":0186
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
      Left            =   3600
      TabIndex        =   6
      Top             =   1335
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   480
      Top             =   1080
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
      Left            =   600
      TabIndex        =   5
      Top             =   1335
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
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Depot Search Wizard"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmPurchasesSearchWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This variable will determine if the DataGrid has been clicked or not
Dim Flag As Boolean


Private Sub Form_Load() 'Form Load Procedure

    Flag = False    'The Flag variable is being initialized to False
    
    Call myViewPurchases    'Calling the myViewPurchases Procedure to interact with the recordset
    
    Set dgrdPurchasesInfo.DataSource = rsViewPurchases 'Setting the DataSource of the DataGrid
    
End Sub



Private Sub cmdClose_Click()    'This procedure will close the Wizard

    Unload Me   'Unloading the Wizard
    
End Sub

Private Sub dgrdPurchasesInfo_Click()    'This procedure is executed if the user clicks the DataGrid
    
    'Setting the Flag variable to True, to indicate that the user
    'has clicked the DataGrid
    Flag = True
    
End Sub



Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    On Error GoTo errorSearch
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsViewPurchases
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[Purchase_ID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[Purchase_Date] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
                    .Filter = "[Customer_ID] Like '" & txtSearch.Text & "%" & "'"
                Case 3:
                    .Filter = "[Item_ID] Like '" & txtSearch.Text & "%" & "'"
            End Select
    
        End With
        
        Set dgrdPurchasesInfo.DataSource = rsViewPurchases    'Setting the DataSource of the DataGrid
            
    Else
        
        Form_Load   'Calling the Form_Load Procedure
        
    End If
    
    Exit Sub

errorSearch:
    MsgBox Err.Description & "" & Err.Number, vbCritical

    
End Sub


Private Sub cmdApply_Click()    'This code is executed when the user clicks the Apply Button
    
On Error GoTo errorApply

    'Here, I am checkin to see if the user has chosen a record
    If Flag = True And rsViewPurchases.RecordCount > 0 Then
    
        With rsViewPurchases
        
            'Reset the textfields with the selected record
            frmServicingDetails.txtPurchaseID.Text = .Fields(0).Value
            frmServicingDetails.txtCustomerID.Text = .Fields(2).Value
            frmServicingDetails.txtItemID.Text = .Fields(3).Value
            
            Unload Me   'Unload the Wizard
            
        End With
        
        Call ItemDetails
        With rsItemDetails
            .MoveFirst
            Do While .EOF = False
                If .Fields(0).Value = frmServicingDetails.txtItemID.Text Then
                    frmServicingDetails.txtItemType.Text = .Fields(1).Value
                End If
                .MoveNext
            Loop
        End With
        
            
    
    Else    'Displaying an error message, asking the user to choose a record
    
        MsgBox "Please Select a Record First!", vbExclamation, "No Record Selected!"
        Exit Sub
        
    End If
    
    Exit Sub
    
errorApply:
    MsgBox Err.Description & "" & Err.Number, vbCritical

End Sub





