VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmServicingDetails 
   Caption         =   "Servicing Details Interface"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmServicingDetails.frx":0000
   ScaleHeight     =   10080
   ScaleWidth      =   12375
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboDuration 
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
      ItemData        =   "frmServicingDetails.frx":1EDC4
      Left            =   2760
      List            =   "frmServicingDetails.frx":1EDD1
      TabIndex        =   29
      Text            =   "----------SELECT-----------"
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtItemType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   4680
      Width           =   2295
   End
   Begin VB.PictureBox picInvalidKeypressMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3840
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   240
         TabIndex        =   26
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.TextBox txtItemID 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox txtServiceDepotID 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtServiceID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtPurchaseID 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtNetTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   0
      Top             =   5040
   End
   Begin VB.CommandButton cmdAdd 
      DisabledPicture =   "frmServicingDetails.frx":1EDDE
      Height          =   855
      Left            =   4200
      Picture         =   "frmServicingDetails.frx":1F1E0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmServicingDetails.frx":21F24
      Height          =   855
      Left            =   6600
      Picture         =   "frmServicingDetails.frx":223E3
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdServiceDepotSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "Click Here to select a Service Depot"
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmdPurchaseSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "Click Here to select a Purchase Record"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Left            =   2760
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmServicingDetails.frx":25127
      Height          =   855
      Left            =   5400
      Picture         =   "frmServicingDetails.frx":255A5
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox txtPurchaseDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtCustomerID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtTotalCost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   6600
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid dgrdServicingInfo 
      Height          =   3855
      Left            =   6120
      TabIndex        =   14
      Top             =   2280
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6800
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
      Caption         =   "Service Agreement Information Table"
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
   Begin VB.Label lblItemType 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type"
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
      Left            =   1080
      TabIndex        =   28
      Top             =   4725
      Width           =   1575
   End
   Begin VB.Label lblPurchaseID 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   24
      Top             =   3285
      Width           =   1335
   End
   Begin VB.Label lblItemID 
      BackStyle       =   0  'Transparent
      Caption         =   "Item ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   23
      Top             =   4245
      Width           =   1335
   End
   Begin VB.Label lblServiceID 
      BackStyle       =   0  'Transparent
      Caption         =   "Service  ID"
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
      Left            =   1080
      TabIndex        =   22
      Top             =   2325
      Width           =   1575
   End
   Begin VB.Label lblServiceDepotID 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Depot ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   2805
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   5295
      Left            =   600
      Top             =   1920
      Width           =   5055
   End
   Begin VB.Label lblNetTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "NET TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   6750
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   5295
      Left            =   5880
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   5880
      X2              =   11400
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000001&
      Height          =   1335
      Left            =   3960
      Top             =   7800
      Width           =   3855
   End
   Begin VB.Label lblQuantity 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Left            =   1080
      TabIndex        =   19
      Top             =   6165
      Width           =   1575
   End
   Begin VB.Label lblDuration 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration"
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
      Left            =   1080
      TabIndex        =   18
      Top             =   5685
      Width           =   1575
   End
   Begin VB.Label lblPurchaseDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   5205
      Width           =   1335
   End
   Begin VB.Label lblCustomerID 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   3765
      Width           =   1695
   End
   Begin VB.Label lblTotalCost 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   6645
      Width           =   1575
   End
End
Attribute VB_Name = "frmServicingDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------------------------
'Process Management System - Inspirat Edition
'Form Name: Servicing Details Interface
'Programmer: Ahamed Imran Sheriff
'Quality Assurance Engineer (Testing): Salvin Ali Saleh
'Start Date: 27/01/2009
'Date Of Last Modification: 27/01/2009
'The Name Of The Database Being Accessed: dbECI
'The Name/s Of The Database Table/s Being Accessed: dbo.ServiceAgreementDetails
'------------------------------------------------------------------------------


Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean

'The following variables will be used to autogenerate the Purchase ID to be
'displayed on the Purchase Details
Dim iNumOfPurchases As Integer  'This variable holds the number of records in the table
Dim strDisplay As String  'This variable will eventually hold the Purchase ID to be autogenerated

'The following variables will be used to autogenerate the Purchase ID to be
'displayed on the Purchase Details form on the second purchase
Dim iNumberOfPurchases As Integer  'This variable holds the number of records in the table
Dim strCode As String  'This variable will eventually hold the Purchase ID to be autogenerated



Private Sub cmdAdd_Click()

    On Error GoTo errorAdd

    enableAllFields     'Calling a Private Function To Enable All Fields
    
    dgrdServicingInfo.Enabled = True    'Enabling the datagrid
    
    cmdServiceDepotSearch.Enabled = True 'Enabling the Service Depot Search wizard button
    
    cmdPurchaseSearch.Enabled = True    'Enabling the Purchases Search wizard button
        
    cmdSave.Enabled = True  'Enabling the Save button
    
    cmdAdd.Enabled = False  'Disabling the Add New button
    
    Call Service_Details    'Calling the Service_Details Procedure to interact with the recordset

    'Generate Service ID By Utilizing the ServiceAgreementDetails Table
    With rsServiceDetails

        If .RecordCount = 0 Then    'If there are no records in the table

            strDisplay = "SA001"

        Else

            'Calculating the number of records and storing in a variable
            iNumOfPurchases = .RecordCount
            iNumOfPurchases = iNumOfPurchases + 1   'incrementing the number by 1

            'The following block of code will generate the ID according
            'to the number of records in the ServiceAgreementDetails Table
            If iNumOfPurchases < 10 Then
                strDisplay = "SA00" & iNumOfPurchases
            ElseIf iNumOfPurchases < 100 Then
                strDisplay = "SA0" & iNumOfPurchases
            ElseIf iNumOfPurchases < 1000 Then
                strDisplay = "SA" & iNumOfPurchases
            End If

        End If
        
        txtServiceID.Text = strDisplay 'Displaying the generated ID in the textfield

        .Requery    'Requerying the Table
        
        .AddNew

    End With
    
    Exit Sub
    
errorAdd:
    MsgBox Err.Description & "" & Err.Number, vbCritical

End Sub


Private Function saveProcedure()    'This procedure will save the record into the database.


    With rsServiceDetails
        
        '.AddNew     'Adding a new recordset

        
        'Save the user-entered data into the recordset
        .Fields(0) = txtServiceID.Text
        .Fields(1) = txtServiceDepotID.Text
        .Fields(2) = txtPurchaseID.Text
        .Fields(3) = txtCustomerID.Text
        .Fields(4) = txtPurchaseDate.Text
        .Fields(5) = cboDuration.Text
        .Fields(6) = txtQuantity.Text
        .Fields(7) = txtTotalCost.Text


        .Update

        'Display Success Message
        MsgBox "The Record Was Added Successfully!", vbInformation, "Succesful Save Procedure!"

        .Requery    'Requerying the Table
        
    End With

End Function


Private Sub cmdClose_Click()

    'Obtaining confirmation from the user
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If

End Sub


Private Sub cmdServiceDepotSearch_Click()

    frmServiceDepotSearchWizard.Show 'Show the search wizard.

End Sub

Private Sub cmdPurchaseSearch_Click()  'On click of the Inpatients Search Wizard Button

    frmPurchasesSearchWizard.Show 'Show the search wizard.

End Sub



Private Sub cmdSave_Click()

    On Error GoTo errorSave

    If textfieldsValidations = False Then

        If MsgBox("Are You Sure You Wish To Add This Record?", vbYesNo + vbQuestion, "Add This Record?") = vbYes Then

            'Enabling the DataGrid
            dgrdServicingInfo.Enabled = True

            txtNetTotal.Text = Val(txtNetTotal.Text) + Val(txtTotalCost.Text)

            saveProcedure   'Calling a function which will save the record in the database

            Call Servicing_Details_Info    'Calling the Servicing_Details_Info Function

            Set dgrdServicingInfo.DataSource = rsServicingDetailsInfo 'Setting the datasource for the datagrid

        Else

            'Display 'No Modifications' Message
            MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"

        End If

        
        
        'Checking if the user wants to add another record for the same patient
        If MsgBox("Does The Customer Wish To Make Another Purchase?", vbYesNo + vbQuestion, "Make New Purchase Record?") = vbYes Then

            'Clearing All Necessary Textfields
            txtPurchaseID.Text = ""
            txtCustomerID.Text = ""
            txtItemID.Text = ""
            txtItemType.Text = ""
            cboDuration.Text = "----------SELECT-----------"
            txtQuantity.Text = ""
            txtTotalCost.Text = ""
            
            With rsServiceDetails


                'Calculating the number of records and storing in a variable
                iNumberOfPurchases = .RecordCount
                iNumberOfPurchases = iNumberOfPurchases + 1   'incrementing the number by 1
        
                'The following block of code will generate the ID according
                'to the number of records in the ServiceAgreementDetails Table
                If iNumberOfPurchases < 10 Then
                    strCode = "SA00" & iNumberOfPurchases
                ElseIf iNumberOfPurchases < 100 Then
                    strCode = "SA0" & iNumberOfPurchases
                ElseIf iNumberOfPurchases < 1000 Then
                    strCode = "SA" & iNumberOfPurchases
                End If
                
                txtServiceID.Text = strCode 'Displaying the generated ID in the textfield
        
                .Requery    'Requerying the Table
        
                .AddNew     'Adding a new recordset
                
                'Disabling the search buttons for Service Depot
                cmdServiceDepotSearch.Enabled = False

            End With


        Else

            On Error GoTo e
            deReports.Commands("rptServiceInvoice").Parameters(0) = txtCustomerID.Text
            deReports.Commands("rptServiceInvoice").Parameters(1) = DateTime.Date
            rptServiceInvoice.Show
            'deReports.rptInvoice.Close
                
            Unload Me
            Exit Sub
e:
            If Err.Number <> 3704 Then
                MsgBox Err.Description & "" & Err.Number, vbCritical
            End If

        End If

    End If
    
    Exit Sub
    
errorSave:
    MsgBox Err.Description & "" & Err.Number, vbCritical

End Sub


Private Sub Form_Load()

    disableAllFields  'Calling a Private Function To Disable All Fields
    
    cmdServiceDepotSearch.Enabled = False    'Disabling the Service Depot Search wizard button
    
    cmdPurchaseSearch.Enabled = False   'Disabling the Purchase Search wizard button
    
    dgrdServicingInfo.Enabled = False   'Disabling the Datagrid
    
    cmdSave.Enabled = False     'Disabling the Save command button
    
    txtPurchaseDate.Text = DateTime.Date 'Displaying the date in the Purchase Date textfield.
    
    Call Connection 'Calling the Connection function to set up a connection with the database

    
End Sub




Private Sub txtQuantity_Change()

    If txtQuantity.Text = "0" Then
        MsgBox "Error! The Figure Cannot Begin With Zero!", vbCritical, "Cannot Begin Figure With 0!"
        txtQuantity.Text = ""
        Exit Sub
    End If
    
    
    Dim iHold As Integer    'This will store  the number of purchases
    Call Purchase_Details
    With rsPurchaseDetails
        .MoveFirst
        Do While .EOF = False
            If .Fields(0).Value = frmServicingDetails.txtPurchaseID.Text Then
                iHold = .Fields(5).Value
            End If
            .MoveNext
        Loop
    End With
    
    
    If Val(txtQuantity.Text) > iHold Then
        MsgBox "Error! The Figure You Have Entered is Too High!", vbCritical, "Please Check No. of Service Agreements!"
        txtQuantity.Text = ""
        Exit Sub
    End If

    
    'Calculating the Total Cost based on the duration
    'of the Service Agreement
    If Val(cboDuration.Text) = 1 Then
        txtTotalCost.Text = Val(txtQuantity.Text) * 1000
    End If
    
    If Val(cboDuration.Text) = 3 Then
        txtTotalCost.Text = Val(txtQuantity.Text) * 3000
    End If
    
    If Val(cboDuration.Text) = 5 Then
        txtTotalCost.Text = Val(txtQuantity.Text) * 5000
    End If
    
    
End Sub


Private Sub txtQuantity_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only digits

    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 5880    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub



Private Sub tmrErrMsg_Timer()

    Static i As Integer

    If i < 200000 Then     'Validation Msg Viewing Time Period
        picInvalidKeypressMsg.Visible = False
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If

End Sub


Private Function textfieldsValidations() As Boolean  'This function will validate all fields

    Flag = True 'Setting the Flag variable to True


    'Checking if the Service Depot ID textfield is empty
    If txtServiceDepotID.Text = "" Then
        txtServiceDepotID.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtServiceDepotID.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If


    'Checking if the Purchase ID textfield is empty
    If txtPurchaseID.Text = "" Then
        txtPurchaseID.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtCustomerID.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtItemID.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtItemType.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtPurchaseID.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtCustomerID.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtItemID.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtItemType.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    
    'Checking if the Total Cost textfield has been filled in
    If txtTotalCost.Text = "0" Then
        cboDuration.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtQuantity.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtTotalCost.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboDuration.BackColor = &H80000004   'Highlighting the textfield in a different colour
        txtQuantity.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtTotalCost.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If


    'Here, I am checking the state of the Flag variable and if it is False, I am displaying a
    'Message Box to instruct the user to enter data into all highlighted textfields.
    'The Save procedure will also be cancelled
    If Flag = False Then
        MsgBox "Error! Please Fill-in The Highlighted Textfields! They Are Compulsory!", vbCritical, "Please Fill Highlighted Textfields"
        textfieldsValidations = True    'Passing values to the Save procedure
    Else
        textfieldsValidations = False   'Passing values to the Save procedure
    End If
    

End Function


Public Function enableAllFields() 'This function will enable all fields on the interface

    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will enable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Then
        eachField.Enabled = True
    End If

    Next
    
End Function


Private Function disableAllFields() 'This function will disable all fields on the interface

    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will disable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Then
        eachField.Enabled = False
    End If

    Next

End Function

