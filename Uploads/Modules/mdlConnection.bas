Attribute VB_Name = "mdlConnection"

'Global ADO Object Variables
Option Explicit
Public conn As ADODB.Connection
Public rsUserAccount As ADODB.Recordset
Public rsPurchaseDetails As ADODB.Recordset
Public rsPurchaseDetailsInfo As ADODB.Recordset
Public rsRetailStoreDetails As ADODB.Recordset
Public rsCustomerDetails As ADODB.Recordset
Public rsItemDetails As ADODB.Recordset
Public rsServiceDetails As ADODB.Recordset
Public rsServicingDetailsInfo As ADODB.Recordset
Public rsServiceDepotDetails As ADODB.Recordset
Public rsViewPurchaseDetails As ADODB.Recordset
Public rsViewPurchases As ADODB.Recordset




Public Sub Connection()
    
    'Opening a connection to link the database
    
    Set conn = New ADODB.Connection 'Setting up a new connection
    
    'Connection String
    conn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=dbECI;Data Source=.\SQLEXPRESS"

    conn.Open 'Opening Connection

End Sub
    

Public Sub User_Account()

    'The Purpose of this function is to open the recordset "User_Account"
    
    Set rsUserAccount = New ADODB.Recordset
    
    With rsUserAccount
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from UserAccount"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub

Public Sub Purchase_Details()

    'The Purpose of this function is to manage the recordset "PurchaseDetails"
    
    Set rsPurchaseDetails = New ADODB.Recordset
    
    With rsPurchaseDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from PurchaseDetails"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Purchase_Details_Info()

    'The Purpose of this function is to select a Customer's Purchase Records
    
    Set rsPurchaseDetailsInfo = New ADODB.Recordset
    
    With rsPurchaseDetailsInfo
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from PurchaseDetails where [Customer_ID] = '" & frmPurchaseDetails.txtCustomerID.Text & "' AND [Purchase_Date] = '" & frmPurchaseDetails.txtPurchaseDate.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub RetailStoreDetails()

    'The Purpose of this function is to manage the recordset "RetailStoreView"
    
    Set rsRetailStoreDetails = New ADODB.Recordset
    
    With rsRetailStoreDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "select * from RetailStoreView"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub CustomerDetails()

    'The Purpose of this function is to manage the recordset "CustomerDetails"
    
    Set rsCustomerDetails = New ADODB.Recordset
    
    With rsCustomerDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "exec usp_CustomerSearch"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub ItemDetails()

    'The Purpose of this function is to manage the recordset "ItemDetails"
    
    Set rsItemDetails = New ADODB.Recordset
    
    With rsItemDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "exec usp_ItemSearch"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub

Public Sub Service_Details()

    'The Purpose of this function is to manage the recordset "ServiceAgreementDetails"
    
    Set rsServiceDetails = New ADODB.Recordset
    
    With rsServiceDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from ServiceAgreementDetails"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Servicing_Details_Info()

    'The Purpose of this function is to select a Customer's Purchase Records
    
    Set rsServicingDetailsInfo = New ADODB.Recordset
    
    With rsServicingDetailsInfo
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from ServiceAgreementDetails where [Customer_ID] = '" & frmServicingDetails.txtCustomerID.Text & "' AND [Purchase_Date] = '" & frmServicingDetails.txtPurchaseDate.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub ServiceDepot_Details()

    'The Purpose of this function is to manage the recordset "ServiceDepot"
    
    Set rsServiceDepotDetails = New ADODB.Recordset
    
    With rsServiceDepotDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from ServiceDepot"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub ViewPurchaseDetails()

    'The Purpose of this function is to manage the recordset "PurchasesTrigger"
    
    Set rsViewPurchaseDetails = New ADODB.Recordset
    
    With rsViewPurchaseDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from PurchasesTrigger"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub myViewPurchases()

    'The Purpose of this function is to manage the view "PurchasesView"
    
    Set rsViewPurchases = New ADODB.Recordset
    
    With rsViewPurchases
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from PurchasesView"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub








