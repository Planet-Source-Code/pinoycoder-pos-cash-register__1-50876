Attribute VB_Name = "ADOmod"
Option Explicit
Public cn As ADODB.Connection
Public cnUser As ADODB.Connection

Public MSDatabase
Public UserDatabase


Global Const DEFSOURCE = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source="
Public db As ADODB.Connection

Public rsItem As ADODB.Recordset
Public rsCus As ADODB.Recordset
Public rsStore As ADODB.Recordset
Public rsUser As ADODB.Recordset
Public rsCategory As ADODB.Recordset
Public rsSubCategory As ADODB.Recordset
Public rsSupp As ADODB.Recordset
Public rsSale As ADODB.Recordset
Public rsRet As ADODB.Recordset
Public rsDel As ADODB.Recordset

Public DBConnect As Boolean
Public UserConnected As Boolean

Dim strSQl As String

Public Sub OpenDB()
    Set db = New ADODB.Connection
    db.Open DEFSOURCE & App.Path & "\Data\Data.MDB;"
    DEnv.Connection1 = DEFSOURCE & App.Path & "\Data\Data.MDB;"
End Sub

Public Function DataConnect() As Boolean

On Error GoTo OpenErr
Set cn = New ADODB.Connection
cn.CursorLocation = adUseClient

MSDatabase = App.Path & ("\Data\Data.MDB")

    cn.CursorLocation = adUseClient
    cn.PROVIDER = "Microsoft.Jet.OLEDB.3.51; Jet OLEDB:Database Password="
    cn.Open MSDatabase ', Admin
    
  
      
Exit Function

OpenErr:

    MsgBox "Error Opening " & MSDatabase & vbNewLine & Err.Description, vbCritical, "Open Database Error"
    DBConnect = False


End Function


Public Sub OpenData()
    Set rsUser = New ADODB.Recordset
    strSQl = "SELECT *FROM  tblUserInfo"
    rsUser.Open strSQl, cn, adOpenStatic, adLockOptimistic

      
    Set rsItem = New ADODB.Recordset
    rsItem.Open "Select * FROM Item", cn, adOpenStatic, adLockOptimistic

    Set rsStore = New ADODB.Recordset
    rsStore.Open "Select * FROM StoreName", cn, adOpenStatic, adLockOptimistic

    Set rsCategory = New ADODB.Recordset
    rsCategory.Open "Select * FROM Category", cn, adOpenStatic, adLockOptimistic

    Set rsSubCategory = New ADODB.Recordset
    rsSubCategory.Open "Select * FROM SubCategory", cn, adOpenStatic, adLockOptimistic

    Set rsSupp = New ADODB.Recordset
    rsSupp.Open "Select * FROM Supplier", cn, adOpenStatic, adLockOptimistic

    Set rsSale = New ADODB.Recordset
    rsSale.Open "Select * FROM Sale", cn, adOpenStatic, adLockOptimistic
    
    Set rsRet = New ADODB.Recordset
    rsRet.Open "Select * FROM Return", cn, adOpenStatic, adLockOptimistic
    
    Set rsDel = New ADODB.Recordset
    rsDel.Open "Select * FROM Delivery", cn, adOpenStatic, adLockOptimistic

End Sub
Public Function UserConnect() As Boolean

On Error GoTo OpenErr
Set cnUser = New ADODB.Connection
cnUser.CursorLocation = adUseClient

UserDatabase = App.Path & "\Data\User.mdb"

    cnUser.CursorLocation = adUseClient
    cnUser.PROVIDER = "Microsoft.Jet.OLEDB.4.0; Jet OLEDB:Database Password=wolfgang"
    cnUser.Open UserDatabase ', Admin
    UserConnected = True
  
      
Exit Function

OpenErr:

    MsgBox "Error Opening " & MSDatabase & vbNewLine & Err.Description, vbCritical, "Open Database Error"
    UserConnected = False


End Function

