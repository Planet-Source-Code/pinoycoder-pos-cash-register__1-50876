Attribute VB_Name = "DeclareMod"
Public lngRow As Long
Public blnLoad As Boolean
Public blnItemlook As Boolean
Public blnSaleReport As Boolean
Public blnDelReport As Boolean
Public blnReceipt As Boolean

Public Sub Main()
   
    'check if already run
    If App.PrevInstance Then
        ActivatePrevInstance
    End If
    
    'open data connection
    DataConnect
   
    'open table
    OpenData
'    ADOConnect
    'frmSales.Show
    frmMain.Show
   ' If DBConnect = True Then
   '     OpenData
   ' End If
    OpenDB
    
End Sub

Public Sub LoadSale()
   frmSales.Show 1
   
 
End Sub
Public Sub LoadUser()
 frmUser.Show 1
   
 
End Sub
Public Sub LoadStore()
  frmStore.Show 1
   
 
End Sub
Public Sub LoadMaint()
    With frmMain
        .Frame1.Visible = False
        .picMaintenance.Visible = True
    End With
End Sub
Public Sub ReturnMain()
    With frmMain
        .Frame1.Visible = True
        .CoolBar1.Visible = True
        .picMaintenance.Visible = False
        .picInventory.Visible = False
        .Picture3.Visible = False
    End With
    
    
End Sub

Public Sub LoadInvent()
    With frmMain
        .Frame1.Visible = False
        .picInventory.Visible = True
    End With
End Sub

Public Sub LoadRep()
    With frmMain
        .Frame1.Visible = False
        .Picture3.Visible = True
    End With
End Sub
