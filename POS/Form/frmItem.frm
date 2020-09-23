VERSION 5.00
Begin VB.Form frmItem 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Maintenance"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   6375
      Begin VB.ComboBox cboSubCat 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox cboSupp 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1800
         Width           =   4815
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtOnhand 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtDescrip 
         Height          =   315
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   9
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtSku 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "SubCategory :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Price :"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "On Hand :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Category :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Description :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "SKU Number :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   3840
      Width           =   7575
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         Height          =   375
         Left            =   3720
         TabIndex        =   27
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "--->"
         Height          =   375
         Left            =   3240
         TabIndex        =   26
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<---"
         Height          =   375
         Left            =   2760
         TabIndex        =   25
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   500
         Left            =   5400
         Picture         =   "frmItem.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   500
         Left            =   4320
         Picture         =   "frmItem.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1065
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   500
         Left            =   3240
         Picture         =   "frmItem.frx":0AAE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1065
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   500
         Left            =   2160
         Picture         =   "frmItem.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   1065
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   500
         Left            =   1080
         Picture         =   "frmItem.frx":0C92
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1065
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   500
         Left            =   120
         Picture         =   "frmItem.frx":0D84
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   600
         Width           =   945
      End
   End
   Begin VB.Label lblSupcode 
      Caption         =   "Label8"
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSubcat 
      Caption         =   "Label8"
      Height          =   375
      Left            =   600
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   -960
      Picture         =   "frmItem.frx":0E76
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7590
   End
   Begin VB.Label lblSku 
      Caption         =   "Label8"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblCatcode 
      Caption         =   "Label8"
      Height          =   255
      Left            =   4920
      TabIndex        =   22
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnEdit As Boolean
Private Sub cboCategory_Click()
    Dim strTemp
    
    strTemp = "'" & CStr(cboCategory.Text) & "'"
    On Error Resume Next
    rsCategory.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsCategory.Find "Category = " & strTemp, 0, adSearchForward
    
    If rsCategory!Category = CStr(cboCategory.Text) Then
        lblCatcode.Caption = rsCategory!catcode
        LoadSubCat 'load subcategory
    End If
    On Error GoTo 0
    Err.Clear
    Exit Sub
    
ErrorNotOnFile:
    
    DoEvents
    On Error GoTo 0
    Err.Clear
End Sub

Private Sub cboSubCat_Click()
    Dim strTemp
    
    strTemp = "'" & CStr(cboSubCat.Text) & "'"
    On Error Resume Next
    rsSubCategory.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsSubCategory.Find "SubCat = " & strTemp, 0, adSearchForward
    'MsgBox cboSubCat.Text
    If rsSubCategory!subcat = CStr(cboSubCat.Text) Then
        lblSubcat.Caption = rsSubCategory!subcatcode
      
    End If
    On Error GoTo 0
    Err.Clear
    Exit Sub
    
ErrorNotOnFile:
    
    DoEvents
    On Error GoTo 0
    Err.Clear
End Sub


Private Sub cboSupp_Click()
     Dim strTemp
    
    strTemp = "'" & CStr(cboSupp.Text) & "'"
    On Error Resume Next
    rsSupp.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsSupp.Find "SUPPLIER = " & strTemp, 0, adSearchForward
    'MsgBox cboSupp.Text
    If rsSupp!Supplier = CStr(cboSupp.Text) Then
        lblsupcode.Caption = rsSupp!supcode
    End If
    On Error GoTo 0
    Err.Clear
    Exit Sub
    
ErrorNotOnFile:
    
    DoEvents
    On Error GoTo 0
    Err.Clear
End Sub

Private Sub cmdAdd_Click()
    blnStatus False
    ClearText
    txtSku.SetFocus
    If rsCategory.RecordCount > 0 Then
        rsCategory.MoveFirst
        cboCategory = rsCategory!Category
        lblCatcode.Caption = rsCategory!catcode
    End If
    
End Sub

Private Sub cmdCancel_Click()
    blnStatus True
    If rsItem.RecordCount > 0 Then rsItem.MoveFirst
    ClearText
    LoadItem
    blnEdit = False
End Sub

Private Sub cmdDelete_Click()
    Dim intRes As Integer
    Dim blnFound As Boolean
    intRes = MsgBox("Are you sure you want to delete?", vbQuestion + vbYesNo, "Delete Confirmation")
    If intRes = vbNo Then
        'blnStatus True
        Exit Sub
    Else
    
        strSearch = CStr(txtSku.Text)
        blnFound = FindItem()
        If blnFound = True And Not blnEdit = True Then
            rsItem.Delete
            MsgBox "Sku successfuly deleted.", vbInformation, "Deleted"
            ClearText
            If rsItem.EOF And rsItem.RecordCount > 0 Then
                Beep
                rsItem.MoveLast
            
            End If
                LoadItem
                Exit Sub
            End If
    End If

End Sub
Private Sub cmdEdit_Click()
    blnStatus False
    txtSku.SetFocus
    blnEdit = True
    lblSku.Caption = txtSku.Text
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    If rsItem.BOF Then Exit Sub
    If rsItem.RecordCount > 0 Then
        rsItem.MoveFirst
        LoadItem
    End If
End Sub

Private Sub cmdLast_Click()
    If rsItem.EOF Then Exit Sub
    If rsItem.RecordCount > 0 Then
        rsItem.MoveLast
        LoadItem
    End If
End Sub

Private Sub cmdNext_Click()
On Error GoTo errorlast
    If Not rsItem.EOF Then rsItem.MoveNext
    
    If rsItem.EOF And rsItem.RecordCount > 0 Then
        Beep
        rsItem.MoveLast
    End If
        LoadItem
    Exit Sub
errorlast:
    MsgBox Err.Description & Err.Number
End Sub

Private Sub cmdPrev_Click()
On Error GoTo errorfirst
    If Not rsItem.BOF Then rsItem.MovePrevious
    If rsItem.BOF And rsItem.RecordCount > 0 Then
        Beep
        rsItem.MoveFirst
    End If
        LoadItem
    Exit Sub
errorfirst:
    MsgBox Err.Description & Err.Number
End Sub

Private Sub cmdSave_Click()
    Dim intRes As Integer
    If Len(Trim(txtSku.Text)) = 0 Then
        intRes = MsgBox("Please input SKU number!", vbCritical + vbOKCancel, "Invalid SKU")
        If intRes = vbCancel Then
            cmdCancel_Click
            Exit Sub
        End If
        txtSku.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtDescrip.Text)) = 0 Then
        intRes = MsgBox("Please input item description!", vbCritical + vbOKCancel, "Invalid Description")
        If intRes = vbCancel Then
            cmdCancel_Click
            Exit Sub
        End If
        txtDescrip.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtPrice.Text)) = 0 Then
        intRes = MsgBox("Please input item price!", vbCritical + vbOKCancel, "Invalid Price")
        If intRes = vbCancel Then
            cmdCancel_Click
            Exit Sub
        End If
        txtPrice.SetFocus
        Exit Sub
    End If
    If cboCategory.Text = "" Then
        MsgBox "Please select from the Category!", vbCritical + vbOKCancel, "Invalid Category"
        Exit Sub
    End If
    If cboSubCat.Text = "" Then
        MsgBox "Please select from the Subcategory!", vbCritical + vbOKCancel, "Invalid SubCategory"
        Exit Sub
    End If
    If cboSupp.Text = "" Then
        MsgBox "Please select from the Supplier!", vbCritical + vbOKCancel, "Invalid Supllier"
        Exit Sub
    End If
    If blnEdit = True Then
        Dim blnFound As Boolean
        strSearch = CStr(lblSku.Caption)
        blnFound = FindItem()
        If blnFound = True Then
            rsItem.Update "Sku", txtSku.Text
            rsItem.Update "Descrip", txtDescrip.Text
            rsItem.Update "Supplier", cboSupp.Text
            rsItem.Update "Category", cboCategory.Text
            rsItem.Update "CatCode", lblCatcode.Caption
            rsItem.Update "SupCode", lblsupcode.Caption
            rsItem.Update "SubCatcode", lblSubcat.Caption
            If Not cboSubCat.Text = "" Then rsItem.Update "Subcategory", cboSubCat.Text
            If Not txtOnhand.Text = "" Then rsItem.Update "Stack", txtOnhand.Text
            rsItem.Update "Price", txtPrice.Text
            blnStatus True
        'Else
        '    MsgBox "SKU number cannot be change!", vbCritical, "Invalid SKU"
        '    ClearText
        '    blnStatus True
        End If
        blnEdit = False
    Else
        rsItem.AddNew
        rsItem!Sku = txtSku.Text
        rsItem!descrip = txtDescrip.Text
        rsItem!catcode = lblCatcode.Caption
        If Not cboSupp.Text = "" Then rsItem!Supplier = cboSupp.Text
        rsItem!Category = cboCategory.Text
        If Not cboSubCat.Text = "" Then rsItem!SubCategory = cboSubCat.Text
        If Not Len(Trim(txtOnhand.Text)) = 0 Then rsItem!stack = txtOnhand.Text
        rsItem!price = txtPrice.Text
        rsItem.Update
        blnStatus True
    End If
End Sub

Private Sub Form_Load()
     
    If rsItem.RecordCount > 0 Then rsItem.MoveFirst
    If rsCategory.RecordCount > 0 Then rsCategory.MoveFirst
    If rsSubCategory.RecordCount > 0 Then rsSubCategory.MoveFirst
    Do Until rsCategory.EOF
        If Not IsNull(rsCategory!Category) Then cboCategory.AddItem rsCategory!Category
        rsCategory.MoveNext
        DoEvents
    Loop
    LoadSupp
    LoadItem
    'LoadSubCat 'load subcategory
     ' load supplier
    blnStatus True 'enable status
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Public Sub LoadItem()
    On Error Resume Next
    txtSku.Text = rsItem!Sku
    If Not IsNull(rsItem!descrip) Then txtDescrip.Text = rsItem!descrip
    If Not IsNull(rsItem!Category) Then cboCategory.Text = rsItem!Category
    If Not IsNull(rsItem!SubCategory) Then cboSubCat.Text = rsItem!SubCategory
    If Not IsNull(rsItem!Supplier) Then cboSupp.Text = rsItem!Supplier
    txtOnhand.Text = rsItem!stack
    txtPrice.Text = rsItem!price
'    lblCatcode.Caption = rsCategory!catcode
    lblSku.Caption = rsItem!Sku
End Sub

Public Sub blnStatus(blnEnable As Boolean)
    cmdAdd.Enabled = blnEnable
    cmdEdit.Enabled = blnEnable
    cmdDelete.Enabled = blnEnable
    cmdExit.Enabled = blnEnable
    cmdSave.Enabled = Not blnEnable
    cmdCancel.Enabled = Not blnEnable
    Frame2.Enabled = Not blnEnable
    
End Sub

Public Sub ClearText()
    txtSku.Text = ""
    txtDescrip.Text = ""
    txtOnhand.Text = ""
    txtPrice.Text = ""
    lblCatcode = ""
End Sub

Public Sub LoadSubCat()
    cboSubCat.Clear
    rsSubCategory.Filter = "(Catcode ='" & lblCatcode.Caption & "')"
    Do Until rsSubCategory.EOF
        If Not IsNull(rsSubCategory!subcat) Then cboSubCat.AddItem rsSubCategory!subcat
        rsSubCategory.MoveNext
        DoEvents
    Loop
    If rsSubCategory.RecordCount > 0 Then
        rsSubCategory.MoveFirst
        cboSubCat.Text = rsSubCategory!subcat
    End If
End Sub

Public Sub LoadSupp()
    If rsSupp.RecordCount > 0 Then rsSupp.MoveFirst
    Do Until rsSupp.EOF
        If Not IsNull(rsSupp!Supplier) Then cboSupp.AddItem rsSupp!Supplier
        rsSupp.MoveNext
        DoEvents
    Loop
End Sub


Private Sub txtSku_KeyPress(KeyAscii As Integer)
    Const conZero As Integer = 48, conNine As Integer = 57
        Const conBackSpace As Integer = 8
        If (KeyAscii < conZero Or KeyAscii > conNine) And KeyAscii <> conBackSpace Then
            KeyAscii = 0
        End If
End Sub
Public Function FindItem() As Boolean
    Dim strTemp
    
    strTemp = "'" & strSearch & "'"
    On Error Resume Next
    rsItem.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsItem.Find "Sku = " & strTemp, 0, adSearchForward
    
    If rsItem!Sku = strSearch Then FindItem = True       'found
    On Error GoTo 0
    Err.Clear
    Exit Function
    
ErrorNotOnFile:
 '   MsgBox "Error =   " & Err.Number & Err.Description
    FindItem = False      'not found
    DoEvents
    On Error GoTo 0
    Err.Clear
End Function

Private Sub txtSku_LostFocus()
    Dim blnFound As Boolean
    strSearch = CStr(txtSku.Text)
    blnFound = FindItem()
    If blnFound = True And Not lblSku.Caption = txtSku.Text Then
        LoadItem
        blnStatus True
        MsgBox "Sku number already exist!", vbInformation, "Item"
    End If
                
End Sub
