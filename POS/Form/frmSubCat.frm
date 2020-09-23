VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subcategory Maintenance"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmSubCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6060
      Left            =   0
      ScaleHeight     =   6030
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   0
      Width           =   5200
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   4935
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtSub 
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtSubCatCode 
            DataField       =   "itemcode"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Subcategory :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Code :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Category :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   0
         TabIndex        =   2
         Top             =   4800
         Width           =   5535
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   500
            Left            =   120
            Picture         =   "frmSubCat.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   650
            Width           =   4815
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   500
            Left            =   3960
            Picture         =   "frmSubCat.frx":09BC
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   500
            Left            =   3000
            Picture         =   "frmSubCat.frx":0AAE
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "&Delete"
            Height          =   500
            Left            =   2040
            Picture         =   "frmSubCat.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   500
            Left            =   1080
            Picture         =   "frmSubCat.frx":0C92
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   500
            Left            =   120
            Picture         =   "frmSubCat.frx":0D84
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   950
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CatCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Category"
            Object.Width           =   5362
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   0
         Picture         =   "frmSubCat.frx":0E76
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5145
      End
      Begin VB.Label lblSubCatcode 
         Caption         =   "Label2"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblCatCode 
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSubCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnEdit As Boolean
Dim lngIndex As Long
Private Sub cboCategory_Click()
    txtSub.Text = ""
    lblSubCatcode.Caption = ""
    txtSubCatCode.Text = ""
  Dim strTemp
    
    strTemp = "'" & CStr(cboCategory.Text) & "'"
    On Error Resume Next
    rsCategory.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsCategory.Find "Category = " & strTemp, 0, adSearchForward
    
    If rsCategory!Category = CStr(cboCategory.Text) Then
        lblCatcode.Caption = rsCategory!catcode
        LoadSub 'load subcategory
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
    txtSub.Text = ""
    txtSubCatCode.Text = ""
    txtSub.SetFocus
    cboCategory.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    
    blnStatus True
    If Not rsSubCategory.EOF Then rsSubCategory.MoveNext
    LoadSub
    cboCategory.Enabled = True
End Sub

Private Sub cmdDel_Click()
    Dim blnFound As Boolean
    Dim intRes As Integer
    If txtSubCatCode.Text = "" Then Exit Sub
    intRes = MsgBox("Are you sure you want to delete the Subcategory?", vbQuestion + vbYesNo, "Delete Confirmation")
    If intRes = vbNo Then Exit Sub
    strSearch = CStr(txtSubCatCode.Text)
    blnFound = FindSub()
    If blnFound = True Then
        rsSubCategory.Delete
        rsSubCategory.Requery
        MsgBox "Subcategory successfuly deleted.", vbInformation, "Deleted"
        txtSubCatCode.Text = ""
        txtSub.Text = ""
        LoadSub
   End If
End Sub

Private Sub cmdEdit_Click()
    If txtSubCatCode.Text = "" Then Exit Sub
    blnStatus False
    cboCategory.Enabled = False
    txtSub.SetFocus
    blnEdit = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim blnFound As Boolean
    If txtSub.Text = "" Then
        MsgBox "Please input Subcategory!", vbCritical, "Invalid Subcategory"
        txtSub.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If blnEdit = True Then
         
    strSearch = CStr(txtSubCatCode.Text)
    blnFound = FindSub()
        If blnFound = True Then
            rsSubCategory.Update "Subcat", txtSub.Text
            blnEdit = False
        End If
    Else
        rsSubCategory.AddNew
        rsSubCategory!catcode = lblCatcode.Caption
        rsSubCategory!subcat = txtSub.Text
        rsSubCategory.Update
        rsSubCategory.Requery
        blnStatus True
    End If
        cboCategory.Enabled = True
        LoadSub
        blnStatus True
End Sub

Private Sub Form_Load()
     blnStatus True
    LoadCategory
    LoadSub
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Public Sub LoadSub()
    ListView1.ListItems.Clear
    Dim itm As ListItem
    rsSubCategory.Filter = "(Catcode ='" & lblCatcode.Caption & "')"
    If Not rsSubCategory.RecordCount < 1 Then rsSubCategory.MoveFirst
    Do Until rsSubCategory.EOF
        If Not IsNull(rsSubCategory!subcatcode) Then Set itm = ListView1.ListItems.Add(, , rsSubCategory!subcatcode)
        If Not IsNull(rsSubCategory!subcat) Then itm.SubItems(1) = rsSubCategory!subcat
        rsSubCategory.MoveNext
        DoEvents
    Loop
    If rsSubCategory.RecordCount > 0 Then
        rsSubCategory.MoveFirst
        txtSub.Text = rsSubCategory!subcat
        lblSubCatcode.Caption = rsSubCategory!subcatcode
        txtSubCatCode.Text = rsSubCategory!subcatcode
    End If
End Sub

Public Sub LoadCategory()
If rsCategory.RecordCount > 0 Then rsCategory.MoveFirst
    Do Until rsCategory.EOF
        If Not IsNull(rsCategory!Category) Then cboCategory.AddItem rsCategory!Category
        rsCategory.MoveNext
        DoEvents
    Loop
    If cboCategory.ListCount > 0 Then
        rsCategory.MoveFirst
        cboCategory.Text = rsCategory!Category
        lblCatcode.Caption = rsCategory!catcode
    End If
End Sub

Public Sub blnStatus(blnEnable As Boolean)
    ListView1.Enabled = blnEnable
    cmdAdd.Enabled = blnEnable
    cmdEdit.Enabled = blnEnable
    cmdDel.Enabled = blnEnable
    cmdExit.Enabled = blnEnable
    cmdSave.Enabled = Not blnEnable
    cmdCancel.Enabled = Not blnEnable
    txtSub.Enabled = Not blnEnable
End Sub

Private Sub ListView1_Click()
    If ListView1.ListItems.count = 0 Then Exit Sub
    txtSubCatCode.Text = ListView1.SelectedItem.Text
    lblSubCatcode.Caption = ListView1.SelectedItem.Text
    txtSub.Text = ListView1.SelectedItem.ListSubItems(1).Text
End Sub
Public Function FindSub() As Boolean
    Dim strTemp
    
    strTemp = "'" & strSearch & "'"
    On Error Resume Next
    rsSubCategory.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsSubCategory.Find "SubCatCode = " & strTemp, 0, adSearchForward
    
    If rsSubCategory!subcatcode = strSearch Then FindSub = True       'found
    On Error GoTo 0
    Err.Clear
    Exit Function
    
ErrorNotOnFile:
 '   MsgBox "Error =   " & Err.Number & Err.Description
    FindSub = False      'not found
    DoEvents
    On Error GoTo 0
    Err.Clear
End Function

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            lngIndex = ListView1.SelectedItem.Index + 1
            If lngIndex = ListView1.ListItems.count + 1 Then
                lngIndex = ListView1.SelectedItem.Index
            End If
            txtSubCatCode.Text = ListView1.ListItems(lngIndex).Text
            lblSubCatcode.Caption = ListView1.ListItems(lngIndex).Text
            txtSub.Text = ListView1.ListItems(lngIndex).ListSubItems(1).Text
    End Select
   
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            lngIndex = ListView1.SelectedItem.Index
            If lngIndex = ListView1.ListItems.count - 1 Then
                    lngIndex = ListView1.SelectedItem.Index
            End If
            txtSubCatCode.Text = ListView1.ListItems(lngIndex).Text
            lblSubCatcode.Caption = ListView1.ListItems(lngIndex).Text
            txtSub.Text = ListView1.ListItems(lngIndex).ListSubItems(1).Text
    End Select
End Sub
