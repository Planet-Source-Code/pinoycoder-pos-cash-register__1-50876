VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category Maintenance"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmCategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6060
      Left            =   0
      ScaleHeight     =   6030
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   0
      Width           =   5200
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5106
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
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   0
         TabIndex        =   6
         Top             =   4800
         Width           =   5535
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   500
            Left            =   120
            Picture         =   "frmCategory.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   500
            Left            =   1080
            Picture         =   "frmCategory.frx":09BC
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "&Delete"
            Height          =   500
            Left            =   2040
            Picture         =   "frmCategory.frx":0AAE
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   500
            Left            =   3000
            Picture         =   "frmCategory.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   500
            Left            =   3960
            Picture         =   "frmCategory.frx":0C92
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   500
            Left            =   120
            Picture         =   "frmCategory.frx":0D84
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   650
            Width           =   4815
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   4935
         Begin VB.TextBox txtCatCode 
            DataField       =   "itemcode"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtCategory 
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Category :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Code :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   -240
         Picture         =   "frmCategory.frx":0E76
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5400
      End
      Begin VB.Label lblCategory 
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnEdit As Boolean
Dim lngIndex As Long
Public Sub LoadCat()
    Dim itm As ListItem
    ListView1.ListItems.Clear
    If Not rsCategory.RecordCount < 1 Then rsCategory.MoveFirst
    Do Until rsCategory.EOF
        If Not IsNull(rsCategory!catcode) Then Set itm = ListView1.ListItems.Add(, , rsCategory!catcode)
        If Not IsNull(rsCategory!catcode) Then itm.SubItems(1) = rsCategory!Category
        rsCategory.MoveNext
        DoEvents
    Loop
End Sub

Private Sub cmdAdd_Click()
    blnStatus False
    txtCategory.SetFocus
    txtCategory.Text = ""
    txtCatCode.Text = ""
End Sub

Private Sub cmdCancel_Click()
    txtCategory.Text = ""
    txtCatCode.Text = ""
    blnStatus True
    If Not rsCategory.EOF Then rsCategory.MoveNext
    LoadCat
End Sub

Private Sub cmdDel_Click()
     Dim blnFound As Boolean
     Dim intRes As Integer
    intRes = MsgBox("Are you sure you want to delete the category?", vbQuestion + vbYesNo, "Delete Confirmation")
    If intRes = vbNo Then Exit Sub
    strSearch = CStr(txtCatCode.Text)
    blnFound = FindCat()
    If blnFound = True Then
        rsCategory.Delete
        rsCategory.Requery
        MsgBox "Category successfuly deleted.", vbInformation, "Deleted"
        LoadCat
        txtCategory.Text = ""
        txtCatCode.Text = ""
    End If
End Sub

Private Sub cmdEdit_Click()
    blnStatus False
    txtCategory.SetFocus
    blnEdit = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim blnFound As Boolean
    If blnEdit = True Then
         
    strSearch = CStr(txtCatCode.Text)
    blnFound = FindCat()
        rsCategory.Update "Category", txtCategory.Text
        rsCategory.Requery
        blnStatus True
        blnEdit = False
    Else
        rsCategory.AddNew
        rsCategory!Category = txtCategory.Text
        rsCategory.Update
        rsCategory.Requery
        blnStatus True
    End If
        txtCatCode.Text = ""
        txtCategory.Text = ""
        lblCategory.Caption = ""
        LoadCat
End Sub

Private Sub Form_Load()
    
    LoadCat
    
    blnStatus True
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub


Private Sub ListView1_Click()
    txtCatCode.Text = ListView1.SelectedItem.Text
    txtCategory.Text = ListView1.SelectedItem.ListSubItems(1).Text
    lblCategory.Caption = ListView1.SelectedItem.ListSubItems(1).Text
End Sub

Public Sub blnStatus(blnEnable As Boolean)
    ListView1.Enabled = blnEnable
    cmdAdd.Enabled = blnEnable
    cmdEdit.Enabled = blnEnable
    cmdDel.Enabled = blnEnable
    cmdExit.Enabled = blnEnable
    cmdSave.Enabled = Not blnEnable
    cmdCancel.Enabled = Not blnEnable
    Frame3.Enabled = Not blnEnable
End Sub
Public Function FindCat() As Boolean
    Dim strTemp
    
    strTemp = "'" & strSearch & "'"
    On Error Resume Next
    rsCategory.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsCategory.Find "CatCode = " & strTemp, 0, adSearchForward
    
    If rsCategory!catcode = strSearch Then FindCat = True       'found
    On Error GoTo 0
    Err.Clear
    Exit Function
    
ErrorNotOnFile:
 '   MsgBox "Error =   " & Err.Number & Err.Description
    FindCat = False      'not found
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
        txtCatCode.Text = ListView1.ListItems(lngIndex).Text
        txtCategory.Text = ListView1.ListItems(lngIndex).ListSubItems(1).Text
        lblCategory.Caption = ListView1.ListItems(lngIndex).ListSubItems(1).Text
    End Select
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
        Case vbKeyUp
        lngIndex = ListView1.SelectedItem.Index
        If lngIndex = ListView1.ListItems.count - 1 Then
            lngIndex = ListView1.SelectedItem.Index
        End If
        txtCatCode.Text = ListView1.ListItems(lngIndex).Text
        txtCategory.Text = ListView1.ListItems(lngIndex).ListSubItems(1).Text
        lblCategory.Caption = ListView1.ListItems(lngIndex).ListSubItems(1).Text
    End Select
End Sub

Private Sub txtCategory_LostFocus()
    Dim blnFound As Boolean
    strSearch = CStr(txtCatCode.Text)
    blnFound = FindCat()
    If blnFound = True And txtCategory.Text = lblCategory.Caption Then
        LoadCat
        blnStatus False
        MsgBox "Category already exist!", vbInformation, "Category"
    End If
End Sub
