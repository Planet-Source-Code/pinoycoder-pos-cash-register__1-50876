VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSupplier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Maintenance"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
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
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4935
         Begin VB.TextBox txtContact 
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox txtPerson 
            Height          =   285
            Left            =   1440
            TabIndex        =   14
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Left            =   1440
            TabIndex        =   13
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtSupplier 
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtSupCode 
            DataField       =   "itemcode"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   -120
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Contact Person :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Contact Number:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Address :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supplier :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   11
            Top             =   240
            Width           =   735
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
         TabIndex        =   1
         Top             =   4800
         Width           =   5535
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   500
            Left            =   120
            Picture         =   "frmSupplier.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   650
            Width           =   4815
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   500
            Left            =   3960
            Picture         =   "frmSupplier.frx":09BC
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   500
            Left            =   3000
            Picture         =   "frmSupplier.frx":0AAE
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "&Delete"
            Height          =   500
            Left            =   2040
            Picture         =   "frmSupplier.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   500
            Left            =   1080
            Picture         =   "frmSupplier.frx":0C92
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   950
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   500
            Left            =   120
            Picture         =   "frmSupplier.frx":0D84
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   950
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3836
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
            Text            =   "SupCode"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Supplier"
            Object.Width           =   5891
         EndProperty
      End
      Begin VB.Label lblSupplier 
         Height          =   135
         Left            =   3960
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   0
         Picture         =   "frmSupplier.frx":0E76
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnEdit As Boolean
Dim lngIndex As Long

Private Sub cmdCancel_Click()
    ClearText
    blnStatus True
    If Not rsCategory.EOF Then rsCategory.MoveNext
    LoadSupplier
End Sub

Private Sub cmdDel_Click()
    Dim blnFound As Boolean
    Dim intRes As Integer
    If txtSupCode.Text = "" Then Exit Sub
    intRes = MsgBox("Are you sure you want to delete the supplier?", vbQuestion + vbYesNo, "Delete Confirmation")
    If intRes = vbNo Then Exit Sub
    strSearch = CStr(txtSupCode.Text)
    blnFound = FindSupp()
    If blnFound = True Then
        rsSupp.Delete
        rsSupp.Requery
        MsgBox "Category successfuly deleted.", vbInformation, "Deleted"
        LoadSupplier
        ClearText
    End If
End Sub

Private Sub cmdEdit_Click()
    If txtSupCode.Text = "" Then Exit Sub
    blnStatus False
    txtSupplier.SetFocus
    blnEdit = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim blnFound As Boolean
    If txtSupplier.Text = "" Then
        MsgBox "Please input Supplier Name!", vbCritical, "Invalid Supplier"
        txtSupplier.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If blnEdit = True Then
         
    strSearch = CStr(txtSupCode.Text)
    blnFound = FindSupp()
        rsSupp.Update "Supplier", txtSupplier.Text
        If txtAddress.Text = "" Then
            rsSupp.Update "Address", Null
        Else
            rsSupp.Update "Address", txtAddress.Text
        End If
        If txtPerson.Text = "" Then
            rsSupp.Update "Contact", Null
        Else
            rsSupp.Update "Contact", txtPerson.Text
        End If
        If txtContact.Text = "" Then
            rsSupp.Update "Telno", Null
        Else
            rsSupp.Update "Telno", txtContact.Text
        End If
        rsSupp.Requery
        blnStatus True
        blnEdit = False
    Else
        rsSupp.AddNew
        rsSupp!Supplier = txtSupplier.Text
        If Not txtAddress.Text = "" Then rsSupp!address = txtAddress.Text
        If Not txtPerson.Text = "" Then rsSupp!contact = txtPerson.Text
        If Not txtContact.Text = "" Then rsSupp!telno = txtContact.Text
        rsSupp.Update
        rsSupp.Requery
        blnStatus True
    End If
        ClearText
        LoadSupplier
End Sub

Private Sub Form_Load()
    If Not rsSupp.BOF And rsSupp.RecordCount > 0 Then
        rsSupp.MoveFirst
        txtSupplier.Text = rsSupp!Supplier
        txtSupCode.Text = rsSupp!supcode
        If Not IsNull(rsSupp!address) Then txtPerson.Text = rsSupp!address
        If Not IsNull(rsSupp!contact) Then txtContact.Text = rsSupp!contact
        If Not IsNull(rsSupp!telno) Then txtAddress.Text = rsSupp!telno
    End If
    LoadSupplier
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Public Sub LoadSupplier()
   Dim itm As ListItem
    ListView1.ListItems.Clear
    If Not rsSupp.RecordCount < 1 Then rsSupp.MoveFirst
    Do Until rsSupp.EOF
        If Not IsNull(rsSupp!supcode) Then Set itm = ListView1.ListItems.Add(, , rsSupp!supcode)
        If Not IsNull(rsSupp!Supplier) Then itm.SubItems(1) = rsSupp!Supplier
        rsSupp.MoveNext
        DoEvents
    Loop
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
Private Sub cmdAdd_Click()
    blnStatus False
    ClearText
    txtSupplier.SetFocus
   
End Sub

Public Sub ClearText()
    txtSupplier.Text = ""
    txtPerson.Text = ""
    txtContact.Text = ""
    txtAddress.Text = ""
    txtSupCode.Text = ""
End Sub
Public Function FindSupp() As Boolean
    Dim strTemp
    
    strTemp = "'" & strSearch & "'"
    On Error Resume Next
    rsSupp.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsSupp.Find "SupCode = " & strTemp, 0, adSearchForward
    
    If rsSupp!supcode = strSearch Then FindSupp = True       'found
    On Error GoTo 0
    Err.Clear
    Exit Function
    
ErrorNotOnFile:
 '   MsgBox "Error =   " & Err.Number & Err.Description
    FindSupp = False      'not found
    DoEvents
    On Error GoTo 0
    Err.Clear
End Function


Private Sub ListView1_Click()
    ClearText
    Dim blnFound As Boolean
    strSearch = CStr(ListView1.SelectedItem.Text)
    blnFound = FindSupp()
    If blnFound = True Then
        txtSupCode.Text = ListView1.SelectedItem.Text
        txtSupplier.Text = ListView1.SelectedItem.ListSubItems(1).Text
        lblSupplier.Caption = ListView1.SelectedItem.ListSubItems(1).Text
        If Not IsNull(rsSupp!address) Then txtAddress.Text = rsSupp!address
        If Not IsNull(rsSupp!contact) Then txtPerson.Text = rsSupp!contact
        If Not IsNull(rsSupp!telno) Then txtContact.Text = rsSupp!telno
    End If
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            
            lngIndex = ListView1.SelectedItem.Index + 1
            
            If lngIndex = ListView1.ListItems.count + 1 Then
                lngIndex = ListView1.SelectedItem.Index
            End If
            ClearText
            Dim blnFound As Boolean
            strSearch = CStr(ListView1.ListItems(lngIndex).Text)
            blnFound = FindSupp()
            If blnFound = True Then
                txtSupCode.Text = ListView1.ListItems(lngIndex).Text
                txtSupplier.Text = ListView1.ListItems(lngIndex).ListSubItems(1).Text
                lblSupplier.Caption = ListView1.ListItems(lngIndex).ListSubItems(1).Text
                If Not IsNull(rsSupp!address) Then txtAddress.Text = rsSupp!address
                If Not IsNull(rsSupp!contact) Then txtPerson.Text = rsSupp!contact
                If Not IsNull(rsSupp!telno) Then txtContact.Text = rsSupp!telno
            End If
            
    End Select
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            
            lngIndex = ListView1.SelectedItem.Index
            
            If lngIndex = ListView1.ListItems.count - 1 Then
                lngIndex = ListView1.SelectedItem.Index
            End If
            ClearText
            Dim blnFound As Boolean
            strSearch = CStr(ListView1.ListItems(lngIndex).Text)
            blnFound = FindSupp()
            If blnFound = True Then
                txtSupCode.Text = ListView1.ListItems(lngIndex).Text
                txtSupplier.Text = ListView1.ListItems(lngIndex).ListSubItems(1).Text
                lblSupplier.Caption = ListView1.ListItems(lngIndex).ListSubItems(1).Text
                If Not IsNull(rsSupp!address) Then txtAddress.Text = rsSupp!address
                If Not IsNull(rsSupp!contact) Then txtPerson.Text = rsSupp!contact
                If Not IsNull(rsSupp!telno) Then txtContact.Text = rsSupp!telno
            End If
            
    End Select
End Sub
