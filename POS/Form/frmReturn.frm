VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return To Supplier"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmReturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6495
      ScaleWidth      =   8175
      TabIndex        =   5
      Top             =   0
      Width           =   8175
      Begin VB.Frame Frame5 
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         Height          =   1695
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   2535
         Begin VB.TextBox txtQtyS 
            Height          =   300
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtCountS 
            Height          =   300
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtRef 
            Height          =   300
            Left            =   1200
            TabIndex        =   1
            Top             =   600
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DDate 
            Height          =   315
            Left            =   1200
            TabIndex        =   22
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   61014017
            CurrentDate     =   37368
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Qty :"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Count :"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ref No. :"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   9015
         TabIndex        =   16
         Top             =   5640
         Width           =   9015
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   615
            Left            =   4680
            Picture         =   "frmReturn.frx":08CA
            TabIndex        =   35
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   615
            Left            =   7200
            Picture         =   "frmReturn.frx":09BC
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   615
            Left            =   3480
            Picture         =   "frmReturn.frx":0AAE
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   615
            Left            =   2640
            Picture         =   "frmReturn.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "&Delete"
            Height          =   615
            Left            =   1800
            Picture         =   "frmReturn.frx":0C92
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   615
            Left            =   960
            Picture         =   "frmReturn.frx":0D84
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   615
            Left            =   120
            Picture         =   "frmReturn.frx":0E76
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "Po&st"
            Height          =   615
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   2880
         TabIndex        =   6
         Top             =   840
         Width           =   5175
         Begin VB.TextBox txtSupplier 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox txtDescrip 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCategory 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   960
            Width           =   3735
         End
         Begin VB.TextBox txtQty 
            Height          =   300
            Left            =   3840
            TabIndex        =   3
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtPrice 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtSku 
            Height          =   300
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sku:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Category :"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier :"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Qty :"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   3360
            TabIndex        =   11
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Price :"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   720
            TabIndex        =   10
            Top             =   1680
            Width           =   615
         End
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2535
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sku"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Price"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ref No."
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   810
         Left            =   600
         Picture         =   "frmReturn.frx":0F68
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label lblsupcode 
         Caption         =   "Label1"
         Height          =   255
         Left            =   6360
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblSupcatcode 
         Caption         =   "Label1"
         Height          =   255
         Left            =   5040
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblCatcode 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3600
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnAdd As Boolean
Dim intSalCount As Integer
Dim lngIndex As Long
Private Sub cmdAdd_Click()
    blnAdd = True
    blnStatus False
    Cleartxt
    If ListView3.ListItems.count = 0 Then
        txtRef.SetFocus
    Else
        txtSku.SetFocus
    End If
End Sub
Public Sub blnStatus(blnEnable As Boolean)
    ListView3.Enabled = blnEnable
    cmdAdd.Enabled = blnEnable
    cmdEdit.Enabled = blnEnable
    cmdDel.Enabled = blnEnable
    cmdExit.Enabled = blnEnable
    cmdSave.Enabled = Not blnEnable
    cmdCancel.Enabled = Not blnEnable
   Frame6.Enabled = Not blnEnable
   Frame5.Enabled = Not blnEnable
End Sub

Private Sub cmdCancel_Click()
    Cleartxt
    blnStatus True
    blnAdd = False
End Sub

Private Sub cmdClear_Click()
    Dim intRes As Integer
    intRes = MsgBox("Are you sure you want to clear all data?", vbQuestion + vbYesNo, "Clear Confirmation")
    If intRes = vbNo Then Exit Sub
    ListView3.ListItems.Clear
    Cleartxt
End Sub

Private Sub cmdDel_Click()
    Dim intRes As Integer
    intRes = MsgBox("Are you sure you want to delete the selected item?", vbQuestion + vbYesNo, "Delete Confirmation")
    If intRes = vbNo Then Exit Sub
    txtCountS.Text = ListView3.ListItems.count
    txtQtyS.Text = CLng(txtQtyS.Text) - CLng(txtQty.Text)
    ListView3.ListItems.Remove (lngIndex)
    Cleartxt
End Sub

Private Sub cmdEdit_Click()
    blnStatus False
    txtSku.SetFocus
End Sub

Private Sub cmdExit_Click()
    Dim intRes As Integer
    If ListView3.ListItems.count > 0 Then
        intRes = MsgBox("Are you sure you want exit now?" & vbCrLf & vbCrLf & "If you exit all data will be removed!", vbQuestion + vbYesNo, "Exit Confirmation")
        If intRes = vbNo Then Exit Sub
        Unload Me
    End If
        Unload Me
End Sub

Private Sub cmdPost_Click()
    Dim lngRec As Long
    Dim blnFound As Boolean
    Dim intRes As Integer
    intRes = MsgBox("Post data?", vbQuestion + vbYesNo, "Post Confirmation")
    If intRes = vbNo Then Exit Sub
    
    With rsRet
       
    For lngRec = 1 To ListView3.ListItems.count
        strSearch = CStr(Trim(ListView3.ListItems(lngRec).Text))
        blnFound = FindItem()
        If blnFound = True Then
            .AddNew
            !Sku = rsItem!Sku
            !descrip = rsItem!descrip
            !price = rsItem!price
            If Not IsNull(rsItem!supcode) Then !supcode = rsItem!supcode
            If Not IsNull(rsItem!subcatcode) Then !subcatcode = lblSupcatcode.Caption
            If Not IsNull(rsItem!catcode) Then !catcode = rsItem!catcode
            !dateentry = ListView3.ListItems(lngRec).ListSubItems(5).Text
            !Qty = ListView3.ListItems(lngRec).ListSubItems(3).Text
            If Not Len(Trim(ListView3.ListItems(lngRec).ListSubItems(4).Text)) = 0 Then !refno = ListView3.ListItems(lngRec).ListSubItems(4).Text
            !datepost = SQLDate(Now)
            !usercode = frmMain.lblUserCode
            .Update
            rsItem.Update "stack", CLng(rsItem!stack) - CLng(ListView3.ListItems(lngRec).ListSubItems(3).Text)
        End If
        DoEvents
    Next lngRec
    End With
    MsgBox "Successfuly posted!", vbInformation
    ListView3.ListItems.Clear
    Cleartxt
End Sub

Private Sub cmdSave_Click()
    If txtSku.Text = "" Then
        txtSku.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If Len(Trim(txtQty.Text)) = 0 Then
        MsgBox "Please input quantity!", vbExclamation + vbOKOnly, "Invalid Quantity"
        txtQty.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    SaveReturn
    blnStatus True
    Cleartxt
    cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
    blnStatus True
    DDate.Value = Now
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Public Sub Cleartxt()
    txtSku.Text = ""
    txtDescrip.Text = ""
    txtCategory.Text = ""
    txtSupplier.Text = ""
    txtPrice.Text = ""
    txtQty.Text = ""
    lblCatcode.Caption = ""
    lblSupcatcode.Caption = ""
    lblsupcode.Caption = ""
End Sub


Private Sub ListView3_Click()
    If ListView3.ListItems.count > 0 Then
        lngIndex = ListView3.SelectedItem.Index
        DisplayItem
    End If
End Sub


Private Sub txtQty_KeyPress(KeyAscii As Integer)
    udp_Rtrn KeyAscii
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
    udp_Rtrn KeyAscii
End Sub

Private Sub txtSku_KeyPress(KeyAscii As Integer)
    udp_Rtrn KeyAscii
End Sub

Private Sub txtSku_LostFocus()
    Dim blnFound As Boolean
    Dim lntRes As Long
    strSearch = CStr(txtSku.Text)
    blnFound = FindItem()
    If blnFound = True Then
        LoadText
        txtQty.SetFocus
    Else
        lntRes = MsgBox("Sku not found!" & vbCrLf & vbCrLf & "Continue anyway?", vbQuestion + vbYesNo, "Invalid Sku")
        If lntRes = vbNo Then
            Cleartxt
            blnStatus True
            Exit Sub
        Else
            txtSku.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
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


Public Sub LoadText()
      txtSku.Text = rsItem!Sku
    txtDescrip.Text = rsItem!descrip
    txtCategory.Text = rsItem!Category
    txtSupplier.Text = rsItem!Supplier
    txtPrice.Text = rsItem!price
    If Not IsNull(rsItem!catcode) Then lblCatcode.Caption = rsItem!catcode
    If Not IsNull(rsItem!subcatcode) Then lblSupcatcode.Caption = rsItem!subcatcode
    If Not IsNull(rsItem!supcode) Then lblsupcode.Caption = rsItem!supcode
    txtQty.Text = ""
End Sub

Public Sub SaveReturn()
    Dim itm As ListItem
          
    If blnAdd = True Then
        Set itm = ListView3.ListItems.Add(, , txtSku.Text)
        itm.SubItems(1) = txtDescrip.Text
        itm.SubItems(2) = txtPrice.Text
        itm.SubItems(3) = txtQty.Text
        If Not txtRef.Text = "" Then itm.SubItems(4) = txtRef.Text
        itm.SubItems(5) = SQLDate(DDate)
        txtCountS.Text = ListView3.ListItems.count
        txtQtyS.Text = CLng(txtQtyS.Text) + CLng(txtQty.Text)
        blnAdd = False
    Else
        With ListView3.ListItems(lngIndex)
        txtQtyS.Text = CLng(txtQtyS.Text) + txtQty.Text - CLng(ListView3.SelectedItem.ListSubItems(3).Text)
        .Text = txtSku.Text
        .ListSubItems(1).Text = txtDescrip.Text
        .ListSubItems(2).Text = txtPrice.Text
        .ListSubItems(3).Text = txtQty.Text
        If Not txtRef.Text = "" Then .ListSubItems(4).Text = txtRef.Text
        .ListSubItems(5).Text = SQLDate(DDate)
        End With
    End If
    
    
End Sub

Public Sub DisplayItem()
    Dim blnFound As Boolean
    strSearch = CStr(Trim(ListView3.SelectedItem.Text))
    blnFound = FindItem()
    If blnFound = True Then
        With rsItem
        txtSku.Text = !Sku
        txtDescrip.Text = !descrip
        If Not IsNull(!Supplier) Then txtSupplier.Text = !Supplier
        If Not IsNull(!Category) Then txtCategory.Text = !Category
        If Not IsNull(!catcode) Then lblCatcode.Caption = !catcode
        txtPrice.Text = !price
        txtQty.Text = ListView3.SelectedItem.ListSubItems(3).Text
        txtRef.Text = ListView3.SelectedItem.ListSubItems(4).Text
        End With
    Else
    MsgBox ListView3.SelectedItem.Text
    End If
End Sub
