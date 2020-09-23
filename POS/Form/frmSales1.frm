VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Point of Sale"
   ClientHeight    =   8520
   ClientLeft      =   765
   ClientTop       =   705
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   8145
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3422
            MinWidth        =   3422
            Text            =   "Status :"
            TextSave        =   "Status :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17358
            MinWidth        =   17358
            Text            =   "F1 - Cancel          F2 - Void          F3 - Return          F4 - Item Look-up"
            TextSave        =   "F1 - Cancel          F2 - Void          F3 - Return          F4 - Item Look-up"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6600
      ScaleHeight     =   480
      ScaleWidth      =   4200
      TabIndex        =   7
      Top             =   7440
      Width           =   4200
      Begin MSForms.CommandButton Command2 
         Height          =   435
         Left            =   1245
         TabIndex        =   8
         Top             =   15
         Width           =   1365
         Caption         =   " Cancel"
         PicturePosition =   327683
         Size            =   "2408;767"
         Picture         =   "frmSales.frx":000C
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   0
      MousePointer    =   12  'No Drop
      ScaleHeight     =   8055
      ScaleWidth      =   12015
      TabIndex        =   2
      Top             =   0
      Width           =   12015
      Begin VB.Frame Frame1 
         Caption         =   "Status"
         Height          =   2775
         Left            =   240
         TabIndex        =   1
         Top             =   5160
         Width           =   4455
         Begin VB.TextBox Text3 
            Height          =   975
            Left            =   720
            TabIndex        =   14
            Text            =   "Text3"
            Top             =   600
            Width           =   2895
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   4095
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   7223
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4920
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   7680
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   7140
         TabIndex        =   6
         Top             =   6120
         Width           =   1710
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDate 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblQty 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   4680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Amount Paid:"
         Height          =   195
         Left            =   5790
         TabIndex        =   5
         Top             =   6120
         Width           =   1155
      End
      Begin VB.Line Line2 
         X1              =   7170
         X2              =   8820
         Y1              =   5565
         Y2              =   5565
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "xxxx.xxx MRf"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7305
         TabIndex        =   4
         Top             =   5670
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total For This Invoice:"
         Height          =   195
         Left            =   5295
         TabIndex        =   3
         Top             =   5670
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lngIndex As Long

Private Sub CcmdOk_Click()
   Dim strTemp
    
    strTemp = "'" & CStr(txtCusCode.Text) & "'"
    On Error Resume Next
    rsCus.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsCus.Find "Cuscode = " & strTemp, 0, adSearchForward
    
    If CStr(rsCus!Cuscode) = CStr(txtCusCode.Text) Then
        txtName.Text = rsCus!FullName
        Picture2.Enabled = True
        Grid.SetFocus
        Grid.Col = 0: Grid.Row = 1
        Grid_EnterCell
        If Not CDate(rsAdmin!Date) = Format(Now, "mm/dd/yyyy") Then
            rsAdmin.Update "Date", Format(Now, "mm/dd/yyyy")
            rsAdmin.Update "Dr", "1"
            txtDr.Text = Format(Now, "mmddyy" & rsAdmin!Dr)
        Else
            txtDr.Text = Format(Now, "mmddyy" & rsAdmin!Dr)
            'rsAdmin.Update "Dr", rsAdmin!Dr + 1
        End If
    End If
    On Error GoTo 0
    Err.Clear
    Exit Sub
    
ErrorNotOnFile:
    MsgBox "Customer Code not found!", vbCritical, "Invalid Code"
    txtCusCode.SetFocus
    SendKeys "{home}+{end}"
     'not found
    'DoEvents
    'On Error GoTo 0
    'Err.Clear
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Grid.Text = ""
If Not Grid.Rows = 2 Then
    Grid.RemoveItem (Grid.Row)
    Grid.Refresh
'Fancy
'Command2.SetFocus
End If
End Sub

Private Sub Form_Activate()
    '// set focus on the combo, before that initilize grid too(interface bug fix)
    Grid.Col = 0: Grid.Row = 1
    'Grid_EnterCell
    DoEvents
    'Text3.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Grid.SetFocus
        Grid.Col = 0: Grid.Row = 1
        Grid_EnterCell
    End If
End Sub

Private Sub Form_Load()
    '// initilize form and do the setup
    Grid.Cols = 6
    Grid.Rows = 200
    Grid.Row = 0
    Grid.Col = 0: Grid.Text = "Sku"
    Grid.Col = 1: Grid.Text = "Description"
    Grid.Col = 2: Grid.Text = "Price"
    Grid.Col = 3: Grid.Text = "Quantity"
    Grid.Col = 4: Grid.Text = "Total"
    Grid.Col = 5: Grid.Text = "S"
    Grid.ColWidth(0) = 1700
    Grid.ColWidth(1) = 3900
    Grid.ColWidth(2) = 1200
    Grid.ColWidth(3) = 1000
    Grid.ColWidth(4) = 1200
    Grid.ColWidth(5) = 300
    Text2.Text = Empty
    Text2.Visible = False
    Grid.Rows = 2
    Label8.Caption = "0.00"

'    DTPicker2.Value = Now + 14 '// by default now we give 14 days for credit
    OpenData
    '// fill teh combo with customers names
    '-
'    Set Combo1.RowSource = CusRS
'    Combo1.ListField = "Name"
'    Combo1.DataField = "Name"
'    Set Combo1.DataSource = CusRS
    '-
    '// generate the next invoice number (last inv # + 1 is the trick)
    Dim newInvNo As Integer
   ' Set SetHRS = New ADODB.Recordset
'    SetHRS.Open "SELECT InvNo FROM Settings", db, adOpenStatic, adLockOptimistic
'    If Not SetHRS.EOF Then
'        newInvNo = SetHRS!InvNo
'    Else
'        newInvNo = 1
'    End If
'    SetHRS.Close
    
    'text3.text = newInvNo '// display the new inv #
    
End Sub


Private Sub Grid_Click()
' GRID.RemoveItem(GRID.Index )
End Sub

Public Sub Grid_EnterCell()
    '// when click on cell
    Select Case Grid.Col
        Case 0
            With Text2
                .Move Grid.CellLeft + Grid.Left, _
                Grid.CellTop + Grid.Top, Grid.CellWidth - 25, Grid.CellHeight - 25
                .Text = Grid.Text
                If Len(.Text) > 0 Then
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End If
                .Visible = True
                .ZOrder 0
                'If Grid.Row Mod 2 = 0 Then
                '    Text2.BackColor = RGB(174, 245, 214) '// lets make the grid color diff, every other grid
                'Else
                '    Text2.BackColor = RGB(255, 255, 255)
                'End If
                .SetFocus
                
            End With
            
    End Select
End Sub
Public Sub Grid_Qty()
    '// when click on cell
    Select Case Grid.Col
        Case 3
            With Text2
                .Move Grid.CellLeft + Grid.Left, _
                Grid.CellTop + Grid.Top, Grid.CellWidth - 25, Grid.CellHeight - 25
                .Text = Grid.Text
                If Len(.Text) > 0 Then
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End If
                .Visible = True
                .ZOrder 0
                'If Grid.Row Mod 2 = 0 Then
                '    Text2.BackColor = RGB(174, 245, 214) '// lets make the grid color diff, every other grid
                'Else
                '    Text2.BackColor = RGB(255, 255, 255)
                'End If
                .SetFocus
                
            End With
            
    End Select
End Sub

Private Sub Grid_GotFocus()
    'Grid_EnterCell
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Grid_EnterCell
    End If
End Sub

Private Sub mnuCancel_Click()
MsgBox "cancel"
End Sub

Private Sub mnuPostVoid_Click()
    MsgBox "postvoid"
End Sub

Private Sub mnuVoid_Click()
    MsgBox "void"
    
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim Qty As Long, Rate As Currency, Total As Currency
    Dim lr As Integer, lTotal As Double
Dim blnFound As Boolean
    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With Text2
                    .Text = Empty
                    .Visible = False
                End With
                Grid.SetFocus
            Case vbKeyLeft
                '// move left
                'If Grid.Col = 0 Or Grid.Col = 1 Or Grid.Col = 2 And Text2.SelLength > 0 Then
                If Grid.Col = 0 Or Grid.Col = 3 And Text2.SelLength > 0 Then
                    With Text2
                        If Not .Text = Empty Then
                            Grid.Text = .Text
                        End If
                        .Visible = False
                        .Text = Empty
                    End With
                    If Grid.Col = 3 Then
                        Grid.Col = 0
                   ' ElseIf Grid.Col = 1 Then
                    '    Grid.Col = 0
                    Else
                        Grid.Col = 0
                    End If
                    Grid_EnterCell
                End If
            Case vbKeyRight
                '// move right
                If Grid.Col = 0 Or Grid.Col = 3 And Text2.SelLength > 0 Then
'                If Grid.Col = 0 Or Grid.Col = 1 Or Grid.Col = 2 And Text2.SelLength > 0 Then
 
                    With Text2
                        If Not .Text = Empty Then
                            Grid.Text = .Text
                        End If
                        .Visible = False
                        .Text = Empty
                    End With
                    If Grid.Col = 0 Then
                        Grid.Col = 3
                    'ElseIf Grid.Col = 1 Then
                    '    Grid.Col = 2
                    'Else
                    '    Grid.Col = 1
                    End If
                    Grid_EnterCell
                End If
            Case vbKeyDown
                '// move down until last row, if last move to first
                With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                If Not Grid.Row = Grid.Rows - 1 Then
                    Grid.Row = Grid.Row + 1
                    Grid_EnterCell
                Else
                    Grid.Row = 1
                    Grid_EnterCell
                End If
            Case vbKeyUp
                '// move up until first row -1, if first then move last
                With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                If Not Grid.Row = 1 Then
                    Grid.Row = Grid.Row - 1
                    Grid_EnterCell
                Else
                    Grid.Row = Grid.Rows - 1
                    Grid_EnterCell
                End If
    
            Case vbKeyReturn
                '// when enter is pressed, move to next col
                With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
               ' Grid.Col = 0
                Select Case Grid.Col
     '               Case 0
     '                   If Not lblQty.Visible = True Then
     '                       If Not Grid.Row = Grid.Rows - 1 Then
     '                           If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
     '                               Grid.Row = Grid.Row + 1
     '                           End If
     '                           Grid.Col = 0
     '                           Grid_EnterCell
     '                       Else
     '                       '// we need to add a new row ey, baby
     '                           If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
     '                               Grid.Rows = Grid.Rows + 1
     '                               Grid.Row = Grid.Row + 1
     '                               'Fancy
     '                           End If
     '                           Grid.Col = 0
     '                           Grid_EnterCell
     '                       End If
     '                   Else
     '                       Grid.Col = 3
     '                       Grid_Qty
     '                   End If
                  '      Command3_Click
                  
                    ' If Grid.Col = 0 Then
      '                  strSearch = CStr(Grid.Text)
      '                  blnFound = FindItem()
      '                  If blnFound = True Then
                           
      '                      Grid.Col = 1
      '                      Grid.Text = rsItem!Description
      '                      Grid.Col = 2
      '                      Grid.Text = rsItem!Price
      '                      Grid_EnterCell
      '                      Grid.Col = 3
      '                      Grid_EnterCell
                            
      '                  Else
      '                      MsgBox "Itemcode did not found!", vbCritical, "Invalid Itemcode"
      '                      Grid.Col = 0
      '                      Grid_EnterCell
      '                      Text2.SetFocus
      '                      SendKeys "{home}+{end}"
      '                      Grid.Col = 0
      '                  End If
                      '  Grid.Col = 3
                       ' Grid_EnterCell
                    
                      Case 3
                            If lblQty.Visible = True Then
                        'hmmm! this is tricky , but cool (naa! not at all)
                        If Len(Trim(Grid.Text)) = 0 Then
                            Grid.Text = "1"
                        End If
                        Grid.Col = 3
                        Qty = CLng(Grid.Text)
                        Grid.Col = 2
                        Rate = CCur(Grid.Text)
                        Total = Qty * Rate:
                                              Grid.Col = 4
                        Grid.Text = Format(Total, "###,###,##0.00")
                       DoTotals
                        Grid_EnterCell
                        If Not Grid.Row = Grid.Rows - 1 Then
                            If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                                Grid.Row = Grid.Row + 1
                           End If
                            Grid.Col = 0
                            Grid_EnterCell
                            lblQty.Visible = False
                        Else
                            '// we need to add a new row ey, baby
                           If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                                Grid.Rows = Grid.Rows + 1
                                Grid.Row = Grid.Row + 1
                                'Fancy
                            End If
                            Grid.Col = 0
                            Grid_EnterCell
                            lblQty.Visible = False
                        End If
                        End If
                        
               End Select
       '     Case vbKeyHome
       '         If Not Grid.Col = 0 And Text2.SelLength > 0 Then
       '             With Text2
       '                 If Not .Text = Empty Then
       '                     Grid.Text = .Text
       '                 End If
       '                 .Visible = False
       '                 .Text = Empty
       '             End With
       '             Grid.Col = 0
       '             Grid_EnterCell
       '         End If
       '     Case vbKeyEnd
       '         If Not Grid.Col = 2 And Text2.SelLength > 0 Then
       '             With Text2
       '                 If Not .Text = Empty Then
       '                     Grid.Text = .Text
        '                End If
       '                 .Visible = False
       '                 .Text = Empty
       '             End With
       '             Grid.Col = 2
       '             Grid_EnterCell
       '         End If
        Case vbKeyAdd
            lblQty.Visible = True
        Case vbKeyF4
            frmItemLook.Show 1
            
    End Select
End Sub
Public Sub Fancy()
    '// since this is the last row as we know
    '// so lets add one more(van mor)
    Dim CurrentCell As Integer
    With Grid
        If .Row Mod 2 = 0 Then
            '// trying to make this row diff col
            CurrentCell = .Col
            Dim r As Integer
            For r = 0 To 4
                .Col = r
                .CellBackColor = RGB(174, 245, 214)
            Next
            .Col = CurrentCell
        End If
    End With
End Sub
Private Sub DoTotals()
    '// get the total from all
    Dim CurrentCell As Integer
    Dim CurrentRow As Integer
    
    CurrentCell = Grid.Col
    CurrentRow = Grid.Row
    
    lTotal = 0
    Grid.Col = 4
    For r = 1 To Grid.Rows - 1
        Grid.Row = r
        If Not Grid.Text = Empty Then
            lTotal = lTotal + CDbl(Grid.Text)
        End If
    Next
    Label8.Caption = Format(lTotal, "###,###,##0.00")
    
    DoEvents
    
    Grid.Col = CurrentCell
    Grid.Row = CurrentRow
End Sub
Private Function FindCustNo(cName As String)
    '// find the cust no for a given cust name
    Set tRS = New ADODB.Recordset
    tRS.Open "Select * FROM mstCust WHERE Name='" & cName & "'", db, adOpenStatic, adLockOptimistic
    If tRS.RecordCount > 0 Then
        FindCustNo = tRS!CNum
    Else
        FindCustNo = Empty
    End If
    tRS.Close
End Function
Private Sub WriteHadder()
    '// write hadder data to db
    Set invHRS = New ADODB.Recordset
    invHRS.Open "SELECT * From InvHeadder", db, adOpenStatic, adLockOptimistic
    With invHRS
        .AddNew
        !CNum = FindCustNo(Combo1.Text)
        !InvNo = Text3.Text
        !SalDate = DTPicker1.Value
        !DueDate = DTPicker2.Value
        !Total = CDbl(Label8.Caption)
        !Paid = IIf(Val(Text1.Text) > 0, CDbl(Text1.Text), 0)
        !Settled = IIf(CDbl(Label8.Caption) = CDbl(Text1.Text), True, False)
        .Update
    End With
    Set SetHRS = New ADODB.Recordset
    SetHRS.Open "SELECT * FROM Settings", db, adOpenStatic, adLockOptimistic
    If SetHRS.EOF Or SetHRS.BOF Then SetHRS.AddNew
    SetHRS!InvNo = Val(Text3.Text) + 1
    SetHRS.Update
    SetHRS.Close
End Sub
Private Sub WriteDetails()
    '// update inv details from grid
    Dim r As Integer
    Set invDRS = New ADODB.Recordset
    invDRS.Open "SELECT * From InvDetails", db, adOpenStatic, adLockOptimistic
    
    For r = 1 To Grid.Rows - 1
        Grid.Row = r
        With invDRS
            .AddNew
            !InvNo = Text3.Text
            Grid.Col = 0: !Qty = IIf(Not Grid.Text = Empty, Val(Grid.Text), 1)
            Grid.Col = 1: !Desc = IIf(Grid.Text = Empty, "Misc.", Grid.Text)
            Grid.Col = 2
            If Not Trim(Grid.Text) = Empty Then
                !Rate = CDbl(Grid.Text)
            Else
                !Rate = 0
            End If
            Grid.Col = 3
            If Val(Grid.Text) > 0 Then
                !Total = CDbl(Grid.Text)
                .Update                 '// update only if total is > 0
            Else
                .Cancel
            End If
        End With
    Next
    
End Sub
Private Sub Command1_Click()
    '// check if vaild then alaka zoom! write the data to db
    Dim okayS As Boolean
    okayS = CheckValidInv
    If okayS = False Then
        MsgBox "Oops! Invoice Number Already Taken!", vbInformation
        Text3.SetFocus
        Text3.Text = Empty
        Exit Sub
    End If
    If Val(Text1.Text) < 1 Then Text1.Text = "0"
    If Val(Label8.Caption) > 0 And Val(Text3.Text) > 0 Then
        WriteHadder
        WriteDetails
        Unload Me
        Exit Sub
    End If
    MsgBox "Oops! Data Missing or Invalid", vbCritical
End Sub


Function CheckValidInv() As Boolean
     '   Set tRS = New ADODB.Recordset
     '   tRS.Open "SELECT * FROM InvHeadder WHERE InvNo ='" & Text3.Text & "'", db, adOpenStatic, adLockOptimistic
      '  If tRS.RecordCount > 0 Then
            CheckValidInv = False
      '  Else
            CheckValidInv = True
      '  End If
End Function
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


Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Qty As Long, Rate As Currency, Total As Currency
    Dim lr As Integer, lTotal As Double
    
    Dim blnFound As Boolean
    
    'check for numeric
    Const conZero As Integer = 48, conNine As Integer = 57
    Const conBackSpace As Integer = 8
    If (KeyAscii < conZero Or KeyAscii > conNine) And KeyAscii <> conBackSpace Then
        KeyAscii = 0
    End If
    
    If Len(Trim(Text2.Text)) = 12 Then
        KeyAscii = 0
        With Text2
            If Not .Text = Empty Then
                Grid.Text = .Text
               
            End If
                .Visible = False
                .Text = Empty
        End With
        strSearch = CStr(Grid.Text)
        blnFound = FindItem()
        If blnFound = True Then
            Grid.Col = 1
            Grid.Text = rsItem!Descrip
            Grid.Col = 2
            Grid.Text = rsItem!Price
            
            If lblQty.Visible = True Then
                Grid.Col = 3
                Grid_Qty
            Else
                Grid.Col = 3
                Grid.Text = 1
                Grid_EnterCell
              
                Grid.Col = 3
                Qty = CLng(Grid.Text)
                Grid.Col = 2
                Rate = CCur(Grid.Text)
                Total = Qty * Rate:
                Grid.Col = 4
                Grid.Text = Format(Total, "###,###,##0.00")
                    If Not Grid.Row = Grid.Rows - 1 Then
                        If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                            Grid.Row = Grid.Row + 1
                        End If
                        Grid.Col = 0
                        Grid.Text = ""
                        Text2.Text = ""
                        Grid_EnterCell
                        Grid.Text = ""
                        Text2.Text = ""
                    Else
                    '// we need to add a new row ey, baby
                        If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                            Grid.Rows = Grid.Rows + 1
                            Grid.Row = Grid.Row + 1
                            'Fancy
                        End If
                        Grid.Col = 0
                        Grid.Text = ""
                        Text2.Text = ""
                        Grid_EnterCell
                        Grid.Text = ""
                        Text2.Text = ""
                    End If
            End If
        Else
            MsgBox "Sku did not found!", vbCritical, "Invalid Sku"
            
            Grid.Col = 0
            Grid_EnterCell
            Text2.Text = Empty
            Text2.SetFocus
            SendKeys "{home}+{end}"
            Grid.Col = 0
        End If
    End If
    
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Grid.SetFocus
        Grid.Col = 0: Grid.Row = 1
        Grid_EnterCell
    End If
End Sub

Private Sub Timer1_Timer()
    lblDate.Caption = Format(Now, "mm-dd-yy")
    lblTime.Caption = Format(Now, "hh:mm:ss")
End Sub
