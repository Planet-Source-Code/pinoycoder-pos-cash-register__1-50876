VERSION 5.00
Begin VB.Form frmItemLook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ITEM LOOK-UP"
   ClientHeight    =   4665
   ClientLeft      =   3615
   ClientTop       =   2265
   ClientWidth     =   4695
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Enter Sku Number"
      ForeColor       =   &H00C00000&
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtSubcat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtOnhand 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtSupplier 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtDescrip 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtSku 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "SUBCATEGORY :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "CATEGORY :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "ON HAND :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "SUPPLIER :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "PRESS <ESCAPE> TO EXIT"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "PRESS <ENTER>TO INCLUDE IN TRANSACTION"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "PRICE :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "DESCRIPTION :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SKU  :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmItemLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnFound As Boolean

Private Sub Form_Deactivate()
    'frmSales.Text2.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
 '  frmSales.Grid.Col = 0
 '  frmSales.Grid_EnterCell
End Sub

Private Sub txtSku_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Qty As Long, Rate As Currency, Total As Currency

    
    Select Case KeyCode
        Case vbKeyEscape
            frmSales.RowNo
            Unload Me
            frmSales.RowNo
            blnFound = False
        Case vbKeyReturn
            If blnFound = True Then
                Dim lngSpace As Long
                With frmSales
                If CLng(txtOnhand.Text) < 1 Then
                    MsgBox "Not enough stock on this item", vbCritical, "Invalid Stock"
                    blnFound = False
                    Unload Me
                    Exit Sub
                End If
                Open App.Path & "\temp.txt" For Output As #1
                Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5
                lngSpace = (38 - Len(txtDescrip.Text) - Len(Format(txtPrice.Text, "###,##0.00")) - 1)
                
                    Print #1, txtDescrip.Text & Space(lngSpace) & Format(txtPrice.Text, "###,##0.00")
                    Print #5, txtDescrip.Text & Space(lngSpace) & Format(txtPrice.Text, "###,##0.00")
                    
                      '  Print #1, Label1.Caption & Space(8) & Label4.Caption & " @ " & Label3.Caption
                       ' Print #5, Label1.Caption & Space(8) & Label4.Caption & " @ " & Label3.Caption
                    .Grid.Col = 0
                    .Grid.Text = txtSku.Text
                    .Grid.Col = 1
                    .Grid.Text = txtDescrip.Text
                    .Grid.Col = 2
                    .Grid.Text = txtPrice.Text
                    
                    If .lblQty.Visible = True Then
                        blnItemlook = True
                        blnLoad = True
                        Unload Me
                        .Grid.Col = 3
                        .Grid_Qty
                        .RowNo
                        blnFound = False
                        Close #1
                        Close #5
                    Else
                        Print #1, txtSku.Text
                        Print #5, txtSku.Text
                        .Grid.Col = 3
                        .Grid.Text = 1
                      '  .Grid_EnterCell
                        
                        .Grid.Col = 3
                         Qty = CLng(.Grid.Text)
                        .Grid.Col = 2
                        Rate = CCur(.Grid.Text)
                        Total = Qty * Rate:
                        .Grid.Col = 4
                        .Grid.Text = Format(Total, "###,###,##0.00")
                        .DoTotals
                        .DoItems ' Compute Items
                        .Grid_EnterCell
                       .RowNo
                        If Not .Grid.Row = .Grid.Rows - 1 Then
                            If Not .Grid.Text = Empty And CDbl(.Grid.Text) > 0 Then
                                .Grid.Row = .Grid.Row + 1
                            End If
                            
                            Unload Me
                            .Grid.Col = 0
                            .Grid_EnterCell
                            .RowNo
                            .Text2.SetFocus
                            blnLoad = True
                            blnFound = False
                        Else
                        '// we need to add a new row ey, baby
                            If Not .Grid.Text = Empty And CDbl(.Grid.Text) > 0 Then
                                .Grid.Rows = .Grid.Rows + 1
                                .Grid.Row = .Grid.Row + 1
                                '.Fancy
                            End If
                            
                            Unload Me
                            .Grid.Col = 0
                            .Grid_EnterCell
                            .RowNo
                            .Text2.SetFocus
                            blnLoad = True
                            blnFound = False
                        End If
                    Close #1
                    Close #5
                    .DoTotals
                    .Addlist
                    RunBat
                    End If
                                        
                  
                End With
                
            Else
                If txtSku.Text = "" Then Exit Sub
                strSearch = CStr(txtSku.Text)
                blnFound = FindItem()
                If blnFound = True Then
                   Itemload
                Else
                    MsgBox "Sku not found!", vbCritical, "Invalid Sku"
                    txtSku.SetFocus
                    SendKeys "{home}+{end}"
                blnFound = False
                End If
            End If
            Close #1
            Close #5
    End Select
    
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
    
    strSearch = CStr(txtSku.Text)
    blnFound = FindItem()
    If blnFound = True Then
        Itemload
    Else
        MsgBox "Sku not found!", vbCritical, "Invalid Sku"
        txtSku.SetFocus
    End If
End Sub

Public Sub Itemload()
    txtDescrip.Text = rsItem!descrip
    txtPrice.Text = rsItem!price
    txtCategory = rsItem!Category
    If Not IsNull(rsItem!SubCategory) Then txtSubcat.Text = rsItem!SubCategory
    txtSupplier.Text = rsItem!Supplier
    txtOnhand.Text = rsItem!stack
    Label4.Visible = True
    txtSku.Locked = True
    txtSku.SetFocus
End Sub
