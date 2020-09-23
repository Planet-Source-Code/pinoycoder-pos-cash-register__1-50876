VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9225
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   9255
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   9255
         TabIndex        =   23
         Top             =   4680
         Width           =   9255
         Begin VB.CommandButton cmdClose 
            Caption         =   "Exit"
            Height          =   615
            Left            =   8040
            Picture         =   "frmReport.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   615
            Left            =   6360
            Picture         =   "frmReport.frx":03FC
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Printer"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   2895
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Preview"
            Height          =   255
            Left            =   480
            TabIndex        =   22
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   2895
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Summary"
            Height          =   375
            Left            =   360
            TabIndex        =   20
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Detail"
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   3480
         TabIndex        =   11
         Top             =   960
         Width           =   5535
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Selected Date"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "From Start to Present"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   840
            TabIndex        =   14
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22609921
            CurrentDate     =   37497
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   2880
            TabIndex        =   15
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22609921
            CurrentDate     =   37497
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            Height          =   255
            Left            =   2400
            TabIndex        =   17
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sku/Supplier/Category"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   3480
         TabIndex        =   2
         Top             =   2760
         Width           =   5535
         Begin VB.TextBox txtSku2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            TabIndex        =   8
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtSku1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            TabIndex        =   7
            Top             =   1200
            Width           =   1935
         End
         Begin VB.OptionButton optSku 
            BackColor       =   &H00FFFFFF&
            Caption         =   "By Selected Sku"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton optAllSku 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Sku"
            Height          =   195
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optSupp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "By Supplier"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   1680
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox cboSupp 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1680
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            Height          =   255
            Left            =   3000
            TabIndex        =   10
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Image Image1 
         Height          =   810
         Left            =   1680
         Picture         =   "frmReport.frx":0706
         Top             =   0
         Width           =   7575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5550
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16211
            MinWidth        =   9596
            Text            =   "Status :"
            TextSave        =   "Status :"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    blnSaleReport = False
    blnDelReport = False
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim sqlStatement As String
    If blnSaleReport = True Then
        Me.MousePointer = 11
        If Option5.Value = True Then
            If Option1.Value = True Then
                If Not DEnv.rsSumSale.State = adStateClosed Then DEnv.rsSumSale.Close
                If optAllSku.Value = True Then
                    sqlStatement = "SELECT Sale.Sku, Sale.Descrip, Sale.Price, Sum(Sale.Qty) AS SumOfQty From Sale GROUP BY Sale.Sku, Sale.Descrip, Sale.Price;"
                    DEnv.rsSumSale.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RSumSale.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumSale.Sections("ReportHeader").Controls("label8").Caption = "FROM START TO END DATE ON ALL SKU'S"
                ElseIf optSku.Value = True Then
                    sqlStatement = "SELECT Item.Sku, Sale.Descrip, Sale.Price, Sum(Sale.Qty) AS SumOfQty FROM Item INNER JOIN Sale ON Item.Sku = Sale.Sku Where ((Item.Sku >= '" & txtSku1.Text & "') And (Item.Sku <= '" & txtSku2.Text & "' )) GROUP BY Item.Sku, Sale.Descrip, Sale.Price;"
                    DEnv.rsSumSale.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RSumSale.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumSale.Sections("ReportHeader").Controls("label8").Caption = "FROM START TO END DATE ON SELECTED SKU'S"
                End If
                
            ElseIf Option2.Value = True Then
                If optAllSku.Value = True Then
                    sqlStatement = "SELECT Sale.Sku, Sale.Descrip, Sale.Price, Sum(Sale.Qty) AS SumOfQty From Sale WHERE (Date >= #" & SQLDate(DTPicker1) & "#) And (Date <= #" & SQLDate(DTPicker2) & "#) GROUP BY Sale.Sku, Sale.Descrip, Sale.Price;"
                    DEnv.rsSumSale.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RSumSale.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumSale.Sections("ReportHeader").Controls("label8").Caption = "FROM " & SQLDate(DTPicker1) & " TO " & SQLDate(DTPicker2) & " ON ALL SKU'S"
                ElseIf optSku.Value = True Then
                    sqlStatement = "SELECT Item.Sku, Sale.Descrip, Sale.Price, Sum(Sale.Qty) AS SumOfQty FROM Item INNER JOIN Sale ON Item.Sku = Sale.Sku Where ((Item.Sku >= '" & txtSku1.Text & "') And (Item.Sku <= '" & txtSku2.Text & "' )) And (Date >= #" & SQLDate(DTPicker1) & "#) And (Date <= #" & SQLDate(DTPicker2) & "#) GROUP BY Item.Sku, Sale.Descrip, Sale.Price;"
                    DEnv.rsSumSale.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                DoEvents
                RSumSale.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                RSumSale.Sections("ReportHeader").Controls("label8").Caption = "FROM " & SQLDate(DTPicker1) & " TO " & SQLDate(DTPicker2) & " ON SELECTED SKU'S"
               End If
            End If
            
                RSumSale.Show 1
                DEnv.rsSumSale.Close
        Else
            If Option1.Value = True Then
                If Not DEnv.rsDetailSale.State = adStateClosed Then DEnv.rsDetailSale.Close
                If optAllSku.Value = True Then
                    sqlStatement = "SELECT [Sale].[Sku], [Supplier].[SUPPLIER], [Category].[Category], [Sale].[Descrip], [Sale].[Price], [Sale].[Date], [Sale].[Qty] FROM Category INNER JOIN (Supplier INNER JOIN Sale ON [Supplier].[SupCode]=[Sale].[Supcode]) ON [Category].[CatCode]=[Sale].[Catcode];"

                    DEnv.rsDetailSale.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                     RDetSale.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    'RSumSale.Sections("ReportHeader").Controls("label8").Caption = "FROM START TO END DATE ON ALL SKU'S"
                ElseIf optSku.Value = True Then
                    sqlStatement = "SELECT Sale.Sku, Supplier.SUPPLIER, Category.Category, Sale.Descrip, Sale.Price, Sale.Date, Sale.Qty FROM Category INNER JOIN (Supplier INNER JOIN Sale ON Supplier.SupCode = Sale.Supcode) ON Category.CatCode = Sale.Catcode WHERE (((Sale.Sku)>= '" & txtSku1.Text & "') And ((Sale.Sku)<= '" & txtSku2.Text & "'));"
'                    sqlStatement = "SELECT Item.Sku, Sale.Descrip, Sale.Price, Sum(Sale.Qty) AS SumOfQty FROM Item INNER JOIN Sale ON Item.Sku = Sale.Sku Where ((Item.Sku >= '" & txtSku1.Text & "') And (Item.Sku <= '" & txtSku2.Text & "' )) GROUP BY Item.Sku, Sale.Descrip, Sale.Price;"
                    DEnv.rsDetailSale.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RDetSale.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                   ' RSumSale.Sections("ReportHeader").Controls("label8").Caption = "FROM START TO END DATE ON SELECTED SKU'S"
                End If
                
            ElseIf Option2.Value = True Then
                If optAllSku.Value = True Then
                    sqlStatement = "SELECT Sale.Sku, Supplier.SUPPLIER, Category.Category, Sale.Descrip, Sale.Price, Sale.Date, Sale.Qty FROM Category INNER JOIN (Supplier INNER JOIN Sale ON Supplier.SupCode = Sale.Supcode) ON Category.CatCode = Sale.Catcode WHERE (Date >= #" & SQLDate(DTPicker1) & "#) And (Date <= #" & SQLDate(DTPicker2) & "#);"

                    DEnv.rsDetailSale.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RDetSale.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumSale.Sections("ReportHeader").Controls("label8").Caption = "FROM " & SQLDate(DTPicker1) & " TO " & SQLDate(DTPicker2) & " ON ALL SKU'S"
                ElseIf optSku.Value = True Then
                    sqlStatement = "SELECT Sale.Sku, Supplier.SUPPLIER, Category.Category, Sale.Descrip, Sale.Price, Sale.Date, Sale.Qty FROM Category INNER JOIN (Supplier INNER JOIN Sale ON Supplier.SupCode = Sale.Supcode) ON Category.CatCode = Sale.Catcode WHERE (Date >= #" & SQLDate(DTPicker1) & "#) And (Date <= #" & SQLDate(DTPicker2) & "#) And (Sku >= '" & txtSku1.Text & "') And (Sku <= '" & txtSku2.Text & "');"
                    DEnv.rsDetailSale.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RDetSale.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumSale.Sections("ReportHeader").Controls("label8").Caption = "FROM " & SQLDate(DTPicker1) & " TO " & SQLDate(DTPicker2) & " ON SELECTED SKU'S"
                End If
                
            End If
          
                RDetSale.Show 1
                DEnv.rsDetailSale.Close
        End If
    ElseIf blnDelReport = True Then
         Me.MousePointer = 11
        If Option5.Value = True Then
            If Option1.Value = True Then
                If Not DEnv.rsSumDelivery.State = adStateClosed Then DEnv.rsSumDelivery.Close
                If optAllSku.Value = True Then
                    sqlStatement = "SELECT Delivery.Sku, Delivery.Descrip, Delivery.Price, Sum(Delivery.Qty) AS SumOfQty From Delivery GROUP BY Delivery.Sku, Delivery.Descrip, Delivery.Price;"

                    DEnv.rsSumDelivery.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RSumDel.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumDel.Sections("ReportHeader").Controls("label8").Caption = "FROM START TO END DATE ON ALL SKU'S"
                ElseIf optSku.Value = True Then
                    sqlStatement = "SELECT Delivery.Sku, Delivery.Descrip, Delivery.Price, Sum(Delivery.Qty) AS SumOfQty From Delivery Where (((Delivery.Sku) >= '" & txtSku1.Text & "') And ((Delivery.Sku) <= '" & txtSku2.Text & "') ) GROUP BY Delivery.Sku, Delivery.Descrip, Delivery.Price;"
                    DEnv.rsSumDelivery.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RSumDel.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumDel.Sections("ReportHeader").Controls("label8").Caption = "FROM START TO END DATE ON SELECTED SKU'S"
                End If
                
            ElseIf Option2.Value = True Then
                If optAllSku.Value = True Then
                    sqlStatement = "SELECT Delivery.Sku, Delivery.Descrip, Delivery.Price, Sum(Delivery.Qty) AS SumOfQty From Delivery Where (((Delivery.DateEntry) >= #" & SQLDate(DTPicker1) & "#) And ((Delivery.DateEntry) <= #" & SQLDate(DTPicker2) & "#)) GROUP BY Delivery.Sku, Delivery.Descrip, Delivery.Price;"

                    DEnv.rsSumDelivery.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RSumDel.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumDel.Sections("ReportHeader").Controls("label8").Caption = "FROM " & SQLDate(DTPicker1) & " TO " & SQLDate(DTPicker2) & " ON ALL SKU'S"
                ElseIf optSku.Value = True Then
                    sqlStatement = "SELECT Delivery.Sku, Delivery.Descrip, Delivery.Price, Sum(Delivery.Qty) AS SumOfQty From Delivery Where (((Delivery.DateEntry) >= #" & SQLDate(DTPicker1) & "#) And ((Delivery.DateEntry) <= #" & SQLDate(DTPicker2) & "#) And ((Delivery.Sku) >= '" & txtSku1.Text & "') And ((Delivery.Sku) <= '" & txtSku2.Text & "')) GROUP BY Delivery.Sku, Delivery.Descrip, Delivery.Price;"

                    DEnv.rsSumDelivery.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RSumDel.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RSumDel.Sections("ReportHeader").Controls("label8").Caption = "FROM " & SQLDate(DTPicker1) & " TO " & SQLDate(DTPicker2) & " ON SELECTED SKU'S"
               End If
            End If
           
                RSumDel.Show 1
                DEnv.rsSumDelivery.Close
        Else
            If Option1.Value = True Then
                If Not DEnv.rsDetailDel.State = adStateClosed Then DEnv.rsDetailDel.Close
                If optAllSku.Value = True Then
                    sqlStatement = "SELECT Delivery.Sku, Supplier.SUPPLIER, Category.Category, Delivery.Descrip, Delivery.Price, Delivery.Qty, Delivery.DateEntry FROM Supplier INNER JOIN (Category INNER JOIN Delivery ON Category.CatCode = Delivery.Catcode) ON Supplier.SupCode = Delivery.Supcode;"
                    DEnv.rsDetailDel.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                     RDetDel.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                     RDetDel.Sections("ReportHeader").Controls("label9").Caption = "FROM START TO END DATE ON ALL SKU'S"
                ElseIf optSku.Value = True Then
                    sqlStatement = "SELECT Delivery.Sku, Supplier.SUPPLIER, Category.Category, Delivery.Descrip, Delivery.Price, Delivery.Qty, Delivery.DateEntry FROM Supplier INNER JOIN (Category INNER JOIN Delivery ON Category.CatCode = Delivery.Catcode) ON Supplier.SupCode = Delivery.Supcode WHERE (((Delivery.Sku)>= '" & txtSku1.Text & "') And ((Delivery.Sku)<= '" & txtSku2.Text & "'));"
                    DEnv.rsDetailDel.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RDetDel.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RDetDel.Sections("ReportHeader").Controls("label9").Caption = "FROM START TO END DATE ON SELECTED SKU'S"
                End If
                
            ElseIf Option2.Value = True Then
                If optAllSku.Value = True Then
                    sqlStatement = "SELECT Delivery.Sku, Supplier.SUPPLIER, Category.Category, Delivery.Descrip, Delivery.Price, Delivery.Qty, Delivery.DateEntry FROM Supplier INNER JOIN (Category INNER JOIN Delivery ON Category.CatCode = Delivery.Catcode) ON Supplier.SupCode = Delivery.Supcode WHERE (((Delivery.DateEntry) >= #" & SQLDate(DTPicker1) & "#) And ((Delivery.DateEntry) <= #" & SQLDate(DTPicker2) & "#));"

                    DEnv.rsDetailDel.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RDetDel.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RDetDel.Sections("ReportHeader").Controls("label9").Caption = "FROM " & SQLDate(DTPicker1) & " TO " & SQLDate(DTPicker2) & " ON ALL SKU'S"
                ElseIf optSku.Value = True Then
                    sqlStatement = "SELECT Delivery.Sku, Supplier.SUPPLIER, Category.Category, Delivery.Descrip, Delivery.Price, Delivery.Qty, Delivery.DateEntry FROM Supplier INNER JOIN (Category INNER JOIN Delivery ON Category.CatCode = Delivery.Catcode) ON Supplier.SupCode = Delivery.Supcode WHERE (((Delivery.DateEntry) >= #" & SQLDate(DTPicker1) & "#) And ((Delivery.DateEntry) <= #" & SQLDate(DTPicker2) & "#) And ((Delivery.Sku) >= '" & txtSku1.Text & "') And ((Delivery.Sku) <= '" & txtSku2.Text & "'));"
                    DEnv.rsDetailDel.Open sqlStatement, db, adOpenStatic, adLockOptimistic
                    DoEvents
                    RDetDel.Sections("ReportHeader").Controls("lblStore").Caption = rsStore!Name
                    RDetDel.Sections("ReportHeader").Controls("label9").Caption = "FROM " & SQLDate(DTPicker1) & " TO " & SQLDate(DTPicker2) & " ON SELECTED SKU'S"
                End If
                
            End If
            
                RDetDel.Show 1
                DEnv.rsDetailDel.Close
        End If
    End If
        
     Me.MousePointer = 0
                    
End Sub



Private Sub DTPicker1_Change()
    If DTPicker1.Value > DTPicker2.Value Then
        DTPicker2.Value = DTPicker1.Value
    End If
End Sub

Private Sub DTPicker2_Change()
    If DTPicker2.Value > DTPicker1.Value Then
        DTPicker1.Value = DTPicker2.Value
    End If
End Sub
Private Sub Form_Load()
    If blnSaleReport = True Then
    End If
    If rsSupp.RecordCount > 0 Then rsSupp.MoveFirst
    Do Until rsSupp.EOF
        If Not IsNull(rsSupp!Supplier) Then cboSupp.AddItem rsSupp!Supplier
        rsSupp.MoveNext
        DoEvents
    Loop
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    DTPicker1.Value = Now
    DTPicker2.Value = Now
End Sub

Private Sub optAllSku_Click()
    If optAllSku.Value = True Then
        txtSku1.Enabled = False
        txtSku2.Enabled = False
    End If
End Sub

Private Sub optSku_Click()
    If optSku.Value = True Then
        txtSku1.Enabled = True
        txtSku2.Enabled = True
        txtSku1.SetFocus
    End If
End Sub

Private Sub optSupp_Click()
    If optSupp.Value = True Then
        cboSupp.Enabled = True
        txtSku1.Enabled = False
        txtSku2.Enabled = False
    End If
End Sub

Private Sub txtSku1_LostFocus()
    txtSku2.Text = txtSku1.Text
End Sub
