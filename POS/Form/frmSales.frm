VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSales 
   BorderStyle     =   0  'None
   ClientHeight    =   8880
   ClientLeft      =   720
   ClientTop       =   375
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
   Picture         =   "frmSales.frx":000C
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "CASH TENDERED"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   4320
      TabIndex        =   43
      Top             =   5880
      Visible         =   0   'False
      Width           =   3255
      Begin MSComctlLib.ListView ListView1 
         Height          =   1800
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   3175
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TENDERED"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Width           =   1834
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5640
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   120
      ScaleHeight     =   5100
      ScaleWidth      =   7455
      TabIndex        =   37
      Top             =   600
      Width           =   7455
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   5040
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   5100
         Left            =   0
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   0
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8996
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
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
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
      Height          =   2700
      Left            =   7680
      Picture         =   "frmSales.frx":15F94E
      ScaleHeight     =   2700
      ScaleWidth      =   4800
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   4800
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtPass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "frmSales.frx":189C90
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "VOID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   75
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE INPUT MANAGER PASSWORD"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   0
         X2              =   4080
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line1 
         X1              =   4080
         X2              =   4080
         Y1              =   0
         Y2              =   2040
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER 1: AMOUNT"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame frmVoid 
      Caption         =   "VOID COMMAND"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   7680
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   4095
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   240
         ScaleHeight     =   1575
         ScaleWidth      =   2655
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
         Begin VB.Label Label6 
            Caption         =   "<F2 - VOID> TO VOID"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label9 
            Caption         =   "<UP> TO SCROLL UP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   25
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label10 
            Caption         =   "<DOWN> TO SCROLL DOWN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Label9 
            Caption         =   "<ESCAPE> TO EXIT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   1320
            Width           =   3615
         End
      End
   End
   Begin VB.Frame frmStatus 
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   4320
      TabIndex        =   18
      Top             =   5880
      Width           =   3255
      Begin VB.Label lblQty 
         Alignment       =   2  'Center
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblDiscount 
         Alignment       =   2  'Center
         Caption         =   "DISCOUNT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   4095
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   3855
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   3030
            TabIndex        =   17
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "CHANGE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   16
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   3030
            TabIndex        =   15
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "DUE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   14
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   3030
            TabIndex        =   13
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "TENDERED :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   3030
            TabIndex        =   11
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XXX,XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " ITEMS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTOTAL "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   3165
         TabIndex        =   3
         Top             =   1440
         Width           =   795
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      ItemData        =   "frmSales.frx":18A183
      Left            =   7680
      List            =   "frmSales.frx":18A185
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   4215
   End
   Begin MSComctlLib.StatusBar S 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   8640
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      SimpleText      =   "S"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22118
            MinWidth        =   22118
            Text            =   $"frmSales.frx":18A187
            TextSave        =   $"frmSales.frx":18A212
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Index           =   6
      Left            =   11160
      TabIndex        =   53
      Top             =   8280
      Width           =   855
   End
   Begin VB.Image Image14 
      Height          =   525
      Left            =   11160
      Picture         =   "frmSales.frx":18A29D
      Top             =   8100
      Width           =   750
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
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
      Index           =   5
      Left            =   2280
      TabIndex        =   52
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F5"
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
      Index           =   4
      Left            =   4440
      TabIndex        =   51
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F11"
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
      Index           =   3
      Left            =   7800
      TabIndex        =   50
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F9"
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
      Index           =   2
      Left            =   6120
      TabIndex        =   49
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F12"
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
      Index           =   1
      Left            =   9480
      TabIndex        =   48
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
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
      Left            =   3360
      TabIndex        =   47
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   46
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
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
      Index           =   0
      Left            =   120
      TabIndex        =   45
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Image Image13 
      Height          =   525
      Left            =   10560
      Picture         =   "frmSales.frx":18B7A7
      Top             =   8100
      Width           =   615
   End
   Begin VB.Image Image12 
      Height          =   540
      Left            =   9480
      Picture         =   "frmSales.frx":18C8DD
      Top             =   8100
      Width           =   1140
   End
   Begin VB.Image Image11 
      Height          =   525
      Left            =   8880
      Picture         =   "frmSales.frx":18E92F
      Top             =   8100
      Width           =   615
   End
   Begin VB.Image Image10 
      Height          =   540
      Left            =   7800
      Picture         =   "frmSales.frx":18FA65
      Top             =   8100
      Width           =   1140
   End
   Begin VB.Image Image9 
      Height          =   525
      Left            =   7200
      Picture         =   "frmSales.frx":191AB7
      Top             =   8100
      Width           =   615
   End
   Begin VB.Image Image7 
      Height          =   540
      Left            =   6120
      Picture         =   "frmSales.frx":192BED
      Top             =   8100
      Width           =   1140
   End
   Begin VB.Image Image8 
      Height          =   525
      Left            =   5520
      Picture         =   "frmSales.frx":194C3F
      Top             =   8100
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   540
      Left            =   4440
      Picture         =   "frmSales.frx":195D75
      Top             =   8100
      Width           =   1140
   End
   Begin VB.Image Image5 
      Height          =   540
      Left            =   3360
      Picture         =   "frmSales.frx":197DC7
      Top             =   8100
      Width           =   1140
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   2280
      Picture         =   "frmSales.frx":199E19
      Top             =   8100
      Width           =   1140
   End
   Begin VB.Image Image3 
      Height          =   540
      Left            =   1200
      Picture         =   "frmSales.frx":19BE6B
      Top             =   8100
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   120
      Picture         =   "frmSales.frx":19DEBD
      Top             =   8100
      Width           =   1140
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
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
      Left            =   480
      TabIndex        =   42
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label24 
      Caption         =   "Label24"
      Height          =   255
      Left            =   10560
      TabIndex        =   41
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label25 
      Caption         =   "Label25"
      Height          =   255
      Left            =   10560
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   10560
      TabIndex        =   36
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   10560
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   10560
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   10560
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngIndex As Long
Dim lngRow As Long
Dim lngCurrentRow As Long
Dim blnBayad As Boolean
Dim blnCash As Boolean
Dim blnEnd As Boolean
Dim blnPay As Boolean
Dim blnCredit As Boolean
Dim blnGift As Boolean
Dim blnCancel As Boolean
Dim blnVoid As Boolean
Dim blnDiscount As Boolean
Dim blnDiscounted As Boolean
Dim strDiscount As String
Dim strDiscounted As String
Dim blnDiscountAmount As Boolean
Dim blnAllDiscount As Boolean
Dim curItemPrice As Currency
Dim blnVoidTrue As Boolean
Dim blnPayment As Boolean
Dim blnReturn As Boolean

Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
      
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd _
        As Long, lpPoint As PointAPI)
    
Private Declare Function ShowCursor Lib "user32" (ByVal _
        bShow As Long) As Long
Private Type PointAPI
  X As Long
  Y As Long
End Type

    

Private Sub cmdOk_Click()
    Dim itm As ListItem
    Dim lngSpace As Long
    
    If blnPay = True Then
        If blnBayad = True Then
                PrintTotal
                blnBayad = False
                blnPayment = True
        End If
        txtPass.PasswordChar = ""
        Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5
        Open App.Path & "\temp.txt" For Output As #2
        Open App.Path & "\payment.txt" For Append As #1
        
        If blnCash = True Then
            
            If Len(Trim(txtPass)) = 0 Then
                Print #1, "CASH," & Format(Label30.Caption, "#####0.00")
                lngSpace = (38 - Len("CASH") - Len(Format(Label30.Caption, "###,##0.00")) - 1)
                Print #2, "CASH" & Space(lngSpace) & Format(Label30.Caption, "###,##0.00")
                Print #5, "CASH" & Space(lngSpace) & Format(Label30.Caption, "###,##0.00")
            ElseIf CCur(txtPass.Text) > CCur(Label30.Caption) Then
                Print #1, "CASH," & Format(txtPass.Text, "#####0.00")
                lngSpace = (38 - Len("CASH") - Len(Format(txtPass.Text, "###,##0.00")) - 1)
                Print #2, "CASH" & Space(lngSpace) & Format(txtPass.Text, "###,##0.00")
                Print #5, "CASH" & Space(lngSpace) & Format(txtPass.Text, "###,##0.00")
            Else
                Print #1, "CASH," & Format(txtPass.Text, "#####0.00")
                lngSpace = (38 - Len("CASH") - Len(Format(txtPass.Text, "###,##0.00")) - 1)
                Print #2, "CASH" & Space(lngSpace) & Format(txtPass.Text, "###,##0.00")
                Print #5, "CASH" & Space(lngSpace) & Format(txtPass.Text, "###,##0.00")
            End If
                
            Close #1
            Close #5
            Close #2
            Addlist
            RunBat
            blnCash = False
            Picture3.Visible = False
            OpenPay
            blnEnd = False
        End If
    ElseIf blnCancel = True Then
        Dim blnUser As Boolean
        txtPass.PasswordChar = "*"
        strSearch = CStr(txtPass.Text)
        blnUser = FindPass()
        If blnUser = True Then
            blnCancel = False
            TransCancel
        End If
    ElseIf blnVoid = True Then
        txtPass.PasswordChar = "*"
        strSearch = CStr(txtPass.Text)
        blnUser = FindPass()
        If blnUser = True Then
            If Mid(rsUser!UserTaskLevel, 1, 1) = 3 Then
                MsgBox "This password is not authorized to void item", vbCritical, "Void"
                txtPass.SetFocus
                SendKeys "{Home}+{end}"
                Exit Sub
            End If
            blnVoid = False
            blnVoidTrue = True
            Text2.Locked = True
            Grid.Row = Grid.Row - 1
            Grid.Col = 0
            Grid_EnterCell
            FancyTrue
            Picture3.Visible = False
            txtPass.Text = ""
            Picture1.Visible = True
        Else
            MsgBox "Password not found", vbCritical, "Invalid Password"
            txtPass.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    ElseIf blnDiscount = True Then
        txtPass.PasswordChar = "*"
        strSearch = CStr(txtPass.Text)
        blnUser = FindPass()
        If blnUser = True Then
            If Mid(rsUser!UserTaskLevel, 1, 1) = 3 Then
                MsgBox "This password is not authorized to discount", vbCritical, "Void"
                txtPass.SetFocus
                SendKeys "{Home}+{end}"
                Exit Sub
            End If
            txtPass.Text = ""
            Image2.Visible = False
            Label17.Visible = True
            Label11.Caption = "ENTER 2: PERCENT"
            blnDiscounted = True
            blnDiscount = False
            txtPass.PasswordChar = ""
        Else
            MsgBox "Password not found", vbCritical, "Invalid Password"
            txtPass.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
    ElseIf blnDiscounted = True Then
         If Len(Trim(txtPass.Text)) = 0 Then
            MsgBox "Please select between <1> and <2>!", vbCritical, "Discount"
            txtPass.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
         If Not CInt(txtPass.Text) >= 1 Or Not CInt(txtPass.Text) <= 2 Then
            MsgBox "Please select between <1> and <2>!", vbCritical, "Invalid Number"
            txtPass.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
       
        blnDiscounted = False
        strDiscount = txtPass.Text
        txtPass.Text = ""
        Label17.Visible = False
        blnDiscountAmount = True
        If strDiscount = "2" Then
            Label11.Caption = "ENTER PERCENT TO DISCOUNT"
        Else
            Label11.Caption = "ENTER AMOUNT TO DISCOUNT"
        End If
    ElseIf blnDiscountAmount = True And blnAllDiscount = False Then
       ' blnDiscountAmount = False
        strDiscounted = txtPass.Text
        Grid.Row = lngCurrentRow
        Grid.Col = 0
        Grid_EnterCell
        Picture3.Visible = False
        Image2.Visible = True
        lblDiscount.Visible = True
    ElseIf blnDiscountAmount = True And blnAllDiscount = True Then
        Dim curAmount1 As Currency
        If Len(Trim(txtPass.Text)) = 0 Then Exit Sub
        frmStatus.Visible = True
        lblDiscount.Visible = True
            strDiscounted = txtPass.Text
            If blnDiscountAmount = True Then
                Open App.Path & "\temp.txt" For Output As #1
                Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5
  
                        blnDiscountAmount = False
                        curItemPrice = CCur(Label12.Caption)
                            If strDiscount = "1" Then
                                curAmount1 = CCur(strDiscounted)
                                lngSpace = (38 - Len("DISCOUNT") - Len(Format(curAmount1, "###,##0.00")) - 2)
                                Print #1, "DISCOUNT" & Space(lngSpace) & "-" & Format(curAmount1, "###,##0.00")
                                Print #5, "DISCOUNT" & Space(lngSpace) & "-" & Format(curAmount1, "###,##0.00")
                                 
                                Print #1, Space(14) & Format(curItemPrice, "###,##0.00") & " - " & Format(curAmount1, "###,##0.00")
                                Print #5, Space(14) & Format(curItemPrice, "###,##0.00") & " - " & Format(curAmount1, "###,##0.00")
                                
                                
                            Else
                                curAmount1 = CCur(curItemPrice * (CCur(strDiscounted) / 100))
                                lngSpace = (38 - Len("DISCOUNT") - Len(Format(curAmount1, "###,##0.00")) - 2)
                                Print #1, "DISCOUNT" & Space(lngSpace) & "-" & Format(curAmount1, "###,##0.00")
                                Print #5, "DISCOUNT" & Space(lngSpace) & "-" & Format(curAmount1, "###,##0.00")
                                Print #1, Space(14) & Format(curItemPrice, "###,##0.00") & " - " & strDiscounted & "%"
                                Print #5, Space(14) & Format(curItemPrice, "###,##0.00") & " - " & strDiscounted & "%"
                              
                            End If
                                Grid.SetFocus
                                Grid.Row = lngCurrentRow
                                Grid.Col = 0
                                Grid_EnterCell
                                Text2.Text = "DISCOUNT"
                                Grid.Text = "DISCOUNT"
                                Grid.Col = 4
                                Grid.Text = "-" & Format(curAmount1, "###,##0.00")
                Close #1
                Close #5
                Addlist
                DoTotals
                blnPay = True
                blnCash = True
                
                If blnBayad = True Then
                    PrintTotal
                    blnBayad = False
                    blnPayment = True
                End If
                        
                        frmStatus.Visible = False
                        lblDiscount.Visible = False
                        Frame1.Visible = True
                Picture3.Visible = True
                Label26.Caption = "AMOUNT"
                Label11.Caption = "PLEASE ENTER AMOUNT"
                txtPass.SetFocus
                    End If
    End If
    txtPass.Text = ""
End Sub

Private Sub Form_Activate()
    '// set focus on the combo, before that initilize grid too(interface bug fix)
    If blnLoad = True Then
        blnLoad = False
        Grid.Col = 3
        DoEvents
    Else
        Grid.Col = 0
        Grid_EnterCell
        DoEvents
   End If
   ' lblTrans = rsStore!Number
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Grid.SetFocus
        Grid.Col = 0
        Grid_EnterCell
    End If
End Sub

Private Sub Form_Load()
'    Timer1.Enabled = True
    Picture3.Height = 140
    Picture3.Width = 275
    
    '// initilize form and do the setup
    
    Grid.Cols = 6
    Grid.Rows = 200
    Grid.Row = 0
    Grid.Col = 0: Grid.Text = "SKU"
    Grid.Col = 1: Grid.Text = "DESCRIPTION"
    Grid.Col = 2: Grid.Text = "PRICE"
    Grid.Col = 3: Grid.Text = "QTY"
    Grid.Col = 4: Grid.Text = "TOTAL"
    Grid.Col = 5: Grid.Text = ""
    Grid.ColWidth(0) = 1700
    Grid.ColWidth(1) = 2500
    Grid.ColWidth(2) = 1000
    Grid.ColWidth(3) = 500
    Grid.ColWidth(4) = 1000
    Grid.ColWidth(5) = 300
    Text2.Text = Empty
    Text2.Visible = False
    Grid.Rows = 2
    Label8.Caption = "0.00"
    Label15.Caption = "0.00"
   OpenData

'ShowCursor (0)

    
    If Not Dir(App.Path & "\temp.txt") = "" Then Kill App.Path & "\temp.txt"
    If Not Dir(App.Path & "\payment.txt") = "" Then Kill App.Path & "\payment.txt"
    AddTitlle ' print title
    'MsgBox "Please enter to start your transaction", vbInformation, "Point of Sale"
End Sub


Private Sub Form_Unload(Cancel As Integer)
ShowCursor (1)
End Sub

Public Sub Grid_EnterCell()
    '// when click on cell
    Select Case Grid.Col
        Case 0
            With Text2
              '  If Not blnVoid = True Then
                .Move Grid.CellLeft + Grid.Left, _
                Grid.CellTop + Grid.Top, Grid.CellWidth - 25, Grid.CellHeight - 25
                .Text = Grid.Text
                If Len(.Text) > 0 Then
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End If
                .Visible = True
                .ZOrder 0
                If blnVoidTrue = True Then
                     Text2.BackColor = RGB(174, 245, 214) '// lets make the grid color diff, every other grid
                Else
                    Text2.BackColor = RGB(255, 255, 255)
                End If
                .SetFocus
              '  End If
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
                If blnVoidTrue = True Then
                    Text2.BackColor = RGB(174, 245, 214) '// lets make the grid color diff, every other grid
                Else
                    Text2.BackColor = RGB(255, 255, 255)
                End If
                
                .SetFocus
                
            End With
            
    End Select
End Sub
Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Grid_EnterCell
    End If
    If KeyCode = vbKeyEscape Then
        Grid.SetFocus
        Grid.Col = 0: Grid.Row = 1
        Grid_EnterCell
    End If
    If KeyCode = vbKeyTab Then
        Grid.SetFocus
        Grid.Col = 0: Grid.Row = 1
        Grid_EnterCell
        
    End If
End Sub


Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKey1 'CASH
          
            If blnBayad = True Then
                PrintTotal
             
                blnBayad = False
            End If
            blnPay = True
            blnCash = True
            Picture3.Visible = True
            Label26.Caption = "AMOUNT"
            Label11.Caption = "PLEASE ENTER AMOUNT"
            txtPass.SetFocus
        Case vbKeyEscape
            
            
           
        Case vbKeyF1 'CANCEL TRANSACTION
            Frame1.Visible = False
            Picture3.Visible = True
            frmStatus.Visible = True
            txtPass.SetFocus
            lblQty.Visible = True
            lblQty.Caption = "CANCEL TRANSACTION"
            blnCancel = True
            txtPass.PasswordChar = "*"
            Label26.Caption = "CANCEL TRANSACTION"
            Label11.Caption = "PLEASE INPUT MANAGER PASSWORD"
        Case vbKeyF5 'DISCOUNT
            If Not blnBayad = True Then Exit Sub
            txtPass.Text = ""
            blnDiscount = True
            blnAllDiscount = True
            Picture3.Visible = True 'password
            Frame1.Visible = False
            frmStatus.Visible = True
            txtPass.SetFocus
            lblDiscount.Visible = True
            Label26.Caption = "DISCOUNT ITEM"
            Label11.Caption = "PLEASE INPUT MANAGER PASSWORD"
            txtPass.PasswordChar = "*"
            
            
            
    End Select
End Sub



Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim Qty As Long, Rate As Currency, Total As Currency
    Dim lr As Integer, lTotal As Double
    Dim blnFound As Boolean
    Dim CurrentCell As Integer
    Dim CurrentRow As Integer
    Dim StrPass As String
    
    Select Case KeyCode
        Case vbKeyEscape
                '// when esc is pressed cancel and get out
            If blnReturn = True Then
                lblQty.Visible = False
                blnReturn = False
            End If
            
            If blnVoidTrue = True Then
                FancyFalse
                frmVoid.Visible = False
                blnVoidTrue = False
                Text2.Locked = False
                lblQty.Visible = False
                Grid.Row = lngCurrentRow
                Grid.Col = 0
                Grid_EnterCell
            End If
        Case vbKeyDown
                '// move down until last row, if last move to first
            If blnVoidTrue = True Then
                With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                If Not Grid.Row = Grid.Rows - 1 Then
                    
                    FancyFalse
                    Grid.Row = Grid.Row + 1
                    Grid_EnterCell
                    FancyTrue
                    
                Else
                    FancyFalse
                    Grid.Row = 1
                    Grid_EnterCell
                    FancyTrue
                    
                End If
            End If
        Case vbKeyUp
                '// move up until first row -1, if first then move last
            If blnVoidTrue = True Then
                 With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                If Not Grid.Row = 1 Then
                    FancyFalse
                    Grid.Row = Grid.Row - 1
                    Grid_EnterCell
                    FancyTrue
                Else
                    FancyFalse
                    Grid.Row = Grid.Rows - 1
                    Grid_EnterCell
                    FancyTrue
                End If
            End If
        Case vbKeyReturn
                '// when enter is pressed, move to next col
            If Not blnVoidTrue = True Then
                With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
               ' Grid.Col = 0
            Select Case Grid.Col
                Case 0
                    If Len(Trim(Grid.Text)) = 0 Then
                        Grid.Col = 0
                        Grid_EnterCell
                        Exit Sub
                    End If
               
                    strSearch = CStr(Grid.Text)
                    blnFound = FindItem()
                    If blnFound = True Then
                        If rsItem!stack < 1 Then
                            MsgBox "Not enough stock for this item", vbCritical, "Invalid Stock"
                            Grid.Col = 0
                            Grid.Text = ""
                            Grid_EnterCell
                            Exit Sub
                        End If
                        Grid.Col = 1
                        Grid.Text = rsItem!descrip
                        Grid.Col = 2
                        If blnDiscountAmount = True Then
                            blnDiscountAmount = False
                            If strDiscount = "1" Then
                                Grid.Text = rsItem!price - CCur(strDiscounted)
                            Else
                                Grid.Text = (rsItem!price - (rsItem!price * (strDiscounted / 100)))
                            End If
                        Else
                            Grid.Text = rsItem!price
                        End If
                        'curItemPrice = rsItem!Price
                        
                        If lblQty.Visible = True Then
                            Grid.Col = 3
                            Grid_Qty
                        Else
                            Grid.Col = 3
                            If blnReturn = True Then
                                Grid.Text = -1
                            Else
                                Grid.Text = 1
                            End If
                            'Grid_EnterCell
              
                            Grid.Col = 3
                            Qty = CLng(Grid.Text)
                            Grid.Col = 2
                            Rate = CCur(Grid.Text)
                            
                            Total = Qty * Rate:
                            Grid.Col = 4
                            Grid.Text = Format(Total, "###,###,##0.00")
                            CopyToLabel
                            DoTotals
                            DoItems ' Compute Items
                            Grid_EnterCell
                            
                            If Not Grid.Row = Grid.Rows - 1 Then
                                If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                                    Grid.Row = Grid.Row + 1
                                End If
                                Grid.Col = 0
                                Grid_EnterCell
                                RowNo
                            Else
                            '// we need to add a new row ey, baby
                                If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                                    Grid.Rows = Grid.Rows + 1
                                    Grid.Row = Grid.Row + 1
                                    'Fancy
                                End If
                                Grid.Col = 0
                        
                                Grid_EnterCell
                                RowNo
                            End If
                            PrintItem
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
                Case 3
                    If lblQty.Visible = True Then
                        'hmmm! this is tricky , but cool (naa! not at all)
                        If Len(Trim(Grid.Text)) = 0 Then
                            Grid.Text = "1"
                        End If
                        If blnReturn = True Then
                            Grid.Text = -(Grid.Text)
                        End If
                        Grid.Col = 3
                        If rsItem!stack < CLng(Grid.Text) Then
                            MsgBox "Stocks for this item is only " & rsItem!stack, vbCritical, "Invalid Stock"
                            Grid.Col = 3
                            Grid.Text = ""
                            Grid_Qty
                            Exit Sub
                        End If
                        Grid.Col = 3
                        Qty = CLng(Grid.Text)
                        Grid.Col = 2
                        Rate = CCur(Grid.Text)
                        Total = Qty * Rate:
                        Grid.Col = 4
                        Grid.Text = Format(Total, "###,###,##0.00")
                       ' curItemPrice = Format(Total, "###,###,##0.00")
                        CopyToLabel 'Copy Grid
                        DoTotals
                        DoItems ' Compute Items
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
                        PrintItem
                    End If
                        
               End Select
            End If
            lblDiscount.Visible = False
        Case vbKeyAdd
            lblQty.Visible = True
            lblQty.Caption = "QUANTITY"
        Case vbKeyF4
            frmItemLook.Show 1
        Case vbKeyTab
            Select Case Grid.Col
            Case 0
                Text2.SetFocus
                'Grid.SetFocus
                Grid.Col = 0
                Grid_EnterCell
                'MsgBox ""
            End Select
         Case vbKeyF2 'VOID
            'MsgBox ""
            If blnVoidTrue = True Then
                Grid.Col = 0
                If Grid.Text = "" Or Grid.Text = "DISCOUNT" Then
                    Grid.Col = 0
                    Grid_EnterCell
                    Exit Sub
                End If
                
                Grid.Col = 5
                If Grid.Text = "V" Or Grid.Text = "-" Then
                    Grid.Col = 0
                    Grid_EnterCell
                    Exit Sub
                End If
                Grid.Col = 5
                Grid.Text = "-"
                Grid.Col = 0
                Grid_EnterCell
                
                blnVoidTrue = False
                Text2.Locked = False
                FancyFalse
                Grid.Col = CurrentCell
                Grid.Row = lngCurrentRow
                
                Dim lngSpace As Long
                Open App.Path & "\temp.txt" For Output As #1
                Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5

                Grid.Col = 0
                Grid.Text = Label1.Caption
                Grid.Col = 1
                Grid.Text = Label2.Caption
                Grid.Col = 2
                Grid.Text = Label3.Caption
                Grid.Col = 3
                Grid.Text = "-" & Label4.Caption
                Grid.Col = 5
                Grid.Text = "V"
                
                lngSpace = (38 - Len(Label2.Caption) - Len(Format(Label3.Caption, "###,##0.00")) - 2)
                Print #1, Label2.Caption & Space(lngSpace) & "-" & Format(Label3.Caption, "###,##0.00") & "V"
                Print #5, Label2.Caption & Space(lngSpace) & "-" & Format(Label3.Caption, "###,##0.00") & "V"
                               
                Print #1, Label1.Caption & Space(7) & Label4.Caption & " @ " & "-" & Format(Label3.Caption, "###,##0.00")
                Print #5, Label1.Caption & Space(7) & Label4.Caption & " @ " & "-" & Format(Label3.Caption, "###,##0.00")
                Close #1
                Close #5
                Addlist
                RunBat
                
                Grid.Col = 3
                Qty = CLng(Grid.Text)
                Grid.Col = 2
                Rate = CCur(Grid.Text)
                Total = Qty * Rate:
                Grid.Col = 4
                Grid.Text = Format(Total, "###,###,##0.00")
                DoTotals
                DoItems ' Compute Items
                Grid.Col = 0
                Grid_EnterCell
                
                Grid.Rows = Grid.Rows + 1
                Grid.Row = Grid.Row + 1
               
                Grid.Col = 0
                Grid_EnterCell
                frmVoid.Visible = False
                lblQty.Visible = False
                lblQty.Caption = "QUANTITY"
            Else
                'If blnDiscount = True Or blnBayad = True Or blnPay = True Then Exit Sub
                If Grid.Rows = 2 Then Exit Sub
                txtPass.Text = ""
                Picture3.Visible = True
                txtPass.SetFocus
                CurrentCell = Grid.Col
                lngCurrentRow = Grid.Row
                frmVoid.Visible = True
'                fComand.Visible = False
                lblQty.Visible = True
                lblQty.Caption = "VOID ITEM"
                blnVoid = True
                txtPass.PasswordChar = "*"
                Label26.Caption = "VOID ITEM"
                Label11.Caption = "PLEASE INPUT MANAGER PASSWORD"

            End If
        Case vbKeyF9 'logoff
            If Not Grid.Rows = 2 Then Exit Sub
            If UserTaskLevel = 3 Then
                frmMain.txtUserID.Text = ""
                frmMain.txtFullName.Text = ""
                frmMain.txtPassword.Text = ""
                
            End If
            ShowCursor (1)
            Unload Me
            frmMain.cmdOk.Default = False
            'frmMain.txtUserID.SetFocus
            
        Case vbKeyF12 'Payment
            If blnVoidTrue = True Or blnDiscount = True Then Exit Sub
            If Grid.Rows = 2 Then Exit Sub
            txtPass.Text = ""
            txtPass.PasswordChar = ""
            lngCurrentRow = Grid.Row
            blnBayad = True
            Picture6.Visible = True
            Frame1.Visible = True
            ListView1.SetFocus
            frmStatus.Visible = False
            Text2.Locked = True
            
            blnPay = True
            blnCash = True
            Picture3.Visible = True
            Label26.Caption = "AMOUNT"
            Label11.Caption = "PLEASE ENTER AMOUNT"
            txtPass.SetFocus
            
        Case vbKeyF1 ' Cancel
            If blnVoidTrue = True Or blnDiscount = True Then Exit Sub
            If Grid.Rows = 2 Then Exit Sub
            txtPass.Text = ""
            Picture3.Visible = True 'password
            txtPass.SetFocus
            CurrentCell = Grid.Col
            lngCurrentRow = Grid.Row
            lblQty.Visible = True
            lblQty.Caption = "CANCEL TRANSACTION"
            blnCancel = True
            txtPass.PasswordChar = "*"
            Label26.Caption = "CANCEL TRANSACTION"
            Label11.Caption = "PLEASE INPUT MANAGER PASSWORD"
        Case vbKeyF5 'Discount
            If blnVoidTrue = True Then Exit Sub
            txtPass.Text = ""
            lngCurrentRow = Grid.Row
            blnDiscount = True
            Picture3.Visible = True 'password
            txtPass.SetFocus
            lblDiscount.Visible = True
            Label26.Caption = "DISCOUNT ITEM"
            Label11.Caption = "PLEASE INPUT MANAGER PASSWORD"
            txtPass.PasswordChar = "*"
        Case vbKeyF3 'Return
            If blnVoidTrue = True Then Exit Sub
            lblQty.Visible = True
            lblQty.Caption = "RETURN"
            blnReturn = True
        Case vbKeyF11
            If blnVoidTrue = True Or blnDiscount = True Then Exit Sub
            If Grid.Rows = 2 Then Exit Sub
            txtPass.Text = ""
            txtPass.PasswordChar = ""
            lngCurrentRow = Grid.Row
            blnBayad = True
            Picture6.Visible = True
            Frame1.Visible = True
            ListView1.SetFocus
            frmStatus.Visible = False
            Text2.Locked = True
            
            txtPass.Text = ""
            blnDiscount = True
            blnAllDiscount = True
            Picture3.Visible = True 'password
            Frame1.Visible = False
            frmStatus.Visible = True
            txtPass.SetFocus
            lblDiscount.Visible = True
            Label26.Caption = "DISCOUNT ALL"
            Label11.Caption = "PLEASE INPUT MANAGER PASSWORD"
            txtPass.PasswordChar = "*"
        End Select
        
End Sub

Public Sub DoTotals()
    '// get the total from all
    Dim CurrentCell As Integer
    Dim CurrentRow As Integer
    Dim lTotal As Currency
    Dim r As Long
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
    Label12.Caption = Format(lTotal, "###,###,##0.00")
    Label30.Caption = Format(lTotal, "###,###,##0.00")
    Label15.Caption = Format(lTotal, "###,###,##0.00")
    DoEvents
    
    Grid.Col = CurrentCell
    Grid.Row = CurrentRow
    If blnReturn = True Then blnReturn = False
End Sub
Public Sub DoItems()
    '// get the total from all
    Dim CurrentCell As Integer
    Dim CurrentRow As Integer
    Dim lTotal As Currency
    Dim r As Long
    CurrentCell = Grid.Col
    CurrentRow = Grid.Row
    
    lTotal = 0
    Grid.Col = 3
    For r = 1 To Grid.Rows - 1
        Grid.Row = r
        If Not Grid.Text = Empty Then
            lTotal = lTotal + CCur(Grid.Text)
        End If
    Next
    Label16.Caption = lTotal
    
    DoEvents
    
    Grid.Col = CurrentCell
    Grid.Row = CurrentRow
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
Private Sub Text2_KeyPress(KeyAscii As Integer)
    'check for numeric
    Const conZero As Integer = 48, conNine As Integer = 57
    Const conBackSpace As Integer = 8
    If (KeyAscii < conZero Or KeyAscii > conNine) And KeyAscii <> conBackSpace Then
        KeyAscii = 0
    End If
    
    Dim Qty As Long, Rate As Currency, Total As Currency
    Dim lr As Integer, lTotal As Double
    
    Dim blnFound As Boolean
    
    If Len(Trim(Text2.Text)) = 13 Then
       ' KeyAscii = 0
        With Text2
            If Not .Text = Empty Then
                Grid.Text = .Text
               
            End If
                .Visible = False
                .Text = Empty
        End With
        If Len(Trim(Text2.Text)) <> 0 Then Exit Sub
        strSearch = CStr(Grid.Text)
        blnFound = FindItem()
        If blnFound = True Then
            Grid.Col = 1
            Grid.Text = rsItem!descrip
            Grid.Col = 2
            Grid.Text = rsItem!price
            
            If lblQty.Visible = True Then
                Grid.Col = 3
                Grid_Qty
            Else
                Grid.Col = 3
                Grid.Text = 1
                'Grid_EnterCell
              
                Grid.Col = 3
                Qty = CLng(Grid.Text)
                Grid.Col = 2
                Rate = CCur(Grid.Text)
                Total = Qty * Rate:
                Grid.Col = 4
                Grid.Text = Format(Total, "###,###,##0.00")
                
                CopyToLabel
                
                DoTotals   ' Compute Total
                DoItems ' Compute Items
                Grid_EnterCell
                    If Not Grid.Row = Grid.Rows - 1 Then
                        If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                            Grid.Row = Grid.Row + 1
                        End If
                        KeyAscii = 0
                        Grid.Col = 0
                        Grid_EnterCell
                        RowNo
                    Else
                    '// we need to add a new row ey, baby
                        If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                            Grid.Rows = Grid.Rows + 1
                            Grid.Row = Grid.Row + 1
                            'Fancy
                        End If
                        KeyAscii = 0
                        Grid.Col = 0
                        
                        Grid_EnterCell
                        RowNo
                    End If
                    PrintItem 'Print to receipt
            End If
        Else
           ' MsgBox "Sku did not found!", vbCritical, "Invalid Sku"
            
            Grid.Col = 0
            Grid_EnterCell
            Text2.Text = Empty
            Text2.SetFocus
            SendKeys "{home}+{end}"
            Grid.Col = 0
        End If
    End If
    
End Sub

Private Sub Timer1_Timer()
    Label18.Caption = rsStore!Name & "           " & Format(Now, "mm-dd-yy " & " hh:mm:ss")
End Sub

Public Sub RowNo()
        '// get the total from all
    Dim CurrentCell As Integer
    Dim CurrentRow As Integer
    
    CurrentCell = Grid.Col
    CurrentRow = Grid.Row
    
    
    Grid.Col = 4
    For lngRow = 1 To Grid.Rows - 1
        Grid.Row = lngRow
     Next
    
    DoEvents
    
    Grid.Col = CurrentCell
    Grid.Row = CurrentRow

End Sub

Public Sub FancyTrue()
    Dim CurrentCell As Integer
    Dim r As Integer
    With Grid
        CurrentCell = .Col
        For r = 0 To 4
            .Col = r
            .CellBackColor = RGB(174, 245, 214)
            If r = 0 Then Label1.Caption = .Text
            If r = 1 Then Label2.Caption = .Text
            If r = 2 Then Label3.Caption = .Text
            If r = 3 Then Label4.Caption = .Text
        Next
        .Col = CurrentCell
    End With
End Sub
Public Sub FancyFalse()
    Dim CurrentCell As Integer
    Dim r As Integer
    With Grid
        CurrentCell = .Col
        For r = 0 To 4
            .Col = r
            .CellBackColor = RGB(255, 255, 255)
        Next
        .Col = CurrentCell
    End With
End Sub

Public Sub CopyItems()
    '// get the total from all
    Dim strSku  As String
    Dim lngQty As Long
    Dim strS As String
    Dim blnFound As Boolean
    Dim r As Long
   
    Grid.Col = 0
    For r = 1 To Grid.Rows - 1
        Grid.Row = r
        Grid.Col = 0
        If Not Grid.Text = "" Then
            If Not Grid.Text = "DISCOUNT" Then
        Grid.Col = 0
        strSku = Grid.Text
        Grid.Col = 3
        lngQty = Grid.Text
        Grid.Col = 5
        strS = Grid.Text
        Grid.Col = 0
             
        If Not strS = "-" Or Not strS = "V" Then
            strSearch = CStr(strSku)
            blnFound = FindItem()
            If blnFound = True Then
                rsSale.AddNew
                rsSale!Sku = rsItem!Sku
                rsSale!catcode = rsItem!catcode
                rsSale!subcatcode = rsItem!subcatcode
                rsSale!supcode = rsItem!supcode
                rsSale!descrip = rsItem!descrip
                rsSale!price = rsItem!price
                rsSale!Qty = lngQty
                rsSale!Date = SQLDate(Now)
                rsSale!usercode = frmMain.lblUserCode.Caption
                rsSale.Update
                rsItem.Update "Stack", CLng(rsItem!stack) - CLng(lngQty)
                End If
            End If
         End If
        End If
    Next r
    
    DoEvents
  
End Sub
Public Function FindPass() As Boolean
    Dim strTemp
    
    strTemp = "'" & strSearch & "'"
    On Error Resume Next
    rsUser.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsUser.Find "UserPassword= " & strTemp, 0, adSearchForward
    
    If rsUser!UserPassword = strSearch Then FindPass = True       'found
    On Error GoTo 0
    Err.Clear
    Exit Function
    
ErrorNotOnFile:
      FindPass = False      'not found
    DoEvents
    On Error GoTo 0
    Err.Clear
End Function



Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
       
        Case vbKeyEscape
            If blnPay = True Then
                If blnPayment = True Then Exit Sub
                Picture3.Visible = False
                blnCash = False
                blnPay = False
                blnBayad = False
                Picture6.Visible = False
                Frame1.Visible = False
                frmStatus.Visible = True
                Text2.Locked = False
                Grid.Row = lngCurrentRow
                Grid.Col = 0
                Grid_EnterCell
            
           ElseIf blnVoid = True Then
                frmVoid.Visible = False
                blnVoid = False
                Grid.Row = lngCurrentRow
                Grid.Col = 0
                Grid_EnterCell
                Picture3.Visible = False
                lblQty.Visible = False
            ElseIf blnCancel = True And blnBayad = True Or blnPay = True Then
                blnCancel = False
                lblQty.Visible = False
                Picture3.Visible = False
                Frame1.Visible = True
                frmStatus.Visible = False
                ListView1.SetFocus
            ElseIf blnCancel = True Then
                blnCancel = False
                lblQty.Visible = False
                Picture3.Visible = False
                Frame1.Visible = False
                frmStatus.Visible = True
                Grid.Row = lngCurrentRow
                Grid.Col = 0
                Grid_EnterCell
                
            ElseIf blnAllDiscount = True Then
                frmStatus.Visible = False
                Frame1.Visible = False
                Picture3.Visible = False
                blnAllDiscount = False
                blnCash = False
                blnPay = False
                blnBayad = False
                blnDiscount = False
                blnDiscounted = False
                blnDiscountAmount = False
                lblDiscount.Visible = False
                blnAllDiscount = False
                Picture6.Visible = False
                Frame1.Visible = False
                frmStatus.Visible = True
                Text2.Locked = False
                Grid.Row = lngCurrentRow
                Grid.Col = 0
                Grid_EnterCell
                
            ElseIf blnCancel = True Then
                lblQty.Visible = False
                Picture3.Visible = False
               ' fComand.Visible = True
                Grid.Row = lngCurrentRow
                Grid.Col = 0
                Grid_EnterCell
                lblQty.Visible = False
            ElseIf blnDiscount = True Or blnDiscounted = True Or blnDiscountAmount = True Then
                Grid.Row = lngCurrentRow
                Grid.Col = 0
                Grid_EnterCell
                Picture3.Visible = False
                blnDiscount = False
                blnDiscounted = False
                blnAllDiscount = False
                blnDiscountAmount = False
                lblDiscount.Visible = False
                Label17.Visible = False
            'ElseIf blnCash = True Or blnCredit = True Or blnGift = True Then
                'blnCash = False
                        
            End If
            
        Case vbKeyTab
            txtPass.SetFocus
            SendKeys "{home}+{end}"
        Case vbKeyF5 'DISCOUNT
            If Not blnBayad = True Then Exit Sub
            If blnPayment = True Then Exit Sub
            If blnBayad = True Then
                PrintTotal
                blnBayad = False
                blnPayment = True
            End If
            txtPass.Text = ""
            blnDiscount = True
            blnAllDiscount = True
            Picture3.Visible = True 'password
            Frame1.Visible = False
            frmStatus.Visible = True
            txtPass.SetFocus
            lblDiscount.Visible = True
            Label26.Caption = "DISCOUNT ITEM"
            Label11.Caption = "PLEASE INPUT MANAGER PASSWORD"
            txtPass.PasswordChar = "*"
                        
    End Select
    
End Sub


Public Sub AddTitlle()
    Dim lngSpace As Long
    Dim lngName As Long
    Dim lngStreet As Long
    Dim lngCity As Long
    Dim lngTin As Long
    
    lngName = Len(rsStore!Name)
    lngStreet = Len(rsStore!Street)
    lngCity = Len(rsStore!City)
    lngTin = Len(rsStore!Tin) + Len("TIN ")
    
    If Not Dir(App.Path & "\temp.txt") = "" Then Kill App.Path & "\temp.txt"
    Open App.Path & "\temp.txt" For Output As #1
    Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5
    
    
    lngSpace = (38 - lngName) / 2
    Print #1, Space(lngSpace) & rsStore!Name & Space(lngSpace)
    Print #5, Space(lngSpace) & rsStore!Name & Space(lngSpace)
    
    lngSpace = (38 - lngStreet) / 2
    Print #1, Space(lngSpace) & rsStore!Street & Space(lngSpace)
    Print #5, Space(lngSpace) & rsStore!Street & Space(lngSpace)
    
    lngSpace = (38 - lngCity) / 2
    Print #1, Space(lngSpace) & rsStore!City & Space(lngSpace)
    Print #5, Space(lngSpace) & rsStore!City & Space(lngSpace)
    
    lngSpace = (38 - lngTin) / 2
    Print #1, Space(lngSpace) & "TIN " & rsStore!Tin & Space(lngSpace)
    Print #5, Space(lngSpace) & "TIN " & rsStore!Tin & Space(lngSpace)
    
    Print #1, Space(38)
    Print #5, Space(38)
    Close #1
    Close #5
    DoEvents
  ' OpenBat
    Addlist 'add to list
    RunBat
    
End Sub

Public Sub PrintItem()
    Dim lngSpace As Long
    
    If Not Dir(App.Path & "\temp.txt") = "" Then Kill App.Path & "\temp.txt"
    Open App.Path & "\temp.txt" For Output As #1
    Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5
    
    lngSpace = (38 - Len(Label2.Caption) - Len(Format(Label24.Caption, "###,##0.00")) - 1)
                
    Print #1, Label2.Caption & Space(lngSpace) & Format(Label24.Caption, "###,##0.00") & Label25.Caption
    Print #5, Label2.Caption & Space(lngSpace) & Format(Label24.Caption, "###,##0.00") & Label25.Caption
    If CLng(Label4.Caption) > 1 Then
        Print #1, Label1.Caption & Space(8) & Label4.Caption & " @ " & Label3.Caption
        Print #5, Label1.Caption & Space(8) & Label4.Caption & " @ " & Label3.Caption
    Else
        Print #1, Label1.Caption
        Print #5, Label1.Caption
        End If
    Close #1
    Close #5
    DoTotals
    Addlist
    RunBat
      
End Sub

Public Sub CopyToLabel()
    Grid.Col = 0
    Label1.Caption = Grid.Text
    Grid.Col = 1
    Label2.Caption = Grid.Text
    Grid.Col = 2
    Label3.Caption = Grid.Text
    Grid.Col = 3
    Label4.Caption = Grid.Text
    Grid.Col = 4
    Label24.Caption = Grid.Text
    Grid.Col = 5
    Label25.Caption = Grid.Text
    Grid.Col = 0
End Sub

Public Sub OpenPay()
    Dim sGet As String
    Dim SP() As String
    Dim itm As ListItem
    Dim curAmount As Currency
    Dim strLine As String
    
    Open App.Path & "\payment.txt" For Input As #1
    ListView1.ListItems.Clear
    Do While Not EOF(1)
        Line Input #1, sGet
        If Not sGet = "" Then
            SP = Split(sGet, ",")
            Set itm = ListView1.ListItems.Add(, , SP(0))
            itm.SubItems(1) = Format(SP(1), "###,###,##0.00")
            curAmount = curAmount + SP(1)
        End If
    Loop
    Close #1
    Label28.Caption = Format(curAmount, "###,##0.00") 'tendered
    Label30.Caption = Format(CCur(Label12.Caption) - CCur(curAmount), "###,###,##0.00") ' amuont due
       
    If CCur(curAmount) >= CCur(Label12) Then
        
        Dim lngSpace As Long
        If blnReceipt = True Then Exit Sub
        blnReceipt = True
        Label30.Caption = "0.00"
        Close #2
        Close #5
        Label32.Caption = Format(CCur(curAmount) - CCur(Label12.Caption)) 'CHANGE
        Open App.Path & "\temp1.txt" For Output As #2
        Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5
        
        Print #2, Space(38)
        Print #5, Space(38)
        
        
        lngSpace = (38 - Len("TOTAL PAYMENT") - Len(Format(curAmount, "###,##0.00")) - 1)
        Print #2, "TOTAL PAYMENT" & Space(lngSpace) & Format(curAmount, "###,##0.00") & " "
        Print #5, "TOTAL PAYMENT" & Space(lngSpace) & Format(curAmount, "###,##0.00") & " "
        
        lngSpace = (38 - Len("CHANGE") - Len(Format(Label32, "###,##0.00")) - 1)
        Print #2, "CHANGE" & Space(lngSpace) & Format(Label32, "###,##0.00") & " "
        Print #5, "CHANGE" & Space(lngSpace) & Format(Label32, "###,##0.00") & " "
        
        Print #2, Space(38)
        Print #5, Space(38)
        
        lngSpace = (38 - 16 - Len("item(s)") - Len(Label16.Caption)) / 2
        Print #2, Space(lngSpace) & "****** " & Label16.Caption & " item(s)" & "******" & Space(lngSpace)
        Print #5, Space(lngSpace) & "****** " & Label16.Caption & " item(s)" & "******" & Space(lngSpace)
        
        Print #2, Space(lngSpace) & "Thank You. Come Again." & Space(lngSpace)
        Print #5, Space(lngSpace) & "Thank You. Come Again." & Space(lngSpace)
        
        lngSpace = (38 - Len("Exchange/Return valid for 48hrs only")) / 2
        Print #2, Space(lngSpace) & "Exchange/Return valid for 48hrs only" & Space(lngSpace)
        Print #5, Space(lngSpace) & "Exchange/Return valid for 48hrs only" & Space(lngSpace)
        
        lngSpace = (38 - Len("W/Pos Receipt and Price Tags still")) / 2
        Print #2, Space(lngSpace) & "W/Pos Receipt and Price Tags still" & Space(lngSpace)
        Print #5, Space(lngSpace) & "W/Pos Receipt and Price Tags still" & Space(lngSpace)
        
        lngSpace = (38 - Len("attached to the item")) / 2
        Print #2, Space(lngSpace) & "attached to the item" & Space(lngSpace)
        Print #5, Space(lngSpace) & "attached to the item" & Space(lngSpace)
      
        Print #2, "#" & rsStore!Number & " " & Format(Now, "mm-dd-yy " & "hh:mm:" & "AM/PM") & " " & frmMain.txtUserID.Text
        Print #5, "#" & rsStore!Number & " " & Format(Now, "mm-dd-yy " & "hh:mm:" & "AM/PM") & " " & frmMain.txtUserID.Text
        
        Print #2, Space(38)
        Print #5, Space(38)
        
        lngSpace = (38 - Len("- THIS IS YOUR OFFICIAL RECEIPT -")) / 2
        Print #2, Space(lngSpace) & "- THIS IS YOUR OFFICIAL RECEIPT -" & Space(lngSpace)
        Print #5, Space(lngSpace) & "- THIS IS YOUR OFFICIAL RECEIPT -" & Space(lngSpace)
        
        Print #2, Space(38)
        Print #5, Space(38)
        
        Close #2
        Close #5
        
        Addlist1
        CopyItems 'copy items on inventory
        Shell App.Path & "\print2.bat", vbHide
        MsgBox "Please press <ENTER> to continue.", vbInformation
        blnPayment = False
        blnVoid = False
        blnDiscount = False
        blnCash = False
        blnVoidTrue = False
        blnBayad = False
        blnCash = False
        blnEnd = False
        blnPay = False
        rsStore.Update "Number", rsStore!Number + 1
        EndOfTrans
        
    Else
        blnPay = True
        blnCash = True
        Picture3.Visible = True
        Label26.Caption = "AMOUNT"
        Label11.Caption = "PLEASE ENTER AMOUNT"
        txtPass.SetFocus
            
    End If
End Sub

Public Sub PrintTotal()
    Dim lngSpace As Long
    Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5
    Open App.Path & "\temp.txt" For Output As #1
                
    Print #1, Space(38)
    Print #5, Space(38)
                
    lngSpace = (38 - Len("TOTAL") - Len(Format(Label12.Caption, "###,##0.00")) - 1)
    Print #1, "TOTAL" & Space(lngSpace) & Format(Label12.Caption, "###,##0.00")
    Print #5, "TOTAL" & Space(lngSpace) & Format(Label12.Caption, "###,##0.00")
                                
    Print #1, "AMOUNT TENDERED"
    Print #5, "AMOUNT TENDERED"
                
    Close #1
    Close #5
    Addlist
    RunBat
  
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If blnDiscounted = True Then
        Const conZero As Integer = 48, conNine As Integer = 57
        Const conBackSpace As Integer = 8
        If (KeyAscii < conZero Or KeyAscii > conNine) And KeyAscii <> conBackSpace Then
            KeyAscii = 0
        End If
     
    End If
End Sub

Private Sub txtPass_LostFocus()
    If blnCash = True Or blnCredit = True Or blnGift = True Then
        txtPass.Text = Format(txtPass.Text, "###,##0.00")
    End If
    
End Sub

Public Sub Addlist()
    Dim strLine As String * 38
    Open App.Path & "\temp.txt" For Input As #3
        Do Until EOF(3)
            Line Input #3, strLine
                If List1.ListCount = 23 Then
                    List1.RemoveItem (0)
                End If
            List1.AddItem strLine
            DoEvents
        Loop
    Close #3
End Sub
Public Sub Addlist1()
    Dim strLine As String * 38
    Open App.Path & "\temp1.txt" For Input As #3
        Do Until EOF(3)
            Line Input #3, strLine
                If List1.ListCount = 23 Then
                    List1.RemoveItem (0)
                End If
            List1.AddItem strLine
            DoEvents
        Loop
    Close #3
End Sub
Public Sub EndOfTrans()
    Grid.Clear
    Grid.Cols = 6
    Grid.Rows = 200
    Grid.Row = 0
    Grid.Col = 0: Grid.Text = "SKU"
    Grid.Col = 1: Grid.Text = "DESCRIPTION"
    Grid.Col = 2: Grid.Text = "PRICE"
    Grid.Col = 3: Grid.Text = "QTY"
    Grid.Col = 4: Grid.Text = "TOTAL"
    Grid.Col = 5: Grid.Text = ""
        
    frmStatus.Visible = True
    List1.Clear
        
    Grid.Rows = 2
    Grid.Row = 1
    Grid.Col = 0
    Grid_EnterCell
    Picture6.Visible = False
    Frame1.Visible = False
    Picture3.Visible = False
    frmStatus.Visible = True
    
    Label28.Caption = "0.00"
    ListView1.ListItems.Clear
    blnEnd = True
    lblQty.Visible = False
    Text2.Locked = False
    AddTitlle '
    Label8.Caption = "0.00"
    Label16.Caption = "0"
    Label15.Caption = "0.00"
    blnReceipt = False
    If Not Dir(App.Path & "\payment.txt") = "" Then Kill App.Path & "\payment.txt"
End Sub
