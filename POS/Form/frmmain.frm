VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "PC - Based Cash Register with Barcode/Scanner"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   6690
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   240
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4704
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B028
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":E4BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1194C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   8640
      ScaleHeight     =   3225
      ScaleWidth      =   5385
      TabIndex        =   35
      Top             =   8400
      Visible         =   0   'False
      Width           =   5415
      Begin MSComctlLib.ListView lstReport 
         Height          =   2895
         Left            =   0
         TabIndex        =   36
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5106
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList5"
         SmallIcons      =   "ImageList5"
         ColHdrIcons     =   "ImageList5"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Report"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   8640
      ScaleHeight     =   3225
      ScaleWidth      =   5385
      TabIndex        =   32
      Top             =   7680
      Visible         =   0   'False
      Width           =   5415
      Begin MSComctlLib.ListView lstInventory 
         Height          =   2895
         Left            =   0
         TabIndex        =   33
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5106
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList4"
         SmallIcons      =   "ImageList4"
         ColHdrIcons     =   "ImageList4"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox picMaintenance 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   8640
      ScaleHeight     =   3225
      ScaleWidth      =   5385
      TabIndex        =   29
      Top             =   6960
      Visible         =   0   'False
      Width           =   5415
      Begin MSComctlLib.ListView lstMaintenance 
         Height          =   2895
         Left            =   0
         TabIndex        =   30
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5106
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ColHdrIcons     =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintenance"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2640
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":14DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":156BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":15F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":16872
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1714E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":17FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1887E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1915A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":19A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1A312
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1ABEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1B4CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1BDA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1CBFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1DA4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1E32A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1EC06
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":20912
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":20C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":21B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2295E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":237B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":24606
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":24EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":257BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   7935
      Begin VB.Frame Frame2 
         Height          =   4695
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   7935
         Begin VB.PictureBox Picture1 
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
            Left            =   2880
            Picture         =   "frmmain.frx":25ADA
            ScaleHeight     =   2700
            ScaleWidth      =   4800
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   4800
            Begin VB.CommandButton cmdCancelChange 
               Caption         =   "Cancel"
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
               Left            =   3600
               TabIndex        =   8
               Top             =   2160
               Width           =   975
            End
            Begin VB.CommandButton cmdOkChange 
               Caption         =   "Ok"
               Enabled         =   0   'False
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
               Left            =   2400
               TabIndex        =   7
               Top             =   2160
               Width           =   975
            End
            Begin VB.TextBox txtNewPass1 
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1560
               PasswordChar    =   "*"
               TabIndex        =   6
               Top             =   1800
               Width           =   3015
            End
            Begin VB.TextBox txtCurrentPass 
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1560
               PasswordChar    =   "*"
               TabIndex        =   4
               Top             =   1080
               Width           =   3015
            End
            Begin VB.TextBox txtNewPass 
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1560
               PasswordChar    =   "*"
               TabIndex        =   5
               Top             =   1440
               Width           =   3015
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   480
               TabIndex        =   22
               Text            =   "Text1"
               Top             =   2280
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   360
               Picture         =   "frmmain.frx":4FE1C
               Top             =   480
               Width           =   480
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Change password"
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
               TabIndex        =   27
               Top             =   75
               Width           =   2775
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Please enter the old password and the new password in the boxes below. Please note that the changes will take place immedeately."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   1200
               TabIndex        =   26
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Current password:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "New password:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Confirm password:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   1800
               Width           =   1335
            End
         End
         Begin VB.PictureBox Picture2 
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
            Left            =   2880
            Picture         =   "frmmain.frx":5030F
            ScaleHeight     =   2700
            ScaleWidth      =   4800
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1920
            Width           =   4800
            Begin VB.TextBox txtUserID 
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1560
               TabIndex        =   0
               Text            =   "ABA"
               Top             =   1080
               Width           =   3015
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Exit"
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
               Left            =   3600
               TabIndex        =   3
               Top             =   2160
               Width           =   975
            End
            Begin VB.CommandButton cmdOk 
               Caption         =   "&Ok"
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
               Left            =   2400
               TabIndex        =   2
               Top             =   2160
               Width           =   975
            End
            Begin VB.TextBox txtPassword 
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1560
               PasswordChar    =   "*"
               TabIndex        =   1
               Text            =   "wolfgang"
               Top             =   1800
               Width           =   3015
            End
            Begin VB.TextBox txtFullName 
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1440
               Width           =   3015
            End
            Begin VB.PictureBox Picture4 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
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
               Height          =   480
               Left            =   360
               Picture         =   "frmmain.frx":7A651
               ScaleHeight     =   480
               ScaleWidth      =   480
               TabIndex        =   14
               Top             =   480
               Width           =   480
            End
            Begin VB.Label lblUsercode 
               Caption         =   "Label13"
               Height          =   255
               Left            =   480
               TabIndex        =   38
               Top             =   2280
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Login"
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
               TabIndex        =   20
               Top             =   75
               Width           =   2775
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "User ID:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Full Name:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Password:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "You will need an user or administrator password to access this software.         Please Login."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   1560
               TabIndex        =   16
               Top             =   360
               Width           =   3135
            End
         End
         Begin VB.Image Image3 
            Height          =   2160
            Left            =   720
            Picture         =   "frmmain.frx":7B293
            Top             =   240
            Width           =   6480
         End
         Begin VB.Image Image1 
            Height          =   3240
            Left            =   0
            Picture         =   "frmmain.frx":7CCAB
            Top             =   1320
            Width           =   3240
         End
      End
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   1080
         Left            =   0
         TabIndex        =   10
         Top             =   5040
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   1905
         BandCount       =   1
         _CBWidth        =   7935
         _CBHeight       =   1080
         _Version        =   "6.7.8988"
         Child1          =   "Toolbar1"
         MinHeight1      =   1020
         Width1          =   6675
         NewRow1         =   0   'False
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   1020
            Left            =   30
            TabIndex        =   11
            Top             =   30
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   1799
            ButtonWidth     =   1799
            ButtonHeight    =   1799
            Style           =   1
            ImageList       =   "ImageList3"
            DisabledImageList=   "ImageList3"
            HotImageList    =   "ImageList3"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   11
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Point of Sale"
                  ImageIndex      =   1
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   1
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   4
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Store Profile"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   4
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Inventory"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   4
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Maintenance"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   4
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Report"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   4
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Security"
                  ImageIndex      =   6
               EndProperty
            EndProperty
            OLEDropMode     =   1
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   6075
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12064
            MinWidth        =   12064
            Picture         =   "frmmain.frx":80477
            Text            =   "Status : Connecting to Database...."
            TextSave        =   "Status : Connecting to Database...."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/18/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:56 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":80D53
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8162F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8194B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8279F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":82ABB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":82F0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8322B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":83547
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":83E23
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":846FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":84FDB
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsSaleko As ADODB.Recordset
Dim xx As Double, yy As Double
Dim intCount As Integer
Dim strLine As String * 91
Dim Tries As Integer

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdCancelChange_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
       
    Set Text1.DataSource = rsUser
    Text1.DataField = "UserPassword"
    
   
    Frame1.Top = (Screen.Height - Frame1.Height) / 2
    Frame1.Left = (Screen.Width - Frame1.Width) / 2
    
    
    Timer1.Enabled = True
    
    
    Dim itmX As ListItem
    Set itmX = lstMaintenance.ListItems.Add(1, "Cat", "Category", 5, 5)
    Set itmX = lstMaintenance.ListItems.Add(2, "Sub", "Sub  Category", 1, 1)
    Set itmX = lstMaintenance.ListItems.Add(3, "Sup", "Supplier", 4, 4)
    Set itmX = lstMaintenance.ListItems.Add(4, "Itm", "Item", 3, 3)
    Set itmX = lstMaintenance.ListItems.Add(5, "Ret", "Return to Main", 2, 2)
    
    Set itmX = lstInventory.ListItems.Add(1, "Del", "Delivery", 2, 2)
    Set itmX = lstInventory.ListItems.Add(2, "Sup", "Return to Supplier", 3, 3)
    Set itmX = lstInventory.ListItems.Add(3, "Ret", "Return to Main", 4, 4)
    
    Set itmX = lstReport.ListItems.Add(1, "Re1", "Delivery", 2, 2)
    Set itmX = lstReport.ListItems.Add(2, "Re2", "Sale", 4, 4)
   ' Set itmX = lstReport.ListItems.Add(3, "Re3", "Return to Supplier", 6, 6)
    Set itmX = lstReport.ListItems.Add(3, "Ret", "Return to Main", 1, 1)
    
    
    picMaintenance.Top = (Screen.Height - picMaintenance.Height) / 2
    picMaintenance.Left = (Screen.Width - picMaintenance.Width) / 2
   
    picInventory.Top = (Screen.Height - picInventory.Height) / 2
    picInventory.Left = (Screen.Width - picInventory.Width) / 2
    Picture3.Top = (Screen.Height - Picture3.Height) / 2
    Picture3.Left = (Screen.Width - Picture3.Width) / 2
    
End Sub


Private Sub lstInventory_Click()
    LoadInventory
End Sub

Private Sub lstInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xx = X
    yy = Y
End Sub

Private Sub lstMaintenance_Click()
    LoadMaintenance
End Sub

Private Sub lstReport_Click()
    LoadReport
End Sub

Private Sub lstReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xx = X
    yy = Y
End Sub

Private Sub Timer1_Timer()
    intCount = intCount + 1
    If intCount = 1 Then
        DoEvents
        StatusBar1.Panels(1).Text = "STATUS : Pls wait opening database in progress..."
        
        
    ElseIf intCount = 2 Then
        Picture1.Enabled = True
        intCount = 0
        Timer1.Enabled = False
        Me.MousePointer = 0
        StatusBar1.Panels(1).Text = "STATUS : Pls.Login"
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
        Case 1: LoadSale
        Case 3: LoadStore
        Case 5: LoadInvent
        Case 7: LoadMaint
        Case 9: LoadRep
        Case 11: LoadUser
       
    End Select
End Sub





Private Sub txtUserID_DblClick()
    txtUserID.Text = ""
    txtFullName.Text = ""
    txtPassword.Text = ""
    CoolBar1.Visible = False
    cmdOk.Default = False
    blnLogin = False
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        
        'On Error Resume Next
        End
    End If
    udp_Rtrn KeyAscii
End Sub

Private Sub txtUserID_LostFocus()
  Dim blnFound As Boolean
  If blnLogin = False Then
    txtUserID.Text = udfUpperName(txtUserID.Text)
    strSearch = CStr(txtUserID.Text)
    blnFound = FindUser
        If blnFound = True Then
            txtFullName.Text = rsUser!UserFirstname & " " & rsUser!UserMiddleInitial & ". " & rsUser!UserLastname
            lblUserCode.Caption = rsUser!usercode
             cmdOk.Default = True
        Else
            MsgBox "User ID not found!", vbCritical + vbOKOnly, "Invalid ID"
            txtUserID.SetFocus
            SendKeys "{home}+{end}"
            txtFullName.Text = ""
            Exit Sub
        End If
    End If
End Sub

Sub PassCheck()
    Dim Password As String
    Dim UserTask As String
    Dim UserExpire As String
    Dim ID As String
    Dim newdate As Date
    Password = Text1.Text  'rsUser!UserPassword
    With frmMain
    If .txtPassword.Text = Password Then 'rsUser!UserPassword Then
        UserPassword = Password
        UserTaskLevel = Left(rsUser!UserTaskLevel, 1)
        UserExpireDate = rsUser!UserExpireDate
            If UserPassword = DEFAULT_PASSWORD Then
                .Picture1.Visible = True
                MsgBox "Your Password is the Default and must be changed.", vbApplicationModal, "Password needs changed"
                .Picture2.Visible = False
                txtCurrentPass.SetFocus
                 blnLogin = True
                Exit Sub
                Tries = 0
            End If
            If UserTaskLevel = 1 Then
                .Toolbar1.Buttons.Item(11).Enabled = True
                CoolBar1.Visible = True
                cmdOk.Default = False
                blnLogin = True
                Tries = 0
            ElseIf UserTaskLevel = 2 Then
                .Toolbar1.Buttons.Item(11).Enabled = False
                CoolBar1.Visible = True
                cmdOk.Default = False
                blnLogin = True
                Tries = 0
            Else
                Beep
                Beep
                Beep
                frmSales.Show 1
                Tries = 0
               ' NUM_TRIES = 0
            End If
        Else
                MsgBox "Your password is incorrect", vbCritical + vbOKOnly, "Invalid Password"
                SendKeys "{home}+{end}"
                txtPassword.SetFocus
            End If
    End With
End Sub
Public Function FindUser() As Boolean
    Dim strTemp
    
    strTemp = "'" & strSearch & "'"
    On Error Resume Next
    rsUser.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsUser.Find "UserId = " & strTemp, 0, adSearchForward
    
    If rsUser!UserID = strSearch Then FindUser = True       'found
    On Error GoTo 0
    Err.Clear
    Exit Function
    
ErrorNotOnFile:
    'txtUserID.SetFocus
    'SendKeys "{home}+{end}"
    'MsgBox "User ID not found"
    
    'txtFullName.Text = ""
    FindUser = False      'not found
    DoEvents
    On Error GoTo 0
    Err.Clear
End Function
Private Sub txtNewPass1_Change()
    Dim count As String
    Dim count1 As String
    count = Len(txtNewPass.Text)
    count1 = Len(txtNewPass1.Text)
    If count1 >= count Then
    cmdOkChange.Enabled = True
    End If
End Sub
Private Sub txtNewPass_Validate(KeepFocus As Boolean)
    Dim count As Integer
    Dim SearchStr As String
    Dim MyStr As String
    
    count = Len(txtNewPass.Text)
    If txtNewPass = txtCurrentPass.Text Then
        SendKeys "{home}+{end}"
        KeepFocus = True
        MsgBox "New password can't be the same as old", vbCritical + vbOKOnly, "New Password"
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If count < MINIMUM_PASSWORD_LENGTH Then
        SendKeys "{home}+{end}"
        KeepFocus = True
        MsgBox "Your password must be at least 8 characters.", vbApplicationModal, "New Password"
        SendKeys "{home}+{end}"
    Exit Sub
    End If
    
    Dim midstr, midpassstr, newmidstr As String
    midstr = txtCurrentPass.Text
    midpassstr = midstr
    newmidstr = Mid(midstr, 3, 4)
    SearchStr = txtNewPass.Text
    MyStr = InStr(1, SearchStr, newmidstr, vbTextCompare)
    If MyStr <> 0 Then
        KeepFocus = True
        MsgBox "The password can not contain similiar words from old password.", vbCritical, "New Password."
        SendKeys "{home}+{end}"
    Exit Sub
    End If
    
End Sub
Private Sub txtCurrentPass_Validate(KeepFocus As Boolean)
    If txtCurrentPass <> UserPassword Then
        KeepFocus = True
        MsgBox "Your currrent password is invalid", vbCritical + vbOKOnly, "Current Password"
        SendKeys "{Home}+{End}"
        txtCurrentPass.SetFocus
    End If
End Sub



Private Sub txtFullName_KeyPress(KeyAscii As Integer)
     udp_Rtrn KeyAscii
End Sub

Private Sub txtNewPass_KeyPress(KeyAscii As Integer)
     udp_Rtrn KeyAscii
End Sub
Private Sub txtCurrentPass_KeyPress(KeyAscii As Integer)
     udp_Rtrn KeyAscii
End Sub
Private Sub cmdOk_Click()
    Dim blnFoundPass As Boolean
    Dim task  As String
    
    
    'If to many tries we want to get rid of this user.
    If Len(Trim(txtUserID.Text)) = 0 Then
        MsgBox "User ID not found", vbCritical + vbOKOnly, "Invalid User ID"
        Exit Sub
    End If
    If Len(Trim(txtPassword.Text)) = 0 Then
        
        MsgBox "Your password is incorrect", vbCritical + vbOKOnly, "Login"
        SendKeys "{home}+{end}"
        txtPassword.SetFocus
        Exit Sub
    End If
    Tries = Tries + 1
    If Tries >= NUM_TRIES Then     'nope, otta hear, you have gotta go.
        MsgBox "YOUR ACCESS HAS BEEN DENIED PLEASE CALL EDP.", vbApplicationModal, "Point Of Sale System"
        End
        
    End If
    strSearch = CStr(txtUserID.Text)
    blnFoundPass = FindUser
        If blnFoundPass = True Then
            Call PassCheck
            UserID = txtUserID.Text
        End If
End Sub


Private Sub cmdOkChange_Click()
               
    If txtNewPass = txtNewPass1 Then
        rsUser.Update "UserPassword", txtNewPass.Text
        If UserTaskLevel = 1 Then
                frmMain.Toolbar1.Buttons.Item(11).Enabled = True
                CoolBar1.Visible = True
                cmdOk.Default = False
                blnLogin = True
                Tries = 0
            ElseIf UserTaskLevel = 2 Then
                frmMain.Toolbar1.Buttons.Item(11).Enabled = False
                CoolBar1.Visible = True
                cmdOk.Default = False
                blnLogin = True
                Tries = 0
            Else
                Picture1.Visible = False
                Beep
                Beep
                Beep
                frmSales.Show 1
                Tries = 0
        End If
            Picture1.Visible = False
            Picture2.Visible = True
    Else
        MsgBox "Your new and comfirmation password do not match!", vbExclamation, "Password Do Not Match"
    End If
   
                
       
End Sub

Private Sub LoadMaintenance()
    On Error GoTo ExitThis
    If lstMaintenance.HitTest(xx, yy).Key = "Cat" Then
        frmCategory.Show 1
    End If
    If lstMaintenance.HitTest(xx, yy).Key = "Sub" Then
        frmSubCat.Show 1
    End If
    If lstMaintenance.HitTest(xx, yy).Key = "Itm" Then
        frmItem.Show 1
    End If
    If lstMaintenance.HitTest(xx, yy).Key = "Sup" Then
        frmSupplier.Show 1
    End If
    If lstMaintenance.HitTest(xx, yy).Key = "Ret" Then
        ReturnMain
    End If
ExitThis:
End Sub

Private Sub lstMaintenance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xx = X
    yy = Y
End Sub



Private Sub LoadInventory()
On Error GoTo ErrorClick
    If lstInventory.HitTest(xx, yy).Key = "Del" Then
        frmDelivery.Show 1
    End If
    If lstInventory.HitTest(xx, yy).Key = "Sup" Then
        frmReturn.Show 1
    End If
    
    If lstInventory.HitTest(xx, yy).Key = "Ret" Then
        ReturnMain
    End If
ErrorClick:
End Sub


Public Sub LoadReport()
On Error GoTo ErrorClick
    If lstReport.HitTest(xx, yy).Key = "Re1" Then
        blnDelReport = True
        frmReport.Caption = "Delivery Report"
        frmReport.Show 1
    End If
    If lstReport.HitTest(xx, yy).Key = "Re2" Then
        blnSaleReport = True
        frmReport.Caption = "Sales Report"
        frmReport.Show 1
    End If
    If lstReport.HitTest(xx, yy).Key = "Re3" Then
        ReturnMain
    End If
    If lstReport.HitTest(xx, yy).Key = "Ret" Then
        ReturnMain
    End If
ErrorClick:
End Sub
