VERSION 5.00
Begin VB.Form frmStore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Store Profile"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmStore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   4815
      Begin VB.TextBox txtPosNo 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtTin 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtContact 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtStreet 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "POS No. :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Tin :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Ctiy :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Street :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Store Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   5535
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   500
         Left            =   120
         Picture         =   "frmStore.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   950
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   500
         Left            =   1080
         Picture         =   "frmStore.frx":0F34
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   950
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   500
         Left            =   2040
         Picture         =   "frmStore.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   950
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   500
         Left            =   3000
         Picture         =   "frmStore.frx":1118
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   950
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   500
         Left            =   3960
         Picture         =   "frmStore.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   950
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   500
         Left            =   120
         Picture         =   "frmStore.frx":12FC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   650
         Width           =   4815
      End
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   -360
      Picture         =   "frmStore.frx":13EE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5400
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdEdit.Enabled = True
    Frame2.Enabled = False
    LoadStore
End Sub

Private Sub cmdEdit_Click()
    cmdEdit.Enabled = False
    Frame2.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    txtName.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Len(Trim(txtName.Text)) = 0 Or Len(Trim(txtStreet.Text)) = 0 Or Len(Trim(txtCity.Text)) = 0 Or Len(Trim(txtContact.Text)) = 0 Or Len(Trim(txtPosNo.Text)) = 0 Then
        MsgBox "All store information is required!", vbInformation, "Save"
        Exit Sub
    End If
    rsStore.Update "Name", txtName.Text
    rsStore.Update "Street", txtStreet.Text
    rsStore.Update "City", txtCity.Text
    rsStore.Update "Contact", txtContact.Text
    rsStore.Update "Tin", txtTin.Text
    rsStore.Update "posno", txtPosNo.Text
    rsStore.Requery
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdEdit.Enabled = True
    Frame2.Enabled = False
End Sub
Private Sub Form_Load()
    LoadStore
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Public Sub LoadStore()
    txtName.Text = rsStore!Name
    txtStreet.Text = rsStore!Street
    txtCity.Text = rsStore!City
    If Not IsNull(rsStore!contact) Then txtContact.Text = rsStore!contact
    txtTin.Text = rsStore!Tin
    txtPosNo.Text = rsStore!Posno
    
End Sub
