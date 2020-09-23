VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Maintenance"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPwdReset 
      Caption         =   "&Reset Password "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   24
      Top             =   5280
      Width           =   825
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   5160
      Width           =   6975
      Begin VB.CommandButton cmdClose 
         Caption         =   "Exit"
         Height          =   500
         Left            =   6000
         Picture         =   "frmUser.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   500
         Left            =   3480
         Picture         =   "frmUser.frx":0F34
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   825
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Save"
         Height          =   500
         Left            =   2640
         Picture         =   "frmUser.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   825
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   500
         Left            =   1800
         Picture         =   "frmUser.frx":1118
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   825
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Edit"
         Height          =   500
         Left            =   960
         Picture         =   "frmUser.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   825
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   500
         Left            =   120
         Picture         =   "frmUser.frx":12FC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   3975
      TabIndex        =   7
      Top             =   840
      Width           =   3975
      Begin MSComCtl2.DTPicker DExpire 
         Height          =   315
         Left            =   2040
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22740993
         CurrentDate     =   37497
      End
      Begin MSComCtl2.DTPicker DActivation 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22740993
         CurrentDate     =   37497
      End
      Begin VB.ComboBox DBComboUserTaskLevel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   3585
      End
      Begin VB.TextBox txtUserNotes 
         DataField       =   "UserNotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3000
         Width           =   3615
      End
      Begin VB.TextBox txtUserID 
         DataField       =   "UserID"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MMMM d, yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   2310
         Width           =   1695
      End
      Begin VB.TextBox txtUserPassword 
         DataField       =   "UserPassword"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2310
         Width           =   1815
      End
      Begin VB.TextBox txtUserLastName 
         DataField       =   "USerLastName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2100
         TabIndex        =   2
         Top             =   360
         Width           =   1675
      End
      Begin VB.TextBox txtUserMiddleInitial 
         DataField       =   "UserMiddleInitial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   1
         Top             =   360
         Width           =   360
      End
      Begin VB.TextBox txtUserFirstName 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "User's First Name"
         Top             =   360
         Width           =   1530
      End
      Begin VB.TextBox txtUserExpireDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   4440
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblUserNotes 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes --Enter anything Regarding this User"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   180
         TabIndex        =   15
         Top             =   2730
         Width           =   3645
      End
      Begin VB.Label lblUserID 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   2040
         Width           =   1890
      End
      Begin VB.Label lblUserPassword 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2070
         Width           =   2070
      End
      Begin VB.Label lblUserExpireDate 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password Expire Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2040
         TabIndex        =   12
         Top             =   1455
         Width           =   2340
      End
      Begin VB.Label lblUserActivationDate 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Activation Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   1455
         Width           =   1845
      End
      Begin VB.Label lblUserTaskLevel 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "User Task Level"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2025
      End
      Begin VB.Label lblUserFullName 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Full User Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   120
         Width           =   2610
      End
   End
   Begin MSComctlLib.ListView lstUser 
      Height          =   3855
      Left            =   3960
      TabIndex        =   23
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full  User Name"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   480
      Picture         =   "frmUser.frx":13EE
      Top             =   0
      Width           =   6540
   End
   Begin VB.Label lblUserCode 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim newdate As String, strSQl As String
Dim blnEdit As Boolean
Dim blnAdd As Boolean


Private Sub cmdCancel_Click()
    ClearText
    blnStatus True
End Sub

Private Sub cmdChange_Click()
    blnStatus False
    txtUserFirstName.SetFocus
    blnEdit = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    DBComboUserTaskLevel.AddItem TASK_LEVEL_1
    DBComboUserTaskLevel.AddItem TASK_LEVEL_2
    DBComboUserTaskLevel.AddItem TASK_LEVEL_3
  
'    newdate = DateAdd("d", EXPIRE_TERM, txtUserExpireDate)
'    DExpire.Value = Format(newdate, "mmmm d, yyyy")
    UpdateUser 'view User
        
    blnStatus True
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub


Private Sub lstUser_Click()
    ClearText
    CheckUser
End Sub

Private Sub txtUserFirstName_LostFocus()
    txtUserFirstName.Text = udfProperName(txtUserFirstName.Text)
End Sub

'makes the user ID, I user 1 from first name, middle initial, and 6 (if there) from Last
Private Sub txtUserLastName_Change()
    Dim blnFound As Boolean
    If blnAdd = True Then
        txtUserID = Left([txtUserFirstName], 1) & Left([txtUserMiddleInitial], 1) & Left([txtUserLastName], 1)
        strSearch = CStr(txtUserID.Text)
        blnFound = FindUser
        If blnFound = True Then
            txtUserID = Left([txtUserFirstName], 1) & Left([txtUserMiddleInitial], 1) & Left([txtUserLastName], 2)
        End If
    End If
End Sub
Private Sub txtUserLastName_KeyPress(KeyAscii As Integer)
    If Not blnEdit = True Then
        txtUserPassword.Text = "password"
        DActivation = Date
        txtUserExpireDate = Date
        newdate = DateAdd("d", EXPIRE_TERM, txtUserExpireDate)
        DExpire = Format(newdate, "mmmm d, yyyy")
    End If
End Sub
Private Sub cmdUpdate_Click()
   ' On Error GoTo Errlbl:
    '
   ' On Error Resume Next
    With rsUser
        If blnAdd = False Then
             Dim strTemp
    
            strTemp = "'" & CStr(lblUserCode.Caption) & "'"
            On Error Resume Next
            rsUser.MoveFirst
    
            On Error GoTo ErrorNotOnFile:
            rsUser.Find "Usercode = " & strTemp, 0, adSearchForward
    
            If rsUser!usercode = lblUserCode.Caption Then
                .Update "UserFirstname", txtUserFirstName.Text
                .Update "UserLastname", txtUserLastName.Text
                .Update "UserTaskLevel", DBComboUserTaskLevel.Text
                .Update "Userpassword", txtUserPassword.Text
                .Update "UserMiddleInitial", txtUserMiddleInitial.Text
                .Update "UserExpireDate", SQLDate(DExpire)
                .Update "UserActivationDate", SQLDate(DActivation)
                .Update "Userid", txtUserID.Text
                If Not txtUserNotes.Text = "" Then .Update "Usernotes", txtUserNotes.Text
                UpdateUser
                blnStatus True
            End If
            On Error GoTo 0
            Err.Clear
            Exit Sub
    
ErrorNotOnFile:
     
            DoEvents
            On Error GoTo 0
            Err.Clear
        
        ElseIf blnAdd = True Then
           ' sgBox ""
            rsUser.AddNew
            rsUser!UserID = txtUserID.Text
            rsUser!UserPassword = "password"
            rsUser!UserTaskLevel = DBComboUserTaskLevel.Text
            rsUser!UserActivationDate = SQLDate(DActivation)
            rsUser!UserExpireDate = SQLDate(DExpire)
            rsUser!UserLastname = txtUserLastName.Text
            rsUser!UserFirstname = txtUserFirstName.Text
            rsUser!UserMiddleInitial = txtUserMiddleInitial.Text
            If Not txtUserNotes.Text = "" Then rsUser!UserNotes = txtUserNotes.Text
            rsUser.Update
            rsUser.Requery
            blnAdd = False
        End If
        
        UpdateUser
        blnStatus True
        Exit Sub
    End With
Errlbl:
    MsgBox Err.Number & Err.Description '"Make sure all information is entered properly", vbApplicationModal, "Information not Correct"
    'rsUser.CancelUpdate
'Resume Next
End Sub
Private Sub cmdPwdReset_Click()
    Dim blnFound As Boolean
    Dim intRes As Integer
    intRes = MsgBox("Are you sure you want to reset the user password?", vbQuestion + vbYesNo, "Reset Confirmation")
    If intRes = vbNo Then Exit Sub
    
    strSearch = CStr(lstUser.SelectedItem.Text)
    blnFound = FindUser
        If blnFound = True Then
            rsUser.Update "Userpassword", "password"
        End If
    txtUserPassword.Text = "password"
    
End Sub

Private Sub cmdMoveNextFor_Click()
    If Not rsUser.EOF Then
        rsUser.MoveNext
    End If
    If rsUser.EOF And rsUser.RecordCount > 0 Then
        rsUser.MoveLast
    End If
End Sub
Private Sub cmdMoveNextBack_Click()
    rsUser.CancelUpdate
    If Not rsUser.BOF Then
        rsUser.MovePrevious
    End If
    If rsUser.BOF And rsUser.RecordCount > 0 Then
        rsUser.MoveFirst
    End If
End Sub
Private Sub cmdDelete_Click()
    Dim DeleteUser As String
    Dim CurrentUser As String
    Dim blnFound As Boolean
    Dim intRes As Integer
    'Setting up details to make sure the current user will not be deleted.
    DeleteUser = txtUserID.Text
    CurrentUser = UserID
    DeleteUser = Format(DeleteUser, "<")
    CurrentUser = Format(CurrentUser, "<")
    
    If DeleteUser = CurrentUser Then
        MsgBox "You cannot delete the current user.", vbCritical, "User Security"
        Exit Sub
    End If
    intRes = MsgBox("Are you sure you want to delete user?", vbQuestion + vbYesNo, "Delete Confirmation")
    If intRes = vbNo Then Exit Sub
    strSearch = CStr(lstUser.SelectedItem.Text)
    blnFound = FindUser
        If blnFound = True Then
            rsUser.Delete
            ClearText
            UpdateUser
        End If
End Sub
Private Sub cmdAdd_Click()
    On Error GoTo ErrLabel
    blnAdd = True
    blnStatus False
    ClearText
    DBComboUserTaskLevel.Text = "1 - Manager"
    txtUserFirstName.SetFocus
    Exit Sub
    
ErrLabel:
    MsgBox Err.Description
End Sub
Sub Update()
On Error GoTo Errlbl:
    rsUser.Update
    Exit Sub
Errlbl:
    MsgBox "Make sure all information is entered properly", vbApplicationModal, "Information not Correct"
    rsUser.CancelUpdate
    Resume Next
End Sub

Public Sub UpdateUser()
    Dim itm As ListItem
    lstUser.ListItems.Clear
      '  rsUser.Requery
    If rsUser.RecordCount > 0 Then rsUser.MoveFirst
    Do Until rsUser.EOF
        If Not IsNull(rsUser!UserID) Then Set itm = lstUser.ListItems.Add(, , rsUser!UserID)
        itm.SubItems(1) = rsUser!UserFirstname & " " & rsUser!UserMiddleInitial & " " & rsUser!UserLastname
        rsUser.MoveNext
    Loop
    'End If
   
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
       
    'txtFullName.Text = ""
    FindUser = False      'not found
    DoEvents
    On Error GoTo 0
    Err.Clear
End Function



Private Sub CheckUser()
    Dim blnFound As Boolean
    strSearch = CStr(lstUser.SelectedItem.Text)
    blnFound = FindUser
        If blnFound = True Then
            DExpire.Value = rsUser!UserExpireDate
            txtUserExpireDate.Text = rsUser!UserExpireDate
            txtUserID.Text = rsUser!UserID
            txtUserFirstName.Text = rsUser!UserFirstname
            txtUserMiddleInitial.Text = rsUser!UserMiddleInitial
            txtUserLastName.Text = rsUser!UserLastname
            txtUserPassword.Text = rsUser!UserPassword
            If Not IsNull(rsUser!UserNotes) Then txtUserNotes.Text = rsUser!UserNotes
            DActivation.Value = rsUser!UserActivationDate
            DBComboUserTaskLevel.Text = rsUser!UserTaskLevel
            lblUserCode.Caption = rsUser!usercode
        End If
End Sub



Private Sub txtUserLastName_LostFocus()
   txtUserLastName.Text = udfProperName(txtUserLastName.Text)
End Sub

Private Sub txtUserMiddleInitial_KeyPress(KeyAscii As Integer)
    If Len(txtUserMiddleInitial.Text) = 2 Then
        Exit Sub
    End If
End Sub

Private Sub txtUserMiddleInitial_LostFocus()

    txtUserMiddleInitial.Text = udfProperName(txtUserMiddleInitial.Text)
End Sub
Public Sub blnStatus(blnEnable As Boolean)
    lstUser.Enabled = blnEnable
    txtUserPassword.Enabled = blnEnable
    cmdPwdReset.Enabled = Not blnEnable
    cmdDelete.Enabled = Not blnEnable
    Picture1.Enabled = Not blnEnable
    cmdUpdate.Enabled = Not blnEnable
    cmdAdd.Enabled = blnEnable
    cmdChange.Enabled = blnEnable
    cmdDelete.Enabled = blnEnable
    cmdCancel.Enabled = Not blnEnable
End Sub


Public Sub ClearText()
    txtUserExpireDate.Text = ""
    txtUserID.Text = ""
    txtUserFirstName.Text = ""
    txtUserMiddleInitial.Text = ""
    txtUserLastName.Text = ""
    txtUserPassword.Text = ""
    txtUserNotes.Text = ""
   ' txtUserActivationDate.Text = ""
    
End Sub
