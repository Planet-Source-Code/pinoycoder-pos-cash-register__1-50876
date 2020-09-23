Attribute VB_Name = "UserMod"
Option Explicit
'Global Variables not neccesarily written in this order
Global UserID As String
Global UserPassword As String
Global UserLName As String
Global UserFName As String
Global UserMInitial As String
Global UserExpireDate As String
Global UserActivationDate As String
Global UserTaskLevel As String
'User specific constants
Global Const EXPIRE_TERM = 365 'password expiration interval in days
Global Const MINIMUM_PASSWORD_LENGTH = 4 'Minimum password length
Global Const DEFAULT_PASSWORD = "password"
'Password entry specifics
Global Const APP_PASSWORD_REQUIRED = True 'Enables password protection disable for development
Global Const NUM_TRIES = 4
Global Const PROVIDER = "Microsoft.Jet.OLEDB.4.0"
'Global Const GBL_USER_CONNECT = "C:\UserInfo.mdb"
'Global Const DB_PASSWORD = "1Ov45FD56g"
Global Const TASK_LEVEL_3 = "3 - Cashier"
Global Const TASK_LEVEL_2 = "2 - Supervisor"
Global Const TASK_LEVEL_1 = "1 - Manager"

Public strSearch As String
Public strSearch1 As Long
Public lngRefNo As Long
Public blnLogin As Boolean

Public Function udfUpperName(ByVal strText As String) As String
' Write the proper name of a text control.
' Changes the first letter of a text in capital.
'
   If strText = "" Then
      udfUpperName = ""
      Exit Function
   End If
   udfUpperName = StrConv(strText, vbUpperCase)
End Function
Public Function SQLDate(ConvertDate As Date) As String
    SQLDate = Format(ConvertDate, "mm/dd/yyyy")
End Function

Public Function udfProperName(ByVal strText As String) As String
' Write the proper name of a text control.
' Changes the first letter of a text in capital.
'
   If strText = "" Then
      udfProperName = ""
      Exit Function
   End If
   udfProperName = StrConv(strText, vbProperCase)
End Function
Public Sub udp_Rtrn(ByVal intKey As Integer)
' Upon hiting the enter key in a text control,
' the cursor tranfers to the next object control.
'
   If intKey = vbKeyReturn Then SendKeys "{TAB}"
   'StrConv intKey, vbUpperCase, 0
End Sub



