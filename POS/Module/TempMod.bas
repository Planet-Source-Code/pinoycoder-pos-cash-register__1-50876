Attribute VB_Name = "TempMod"
Public Type fldReturn
    strSku As String * 13
    lngRef As Long
    lngCat As Long
    lngSub As Long
    lngSupp As Long
    intId As Integer
    strDescrip As String * 30
    lngQty As Long
    ccurPrice As Currency
    dteDate As Date
End Type
Public Type fldDeliver
    strSku As String * 13
    lngRef As Long
    lngCat As Long
    lngSub As Long
    lngSupp As Long
    intId As Integer
    strDescrip As String * 30
    lngQty As Long
    ccurPrice As Currency
    dteDate As Date
End Type

