Attribute VB_Name = "SaleMOd"

Public Sub TransCancel()
    Dim lngSpace As Long
    With frmSales
        Open App.Path & "\jornal.00" & rsStore!Posno For Append As #5
        Open App.Path & "\temp.txt" For Output As #2
        
        Print #2, "SALES CANCELLED"
        Print #5, "SALES CANCELLED"
        
        Print #2, "CANCELLED"
        Print #5, "CANCELLED"
        Print #2, "CANCELLED"
        Print #5, "CANCELLED"
        Print #2, "CANCELLED"
        Print #5, "CANCELLED"
        
        Print #2, "#" & rsStore!Number & " " & Format(Now, "mm-dd-yy " & "hh:mm:" & "AM/PM")
        Print #5, "#" & rsStore!Number & " " & Format(Now, "mm-dd-yy " & "hh:mm:" & "AM/PM")

        Print #2, Space(38)
        Print #5, Space(38)
        
        Print #2, "TRANS CANCELLED #" & Space(6) & rsStore!Number
        Print #5, "TRANS CANCELLED #" & Space(6) & rsStore!Number
        
        Print #2, Space(38)
        Print #5, Space(38)
        
        Print #2, "-----------------------------------"
        Print #5, "-----------------------------------"
        
        Print #2, "MANAGER SIGNATURE"
        Print #5, "MANAGER SIGNATURE"
        
        Print #2, Space(38)
        Print #5, Space(38)
        Print #2, Space(38)
        Print #5, Space(38)
        
        Print #2, "-----------------------------------"
        Print #5, "-----------------------------------"
        
        Print #2, "EMPLOYEE SIGNATURE"
        Print #5, "EMPLOYEE SIGNATURE"
        
        Print #2, Space(38)
        Print #5, Space(38)
        Print #2, Space(38)
        Print #5, Space(38)
        
        Close #2
        Close #5
        
        .Addlist
        
       RunBat
        
        MsgBox "Please press <ENTER> to continue.", vbInformation
        .EndOfTrans
        
        End With
End Sub

Public Sub RunBat()
    Shell App.Path & "\print.bat", vbHide
End Sub
