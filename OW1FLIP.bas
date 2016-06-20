Sub Stop_Button()
End
End Sub

Sub CLEAR()
Set aRange = Sheets("Sheet1").Range("A6.AL50000")
aRange.ClearContents
End Sub


Sub Add_Location()
  Dim NOTES As String
  NOTES = "CREDIT AR"
  Set HE = CreateObject("HostExplorer")
  Set CurrentHost = HE.CurrentHost
  Dim irow As Long
  irow = 6
  
ADD_365:

Do
    If Range("A" & irow).Value = "" Then
    Application.StatusBar = "POE ADD COMPLETE"
    MsgBox "RUN COMPLETE"
    Exit Sub
    End If
    
    file = Range("A" & irow).Value
    Name = Range("B" & irow).Value
    icol = -1
    y = Name
    
    CurrentHost.pause 600
    CurrentHost.Keys (file)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("12^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("9^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("AWG^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("/F^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("/F^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("/F^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("14^M")
    CurrentHost.pause 600
    CurrentHost.Keys (y)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("//^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("Y^M")
    CurrentHost.pause 1600
    CurrentHost.Keys ("/^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("/^M")
    CurrentHost.pause 600

   
    
    Application.StatusBar = "Processing borrower " & file
   
    
irow = irow + 1

Loop
    
End Sub

