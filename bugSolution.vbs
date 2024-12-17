Option Explicit

' Corrected version demonstrating explicit type checking and conversion
Dim strValue, intValue

strValue = "123"
intValue = CInt(strValue) ' Explicit conversion to integer

If IsNumeric(strValue) Then
  If CInt(strValue) = intValue Then
    MsgBox "Values match after explicit conversion."
  End If
Else
  MsgBox "String is not numeric."
End If

'Demonstrates early binding through object creation
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'This will cause an error if the object doesn't exist due to early binding
MsgBox objFSO.GetFile(".\test.txt").Size 