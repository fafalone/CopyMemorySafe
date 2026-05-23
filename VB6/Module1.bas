Attribute VB_Name = "Module1"
Option Explicit


 Public Sub Main()
     Dim p As Long
     Dim v As Long
     ' Intentionally crash
     p = 0
     CopyMemorySafe ByVal p, v, 4
     Debug.Print "Still running"
     MsgBox "hi"
 End Sub

