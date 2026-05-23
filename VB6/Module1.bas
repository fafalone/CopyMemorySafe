Attribute VB_Name = "Module1"
Option Explicit


 Public Sub Main()
     Dim hVeh As LongPtr
     Dim p As LongPtr
     Dim v As Long
     hVeh = AddVectoredExceptionHandler(1, AddressOf VectoredHandler)
     Debug.Print "Handler installed"
     ' Intentionally crash
     p = 0
     CopyMemory ByVal p, v, 4
     Debug.Print "Still running"
     MsgBox "hi"
     If hVeh <> 0 Then
         RemoveVectoredExceptionHandler hVeh
     End If
 End Sub

