VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   The interface to implement for an enumerator
'
Option Explicit

' This returns the current item of the enumeration.
' The current item is called after a successful MoveNext
Public Property Get Current() As Variant: End Property

' Requests the enumerator moves to the next item being enumerated.
' Returns True on success, False otherwise. This is called prior
' to Current. If this is successful, Current will be called
Public Function MoveNext() As Boolean: End Function

' Resets enumerator to begin enumrating from the beginning
Public Sub Reset(): End Sub

' Terminate enumerator instance
Public Sub Terminate(): End Sub
