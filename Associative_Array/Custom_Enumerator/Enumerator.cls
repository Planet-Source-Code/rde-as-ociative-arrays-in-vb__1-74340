VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   Enumerator (User object to be enumerated)
'
Option Explicit

Implements IEnumerator

Private mArray() As Variant
Private mIndex(1 To 9) As Long
Private mInst As Long

'Private Sub Form_Load()
'   Dim cEnum As Enumerator
'   Set cEnum = New Enumerator
'   Dim v, v2
'   For Each v In cEnum
'      Debug.Print v
'      For Each v2 In cEnum
'         Debug.Print v2
'      Next
'   Next
'   Set cEnum = Nothing
'End Sub

Private Sub Class_Initialize()
    Dim i As Long
    ReDim mArray(3)
    For i = LBound(mArray) To UBound(mArray)
       mArray(i) = "Item" & i
    Next
End Sub

' Implement our enumerator

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    mInst = mInst + 1&
    ' Reset the user enumerator instance
    mIndex(mInst) = LBound(mArray) - 1&
    Set NewEnum = CreateEnumerator(Me)
End Function

' Return the current index array element
Private Property Get IEnumerator_Current() As Variant
    If IsObject(mArray(mIndex(mInst))) Then
        Set IEnumerator_Current = mArray(mIndex(mInst))
    Else
        IEnumerator_Current = mArray(mIndex(mInst))
    End If
End Property

' Move to the next array element, if possible
' Return if we successfully moved or not
Private Function IEnumerator_MoveNext() As Boolean
    mIndex(mInst) = mIndex(mInst) + 1&
    IEnumerator_MoveNext = (mIndex(mInst) <= UBound(mArray))
End Function

' Reset the index to a position before the first element
Private Sub IEnumerator_Reset()
    mIndex(mInst) = LBound(mArray) - 1&
End Sub

' Terminate enumerator instance
Private Sub IEnumerator_Terminate()
    mInst = mInst - 1&
End Sub
