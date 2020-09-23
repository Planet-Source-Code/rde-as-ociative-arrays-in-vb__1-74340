Attribute VB_Name = "mEnumerator"

'      Name:    mEnumerator
'      Date:    10/17/2004
'      Author:  Kelly Ethridge
'
'   This module creates lightweight objects that will wrap a user's object that implements the IEnumerator interface.
'   By using lightweight objects, the IEnumVariant interface can easily be implemented, even though it is not VB friendly.
'   The lightweight object forwards the IEnumVariant calls to the IEnumerable interface implemented in the user enumerator.
'
'   To learn more about lightweight objects, you should refer to classic book:
'       Advanced Visual Basic 6 Power Techniques for Everyday Programs
'       By Matthew Curland

Option Explicit

' GUID structure used to identify interfaces
Private Type VBGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function IsEqualGUID Lib "ole32" (rguid1 As VBGUID, rguid2 As VBGUID) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cb As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (pv As Any)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (pDest As Any, ByVal lByteLen As Long)
Private Declare Sub VariantCopy Lib "oleaut32" (pvDest As Variant, pvSrc As Variant)

Private Const ENUM_FINISHED As Long = 1
Private Const E_NOINTERFACE As Long = &H80004002

' This is the type that will wrap the user enumerator.
' When a new IEnumVariant compatible object is created,
' it will have the internal structure of tUserEnum
Private Type tUserEnum
   pVTable As Long
   cRefs As Long
   UserEnum As IEnumerator
End Type

' This is an array of pointers to functions that the
' object's VTable will point to
Private Type VTable
   Functions(0 To 6) As Long
End Type

' The created VTable of function pointers
Private mVTable As VTable

' Pointer to the mVTable memory address
Private mpVTable As Long

' GUIDs to identify IUnknown and IEnumVariant when the interface is queried
Private Const IID_IUnknown_Data1 As Long = 0
Private Const IID_IEnumVariant_Data1 As Long = &H20404
Private IID_IEnumVariant As VBGUID
Private IID_IUnknown As VBGUID

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Public Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Creates the LightWeight object that will wrap the user's enumerator
Public Function CreateEnumerator(ByVal IEnum As IEnumerator) As IUnknown
    Dim This As Long
    Dim Struct As tUserEnum

    If mpVTable = 0& Then Init

    ' Allocate memory to place the new object
    This = CoTaskMemAlloc(LenB(Struct))
    If This = 0& Then Err.Raise 7& ' Out of memory

    ' Fill the structure of the new wrapper object
    With Struct
        Set .UserEnum = IEnum
        .cRefs = 1&
        .pVTable = mpVTable
    End With

    ' Move the structure to the allocated memory to complete the object
    CopyMemory ByVal This, ByVal VarPtr(Struct), LenB(Struct)
    ZeroMemory ByVal VarPtr(Struct), LenB(Struct) ' Clear the structure

    ' Assign the return value to the newly created object
    CopyMemory CreateEnumerator, This, 4&
End Function

' Setup the guids and vtable function pointers
Private Sub Init()
    With IID_IEnumVariant
        .Data1 = &H20404
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With IID_IUnknown
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With mVTable
        .Functions(0) = FuncAddr(AddressOf QueryInterface)
        .Functions(1) = FuncAddr(AddressOf AddRef)
        .Functions(2) = FuncAddr(AddressOf Release)
        .Functions(3) = FuncAddr(AddressOf IEnumVariant_Next)
        .Functions(4) = FuncAddr(AddressOf IEnumVariant_Skip)
        .Functions(5) = FuncAddr(AddressOf IEnumVariant_Reset)
        .Functions(6) = FuncAddr(AddressOf IEnumVariant_Clone)
        mpVTable = VarPtr(.Functions(0))
   End With
End Sub

' Helper to get the function pointers of AddressOf methods
Private Function FuncAddr(ByVal pfn As Long) As Long
    FuncAddr = pfn
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  VTable functions in the IEnumVariant and IUnknown interfaces
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function QueryInterface(This As tUserEnum, riid As VBGUID, pvObj As Long) As Long
    Dim ok As Long

    ' When VB queries the interface, we support only two
    Select Case riid.Data1
        Case IID_IUnknown_Data1     ' IUnknown
            ok = IsEqualGUID(riid, IID_IUnknown)
        Case IID_IEnumVariant_Data1 ' IEnumVariant
            ok = IsEqualGUID(riid, IID_IEnumVariant)
    End Select

    If ok Then
        pvObj = VarPtr(This)
        AddRef This
    Else
        QueryInterface = E_NOINTERFACE
    End If
End Function

' Increment the number of references to the object
Private Function AddRef(This As tUserEnum) As Long
    With This
        .cRefs = .cRefs + 1&
        AddRef = .cRefs
    End With
End Function

' Decrement the number of references to the object,
' checking to see if the last reference was released
Private Function Release(This As tUserEnum) As Long
    With This
        .cRefs = .cRefs - 1&
        Release = .cRefs
        If .cRefs = 0& Then
            ' Clean up the lightweight object
            .UserEnum.Terminate
            Set .UserEnum = Nothing
            ' Release the memory
            CoTaskMemFree VarPtr(This)
        End If
    End With
End Function

' Move to the next element and return it, signaling if we have reached the end
Private Function IEnumVariant_Next(This As tUserEnum, ByVal celt As Long, prgVar As Variant, ByVal pceltFetched As Long) As Long
    If This.UserEnum.MoveNext Then
        VariantCopy prgVar, This.UserEnum.Current

        ' Check if valid (not zero) before we write to that memory location
        If pceltFetched Then CopyMemory ByVal pceltFetched, 1&, 4&
    Else
        IEnumVariant_Next = ENUM_FINISHED
    End If
End Function

' Skip the requested number of elements as long as we don't run out of them
Private Function IEnumVariant_Skip(This As tUserEnum, ByVal celt As Long) As Long
    Do While celt > 0&
        If This.UserEnum.MoveNext = False Then
            IEnumVariant_Skip = ENUM_FINISHED
            Exit Function
        End If
        celt = celt - 1&
    Loop
End Function

' Request the user enumerator to reset
Private Function IEnumVariant_Reset(This As tUserEnum) As Long
   This.UserEnum.Reset
End Function

' We just return a reference to the original object
Private Function IEnumVariant_Clone(This As tUserEnum, ppenum As IUnknown) As Long
   CopyMemory ppenum, VarPtr(This), 4&
   AddRef This
End Function
