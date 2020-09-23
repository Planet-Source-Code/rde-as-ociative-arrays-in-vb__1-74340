VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAssocArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'   cAssocArr - Associative Array
'
'   Self-Balancing Binary-Search AVL Tree with Double-Link List.
'_______________________________________________________________________________
'
'ABSTRACT
'   Modified from Herbert Glarner's BalancedBinaryAVLTree/SatStructures Dll.
'   Keys are internally managed with a balanced binary AVL tree in conjunction
'   with a doubly-linked list. Features very fast insertion, removal, retrieval
'   in key sorted order, and case-sensitive or case-insensitive key lookups.
'_______________________________________________________________________________
'
'DESCRIPTION
'   Unlike VB's Collection object specifying a string Key is always required.
'   Also, specifying an Index is not possible but a perminant Id is returned.
'   It is possible to access the data in key sorted order (ascending as well
'   as descending). All operations are around twice as fast as the Collection
'   object (as far as provided by the Collection).
'
'   Duplicate keys are permitted, but key searches will find any one duplicate
'   only. Enumerating the data will return duplicate keys in arbitrary order
'   (in relation to one another). It is up to the client to manage access to
'   duplicate keys and handling distinguishing issues. When deleting a key
'   which may be duplicated always specify the items perminant internal Id.
'_______________________________________________________________________________
'
'AUTHORS
'   AVL inventors, G.M.Adelson - Velskii And E.M.Landis
'   This software remains (c) 2005 by Herbert Glarner.
'   Rd modified and renamed Text Compare version.
'_______________________________________________________________________________
'
'HISTORY
'   04 Apr 2005    1.0           Creation and Access functionality
'   05 Apr 2005    1.1           Replaced StrComp, Keys stored as Byte Arrays
'   07 Apr 2005                  Fixed a bug in IStrComp
'   13 Nov 2011    1.1.2         Rd modified and renamed version
'   15 Mar 2012    1.2           Added node deletion routine
'   07 Apr 2012    1.3           Converted enumeration to double-link-list
'   08 Apr 2012    1.4           Completely reworked node deletion routine
'_______________________________________________________________________________
'
'FILE
'   AssocArr.cls                 Rd renamed file
'_______________________________________________________________________________
'
'PUBLIC PROPERTIES
'   nNodes = Count               Current item count
'
'PUBLIC METHODS
'   Id   =   Insert (Key, Data)  Returns the perminant internal index
'   Sub      Remove (Key, [Id])  Deletes existing item by key [and id]
'   Sub      Clear               Invalidates all nodes and resets count
'
'Default method returns the data associated with this key
'BINARY SEARCH KEY LOOKUP
'   Data =   Item (Key)          Non-existing returns vbEmpty
'   Item (Key) = Data            Binds new value to existing key
'                               (Non-existing key inserts new item)
'
'This property returns the perminant Id associated with this key
'BINARY SEARCH KEY LOOKUP (pass result to direct access methods below)
'   Id   =   Id (Key)            Non-existing returns -1
'
'Accessing the data as per sorted key order (lexicographically).
'LEXICAL ACCESS (SORTED BY KEY)
'   Data =   FindNext            Ascending, No more items: vbEmpty
'   Data =   FindPrev            Descending, No more items: vbEmpty
'
'These properties accept/return a perminant internal index
'DIRECT ACCESS (RELATIVE BY KEY)
'   Id   =   NextItem (Id)       Ascending, No next item: -1
'   Id   =   PrevItem (Id)       Descending, No previous item: -1
'
'DIRECT ACCESS WITH KNOWN INTERNAL INDEX
'   Key  =   Key (Id)            Non-existing returns vbNullString
'   Data =   Data (Id)           Non-existing returns vbEmpty
'_______________________________________________________________________________
'
'CONSTANTS

'Elements in rNode having no pointer to another index.
Private Const sstNoChild As Long = -1
Private Const sstNoParent As Long = -1

'TextComp results.
Private Const sstKeyIsSmaller As Long = -1
Private Const sstKeyIsEqual As Long = 0
Private Const sstKeyIsLarger As Long = 1

'Used in Balance calculation for Height fields of the UDT.
Private Const ssmbLeftHeavy As Long = -2
Private Const ssmbLeftBalanced As Long = -1
Private Const ssmbBalanced As Long = 0
Private Const ssmbRightBalanced As Long = 1
Private Const ssmbRightHeavy As Long = 2

'Used for navigation direction
Private Const ssmdRight As Long = -1
Private Const ssmdLeft As Long = 0

Private Const DELETED As Long = -2
'_______________________________________________________________________________
'
'UDTS

Private Type rNode
    'The items key and properties of the virtual tree.
    Key As String              'Provided by client
    LeftTree As Long           'Index of left subtree
    RightTree As Long          'Index of right subtree
    PrevNode As Long           'Previous node in key order (linked list)
    NextNode As Long           'Next node in key order (linked list)
    Parent As Long             'Index of the parent entry
    LeftHeight As Long         'This nodes left height (excl. this node)
    RightHeight As Long        'This nodes right height (excl. this node)
End Type
'_______________________________________________________________________________
'
'VARIABLES

'This vector holds the item data.
Private vData() As Variant     'Client specified data

'This vector holds the balanced binary tree.
Private vBinTree() As rNode
Private lBinTreeMax As Long    'The same as UBound(vBinTree)
Private lBinTreeNext As Long   '0=empty, 1=only 0 exists etc.
Private lBinTreeRoot As Long   'Initially 0, later anywhere (due to balancing)
Private lBinTreeCnt As Long    'Current item count

'This vector holds pointers to free indicies in the binary tree.
Private vFreeIdx() As Long

'Last accessed node, stores the continuation points for public enumeration methods
Private lCurNext As Long       'Last FindNext enumerated item
Private lCurPrev As Long       'Last FindPrev enumerated item
'_______________________________________________________________________________

' Text Compare

Private Const MAX_KEY_LEN = 260&

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lLenB As Long)

' SAFEARRAY Header, used in place of the real one to trick VB into letting us access string data in-place
Private Type SAFEARRAY
    cDims      As Integer      ' Count of dimensions in this array
    fFeatures  As Integer      ' Bitfield flags indicating attributes of the array
    cbElements As Long         ' Byte size of each element of the array
    cLocks     As Long         ' Number of times the array has been locked without corresponding unlock. The cLocks field is a reference count that indicates how many times the array has been locked. When there is no lock, you're not supposed to access the array data, which is located in pvData.
    pvData     As Long         ' Pointer to the start of the array data (use only if cLocks > 0)
    cElements  As Long         ' Count of elements in this dimension
    lLbound    As Long         ' The lower-bounding index of this dimension
End Type

Private ThisStr As SAFEARRAY
Private ThisChrs() As Integer
Private ThanStr As SAFEARRAY
Private ThanChrs() As Integer

Private iaLCase() As Long
Private lKeyComp As Long  ' vbTextCompare, vbBinaryCompare
'_______________________________________________________________________________


'CONSTRUCTION AND DESTRUCTION
'
Private Sub Class_Initialize()
    Call Clear(True)
    With ThisStr
        .cDims = 1                ' 1 Dimensional
        .fFeatures = 17           ' Cannot REDIM the array
        .cbElements = 2&          ' This is an integer array
        .cElements = MAX_KEY_LEN  ' Set max key length
        .cLocks = 1&              ' Lock the array
    End With
    CopyMemory ByVal ArrayPtr(ThisChrs), VarPtr(ThisStr), 4&
    With ThanStr
        .cDims = 1                ' 1 Dimensional
        .fFeatures = 17           ' Cannot REDIM the array
        .cbElements = 2&          ' This is an integer array
        .cElements = MAX_KEY_LEN  ' Set max key length
        .cLocks = 1&              ' Lock the array
    End With
    CopyMemory ByVal ArrayPtr(ThanChrs), VarPtr(ThanStr), 4&
    InitLCaseLookupTable iaLCase
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    Do Until i = lBinTreeMax
       If IsObject(vData(i)) Then Set vData(i) = Nothing
       i = i + 1&
    Loop
    Erase vData
    Erase vBinTree
    Erase vFreeIdx
    Erase iaLCase
    CopyMemory ByVal ArrayPtr(ThisChrs), 0&, 4&
    CopyMemory ByVal ArrayPtr(ThanChrs), 0&, 4&
End Sub
'_______________________________________________________________________________


'PUBLIC PROPERTIES

'Returns 0 if there are no nodes at all.
'Note when using a loop for sequential access
'use FindNext/FindPrev or NextItem/PrevItem.
Public Property Get Count() As Long
    Count = lBinTreeCnt
End Property
'_______________________________________________________________________________

'PUBLIC METHODS

'Invalidates all nodes and resets the count
Public Sub Clear(Optional ByVal bKeyCompCaseInsensitive As Boolean = True)
    Dim i As Long
    Do Until i = lBinTreeMax
       If IsObject(vData(i)) Then Set vData(i) = Nothing
       vData(i) = Empty 'vbEmpty
       i = i + 1&
    Loop
    lBinTreeRoot = sstNoChild
    lBinTreeNext = 0&
    lBinTreeCnt = 0&
    lBinTreeMax = 200&
    ReDimVectors 'Create some free nodes for the binary tree.
    lCurNext = sstNoChild
    lCurPrev = sstNoChild
    With vBinTree(sstNoChild)
       .LeftTree = sstNoChild
       .RightTree = sstNoChild
       .PrevNode = sstNoChild
       .NextNode = sstNoChild
    End With
    lKeyComp = Abs(bKeyCompCaseInsensitive) 'vbTextCompare=1, vbBinaryCompare=0
End Sub

'"Insert" associates Data with this Key
'Returns the perminant internal Id of the new key
'This Id can be used to retrieve the Key, Data, NextItem or PrevItem.
Public Function Insert(Key As String, ByVal Data) As Long 'Id
    Dim lCurNode As Long
    Dim lNewNode As Long
    Dim lChild As Long

    '___________________________________________________________________________

    'If the tree is empty, create the first root node directly.
    If lBinTreeCnt = 0& Then
        'Create the very first node, being the first root.
        lNewNode = MakeNode(sstNoParent, Key, Data, 1&)
        'Return its internal index to the caller
        Insert = lNewNode
        Exit Function
    End If
    '___________________________________________________________________________

    'Beginning with the trees root.
    lCurNode = lBinTreeRoot

    'Loop until we create a new leaf.
    Do
        'Comparing new key with existing key
        'Results: -1 when Key smaller, +1 when larger, 0 when equal.
        Select Case TextComp(StrPtr(Key), StrPtr(vBinTree(lCurNode).Key))
          'We test ssmdRight first, because *all* keys qualify only here if they
          'are fed in sorted order; otherwise (random feed) they are almost
          'equally distributed. (First condition is executed slightly faster.)
          Case sstKeyIsLarger      'Right child

            lChild = vBinTree(lCurNode).RightTree
            If lChild = sstNoChild Then
                'Make a new right child node and link it as a child of the
                'current node (becoming the parent now for the new node).
                lNewNode = MakeNode(lCurNode, Key, Data, ssmdRight)

                'Heights of the path need to be updated
                UpdateHeightsIns lCurNode, ssmdRight

                'Return internal index to the caller and leave
                Insert = lNewNode
                Exit Function
            End If

          Case sstKeyIsSmaller     'Left child

            lChild = vBinTree(lCurNode).LeftTree
            If lChild = sstNoChild Then
                'Make a new left child node and link it as a child of the
                'current node (becoming the parent now for the new node).
                lNewNode = MakeNode(lCurNode, Key, Data, ssmdLeft)

                'Heights of the path need to be updated
                UpdateHeightsIns lCurNode, ssmdLeft

                'Return internal index to the caller and leave
                Insert = lNewNode
                Exit Function
            End If

          Case Else        'This key already exists

            lChild = vBinTree(lCurNode).RightTree
            If lChild = sstNoChild Then
                'Make a new right child node and link it as a child of the
                'current node (becoming the parent now for the new node).
                lNewNode = MakeNode(lCurNode, Key, Data, ssmdRight)

                'Heights of the path need to be updated
                UpdateHeightsIns lCurNode, ssmdRight
            Else
                'Follow all left subtrees until there are no more
                Do Until vBinTree(lChild).LeftTree = sstNoChild
                   lChild = vBinTree(lChild).LeftTree
                Loop
                'Make a new left child node and link it as a child of the
                'current node (becoming the parent now for the new node).
                lNewNode = MakeNode(lChild, Key, Data, ssmdLeft)

                'Heights of the path need to be updated
                UpdateHeightsIns lChild, ssmdLeft
            End If

            'Return internal index to the caller and leave
            Insert = lNewNode
            Exit Function
        End Select

        'Following the childs left or right child now.
        lCurNode = lChild
    Loop
    '___________________________________________________________________________

End Function
'___________________________________________________________________________


'"Remove" deletes the specified item by Key (required).
'Optionally specify the items Id if known (faster), and required if duplicates may exist.
Public Sub Remove(Key As String, Optional ByVal Id As Long = -1) ' Created 8 April 2012 ©Rd

   Dim Target As Long
   Dim Branch As Long
   Dim Propagate As Long
   Dim Successor As Long
   '___________________________________________________________________________

   Target = sstNoChild
   If Id = sstNoChild Then
      If lBinTreeCnt > 1& Then
         Target = lBinTreeRoot ' Beginning with the trees root
         Do Until Target = sstNoChild ' Find the node based on passed key
             Select Case TextComp(StrPtr(Key), StrPtr(vBinTree(Target).Key))
               Case sstKeyIsLarger
                  Target = vBinTree(Target).RightTree ' Move right
               Case sstKeyIsSmaller
                  Target = vBinTree(Target).LeftTree  ' Move left
               Case sstKeyIsEqual: Exit Do
             End Select
         Loop ' Following the left or right child

      ElseIf lBinTreeCnt = 1& Then
         If TextComp(StrPtr(Key), StrPtr(vBinTree(lBinTreeRoot).Key)) = sstKeyIsEqual Then
            Clear (lKeyComp = vbTextCompare) ' Delete sole item
         End If
      End If

   ElseIf Id < lBinTreeMax Then
      If TextComp(StrPtr(Key), StrPtr(vBinTree(Id).Key)) = sstKeyIsEqual Then
         Target = Id ' Item specified by internal index
      End If
   End If
   '___________________________________________________________________________

   If Target = sstNoChild Then Exit Sub ' Requested item not found in tree; abort

   Select Case True                                                                              '          8
      ' Select in-order successor or predecessor based on longer branch to minimize unbalancing  '        /   \
      Case vBinTree(Target).RightHeight > vBinTree(Target).LeftHeight                            '       5    [17]
         ' NextNode has a value > deleted node, but is least of the deleted nodes right children '      / \      \
         Successor = vBinTree(Target).NextNode ' Successor moves up                              '     2   7     ^23
         With vBinTree(Successor)                                                                '
            .LeftTree = vBinTree(Target).LeftTree ' Inherit deleted nodes left child (if any)    '          8
             vBinTree(.LeftTree).Parent = Successor                                              '        /   \
            .PrevNode = vBinTree(Target).PrevNode                                                '       5    [17]
             vBinTree(.PrevNode).NextNode = Successor                                            '      / \    /  \
            .LeftHeight = vBinTree(Target).LeftHeight                                            '     2   7  11  ^23
            .RightHeight = vBinTree(Target).RightHeight '-1...                                   '        /          \
            If vBinTree(Target).RightTree <> Successor Then                                      '       6            25
               vBinTree(.RightTree).Parent = .Parent   ' Move up successors right child (if any) '             8
               vBinTree(.Parent).LeftTree = .RightTree ' Replace successor with its right child  '           /   \
              .RightTree = vBinTree(Target).RightTree  ' Inherit deleted nodes right child   |   '         5     [17]
               vBinTree(.RightTree).Parent = Successor                        '             /    '       /  \     /  \
               Propagate = .Parent    ' Prop = 20.L --1                       ' 19.P = 20 _/     '      2    7   11   23
               Branch = ssmdLeft      ' Path left                             ' 20.L = 19        '     / \   /  / \   / \
            Else                                                                                 '    1  4  6 10  13 20 25
               Propagate = Successor  ' Prop = [17].R --1 >> 23.R --1                            '       /   /   /   /    \
               Branch = ssmdRight     ' Path right                                               '      3   9   12  18^   28
            End If                                                                               '                   \
         End With                                                                                '                   19
      ' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
      Case Else 'vBinTree(Target).LeftHeight >= vBinTree(Target).RightHeight                     '           8
         ' PrevNode has a value < deleted node, but is greatest of deleted nodes left children   '         /   \
         Successor = vBinTree(Target).PrevNode                                                   '        5    [17]
         If Successor <> sstNoChild Then                                                         '       /     /
            If vBinTree(Target).LeftTree <> sstNoChild Then ' Predecessor moves up               '      2     11^
               With vBinTree(Successor)                                                          '
                  .RightTree = vBinTree(Target).RightTree ' Inherit deleted nodes right child    '           8
                   vBinTree(.RightTree).Parent = Successor                                       '         /   \
                  .NextNode = vBinTree(Target).NextNode                                          '        5    [17]
                   vBinTree(.NextNode).PrevNode = Successor                                      '       / \    /  \
                  .RightHeight = vBinTree(Target).RightHeight                                    '      2   7  11^  23
                  .LeftHeight = vBinTree(Target).LeftHeight '-1...                               '            /     /
                  If vBinTree(Target).LeftTree <> Successor Then                                 '           10    20
                     vBinTree(.LeftTree).Parent = .Parent    ' Move up predecessors left child (if any)
                     vBinTree(.Parent).RightTree = .LeftTree ' Replace predecessor with its left child         8
                    .LeftTree = vBinTree(Target).LeftTree    ' Inherit deleted nodes left child |'           /   \
                     vBinTree(.LeftTree).Parent = Successor                  '                 / '         5     [17]
                     Propagate = .Parent    ' Prop = 13.R --1                ' 14.P = 13 _ _ _/  '       /  \     /  \
                     Branch = ssmdRight     ' Path right                     ' 13.R = 14         '      2    7   11   23
                  Else                                                                           '     / \   /  / \   / \
                     Propagate = Successor  ' Prop = [17].L --1 >> 11.L --1                      '    1  4  6 10  13 20 25
                     Branch = ssmdLeft      ' Path left                                          '       /   /   / \      \
                  End If                                                                         '      3   9   12 15^    28
               End With                                                                          '                 /
            ' = = = = = = = = = = = = = = = = = = = = = = = = = = =                              '                14
            Else ' Predecessor is not below (so no children)      ' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
               With vBinTree(Target)                                             '
                  vBinTree(.NextNode).PrevNode = Successor                       '            8
                  vBinTree(Successor).NextNode = .NextNode                       '          /   \
                  Branch = (vBinTree(.Parent).RightTree = Target)                '         5    [17]
                  Propagate = .Parent     ' Prop = 8.[L|R] --1                   '        /
                  Successor = sstNoChild  ' Nullify target node                  '       2
               End With                                                          '
            End If                                                               '
         Else ' Target is left-most node of tree (with no right child)           '            8
            With vBinTree(Target)                                                '          /   \
               vBinTree(.NextNode).PrevNode = sstNoChild                         '         5     17
               Propagate = .Parent        ' Prop = 5.L --1                       '        /
               Branch = ssmdLeft          ' Path left                            '      [2]
            End With                                                             '
         End If
   End Select

   With vBinTree(Target)
      vBinTree(Successor).Parent = .Parent
      Select Case True
        Case .Parent = sstNoParent  ' If the deleted node was the trees root,
           lBinTreeRoot = Successor ' the role is taken over by its successor
        Case vBinTree(.Parent).LeftTree = Target ' Deleted node is to parents left
           vBinTree(.Parent).LeftTree = Successor
        Case Else '(.Parent).RightTree = Target  ' Deleted node is to parents right
           vBinTree(.Parent).RightTree = Successor
      End Select

      If lCurNext = Target Then
         ' DELETED preserves enumeration sequence on current deletion
         vBinTree(DELETED).PrevNode = .PrevNode
         vBinTree(DELETED).NextNode = .NextNode
         lCurNext = DELETED
      End If

      If lCurPrev = Target Then
         ' DELETED preserves enumeration sequence on current deletion
         vBinTree(DELETED).PrevNode = .PrevNode
         vBinTree(DELETED).NextNode = .NextNode
         lCurPrev = DELETED
      End If

   End With

   RemoveNode Target
   UpdateHeightsDel Propagate, Branch

   '___________________________________________________________________________

End Sub
'___________________________________________________________________________


'This "Default" property returns the data associated with
'the provided key (or vbEmpty if no such key exists).
Public Property Get Item(Key As String) As Variant 'Data
Attribute Item.VB_UserMemId = 0
    Dim lNode As Long

    'Begin with root node.
    lNode = lBinTreeRoot

    Do Until lNode = sstNoChild 'Find the node based on passed key

        Select Case TextComp(StrPtr(Key), StrPtr(vBinTree(lNode).Key))
          Case sstKeyIsSmaller 'Continue with left subtree
            lNode = vBinTree(lNode).LeftTree

          Case sstKeyIsLarger 'Continue with right subtree
            lNode = vBinTree(lNode).RightTree

          Case sstKeyIsEqual '0 as the last case: Lowest chances to hit.
            If IsObject(vData(lNode)) Then
               Set Item = vData(lNode)
            Else
               Item = vData(lNode)
            End If
            Exit Property   'Found
        End Select

    Loop
    'Not existing
    Item = Empty 'vbEmpty
End Property
'

Public Property Set Item(Key As String, ByVal Data)
    Dim lNode As Long

    'Begin with root node.
    lNode = lBinTreeRoot

    Do Until lNode = sstNoChild 'Find the node based on passed key

        Select Case TextComp(StrPtr(Key), StrPtr(vBinTree(lNode).Key))
          Case sstKeyIsSmaller 'Continue with left subtree
            lNode = vBinTree(lNode).LeftTree

          Case sstKeyIsLarger 'Continue with right subtree
            lNode = vBinTree(lNode).RightTree

          Case sstKeyIsEqual '0 as the last case: Lowest chances to hit.
            Set vData(lNode) = Data
            Exit Property   'Found
        End Select

    Loop
    'Not existing key is added
    Insert Key, Data
End Property
'

Public Property Let Item(Key As String, ByVal Data)
    Dim lNode As Long

    'Begin with root node.
    lNode = lBinTreeRoot

    Do Until lNode = sstNoChild 'Find the node based on passed key

        Select Case TextComp(StrPtr(Key), StrPtr(vBinTree(lNode).Key))
          Case sstKeyIsSmaller 'Continue with left subtree
            lNode = vBinTree(lNode).LeftTree

          Case sstKeyIsLarger 'Continue with right subtree
            lNode = vBinTree(lNode).RightTree

          Case sstKeyIsEqual '0 as the last case: Lowest chances to hit.
            vData(lNode) = Data
            Exit Property   'Found
        End Select

    Loop
    'Not existing key is added
    Insert Key, Data
End Property
'

'This property returns the perminant internal index of the memory associated
'with the provided key (or -1 if no such key exists). With the Id, use the
'very fast direct access functions below to retrieve the associated data.
Public Property Get Id(Key As String) As Long 'Id
    Dim lNode As Long

    'Begin with root node.
    lNode = lBinTreeRoot

    Do Until lNode = sstNoChild 'Find the node based on passed key

        Select Case TextComp(StrPtr(Key), StrPtr(vBinTree(lNode).Key))
          Case sstKeyIsSmaller 'Continue with left subtree
            lNode = vBinTree(lNode).LeftTree

          Case sstKeyIsLarger 'Continue with right subtree
            lNode = vBinTree(lNode).RightTree

          Case sstKeyIsEqual '0 as the last case: Lowest chances to hit.
            Id = lNode       'Found
            Exit Property
        End Select

    Loop
    'Not existing
    Id = sstNoChild
End Property
'___________________________________________________________________________


'Returns the data associated with this keys internal index.
'Returns vbEmpty if there is no such internal index.
Public Property Get Data(ByVal Id As Long) As Variant
    If Id < lBinTreeMax And Id > sstNoChild Then
        If IsObject(vData(Id)) Then
            Set Data = vData(Id)
        Else
            Data = vData(Id)
        End If
    Else
        Data = Empty 'vbEmpty
    End If
End Property

'Returns the key, or a zero-length string if there is no such internal index.
Public Property Get Key(ByVal Id As Long) As String
    If Id < lBinTreeMax And Id > sstNoChild Then
        Key = vBinTree(Id).Key
    End If
End Property
'_______________________________________________________________________________



'Accessing the data as per sorted key order (lexicographically).

'LEXICAL ACCESS (SORTED BY KEY) Ascending

'This method returns the data associated with this item.
'Pass bFindFirst as True to begin sequence.
'No more items returns vbEmpty.
'FindNext cannot be nested within itself.

Public Function FindNext(Optional ByVal bFindFirst As Boolean) As Variant 'Data
   If Not bFindFirst Then
      'Having an index in lCurNext we find the successor of that node.
      lCurNext = vBinTree(lCurNext).NextNode

   Else 'bFindFirst is True to begin sequence
      'Begin with root node.
      lCurNext = lBinTreeRoot

      'Follow all left subtrees until there are no more.
      Do Until vBinTree(lCurNext).LeftTree = sstNoChild
          lCurNext = vBinTree(lCurNext).LeftTree
      Loop
   End If

   'Return result.
   If IsObject(vData(lCurNext)) Then
      Set FindNext = vData(lCurNext)
   Else
      FindNext = vData(lCurNext)
   End If
End Function
'_______________________________________________________________________________


'LEXICAL ACCESS (SORTED BY KEY) Descending

'This method returns the data associated with this item.
'Pass bFindLast as True to begin sequence.
'No more items returns vbEmpty.
'FindPrev cannot be nested within itself.

Public Function FindPrev(Optional ByVal bFindLast As Boolean) As Variant 'Data
   If Not bFindLast Then
      'Having an index in lCurPrev we find the predecessor of that node.
      lCurPrev = vBinTree(lCurPrev).PrevNode

   Else 'bFindLast is True to begin sequence
      'Begin with root node.
      lCurPrev = lBinTreeRoot

      'Follow all right subtrees until there are no more.
      Do Until vBinTree(lCurPrev).RightTree = sstNoChild
          lCurPrev = vBinTree(lCurPrev).RightTree
      Loop
   End If

   'Return result.
   If IsObject(vData(lCurPrev)) Then
      Set FindPrev = vData(lCurPrev)
   Else
      FindPrev = vData(lCurPrev)
   End If
End Function
'_______________________________________________________________________________


'These properties accept/return a perminant internal Id

'DIRECT ACCESS (RELATIVE BY KEY) Ascending

'Pass Id < zero for first item. No next item returns -1
'If you wish to delete items during the enumeration,
'you should record NextItem(Id) before removing Id

Public Property Get NextItem(ByVal Id As Long) As Long 'Id
   If Id < lBinTreeMax And Id > sstNoChild Then
      'Having an index in Id we find the successor of that node.
      Id = vBinTree(Id).NextNode

   Else 'Id < zero for first item
      'Begin with root node.
      Id = lBinTreeRoot

      'Follow all left subtrees until there are no more.
      Do Until vBinTree(Id).LeftTree = sstNoChild
          Id = vBinTree(Id).LeftTree
      Loop
   End If

   NextItem = Id
End Property
'_______________________________________________________________________________

'DIRECT ACCESS (RELATIVE BY KEY) Descending

'Pass Id < zero for last item. No previous item returns -1
'If you wish to delete items during the enumeration,
'you should record PrevItem(Id) before removing Id

Public Property Get PrevItem(ByVal Id As Long) As Long 'Id
   If Id < lBinTreeMax And Id > sstNoChild Then
      'Having an index in Id we find the predecessor of that node.
      Id = vBinTree(Id).PrevNode

   Else 'Id < zero for last item
      'Begin with root node.
      Id = lBinTreeRoot

      'Follow all right subtrees until there are no more.
      Do Until vBinTree(Id).RightTree = sstNoChild
          Id = vBinTree(Id).RightTree
      Loop
   End If

   PrevItem = Id
End Property
'_______________________________________________________________________________


'INTERNAL ROUTINES

Public Sub Validate()
   Call jsw_assert(lBinTreeRoot)
End Sub

' Test function thanks to Julienne S. Walker

' Balanced trees are not trivial data structures, and deletions from binary trees are exceptionally
' tricky; a small change in one part of the tree could cause a violation in a distant part of the tree.
' So it makes sense to have a little tester function to make sure that no violations have occurred.

Private Function jsw_assert(ByVal root As Long) As Long

   Dim ln As Long ' Helpers
   Dim rn As Long
   Dim lh As Long
   Dim rh As Long
   
   Dim lStrComp As Long

   If root = sstNoChild Then
      ' Empty tree case
      jsw_assert = 0
   Else
      ln = vBinTree(root).LeftTree
      rn = vBinTree(root).RightTree

      lh = jsw_assert(ln)
      rh = jsw_assert(rn)

      If Not (ln = sstNoChild) Then
         lStrComp = TextComp(StrPtr(vBinTree(ln).Key), StrPtr(vBinTree(root).Key))
        'If lStrComp <> sstKeyIsSmaller Then '(no duplicates)
         If lStrComp = sstKeyIsLarger Then
            MsgBox "Binary tree violation"      ' Invalid binary search tree
            Exit Function
      End If: End If

      If Not (rn = sstNoChild) Then
         lStrComp = TextComp(StrPtr(vBinTree(rn).Key), StrPtr(vBinTree(root).Key))
        'If lStrComp <> sstKeyIsLarger Then '(no duplicates)
         If lStrComp = sstKeyIsSmaller Then
            MsgBox "Binary tree violation"      ' Invalid binary search tree
            Exit Function
      End If: End If

      If Abs(lh - rh) > 1 Then
         ' Balance height mismatch
         MsgBox "Height violation"     ' Invalid AVL balance
      End If
      If lh > rh Then
         jsw_assert = lh + 1
      Else
         jsw_assert = rh + 1
      End If
   End If

End Function
'

Private Sub RemoveNode(ByVal Index As Long)
    'Index is the element index into the vector.
    If IsObject(vData(Index)) Then
       Set vData(Index) = Nothing
    End If
    vData(Index) = Empty 'vbEmpty
    With vBinTree(Index)
      .Key = vbNullString
      .LeftTree = sstNoChild
      .RightTree = sstNoChild
      .PrevNode = sstNoChild
      .NextNode = sstNoChild
    End With
    lBinTreeCnt = lBinTreeCnt - 1&

    'Free this Index for reuse.
    vFreeIdx(Index) = lBinTreeNext
    lBinTreeNext = Index
End Sub

Private Sub ReDimVectors()
    Dim i As Long
    If lBinTreeMax < 100000 Then
       lBinTreeMax = lBinTreeMax + lBinTreeMax
    Else
       lBinTreeMax = lBinTreeMax + 50000
    End If
    ReDim Preserve vData(-1 To lBinTreeMax) As Variant
    ReDim Preserve vBinTree(-2 To lBinTreeMax) As rNode
    ReDim Preserve vFreeIdx(0 To lBinTreeMax) As Long
    For i = lBinTreeNext To lBinTreeMax
        vFreeIdx(i) = i + 1&
    Next i
End Sub

'Return index to free vector position
Private Function NextFreeIndex() As Long
    NextFreeIndex = lBinTreeNext
    If lBinTreeNext = lBinTreeMax Then ReDimVectors
    lBinTreeNext = vFreeIdx(lBinTreeNext)
End Function

'Creating a new node below a given parent. Returns the index into vBinTree().
Private Function MakeNode(ByVal Parent As Long, Key As String, ByVal Data, ByVal Path As Long) As Long
    Dim lLink As Long
    Dim lFree As Long
    lFree = NextFreeIndex

    If IsObject(Data) Then
        Set vData(lFree) = Data
    Else
        vData(lFree) = Data
    End If
    'lFree is the element index into the vector.
    With vBinTree(lFree)
       .Key = Key
       .LeftTree = sstNoChild
       .RightTree = sstNoChild
        Select Case Path
           Case ssmdRight
              vBinTree(Parent).RightTree = lFree
              lLink = vBinTree(Parent).NextNode
              vBinTree(Parent).NextNode = lFree
              vBinTree(lLink).PrevNode = lFree
             .PrevNode = Parent
             .NextNode = lLink
           Case ssmdLeft
              vBinTree(Parent).LeftTree = lFree
              lLink = vBinTree(Parent).PrevNode
              vBinTree(Parent).PrevNode = lFree
              vBinTree(lLink).NextNode = lFree
             .PrevNode = lLink
             .NextNode = Parent
           Case Else ' Tree Root
              lBinTreeRoot = lFree
             .PrevNode = sstNoChild
             .NextNode = sstNoChild
        End Select
       .Parent = Parent
       .LeftHeight = ssmbBalanced
       .RightHeight = ssmbBalanced
    End With

    lBinTreeCnt = lBinTreeCnt + 1&

    'Return the index as this functions result
    MakeNode = lFree
End Function
'

'BALANCING THE TREE
'Depending on the nodes' actual balances, one of four balancing activities can
'be called (BalanceLL, BalanceRR, BalanceLR or BalanceRL). In a test scenario
'with 1 million nodes with 8 char long RANDOM keys consisting of all upper- and
'lowercase alphabetic characters A...Z and a...z, the balancing routines were
'called so many times:          With 1 mill already SORTED input (key length 4):
'    Rebalances LL 119628            0
'    Rebalances RR 118737       999980
'    Rebalances LR 114924            0
'    Rebalances RL 115005            0
'    Total         468294       999980
Private Sub UpdateHeightsIns(ByVal Parent As Long, ByVal Branch As Long)
    Dim lGrandP As Long
    Dim lNewHeight As Long
    '___________________________________________________________________________
    
    'Must be reverse height, because we might break update for upper levels.
    Do Until Parent = sstNoParent
        'Increase height along path. Direction is either Left or Right.
        'If the Keys are fed sorted, only ssmdRight will appear (if fed in
        'random order, there are almost equal chances for the left and the right
        'path): For this reason, ssmdRight is tested first (slightly faster).
        Select Case Branch

         Case ssmdRight
            lNewHeight = vBinTree(Parent).RightHeight + 1&
            vBinTree(Parent).RightHeight = lNewHeight

            'If Right was *not* responsible for a height increase,
            'this breaks the update (no propagation further up).
            If lNewHeight = vBinTree(Parent).LeftHeight Then Exit Do

         Case Else 'ssmdLeft
            lNewHeight = vBinTree(Parent).LeftHeight + 1&
            vBinTree(Parent).LeftHeight = lNewHeight

            'If Left was *not* responsible for a height increase,
            'this breaks the update (no propagation further up).
            If lNewHeight = vBinTree(Parent).RightHeight Then Exit Do
        End Select

        lGrandP = vBinTree(Parent).Parent ' Record parent now before rotations
        Branch = (vBinTree(lGrandP).RightTree = Parent) ' Path to this position

        'If the shorter branch was added to, they become equal *no height increase* Exit Do^
        'If the branches were equal then the addition *increases the parents height* by one
        'but no balance update is performed by UpdateBalance which returns False so continue.
        'If the longer branch was added to then a rotation will be performed which will shift
        'the addition to the shorter branch instead so no height increase of the longer branch,
        'therefore if True is returned no propagation further up so Exit Do.
        If UpdateBalance(Parent) Then Exit Do

        Parent = lGrandP
    Loop
End Sub

Private Sub UpdateHeightsDel(ByVal Parent As Long, ByVal Branch As Long)
    Dim lGrandP As Long
    Dim lBranch As Long
    Dim lNewHeight As Long
    '___________________________________________________________________________
    
    'Must be reverse height, because we might break update for upper levels.
    Do Until Parent = sstNoParent
        lGrandP = vBinTree(Parent).Parent ' Record parent now before rotations
        lBranch = (vBinTree(lGrandP).RightTree = Parent) ' Path to this position

        'Decrease height along path
        Select Case Branch

         Case ssmdRight
            lNewHeight = vBinTree(Parent).RightHeight - 1&
            vBinTree(Parent).RightHeight = lNewHeight

            'If right was *not* responsible for a height decrease,
            'this breaks the update (no propagation further up).
            If lNewHeight < vBinTree(Parent).LeftHeight Then
                'If True is returned deletion propagates further up
                If UpdateBalance(Parent) = False Then Exit Do
            End If

         Case Else 'ssmdLeft
            lNewHeight = vBinTree(Parent).LeftHeight - 1&
            vBinTree(Parent).LeftHeight = lNewHeight

            'If Left was *not* responsible for a height decrease,
            'this breaks the update (no propagation further up).
            If lNewHeight < vBinTree(Parent).RightHeight Then
                'If True is returned deletion propagates further up
                If UpdateBalance(Parent) = False Then Exit Do
            End If
        End Select

        'If the branches were equal then the deletion causes *no change* to the longer branch
        'or the parent height, also no rotations are performed by UpdateBalance so Exit Do.
        'If the shorter branch was deleted from, they become unbalanced and True is returned
        'where the rotation performed *decreases the longer branch* so propagates further up.
        'If the longer branch was deleted from they become equal so no update required, but the
        '*decrease of the longer branch* reduces the parents height so propagates further up.

        Parent = lGrandP
        Branch = lBranch
    Loop
End Sub

Private Function UpdateBalance(ByVal lNode As Long) As Boolean
    Dim lChild As Long

    'Nodes with a too heavy weight left or right cause a tree balancing
    'activity now. Too heavy is, when there is a height difference of +/-2.
    Select Case vBinTree(lNode).RightHeight - vBinTree(lNode).LeftHeight

      'testing right-heavy first because we want RR to be executed as fast as
      'possible to acknowledge that RR type is executed almost always when keys
      'are fed already sorted to Insert; whereas the distribution of the 4 types is
      'almost the same (totalling in about 44%) when the keys are fed in random order.
      Case ssmbRightHeavy
        'Node is right-heavy (+2). Examine its right child.
        lChild = vBinTree(lNode).RightTree

        Select Case vBinTree(lChild).RightHeight - vBinTree(lChild).LeftHeight

          'Testing ssmbRightBalanced first to take care for the case that
          'keys may come in already sorted, in which case 99.8% of all adds
          'require a RR balancing. If keys are fed random, all 4 types are
          'equally distributed.
          Case ssmbRightBalanced
             BalanceRR lNode   'Right child is right-balanced. RR situation.

          Case ssmbLeftBalanced
             BalanceRL lNode   'Right child is left-balanced. RL situation.

          Case Else
             'Must be ssmbBalanced due to deletion from the parents sibling branch.
             'As these two children are equal a rotaion will increase the sibling
             'branch and decrease only one child here. As no decrease in height for
             'this parent occurs no propagation further up on deletion return False.
             BalanceRR lNode
             Exit Function
        End Select

        'Indicate a rebalance was performed
        UpdateBalance = True

      Case ssmbLeftHeavy
        'Node is left-heavy (-2). Examine its left child.
        lChild = vBinTree(lNode).LeftTree

        Select Case vBinTree(lChild).RightHeight - vBinTree(lChild).LeftHeight

          Case ssmbLeftBalanced
             BalanceLL lNode   'Left child is left-balanced. LL situation.

          Case ssmbRightBalanced
             BalanceLR lNode   'Left child is right-balanced. LR situation.

          Case Else
             'Must be ssmbBalanced due to deletion from the parents sibling branch.
             'As these two children are equal a rotaion will increase the sibling
             'branch and decrease only one child here. As no decrease in height for
             'this parent occurs no propagation further up on deletion return False.
              BalanceLL lNode
             Exit Function
        End Select

        'Indicate a rebalance was performed
        UpdateBalance = True
    End Select

End Function

'_______________________________________________________________________________
'
'INTERNAL ROUTINES TO BALANCE THE BINARY TREE

'LL - A is the node that the rotation is performed on. This rotation is performed
'when A is unbalanced to the left (the left subtree is 2 higher than the right
'subtree) and B is leftheavy (the left subtree of B is 1 higher than the right
'subtree of B). T1, T2 and T3 represent subtrees (a node was added to T1 which
'made B leftheavy and unbalanced A). P is A's parent and listed only because A's
'relationship with P is taken over by B, making relinks neccessary.
'
'        P                        P
'        |                        |
'        A                        B
'       / \      ---------->     / \
'      B  T3                    T1  A
'     / \                      /   / \
'    T1 T2                    *   T2 T3
'   /
'  *
'+--+--------+--------+
'|  | Before | After  |
'+--+--+--+--+--+--+--+
'|  |LT|RT|PA|LT|RT|PA|
'+--+--+--+--+--+--+--+
'|A |B |  |P |T2|  |B |
'|B |  |T2|A |  |A |P |
'|T2|  |  |B |  |  |A |
'|P |  A  |  |  B  |  |
'+--+--+--+--+--+--+--+
'
Private Sub BalanceLL(ByVal lA As Long)
    Dim lP As Long
    Dim lB As Long
    Dim lT2 As Long

    'Indices of old state
    lP = vBinTree(lA).Parent
    lB = vBinTree(lA).LeftTree
    lT2 = vBinTree(lB).RightTree

    'New links
    vBinTree(lA).Parent = lB
    vBinTree(lA).LeftTree = lT2
    If lT2 <> sstNoChild Then vBinTree(lT2).Parent = lA
    vBinTree(lB).Parent = lP
    vBinTree(lB).RightTree = lA

    Select Case True
      Case lP = sstNoParent 'If A was the whole trees root,
         lBinTreeRoot = lB  'then its role is taken over by B now.
      Case vBinTree(lP).LeftTree = lA
         vBinTree(lP).LeftTree = lB 'A was either P's left or right child
      Case Else                     'B is taking that place now.
         vBinTree(lP).RightTree = lB
    End Select

    'New weights
    vBinTree(lA).LeftHeight = vBinTree(lB).RightHeight
    vBinTree(lB).RightHeight = vBinTree(lA).LeftHeight + 1
End Sub
'
'RR - A is the node that the rotation is performed on. This rotation is performed
'when A is unbalanced to the right (the right subtree is 2 higher than the left
'subtree) and B is rightheavy (the right subtree of B is 1 higher than the left
'subtree of B). T1, T2 and T3 represent subtrees (a node was added to T3 which
'made B rightheavy and unbalanced A). P is A's parent and listed only because A's
'relationship with P is taken over by B, making relinks neccessary.
'(This rotation is performed on almost every node if the keys are added in already
'sorted order, and no other rotations are performed at all then.)
'
'        P                        P
'        |                        |
'        A                        B
'       / \      ---------->     / \
'      T1  B                    A  T3
'         / \                  / \   \
'        T2  T3               T1 T2   *
'             \
'              *
'+--+--------+--------+
'|  | Before | After  |
'+--+--+--+--+--+--+--+
'|  |LT|RT|PA|LT|RT|PA|
'+--+--+--+--+--+--+--+
'|A |  |B |P |  |T2|B |
'|B |T2|  |A |A |  |P |
'|T2|  |  |B |  |  |A |
'|P |  A  |  |  B  |  |
'+--+--+--+--+--+--+--+
'
Private Sub BalanceRR(ByVal lA As Long)
    Dim lP As Long
    Dim lB As Long
    Dim lT2 As Long

    'Indices of old state
    lP = vBinTree(lA).Parent
    lB = vBinTree(lA).RightTree
    lT2 = vBinTree(lB).LeftTree

    'New links
    vBinTree(lA).Parent = lB
    vBinTree(lA).RightTree = lT2
    If lT2 <> sstNoChild Then vBinTree(lT2).Parent = lA
    vBinTree(lB).Parent = lP
    vBinTree(lB).LeftTree = lA

    Select Case True
      Case lP = sstNoParent 'If A was the whole trees root,
         lBinTreeRoot = lB  'then its role is taken over by B now.
      Case vBinTree(lP).LeftTree = lA
         vBinTree(lP).LeftTree = lB 'A was either P's left or right child
      Case Else                     'B is taking that place now.
         vBinTree(lP).RightTree = lB
    End Select

    'New weights
    vBinTree(lA).RightHeight = vBinTree(lB).LeftHeight
    vBinTree(lB).LeftHeight = vBinTree(lA).RightHeight + 1
End Sub
'
'LR - C is the node that the rotation is performed on. This rotation is performed
'when C is unbalanced to the left (the left subtree is 2 higher than the right
'subtree), A is rightheavy (the right subtree of A is 1 higher than the left
'subtree of A). T1, T2, T3, and T4 represent subtrees (a node was added
'to B's branch which made B heavy, made A rightheavy and unbalanced C).
'This consists of a single left rotation at node A, followed by a single right
'at node C. P is C's parent and listed only because C's relationship with P is
'taken over by B, making relinks neccessary.
'
'        P                        P                        P
'        |                        |                        |
'        C                        C                        B
'       / \     ---------->      / \      ---------->    /   \
'      A  T5                    B  T5                   A     C
'     / \   \                  / \   \                 / \   / \
'    T1  B  T6                A  T4  T6               T1 T3 T4 T5
'   /   / \                  / \  /\                 /   /\ /\   \
'  T2  T3 T4                T1 T3  *                T2     *     T6
'      /\ /\               /    /\
'        *                T2     *
'+--+--------+--------+
'|  | Before | After  |
'+--+--+--+--+--+--+--+
'|  |LT|RT|PA|LT|RT|PA|
'+--+--+--+--+--+--+--+
'|A |  |B |C |  |T3|B |
'|B |T3|T4|A |A |C |P |
'|C |A |  |P |T4|  |B |
'|T3|  |  |B |  |  |A |
'|T4|  |  |B |  |  |C |
'|P |  C  |  |  B  |  |
'+--+--+--+--+--+--+--+
'
Private Sub BalanceLR(ByVal lC As Long)
    Dim lP As Long
    Dim lA As Long
    Dim lB As Long
    Dim lT3 As Long
    Dim lT4 As Long

    'Indices of old state
    lP = vBinTree(lC).Parent
    lA = vBinTree(lC).LeftTree
    lB = vBinTree(lA).RightTree
    lT3 = vBinTree(lB).LeftTree
    lT4 = vBinTree(lB).RightTree

    'New links
    vBinTree(lA).Parent = lB
    vBinTree(lA).RightTree = lT3
    If lT3 <> sstNoChild Then vBinTree(lT3).Parent = lA
    vBinTree(lC).Parent = lB
    vBinTree(lC).LeftTree = lT4
    If lT4 <> sstNoChild Then vBinTree(lT4).Parent = lC
    vBinTree(lB).Parent = lP
    vBinTree(lB).LeftTree = lA
    vBinTree(lB).RightTree = lC

    Select Case True
      Case lP = sstNoParent 'If C was the whole trees root,
         lBinTreeRoot = lB  'then its role is taken over by B now.
      Case vBinTree(lP).LeftTree = lC
         vBinTree(lP).LeftTree = lB 'C was either P's left or right child
      Case Else                     'B is taking that place now.
         vBinTree(lP).RightTree = lB
    End Select

    'New weights
    vBinTree(lA).RightHeight = vBinTree(lB).LeftHeight
    vBinTree(lB).LeftHeight = vBinTree(lA).LeftHeight + 1
    vBinTree(lC).LeftHeight = vBinTree(lB).RightHeight
    vBinTree(lB).RightHeight = vBinTree(lC).RightHeight + 1
End Sub
'
'RL - A is the node that the rotation is performed on. This rotation is performed
'when A is unbalanced to the right (the right subtree is 2 higher than the left
'subtree), C is leftheavy (the left subtree of C is 1 higher than the right
'subtree of C). T1, T2, T3, and T4 represent subtrees (a node was added
'to B's branch which made B heavy, made C leftheavy and unbalanced A).
'This consists of a single right at node C, followed by a single left at node A.
'P is A's parent and listed only because A's relationship with P is taken over
'by B, making relinks neccessary.
'
'        P                      P                      P
'        |                      |                      |
'        A                      A                      B
'       / \    --------->     /   \    --------->    /   \
'      T1  C                 T1    B                A     C
'     /   / \               /     / \              / \   / \
'    T2  B  T5             T2    T3  C            T1 T3 T4 T5
'       / \   \                 /\  / \          /   /\ /\   \
'      T3 T4  T6                *  T4 T5        T2     *     T6
'      /\ /\                       /\   \
'        *                         *    T6
'+--+--------+--------+
'|  | Before | After  |
'+--+--+--+--+--+--+--+
'|  |LT|RT|PA|LT|RT|PA|
'+--+--+--+--+--+--+--+
'|A |  |C |P |  |T3|B |
'|B |T3|T4|C |A |C |P |
'|C |B |  |A |T4|  |B |
'|T3|  |  |B |  |  |A |
'|T4|  |  |B |  |  |C |
'|P |  A  |  |  B  |  |
'+--+--+--+--+--+--+--+
'
Private Sub BalanceRL(ByVal lA As Long)
    Dim lP As Long
    Dim lB As Long
    Dim lC As Long
    Dim lT3 As Long
    Dim lT4 As Long

    'Indices of old state
    lP = vBinTree(lA).Parent
    lC = vBinTree(lA).RightTree
    lB = vBinTree(lC).LeftTree
    lT3 = vBinTree(lB).LeftTree
    lT4 = vBinTree(lB).RightTree

    'New links
    vBinTree(lA).Parent = lB
    vBinTree(lA).RightTree = lT3
    If lT3 <> sstNoChild Then vBinTree(lT3).Parent = lA
    vBinTree(lB).Parent = lP
    vBinTree(lB).LeftTree = lA
    vBinTree(lB).RightTree = lC
    vBinTree(lC).Parent = lB
    vBinTree(lC).LeftTree = lT4
    If lT4 <> sstNoChild Then vBinTree(lT4).Parent = lC

    Select Case True
      Case lP = sstNoParent 'If A was the whole trees root,
         lBinTreeRoot = lB  'then its role is taken over by B now.
      Case vBinTree(lP).LeftTree = lA
         vBinTree(lP).LeftTree = lB 'A was either P's left or right child
      Case Else                     'B is taking that place now.
         vBinTree(lP).RightTree = lB
    End Select

    'New weights
    vBinTree(lA).RightHeight = vBinTree(lB).LeftHeight
    vBinTree(lB).LeftHeight = vBinTree(lA).LeftHeight + 1
    vBinTree(lC).LeftHeight = vBinTree(lB).RightHeight
    vBinTree(lB).RightHeight = vBinTree(lC).RightHeight + 1
End Sub


' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Function TextComp(ByVal StrPtr_ThisStr As Long, ByVal StrPtr_ThanStr As Long, Optional ByVal iChrPos As Long) As Long
    ThisStr.pvData = StrPtr_ThisStr ' Returns Lesser = -1&, Equal = 0& or Greater = 1&
    ThanStr.pvData = StrPtr_ThanStr
    Select Case lKeyComp
      Case 0:
         Do While ThanChrs(iChrPos) And (ThanChrs(iChrPos) = ThisChrs(iChrPos))
             iChrPos = iChrPos + 1
         Loop
         TextComp = Sgn(ThisChrs(iChrPos) - ThanChrs(iChrPos))
      Case 1:
         Do While ThanChrs(iChrPos) And (iaLCase(ThanChrs(iChrPos)) = iaLCase(ThisChrs(iChrPos)))
             iChrPos = iChrPos + 1
         Loop
         TextComp = Sgn(iaLCase(ThisChrs(iChrPos)) - iaLCase(ThanChrs(iChrPos)))
    End Select
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Function ArrayPtr(Arr) As Long
    Dim iDataType As Long
    On Error GoTo UnInit
    CopyMemory iDataType, Arr, 2&                       ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then           ' if a valid array was passed
        CopyMemory ArrayPtr, ByVal VarPtr(Arr) + 8&, 4& ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array. Thanks to Francesco Balena.
    End If
UnInit:
End Function

' ¤¤ InitLCaseLookupTable ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Initializes the unicode lowercase lookup table to
'  enable in-line case-insensitive textual comparison.
'
'  CharLowerBuffW piArr(lb), iLen    '0.0009881144111891 secs
'  piArr(i) = LukeLCaseI             '0.0009657652020019 secs
'  piArr(i) = AscW(LCase$(ChrW$(i))) '0.0092841408614782 secs
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Sub InitLCaseLookupTable(piArr() As Long)
    Dim i As Long, j As Long
    ReDim piArr(-32768 To 32767) As Long '65535
    For i = -32768 To -224
       piArr(i) = i
    Next
    For i = -223 To 9423
      'piArr(i) = AscW(LCase$(ChrW$(i)))
       j = i
       LukeLCaseI j          'For i = 32768 To 65535
       piArr(i) = j          '   j = i               'For i = 32768 To 65535
    Next                     '   LukeLCaseI j        '   piArr(i) = piArr(i - 65536)
    For i = 9424 To 32767    '   piArr(i) = j        'Next
       piArr(i) = i          'Next
    Next
    ' StrComp compatibility ("_", "@", etc, list before number "0")
    For i = 15 To 31
        piArr(i) = 14&     '14 (ctrl chars)
    Next
    For i = 32 To 47
        piArr(i) = i - 17& '15...30
    Next
    For i = 58 To 64
        piArr(i) = i - 27& '31...37
    Next
    For i = 91 To 96
        piArr(i) = i - 53& '38...43
    Next
    For i = 123 To 126
        piArr(i) = i - 79& '44...47
    Next
End Sub

' ¤¤ LukeLCaseI ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Unicode LCase Conversion
'
'  Sub to enable case insensitive comparison of unicode strings
'  as integer data. Substitutes AscW(LCase$(ChrW$(i))) without
'  any type conversion. With native code, it is extremely fast.
'
'  By: Adapted directly from CharLower by selftaught
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Sub LukeLCaseI(ByRef piChar As Long)

    Const MakeLCase As Long = 32

    Const Neg223 As Long = -223
    Const Neg197 As Long = -197

    Const Zero As Long = 0
    Const One As Long = 1
    Const Two As Long = 2
    Const Eight As Long = 8

    Const n16 As Long = 16
    Const n26 As Long = 26
    Const n37 As Long = 37
    Const n48 As Long = 48
    Const n63 As Long = 63
    Const n65 As Long = 65
    Const n74 As Long = 74
    Const n80 As Long = 80
    Const n86 As Long = 86
    Const n91 As Long = 91
    Const n100 As Long = 100
    Const n112 As Long = 112
    Const n192 As Long = 192
    Const n205 As Long = 205
    Const n215 As Long = 215
    Const n216 As Long = 216
    Const n217 As Long = 217
    Const n218 As Long = 218
    Const n223 As Long = 223
    Const n126 As Long = 126
    Const n128 As Long = 128
    Const n255 As Long = 255
    Const n256 As Long = 256
    Const n304 As Long = 304
    Const n311 As Long = 311
    Const n313 As Long = 313
    Const n328 As Long = 328
    Const n330 As Long = 330
    Const n375 As Long = 375
    Const n376 As Long = 376
    Const n377 As Long = 377
    Const n379 As Long = 379
    Const n381 As Long = 381
    Const n385 As Long = 385
    Const n386 As Long = 386
    Const n388 As Long = 388
    Const n390 As Long = 390
    Const n391 As Long = 391
    Const n392 As Long = 392
    Const n393 As Long = 393
    Const n394 As Long = 394
    Const n395 As Long = 395
    Const n396 As Long = 396
    Const n398 As Long = 398
    Const n399 As Long = 399
    Const n400 As Long = 400
    Const n401 As Long = 401
    Const n402 As Long = 402
    Const n403 As Long = 403
    Const n404 As Long = 404
    Const n406 As Long = 406
    Const n407 As Long = 407
    Const n408 As Long = 408
    Const n409 As Long = 409
    Const n412 As Long = 412
    Const n413 As Long = 413
    Const n415 As Long = 415
    Const n416 As Long = 416
    Const n418 As Long = 418
    Const n420 As Long = 420
    Const n423 As Long = 423
    Const n425 As Long = 425
    Const n428 As Long = 428
    Const n430 As Long = 430
    Const n431 As Long = 431
    Const n433 As Long = 433
    Const n434 As Long = 434
    Const n435 As Long = 435
    Const n437 As Long = 437
    Const n439 As Long = 439
    Const n440 As Long = 440
    Const n444 As Long = 444
    Const n452 As Long = 452
    Const n455 As Long = 455
    Const n458 As Long = 458
    Const n461 As Long = 461
    Const n476 As Long = 476
    Const n477 As Long = 477
    Const n495 As Long = 495
    Const n497 As Long = 497
    Const n499 As Long = 499
    Const n500 As Long = 500
    Const n501 As Long = 501
    Const n506 As Long = 506
    Const n535 As Long = 535
    Const n595 As Long = 595
    Const n596 As Long = 596
    Const n601 As Long = 601
    Const n603 As Long = 603
    Const n608 As Long = 608
    Const n611 As Long = 611
    Const n616 As Long = 616
    Const n617 As Long = 617
    Const n623 As Long = 623
    Const n626 As Long = 626
    Const n629 As Long = 629
    Const n658 As Long = 658
    Const n902 As Long = 902
    Const n904 As Long = 904
    Const n907 As Long = 907
    Const n908 As Long = 908
    Const n910 As Long = 910
    Const n911 As Long = 911
    Const n913 As Long = 913
    Const n930 As Long = 930
    Const n940 As Long = 940
    Const n972 As Long = 972
    Const n994 As Long = 994
    Const n1007 As Long = 1007
    Const n1025 As Long = 1025
    Const n1037 As Long = 1037
    Const n1040 As Long = 1040
    Const n1072 As Long = 1072
    Const n1120 As Long = 1120
    Const n1153 As Long = 1153
    Const n1168 As Long = 1168
    Const n1215 As Long = 1215
    Const n1217 As Long = 1217
    Const n1219 As Long = 1219
    Const n1223 As Long = 1223
    Const n1227 As Long = 1227
    Const n1232 As Long = 1232
    Const n1259 As Long = 1259
    Const n1262 As Long = 1262
    Const n1269 As Long = 1269
    Const n1272 As Long = 1272
    Const n1273 As Long = 1273
    Const n1329 As Long = 1329
    Const n1366 As Long = 1366
    Const n4256 As Long = 4256
    Const n4294 As Long = 4294
    Const n7680 As Long = 7680
    Const n7829 As Long = 7829
    Const n7840 As Long = 7840
    Const n7929 As Long = 7929
    Const n7944 As Long = 7944
    Const n7952 As Long = 7952
    Const n7960 As Long = 7960
    Const n7966 As Long = 7966
    Const n7976 As Long = 7976
    Const n7984 As Long = 7984
    Const n7992 As Long = 7992
    Const n8000 As Long = 8000
    Const n8008 As Long = 8008
    Const n8014 As Long = 8014
    Const n8025 As Long = 8025
    Const n8032 As Long = 8032
    Const n8040 As Long = 8040
    Const n8048 As Long = 8048
    Const n8120 As Long = 8120
    Const n8121 As Long = 8121
    Const n8122 As Long = 8122
    Const n8123 As Long = 8123
    Const n8136 As Long = 8136
    Const n8140 As Long = 8140
    Const n8152 As Long = 8152
    Const n8153 As Long = 8153
    Const n8154 As Long = 8154
    Const n8155 As Long = 8155
    Const n8165 As Long = 8165
    Const n8168 As Long = 8168
    Const n8169 As Long = 8169
    Const n8170 As Long = 8170
    Const n8171 As Long = 8171
    Const n8172 As Long = 8172
    Const n8184 As Long = 8184
    Const n8185 As Long = 8185
    Const n8186 As Long = 8186
    Const n8187 As Long = 8187
    Const n8544 As Long = 8544
    Const n8560 As Long = 8560
    Const n9397 As Long = 9397
    Const n9423 As Long = 9423
    Const n32768 As Long = 32768
    Const n65536 As Long = 65536

    If piChar < Neg223 Then ' -32768 .. -222
    ElseIf piChar > n9423 Then ' 9424..(65535)
        If piChar < n32768 Then ' 9424..32767
        Else 'piChar > n32767  ' 32768..65535
            Dim i As Long
            i = piChar - n65536
            If i >= Neg223 And i <= n9423 Then
                LukeLCaseI i
            End If
            piChar = i
        End If
    ElseIf piChar < n4294 Then '..4293
        If piChar > n1366 And piChar < n4256 Then ' 1367..4255
        ElseIf piChar < n416 Then '..415
            If piChar < n192 Then '..191
                If piChar < Neg197 Then
                    piChar = piChar + MakeLCase
                ElseIf piChar < n65 Then
                ElseIf piChar < n91 Then
                    piChar = piChar + MakeLCase
                End If
            ElseIf piChar < n215 Then
                piChar = piChar + MakeLCase
            ElseIf piChar < n216 Then
            ElseIf piChar < n223 Then ' 216..222
                piChar = piChar + MakeLCase
            ElseIf piChar < n256 Then
            ElseIf piChar < n304 Then ' 256..303
                If piChar Mod Two = Zero Then piChar = piChar + One
            ElseIf piChar = n304 Then
            ElseIf piChar < n311 Then ' 305..310
                If piChar Mod Two = Zero Then piChar = piChar + One
            ElseIf piChar < n313 Then
            ElseIf piChar < n328 Then ' 313..327
                If piChar Mod Two Then piChar = piChar + One
            ElseIf piChar < n330 Then
            ElseIf piChar < n375 Then ' 330..374
                If piChar Mod Two = Zero Then piChar = piChar + One
            ElseIf piChar = n375 Then
            ElseIf piChar = n376 Then
                piChar = n255
            ElseIf piChar = n377 Or piChar = n379 Or piChar = n381 Then
                piChar = piChar + One
            ElseIf piChar < n385 Then
            ElseIf piChar = n385 Then
                piChar = n595
            ElseIf piChar = n386 Or piChar = n388 Then
                piChar = piChar + One
            ElseIf piChar = n390 Then
                piChar = n596
            ElseIf piChar = n391 Then
                piChar = n392
            ElseIf piChar = n393 Or piChar = n394 Then
                piChar = piChar + n205
            ElseIf piChar = n395 Then
                piChar = n396
            ElseIf piChar < n398 Then
            ElseIf piChar = n398 Then
                piChar = n477
            ElseIf piChar = n399 Then
                piChar = n601
            ElseIf piChar = n400 Then
                piChar = n603
            ElseIf piChar = n401 Then
                piChar = n402
            ElseIf piChar = n403 Then
                piChar = n608
            ElseIf piChar = n404 Then
                piChar = n611
            ElseIf piChar = n406 Then
                piChar = n617
            ElseIf piChar = n407 Then
                piChar = n616
            ElseIf piChar = n408 Then
                piChar = n409
            ElseIf piChar < n412 Then
            ElseIf piChar = n412 Then
                piChar = n623
            ElseIf piChar = n413 Then
                piChar = n626
            ElseIf piChar = n415 Then
                piChar = n629
            End If
        ElseIf piChar < n902 Then
            If piChar < n535 Then '..534
                If piChar = n416 Or piChar = n418 Or piChar = n420 Or piChar = n423 Then
                    piChar = piChar + One
                ElseIf piChar < n425 Then
                ElseIf piChar = n425 Or piChar = n430 Then
                    piChar = piChar + n218
                ElseIf piChar = n428 Or piChar = n431 Or piChar = n435 Or piChar = n437 Then
                    piChar = piChar + One
                ElseIf piChar = n433 Or piChar = n434 Then
                    piChar = piChar + n217
                ElseIf piChar = n439 Then
                    piChar = n658
                ElseIf piChar = n440 Or piChar = n444 Then
                    piChar = piChar + One
                ElseIf piChar < n452 Then
                ElseIf piChar = n452 Or piChar = n455 Or piChar = n458 Then
                    piChar = piChar + Two
                ElseIf piChar < n461 Then
                ElseIf piChar < n476 Then ' 461..475
                    If piChar Mod Two Then piChar = piChar + One
                ElseIf piChar <= n477 Then
                ElseIf piChar < n495 Then ' 478..494
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar = n497 Then
                    piChar = n499
                ElseIf piChar = n500 Then
                    piChar = n501
                ElseIf piChar < n506 Then
                Else 'If piChar < n535 Then ' 506..534
                    If piChar Mod Two = Zero Then piChar = piChar + One
                End If
            End If
        ElseIf piChar < n1329 Then '..1328
            If piChar < n1273 Then '..1272
                If piChar = n902 Then
                    piChar = n940
                ElseIf piChar < n904 Then
                ElseIf piChar < n907 Then ' 904..906
                    piChar = piChar + n37
                ElseIf piChar = n908 Then
                    piChar = n972
                ElseIf piChar = n910 Or piChar = n911 Then
                    piChar = piChar + n63
                ElseIf piChar < n913 Then
                ElseIf piChar = n930 Then
                ElseIf piChar < n940 Then ' 913..939
                    piChar = piChar + MakeLCase
                ElseIf piChar < n994 Then
                ElseIf piChar < n1007 Then ' 994..1006
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar < n1025 Then
                ElseIf piChar = n1037 Then
                ElseIf piChar < n1040 Then ' 1025..1039
                    piChar = piChar + n80
                ElseIf piChar < n1072 Then ' 1040..1071
                    piChar = piChar + MakeLCase
                ElseIf piChar < n1120 Then
                ElseIf piChar < n1153 Then ' 1120..1152
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar < n1168 Then
                ElseIf piChar < n1215 Then ' 1168..1214
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar = n1217 Or piChar = n1219 Or piChar = n1223 Or piChar = n1227 Then
                    piChar = piChar + One
                ElseIf piChar < n1232 Then
                ElseIf piChar < n1259 Then ' 1232..1258
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar < n1262 Then
                ElseIf piChar < n1269 Then ' 1262..1268
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar = n1272 Then
                    piChar = n1273
                End If
            End If
        Else ' 1329..1366 | 4256..4293
            piChar = piChar + n48
        End If
    ElseIf piChar < n8048 Then ' 4294..8047
        If piChar < n7680 Then
        ElseIf piChar < n7829 Then ' 7680..7828
            If piChar Mod Two = Zero Then piChar = piChar + One
        ElseIf piChar < n7840 Then
        ElseIf piChar < n7929 Then ' 7840..7928
            If piChar Mod Two = Zero Then piChar = piChar + One
        ElseIf piChar < n7944 Then
        ElseIf piChar < n7952 Then ' 7944..7951
            piChar = piChar - Eight
        ElseIf piChar < n7960 Then
        ElseIf piChar < n7966 Then ' 7960..7965
            piChar = piChar - Eight
        ElseIf piChar < n7976 Then
        ElseIf piChar < n7984 Then ' 7976..7983
            piChar = piChar - Eight
        ElseIf piChar < n7992 Then
        ElseIf piChar < n8000 Then ' 7992..7999
            piChar = piChar - Eight
        ElseIf piChar < n8008 Then
        ElseIf piChar < n8014 Then ' 8008..8013
            piChar = piChar - Eight
        ElseIf piChar < n8025 Then
        ElseIf piChar < n8032 Then ' 8025..8031
            If piChar Mod Two Then piChar = piChar - Eight
        ElseIf piChar < n8040 Then
        ElseIf piChar < n8048 Then ' 8040..8047
            piChar = piChar - Eight
        End If
   ElseIf piChar < n8120 Then '..8119
   Else ' piChar < n9424 Then '..9423
        If piChar = n8120 Or piChar = n8121 Then
            piChar = piChar - Eight
        ElseIf piChar = n8122 Or piChar = n8123 Then
            piChar = piChar - n74
        ElseIf piChar < n8136 Then
        ElseIf piChar < n8140 Then ' 8136..8139
            piChar = piChar - n86
        ElseIf piChar < n8152 Then
        ElseIf piChar = n8152 Or piChar = n8153 Then
            piChar = piChar - Eight
        ElseIf piChar = n8154 Or piChar = n8155 Then
            piChar = piChar - n100
        ElseIf piChar < n8168 Then
        ElseIf piChar = n8168 Or piChar = n8169 Then
            piChar = piChar - Eight
        ElseIf piChar = n8170 Or piChar = n8171 Then
            piChar = piChar - n112
        ElseIf piChar = n8172 Then
            piChar = n8165
        ElseIf piChar < n8184 Then
        ElseIf piChar = n8184 Or piChar = n8185 Then
            piChar = piChar - n128
        ElseIf piChar = n8186 Or piChar = n8187 Then
            piChar = piChar - n126
        ElseIf piChar < n8544 Then
        ElseIf piChar < n8560 Then ' 8544..8559
            piChar = piChar + n16
        ElseIf n9397 < piChar Then 'And piChar < n9424
            piChar = piChar + n26
        End If
    End If
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

'Public Function IsGoodLCaseI() As Boolean
'    Dim i As Long
'    Dim iConv As Long
'    For i = -32768 To 65535
'        iConv = i
'        LukeLCaseI iConv
'        IsGoodLCaseI = AscW(LCase$(ChrW$(i))) = iConv
'        Debug.Assert IsGoodLCaseI
'        If Not IsGoodLCaseI Then Exit Function
'    Next
'End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
