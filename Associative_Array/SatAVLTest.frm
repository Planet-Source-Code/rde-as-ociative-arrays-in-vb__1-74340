VERSION 5.00
Begin VB.Form SatAVLTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SatAVL Test and comparison with Collection"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   525
      Left            =   90
      TabIndex        =   39
      Top             =   2100
      Width           =   5955
      Begin VB.Label lblStatus 
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test results"
      Height          =   3075
      Left            =   90
      TabIndex        =   12
      Top             =   2730
      Width           =   5925
      Begin VB.Label AVLDestroy 
         Caption         =   "0"
         Height          =   255
         Left            =   3675
         TabIndex        =   38
         Top             =   2760
         Width           =   1665
      End
      Begin VB.Label CollDestroy 
         Caption         =   "0"
         Height          =   255
         Left            =   1770
         TabIndex        =   37
         Top             =   2760
         Width           =   1665
      End
      Begin VB.Label Label28 
         Caption         =   "Destroy"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label AVLSeqData 
         Caption         =   "0"
         Height          =   255
         Left            =   3675
         TabIndex        =   35
         Top             =   2430
         Width           =   1665
      End
      Begin VB.Label Label26 
         Caption         =   "(omitted, too slow)"
         Height          =   255
         Left            =   1770
         TabIndex        =   34
         Top             =   2430
         Width           =   1665
      End
      Begin VB.Label Label25 
         Caption         =   "Sequential, Data"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label AVLSeqKey 
         Caption         =   "0"
         Height          =   255
         Left            =   3675
         TabIndex        =   32
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label Label23 
         Caption         =   "n.a."
         Height          =   255
         Left            =   1770
         TabIndex        =   31
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label Label22 
         Caption         =   "Sequential, Key"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Label AVLDown 
         Caption         =   "0"
         Height          =   255
         Left            =   3675
         TabIndex        =   29
         Top             =   1770
         Width           =   1665
      End
      Begin VB.Label Label20 
         Caption         =   "n.a."
         Height          =   255
         Left            =   1770
         TabIndex        =   28
         Top             =   1770
         Width           =   1665
      End
      Begin VB.Label Label19 
         Caption         =   "Highest/Lower"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   1770
         Width           =   1575
      End
      Begin VB.Label AVLUp 
         Caption         =   "0"
         Height          =   255
         Left            =   3675
         TabIndex        =   26
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label17 
         Caption         =   "n.a."
         Height          =   255
         Left            =   1770
         TabIndex        =   25
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label16 
         Caption         =   "Lowest/Higher"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label AVLDirDesc 
         Caption         =   "0"
         Height          =   255
         Left            =   3675
         TabIndex        =   23
         Top             =   1110
         Width           =   1665
      End
      Begin VB.Label CollDirDesc 
         Caption         =   "0"
         Height          =   255
         Left            =   1770
         TabIndex        =   22
         Top             =   1110
         Width           =   1665
      End
      Begin VB.Label Label13 
         Caption         =   "Direct descending"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label AVLDirAsc 
         Caption         =   "0"
         Height          =   255
         Left            =   3675
         TabIndex        =   20
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label CollDirAsc 
         Caption         =   "0"
         Height          =   255
         Left            =   1770
         TabIndex        =   19
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label Label10 
         Caption         =   "Direct ascending"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label AVLAdd 
         Caption         =   "0"
         Height          =   255
         Left            =   3675
         TabIndex        =   17
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label CollAdd 
         Caption         =   "0"
         Height          =   255
         Left            =   1770
         TabIndex        =   16
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label Label7 
         Caption         =   "Add"
         Height          =   165
         Left            =   180
         TabIndex        =   15
         Top             =   510
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "SatAVL w/ DLList TextComp"
         Height          =   225
         Left            =   3675
         TabIndex        =   14
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "Collection"
         Height          =   225
         Left            =   1770
         TabIndex        =   13
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mass test settings"
      Height          =   1935
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5925
      Begin VB.CommandButton cmdForEach 
         Caption         =   "&For Each"
         Height          =   315
         Left            =   4605
         TabIndex        =   43
         Top             =   1440
         Width           =   1125
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Delete Test"
         Height          =   315
         Left            =   4605
         TabIndex        =   42
         Top             =   1065
         Width           =   1125
      End
      Begin VB.CheckBox chkTxtCmp 
         Caption         =   "&Text Compare"
         Height          =   195
         Left            =   4215
         TabIndex        =   41
         Top             =   330
         Width           =   1440
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start Test"
         Height          =   315
         Left            =   4605
         TabIndex        =   11
         Top             =   690
         Width           =   1125
      End
      Begin VB.CheckBox chkDataLen 
         Caption         =   "Random 1-1024"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   1470
         Width           =   1545
      End
      Begin VB.TextBox txtDataLen 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1380
         TabIndex        =   9
         Text            =   "32"
         Top             =   1440
         Width           =   1005
      End
      Begin VB.CheckBox chkKeySorted 
         Caption         =   "Sorted ascending (unchecked=random)"
         Height          =   285
         Left            =   1380
         TabIndex        =   7
         Top             =   690
         Width           =   3255
      End
      Begin VB.CheckBox chkKeyLen 
         Caption         =   "Random 8-24"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txtKeyLen 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1380
         TabIndex        =   4
         Text            =   "8"
         Top             =   1050
         Width           =   1005
      End
      Begin VB.TextBox txtNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1380
         TabIndex        =   2
         Text            =   "100000"
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Data length"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "Key order"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Key length"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Elements"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   1125
      End
   End
End
Attribute VB_Name = "SatAVLTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oColl As Collection
Dim oAArr As cAssocArr

Private Declare Sub QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency)
Private Declare Sub QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency)

Private fTxtCmpFlag As Boolean

Private Sub chkTxtCmp_Click()
   fTxtCmpFlag = chkTxtCmp.Value
End Sub

Private Sub chkDataLen_Click()
    If chkDataLen.Value Then
        txtDataLen.Enabled = False
    Else
        txtDataLen.Enabled = True
    End If
End Sub

Private Sub chkKeyLen_Click()
    If chkKeyLen.Value Then
        txtKeyLen.Enabled = False
    Else
        txtKeyLen.Enabled = True
    End If
End Sub

Private Sub chkKeySorted_Click()
    If chkKeySorted.Value Then
        txtKeyLen.Enabled = False
        chkKeyLen.Enabled = False
    Else
        txtKeyLen.Enabled = True
        chkKeyLen.Enabled = True
    End If
End Sub

Private Sub cmdForEach_Click()

   ' For Each 2 nesting example
   Dim sMsg As String
   Dim cAA As cAssocArr
   Set cAA = New cAssocArr

   cAA("A") = "Item1"
   cAA("B") = "Item2"
   cAA("C") = "Item3"
   cAA("D") = "Item4"

   Dim v1, v2
   For Each v1 In cAA
      sMsg = sMsg & v1 & vbCr
      For Each v2 In cAA
         sMsg = sMsg & "   " & v2 & vbCr
         If v2 = "Item2" Then
            cAA.Remove "b"
            sMsg = sMsg & "   Del " & v2 & vbCr
         End If
      Next
   Next
   Set cAA = Nothing
   MsgBox sMsg
End Sub

Private Sub cmdStart_Click()
    Const Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim lNum As Long
    Dim lCnt As Long
    Dim sAllKeys() As String
    Dim lK1 As Long, lK2 As Long, lK3 As Long, lK4 As Long, lK5 As Long
    Dim lKeyLen As Long
    Dim sKey As String
    Dim lKeyPos As Long
    Dim lDataLen As Long
    Dim bRndDataLen As Boolean
    Dim timStart As Double
    Dim timEnd As Double
    Dim timDuration As Double
    Dim sData As String
    Dim lId As Long

    'Get data from form
    lNum = CLng(txtNum.Text)
    If lNum > 2000000 Then
        'In VB, we can't allocate arrays with more than 2 mill elements
        MsgBox "Array size max 2 million under VB"
        Exit Sub
    ElseIf lNum < 1 Then
        MsgBox "Array size not specified"
        Exit Sub
    End If
    ReDim sAllKeys(1 To lNum) As String
    
    cmdStart.Enabled = False
    cmdStart.Refresh
    
    'First the key generation overhead (outside of time measurement): Same keys
    'for Collection and SatAVL.
    lblStatus = "Generating and storing keys..."
    lblStatus.Refresh
    If chkKeySorted Then
        For lK1 = 1 To 26: For lK2 = 1 To 26: For lK3 = 1 To 26: For lK4 = 1 To 26: For lK5 = 1 To 26
            sKey = Mid$(Alpha, lK1, 1) & Mid$(Alpha, lK2, 1) & Mid$(Alpha, lK3, 1) & Mid$(Alpha, lK4, 1) & Mid$(Alpha, lK5, 1)
            lCnt = lCnt + 1: sAllKeys(lCnt) = sKey
            If lCnt = lNum Then GoTo KeysGenerated
        Next: Next: Next: Next: Next
KeysGenerated:
    Else
        lKeyLen = CLng(txtKeyLen.Text)
        For lCnt = 1 To lNum
            If chkKeyLen.Value Then lKeyLen = Int(Rnd() * 17) + 8
            'Generate random key
            sKey = Space$(lKeyLen)
            For lKeyPos = 1 To lKeyLen
                Mid$(sKey, lKeyPos, 1) = Mid$(Alpha, Int(Rnd() * 26 + 1), 1)
            Next lKeyPos
            'Storing the generated key to have something to retrieve
            sAllKeys(lCnt) = sKey
        Next lCnt
    End If
    
    lDataLen = CLng(txtDataLen.Text)
    bRndDataLen = (chkDataLen.Value = 1)
    If bRndDataLen Then
       sData = Space$(1024&)
    Else
       sData = Space$(lDataLen)
    End If
    'ALL DATA READY FOR TESTS (except data length creation when random)
    
    
    'C O L L E C T I O N
    Set oColl = New Collection
    
    'Add
    lblStatus = "Collection.Add..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        'Calculate random data length if wanted
        If bRndDataLen Then lDataLen = Rnd() * 1024& + 1&
        
        'Add the item as per key array, reserving data space as required
        oColl.Add Left$(sData, lDataLen), sAllKeys(lCnt)
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    CollAdd.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    CollAdd.Refresh
    
    'Item (direct ascending)
    lblStatus = "Collection.Item (Direct Ascending)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        sData = oColl.Item(sAllKeys(lCnt))
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    CollDirAsc.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    CollDirAsc.Refresh
    
    'Item (direct descending)
    lblStatus = "Collection.Item (Direct Descending)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = lNum To 1 Step -1
        sData = oColl.Item(sAllKeys(lCnt))
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    CollDirDesc.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    CollDirDesc.Refresh
    
    'Destroy
    lblStatus = "Coll.Destroy..."
    lblStatus.Refresh
    timStart = QPTimer
    Set oColl = Nothing
    timEnd = QPTimer
    timDuration = timEnd - timStart
    CollDestroy.Caption = Format(timDuration, "0.000") & "s "
    CollDestroy.Refresh




    If bRndDataLen Then
       sData = Space$(1024&)
    Else
       sData = Space$(lDataLen)
    End If

    'S A T A V L 2 TextComp
    Set oAArr = New cAssocArr
   'oAArr.Clear True

    oAArr.Clear fTxtCmpFlag

    'Add
    lblStatus = "Associative.Add..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        'Calculate random data length if wanted
        If bRndDataLen Then lDataLen = Rnd() * 1024& + 1&

        'Add the item as per key array, reserving data space as required
        lId = oAArr.Insert(sAllKeys(lCnt), Left$(sData, lDataLen))
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLAdd.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLAdd.Refresh
    
    'Item (direct ascending)
    lblStatus = "Associative.Item (Direct Ascending)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        sData = oAArr.Item(sAllKeys(lCnt))
        If sData = "" Then Stop
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDirAsc.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLDirAsc.Refresh
    
    'Item (direct descending)
    lblStatus = "Associative.Item (Direct Descending)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = lNum To 1 Step -1
        sData = oAArr.Item(sAllKeys(lCnt))
        If sData = "" Then Stop
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDirDesc.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLDirDesc.Refresh
    
    'Lowest/higher
    lblStatus = "Associative.Sorted (Lowest/Higher)..."
    lblStatus.Refresh
    timStart = QPTimer
    sData = oAArr.FindNext(True)
    Do While sData <> ""
        sData = oAArr.FindNext
    Loop
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLUp.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLUp.Refresh
    
    'Highest/Lower
    lblStatus = "Associative.Sorted (Highest/Lower)..."
    lblStatus.Refresh
    timStart = QPTimer
    sData = oAArr.FindPrev(True)
    Do While sData <> ""
        sData = oAArr.FindPrev
    Loop
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDown.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLDown.Refresh
    
    'Sequential (Key)
    lblStatus = "Associative.Sequential (Key)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 0 To lNum - 1
        sKey = oAArr.Key(lCnt)
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLSeqKey.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLSeqKey.Refresh
    
    'Sequential (Data)
    lblStatus = "Associative.Sequential (Data)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 0 To lNum - 1
        sData = oAArr.Data(lCnt)
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLSeqData.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLSeqData.Refresh

Call oAArr.Validate
Call oAArr.Remove(sAllKeys(lNum \ 2))
Call oAArr.Validate

    'Destroy
    lblStatus = "Associative.Destroy..."
    lblStatus.Refresh
    timStart = QPTimer
    Set oAArr = Nothing
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDestroy.Caption = Format(timDuration, "0.000") & "s "
    AVLDestroy.Refresh
    
    lblStatus = "Ready"
    cmdStart.Enabled = True
End Sub




Private Sub Form_Unload(Cancel As Integer)
    'Make sure we destroy the objects (redundant)
    Set oColl = Nothing
    Set oAArr = Nothing
End Sub


'Source for Timer: http://vb-tec.de/timer.htm (German)
Public Function QPTimer() As Double
  Static Takt As Currency
  Dim Dauer As Currency
  
  If Takt = 0 Then
    'einmal die Taktfrequenz bestimmen:
    QueryPerformanceFrequency Takt
  End If
  
  'aktuelle Zeit holen:
  QueryPerformanceCounter Dauer
  
  'Zeit in Sekunden umrechnen:
  QPTimer = Dauer / Takt
End Function



Private Sub cmdRemove_Click()
    Const Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim lNum As Long
    Dim lCnt As Long
    Dim sAllKeys() As String
    Dim lK1 As Long, lK2 As Long, lK3 As Long, lK4 As Long, lK5 As Long
    Dim lKeyLen As Long
    Dim sKey As String
    Dim lKeyPos As Long
    Dim lDataLen As Long
    Dim bRndDataLen As Boolean
    Dim timStart As Double
    Dim timEnd As Double
    Dim timDuration As Double
    Dim lAddr As Long
    Dim sData As String
    Dim lId As Long


    'Get data from form
    lNum = CLng(txtNum.Text)
    If lNum > 2000000 Then
        'In VB, we can't allocate arrays with more than 2 mill elements
        MsgBox "Array size max 2 million under VB"
        Exit Sub
    ElseIf lNum < 1 Then
        MsgBox "Array size not specified"
        Exit Sub
    End If
    ReDim sAllKeys(1 To lNum) As String
    
   
    'First the key generation overhead (outside of time measurement): Same keys
    'for Collection and SatAVL.
    lblStatus = "Generating and storing keys..."
    lblStatus.Refresh
    If chkKeySorted Then
        For lK1 = 1 To 26: For lK2 = 1 To 26: For lK3 = 1 To 26: For lK4 = 1 To 26: For lK5 = 1 To 26
            sKey = Mid$(Alpha, lK1, 1) & Mid$(Alpha, lK2, 1) & Mid$(Alpha, lK3, 1) & Mid$(Alpha, lK4, 1) & Mid$(Alpha, lK5, 1)
            lCnt = lCnt + 1: sAllKeys(lCnt) = sKey
            If lCnt = lNum Then GoTo KeysGenerated
        Next: Next: Next: Next: Next
KeysGenerated:
    Else
        lKeyLen = CLng(txtKeyLen.Text)
        For lCnt = 1 To lNum
            If chkKeyLen.Value Then lKeyLen = Int(Rnd() * 17) + 8
            'Generate random key
            sKey = Space$(lKeyLen)
            For lKeyPos = 1 To lKeyLen
                Mid$(sKey, lKeyPos, 1) = Mid$(Alpha, Int(Rnd() * 26 + 1), 1)
            Next lKeyPos
            'Storing the generated key to have something to retrieve
            sAllKeys(lCnt) = sKey
        Next lCnt
    End If
    
    bRndDataLen = (chkDataLen.Value = 1)
    lDataLen = CLng(txtDataLen.Text)
    'ALL DATA READY FOR TESTS (except data length creation when random)




    Set oAArr = New cAssocArr
   'oAArr.Clear True

    oAArr.Clear fTxtCmpFlag

    'Add
    lblStatus = "Associative.Add..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        'Calculate random data length if wanted
        If bRndDataLen Then lDataLen = Rnd() * 1024& + 1&

        'Add the item as per key array, reserving data space as required
        lId = oAArr.Insert(sAllKeys(lCnt), lDataLen)
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLAdd.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLAdd.Refresh

    'Remove
    lblStatus = "Associative.Remove..."

Call oAArr.Validate
oAArr.Remove oAArr.Key(lNum \ 2)
Call oAArr.Validate
oAArr.Remove oAArr.Key(lNum \ 3)
Call oAArr.Validate
oAArr.Remove oAArr.Key(lNum \ 4)
Call oAArr.Validate

    'Destroy
    lblStatus = "Associative.Destroy..."
    lblStatus.Refresh
    timStart = QPTimer
    Set oAArr = Nothing
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDestroy.Caption = Format(timDuration, "0.000") & "s "
    AVLDestroy.Refresh

    lblStatus = "Ready"
    cmdStart.Enabled = True

End Sub


