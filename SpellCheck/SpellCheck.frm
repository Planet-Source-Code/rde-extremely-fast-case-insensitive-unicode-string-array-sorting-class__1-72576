VERSION 5.00
Begin VB.Form frmSpellCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Spell Checker Using Soundex and Levinshtein Distance"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SpellCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   1065
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1065
   End
   Begin VB.CommandButton cmdCreateList 
      Caption         =   "Generate Word List"
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox lstLdMatches 
      Height          =   2205
      Left            =   2940
      TabIndex        =   8
      Top             =   180
      Width           =   2775
   End
   Begin VB.HScrollBar LdVal 
      Height          =   235
      Left            =   105
      TabIndex        =   4
      Top             =   1065
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox txtInputWord 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   60
      X2              =   2840
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   60
      X2              =   2840
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label lblLDMatchCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1470
      TabIndex        =   10
      Top             =   825
      Width           =   1185
   End
   Begin VB.Label lblLDVal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2340
      TabIndex        =   9
      Top             =   3705
      Width           =   375
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   2115
      Width           =   2550
   End
   Begin VB.Label lblAccuracy 
      Caption         =   "Adjust accuracy"
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblEnterWord 
      Caption         =   "Enter a word"
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form requires the following type library:

' Reference=*\G{C878CB53-7E75-4115-BD13-EECBC9430749}#1.0#0#MemAPIs.tlb#Memory APIs

' Or uncomment the following declares:

'Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal lByteLen As Long)
'Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
'Private Declare Function AllocStrSpPtr Lib "oleaut32" Alias "SysAllocStringLen" (ByVal lStrPtr As Long, ByVal lLen As Long) As Long

Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As VbFileAttribute
Private Declare Function GetInputState Lib "user32" () As Long

Private Const INVALID_FILE_ATTRIBUTES As Long = &HFFFFFFFF

Private Const SUBooRANGE As Long = &H9
Private Const MAX_PATH As Long = 260

Private aWords() As String
Private lA_sdx() As Long
Private lA_rev() As Long
Private lA_len() As Long

Private arrLDMatches() As Long
Private arrLDItemCnt() As Long

'Private strClosestMatch As String
Private sAppPathSlash As String

Private mTag As Long
Private mCnt As Long

Private Sub SortWordsFile()
    Dim cSort As cQS5_tcm
    lblStatus.Caption = " Sorting..."
    lblStatus.Refresh
    Set cSort = New cQS5_tcm
    cSort.SortMethod = TextCompare
    cSort.SortOrder = Ascending
    cSort.BlizzardStringSort aWords, 1&, mCnt, False
    Set cSort = Nothing
End Sub

Public Property Get CorrectedWord() As String
    CorrectedWord = txtInputWord
End Property

Public Property Let CorrectedWord(sSpellWord As String)
    txtInputWord = sSpellWord
'    If Clipboard.GetFormat(vbCFText) Then
'        Text1.SelText = Clipboard.GetText(vbCFText)
'    End If
End Property

Private Sub Form_Load()
    sAppPathSlash = App.Path
    If Right$(sAppPathSlash, 1) <> "\" Then sAppPathSlash = sAppPathSlash & "\"
    LdVal.Enabled = False
    Me.Show
    Me.Refresh
    If IsValidData Then
        txtInputWord.Visible = True
        txtInputWord.SetFocus
        lblEnterWord.Visible = True
        lblAccuracy.Visible = True
        LdVal.Visible = True
        lblStatus.Caption = " Preparing..."
        Me.Refresh
        GetDataFiles
        GetWordsFile
        lblStatus.Caption = " Ready..."
        lblStatus.Refresh
    Else
        cmdCreateList.Visible = True
        cmdCreateList.SetFocus
    End If
End Sub

Private Function IsValidData() As Boolean
    Dim lFileLen As Long ' Enable updating of new words file
    If FileExists(sAppPathSlash & "words.dat") Then
        If FileExists(sAppPathSlash & "sdx_r.dat") Then
            lFileLen = CLng(GetSetting(App.EXEName, "Words Data", "File Size", "-1"))
            IsValidData = (lFileLen = FileLen(sAppPathSlash & "words.dat"))
        End If
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Must Unload this form at program termination
    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub cmdClose_Click()
    If cmdClose.Caption = "Close" Then
        ' Must Unload this form at program termination
        Me.Hide

'MsgBox CorrectedWord
'Unload Me

    Else
        cmdClose.Caption = "Close"
        mTag = (mTag Mod MAX_PATH) + 1&

        lblStatus.Caption = " Processing stopped..."
        lblStatus.Refresh

    End If
End Sub

Private Sub cmdSelect_Click()
    ' Must Unload this form at program termination
    Me.Hide
    If lstLdMatches.SelCount Then txtInputWord = lstLdMatches.Text

'    Clipboard.Clear
'    Clipboard.SetText txtInputWord, vbCFText

'MsgBox CorrectedWord
'Unload Me
End Sub

Private Sub lstLdMatches_DblClick()
    ' Must Unload this form at program termination
    Me.Hide
    txtInputWord = lstLdMatches.Text 'lstLdMatches.List(lstLdMatches.ListIndex)

'    Clipboard.Clear
'    Clipboard.SetText txtInputWord, vbCFText

'MsgBox CorrectedWord
'Unload Me
End Sub

Private Sub cmdCreateList_Click()
    Dim lFileLen As Long
    On Error GoTo ErrHandler
    cmdCreateList.Enabled = False
    lblStatus.Caption = " Preparing..."
    lblStatus.Refresh
    mCnt = 0&
    GetWordsFile
    GenerateData
    SaveDataFiles
    lFileLen = FileLen(sAppPathSlash & "words.dat")
    SaveSetting App.EXEName, "Words Data", "File Size", CStr(lFileLen)
    cmdCreateList.Visible = False
    txtInputWord.Visible = True
    txtInputWord.SetFocus
    lblEnterWord.Visible = True
    lblAccuracy.Visible = True
    LdVal.Visible = True
ErrHandler:
End Sub

Private Sub txtInputWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0&
    ElseIf KeyAscii = vbKeySpace Then
        KeyAscii = 0&
        Beep
    End If
End Sub

Private Sub txtInputWord_Change()
    Dim strInput As String
    strInput = LCase$(Trim$(txtInputWord.Text))
    If LenB(strInput) = 0& Then Exit Sub
    mTag = (mTag Mod MAX_PATH) + 1&
    cmdClose.Caption = "Stop"
    Call GetMatches(strInput, mTag)
End Sub

Private Sub GenerateData()
    Dim tmpStr As String, i As Long
    Dim lpStr As Long, lp As Long
    Screen.MousePointer = vbHourglass

    ReDim lA_sdx(1& To mCnt) As Long
    ReDim lA_rev(1& To mCnt) As Long
    ReDim lA_len(1& To mCnt) As Long

    lpStr = VarPtr(tmpStr)       ' Cache pointer to the string variable
    lp = VarPtr(aWords(1&)) - 4& ' Cache pointer to the array (one based)

    For i = 1& To mCnt
        '// Read words from the array and add to the database
        'tmpStr = aWords(i) ' Grab current value into variable
        CopyMemByV lpStr, lp + (i * 4&), 4&

        lA_sdx(i) = GetSoundexWord(tmpStr)
        lA_rev(i) = GetSoundexWordR(tmpStr)
        lA_len(i) = Len(tmpStr)

        '// Prevent the UI from freezing up
        If i Mod 1000 = 0& Then
            lblStatus.Caption = " Adding..." & i & " words"
            lblStatus.Refresh
        End If
    Next i

    CopyMemByR ByVal lpStr, 0&, 4& ' De-reference pointer to variable

    lblStatus.Caption = " Database created - " & i - 1& & " words added"
    Screen.MousePointer = vbDefault
    Me.Refresh
End Sub

Private Function SaveTextFile(sFileSpec As String, sText As String) As Long
    On Error GoTo SaveFileError
    Dim iFile As Integer
    iFile = FreeFile
    Open sFileSpec For Output Access Write Lock Write As #iFile
      Print #iFile, sText;
SaveFileError:
    Close #iFile
    SaveTextFile = Err
End Function

Private Sub GetWordsFile()
    Dim iFile As Integer, numcount As Long
    Dim sFile As String, iLen As Long
    Dim iSubLen As Long, lCnt As Long
    Dim idx1 As Long, idx2 As Long
    Dim lA() As Long
    Const PBrk As String = vbCrLf & vbCrLf

    On Error GoTo ErrorHandler
    iFile = FreeFile
    Screen.MousePointer = vbHourglass

    ' Open in binary mode
    Open sAppPathSlash & "words.dat" For Binary Access Read Lock Write As #iFile

        ' Get the data length
        iLen = LOF(iFile)
        
        ' Allocate the length first: sFile = Space$(iLen)
       'CopyMemByV VarPtr(sFile), VarPtr(AllocStrSpPtr(0&, iLen)), 4&
        sFile = AllocStr(vbNullString, iLen)

        ' Get the file in one chunk
        Get #iFile, 1&, sFile
        ' sFile = File text converted to Unicode

    Close #iFile ' Close the file

    If Not mCnt = 0& Then
        ReDim aWords(1& To mCnt) As String
        idx1 = 1&
        ' Read words from the file and add to the array
        For numcount = 1& To mCnt
            iSubLen = lA_len(numcount)
           'CopyMemByV VarPtr(aWords(numcount)), VarPtr(AllocStrSpPtr(0&, iSubLen)), 4&
            aWords(numcount) = AllocStr(vbNullString, iSubLen)
            Mid$(aWords(numcount), 1&) = Mid$(sFile, idx1, iSubLen)
            idx1 = idx1 + iSubLen + 2&
        Next
    Else
        If (InStr(1&, sFile, vbCrLf) = 0&) Then ' Fix UNIX line ends
            sFile = Replace$(sFile, vbLf, vbCrLf)
        End If
        Do While Not (InStr(1&, sFile, PBrk) = 0&) ' Remove blank lines
            sFile = Replace$(sFile, PBrk, vbCrLf)
        Loop
        If Not (iLen = Len(sFile)) Then ' If file has changed
            SaveTextFile sAppPathSlash & "words.dat", sFile
            iLen = Len(sFile)
        End If
        lCnt = 1000000
        ReDim lA(1& To lCnt) As Long
        idx1 = 1&
        ' Index words from the file
        Do While (idx1 < iLen)
            idx2 = InStr(idx1, sFile, vbCrLf)
            If (idx2 = 0&) Then idx2 = iLen + 1&
            numcount = numcount + 1&
            lA(numcount) = idx2
            idx1 = idx2 + 2&
        Loop
        lCnt = numcount
        ReDim aWords(1& To lCnt) As String
        idx1 = 1&
        ' Read words from the file and add to the array
        For numcount = 1& To lCnt
            idx2 = lA(numcount)
            iSubLen = idx2 - idx1
           'CopyMemByV VarPtr(aWords(numcount)), VarPtr(AllocStrSpPtr(0&, iSubLen)), 4&
            aWords(numcount) = AllocStr(vbNullString, iSubLen)
            Mid$(aWords(numcount), 1&) = Mid$(sFile, idx1, iSubLen)
            idx1 = idx2 + 2&
        Next
        mCnt = lCnt
        SortWordsFile
        SaveTextFile sAppPathSlash & "words.dat", Join(aWords, vbCrLf)
    End If

ErrorHandler:
    Screen.MousePointer = vbDefault
    If Err = SUBooRANGE Then
        lCnt = lCnt + 1000000
        ReDim Preserve lA(1& To lCnt) As Long
        Resume
    ElseIf Err Then
        Close #iFile ' Close the file
        MsgBox "Error - " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub SaveDataFiles()
    On Error GoTo Abort
    Dim iFile As Integer

    Screen.MousePointer = vbHourglass

    iFile = FreeFile
    Open sAppPathSlash & "sdx_f.dat" For Binary Access Write Lock Write As #iFile
        Put #iFile, 1&, lA_sdx()
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "sdx_r.dat" For Binary Access Write Lock Write As #iFile
        Put #iFile, 1&, lA_rev()
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "len_w.dat" For Binary Access Write Lock Write As #iFile
        Put #iFile, 1&, lA_len()
Abort:
    Close #iFile
    Screen.MousePointer = vbDefault
End Sub

Private Sub GetDataFiles()
    On Error GoTo Abort
    Dim iFile As Integer
    Screen.MousePointer = vbHourglass

    mCnt = FileLen(sAppPathSlash & "len_w.dat") \ 4&

    ReDim lA_sdx(1& To mCnt) As Long
    ReDim lA_rev(1& To mCnt) As Long
    ReDim lA_len(1& To mCnt) As Long

    iFile = FreeFile
    Open sAppPathSlash & "sdx_f.dat" For Binary Access Read Lock Write As #iFile
        Get #iFile, 1&, lA_sdx()
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "sdx_r.dat" For Binary Access Read Lock Write As #iFile
        Get #iFile, 1&, lA_rev()
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "len_w.dat" For Binary Access Read Lock Write As #iFile
        Get #iFile, 1&, lA_len()
Abort:
    Close #iFile ' Close the file
    Screen.MousePointer = vbDefault
End Sub

Private Sub GetMatches(ByVal strInput As String, ByVal lTag As Long)
    Dim lMatches() As Long, strMatch As String
    Dim lSndex As Long, lRevSndex As Long
    Dim lenTmp As Long, LD As Long, LdMax As Long
    Dim Index As Long, Total As Long
    Dim lpStr As Long, lp As Long

    On Error GoTo ExitSub

    lpStr = VarPtr(strMatch)     ' Cache pointer to the string variable
    lp = VarPtr(aWords(1&)) - 4& ' Cache pointer to the array (one based)

    lblStatus.Caption = " Processing items..."
    lblStatus.Refresh

    '// Get the soundex of the input word
    lSndex = GetSoundexWord(strInput)

    '// Get the soundex of the input word reversed
    lRevSndex = GetSoundexWordR(strInput)

    If GetInputState <> 0& Then DoEvents
    If Not lTag = mTag Then GoTo ExitSub

    ReDim lMatches(1& To mCnt) As Long
    
    '// Find all entries in the database which match the soundex of the input word
    For Index = 1& To mCnt
        If GetInputState <> 0& Then DoEvents
        If Not lTag = mTag Then GoTo ExitSub

        If (lA_sdx(Index) = lSndex) Or (lA_rev(Index) = lRevSndex) Then
            Total = Total + 1&
            lMatches(Total) = Index

            '// The max length of returned matches is the
            '// maximum Leveshtein distance we can have
            lenTmp = lA_len(Index)
            If lenTmp > LdMax Then LdMax = lenTmp
        End If
    Next

    lstLdMatches.Clear
    lblLDMatchCount.Caption = "0"

    ReDim arrLDMatches(LdMax, Total) As Long
    ReDim arrLDItemCnt(LdMax) As Long
    LdMax = 0&

    '// Walk through all soundex matches
    For Index = 1& To Total
        If GetInputState <> 0& Then DoEvents
        If Not lTag = mTag Then GoTo ExitSub

        '// The one and only access to the strings
        'strMatch = aWords(lMatches(Index)) ' Grab current word
        CopyMemByV lpStr, lp + (lMatches(Index) * 4&), 4&

        '// Get all Levenshtein distances
        LD = GetLevenshteinDistance(strInput, strMatch)

        '// Add better matches up
        arrLDMatches(LD, arrLDItemCnt(LD)) = lMatches(Index)
        arrLDItemCnt(LD) = arrLDItemCnt(LD) + 1&

        '// Record maximum Leveshtein distance
        If LD > LdMax Then LdMax = LD
    Next

    If GetInputState <> 0& Then DoEvents
    If Not lTag = mTag Then GoTo ExitSub

    lblStatus.Caption = " " & Total & " items found"
    lblStatus.Refresh

    '// Determine lowest Levenshtein distance that produced matches
    For Index = 0& To LdMax
        If arrLDItemCnt(Index) <> 0& Then Exit For
    Next

'    If arrLDItemCnt(Index) = 1& Then
'        strClosestMatch = aWords(arrLDMatches(Index, 0&))
'    Else
'        strClosestMatch = vbNullString
'    End If

    With LdVal
        .Min = Index
        .Max = LdMax
         ' Call Change if Value will not change
        If .Value = Index Then LdVal_Change Else .Value = Index
        .Enabled = True
    End With

    cmdClose.Caption = "Close"
ExitSub:

    CopyMemByR ByVal lpStr, 0&, 4& ' De-reference pointer to variable
End Sub

Private Sub LdVal_Change()
    Dim i As Long, j As Long

    On Error GoTo ExitSub
    lstLdMatches.Clear

    '// Add all Levenshtein distances up to the scroll(threshold) value
    For i = 0& To LdVal.Value
        For j = 0& To arrLDItemCnt(i) - 1&
            lstLdMatches.AddItem aWords(arrLDMatches(i, j))
        Next
    Next

    lblLDMatchCount.Caption = lstLdMatches.ListCount & " matches"

    LdVal_Scroll
ExitSub:
End Sub

Private Sub LdVal_Scroll()
    lblLDVal.Caption = LdVal.Value
End Sub

Private Function FileExists(sFileSpec As String) As Boolean
    Dim Attribs As VbFileAttribute
    Attribs = GetAttributes(sFileSpec)
    If (Attribs <> INVALID_FILE_ATTRIBUTES) Then
        FileExists = ((Attribs And vbDirectory) <> vbDirectory)
    End If
End Function

'// Russell Soundex

'// From Wikipedia, the free encyclopedia

'// Soundex is a phonetic algorithm for indexing names by their
'// sound when pronounced in English.

'// Soundex is the most widely known of all phonetic algorithms and
'// is often used (incorrectly) as a synonym for "phonetic algorithm".

'// The basic aim is for names with the same pronunciation to be
'// encoded to the same signature so that matching can occur despite
'// minor differences in spelling.

'// The Soundex code for a name consists of a letter followed by three
'// numbers: the letter is the first letter of the name, and the numbers
'// encode the remaining consonants.

'// Similar sounding consonants share the same number so, for example,
'// the labial B, F, P and V are all encoded as 1.

'// If two or more letters with the same number were adjacent in the
'// original name, or adjacent except for any intervening vowels, then
'// all are omitted except the first.

'// Vowels can affect the coding, but are never coded directly unless
'// they appear at the start of the name.

'// Russell Soundex for Spell Checking

'// This particular version of the Soundex algorithm has been adapted
'// from the original design in an attempt to more reliably facilitate
'// word matching for a generic English language spell checker.

'// Normally, each Soundex begins with the first letter of the given
'// name and only subsequent letters are used to produce the phonetic
'// signature, so only names beginning with the same first letter are
'// compared for similar pronunciation using the standard algorithm.

'// For example, one may seek the correct spelling for "upholstery" and
'// may inadvertently type "apolstry", "apolstery", or even "apholstery"
'// but would still not retrieve the correct spelling for this word.

'// Therefore, this version of the Soundex algorithm has been modified
'// to allow the matching of words that start with differing first
'// letters so as not to assume that the first letter is always known.

'// Consequently, encoding begins with the first letter of the word.

'// Because of this change, many more similarly spelled words are
'// returned as a match, so the Soundex's length has also been
'// extended from three numbers to four to produce a more unique
'// phonetic signature.

'// Returns the 4 character Soundex code for an English word.
Private Function GetSoundexWord(argWord As String) As Long
    Dim replaceMask(0& To 6&) As Boolean
    Dim bytSoundex(1& To 4&) As Byte
    Dim sWord As String, code As Byte
    Dim i As Long, j As Long

    If LenB(argWord) = 0& Then Exit Function

    '// Normalise it to remove ambiguity
    sWord = LCase$(argWord)

    '// Replacement
    '   [a, e, h, i, o, u, w, y] = 0
    '   [b, f, p, v] = 1
    '   [c, g, j, k, q, s, x, z] = 2
    '   [d, t] = 3
    '   [l] = 4
    '   [m, n] = 5
    '   [r] = 6

    replaceMask(0&) = True '// do nothing
    For i = 1& To Len(sWord)
        Select Case MidI(sWord, i) 'Mid$(sWord, i, 1)
              ' "a", "e", "h", "i", "o", "u", "w", "y"
            Case 97, 101, 104, 105, 111, 117, 119, 121:  code = 0& '// do nothing
              
              ' "b", "f", "p", "v"
            Case 98, 102, 112, 118:                      code = 1& '// key labials
              
              ' "c", "g", "j", "k", "q", "s", "x", "z"
            Case 99, 103, 106, 107, 113, 115, 120, 122:  code = 2&
            
            Case 100, 116: code = 3&   ' "d", "t"
            Case 108:      code = 4&   ' "l"
            Case 109, 110: code = 5&   ' "m", "n"
            Case 114:      code = 6&   ' "r"
        End Select

        If replaceMask(code) Then '// do nothing if already recorded
        Else '// add new code
            replaceMask(code) = True
            j = j + 1&
            bytSoundex(j) = code
        End If

        If j = 4& Then Exit For
    Next i

    '// Return the first four values (padded with 0's)
    CopyMemByR GetSoundexWord, bytSoundex(1&), 4&
End Function

'// Returns the 4 character Soundex code for an English word
'// but from right to left.
Private Function GetSoundexWordR(argWord As String) As Long
    Dim replaceMask(0& To 6&) As Boolean
    Dim bytSoundex(1& To 4&) As Byte
    Dim sWord As String, code As Byte
    Dim i As Long, j As Long

    If LenB(argWord) = 0& Then Exit Function

    '// Normalise it to remove ambiguity
    sWord = LCase$(argWord)

    '// Replacement
    '   [a, e, h, i, o, u, w, y] = 0
    '   [b, f, p, v] = 1
    '   [c, g, j, k, q, s, x, z] = 2
    '   [d, t] = 3
    '   [l] = 4
    '   [m, n] = 5
    '   [r] = 6

    replaceMask(0&) = True '// do nothing
    For i = Len(sWord) To 1& Step -1
        Select Case MidI(sWord, i) 'Mid$(sWord, i, 1)
              ' "a", "e", "h", "i", "o", "u", "w", "y"
            Case 97, 101, 104, 105, 111, 117, 119, 121:  code = 0& '// do nothing
              
              ' "b", "f", "p", "v"
            Case 98, 102, 112, 118:                      code = 1& '// key labials
              
              ' "c", "g", "j", "k", "q", "s", "x", "z"
            Case 99, 103, 106, 107, 113, 115, 120, 122:  code = 2&
            
            Case 100, 116: code = 3&   ' "d", "t"
            Case 108:      code = 4&   ' "l"
            Case 109, 110: code = 5&   ' "m", "n"
            Case 114:      code = 6&   ' "r"
        End Select

        If replaceMask(code) Then '// do nothing if already recorded
        Else '// add new code
            replaceMask(code) = True
            j = j + 1&
            bytSoundex(j) = code
        End If

        If j = 4& Then Exit For
    Next i

    '// Return the first four soundex values
    CopyMemByR GetSoundexWordR, bytSoundex(1&), 4&
End Function

'// Returns the Minimum of 3 numbers
Private Function min3(ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long) As Long
    If n1 < n2 Then min3 = n1 Else min3 = n2
    If n3 < min3 Then min3 = n3
End Function

'// Levenshtein Distance

'// From Wikipedia, the free encyclopedia

'// In information theory and computer science, the Levenshtein distance
'// or edit distance between two strings is given by the minimum number
'// of operations needed to transform one string into the other, where
'// an operation is an insertion, deletion, or substitution of a single
'// character.

'// It is named after Vladimir Levenshtein, who considered this distance
'// in 1965.

'// It is useful in applications that need to determine how similar two
'// strings are, such as spell checkers.

'// It can be considered a generalisation of the Hamming distance, which
'// is used for strings of the same length and only considers substitution
'// edits.

'// There are also further generalisations of the Levenshtein distance
'// that consider, for example, exchanging two characters as an operation,
'// like in the Damerau-Levenshtein distance algorithm.

'// int LevenshteinDistance(char str1[1..lenStr1], char str2[1..lenStr2])
'//    // d is a table with lenStr1+1 rows and lenStr2+1 columns
'//    declare int d[0..lenStr1, 0..lenStr2]
'//    // i and j are used to iterate over str1 and str2
'//    declare int i, j, cost
'//
'//    for i from 0 to lenStr1
'//        d i, 0:=i
'//    for j from 0 to lenStr2
'//        d 0, j:=j
'//
'//    for i from 1 to lenStr1
'//        for j from 1 to lenStr2
'//            if str1[i] = str2[j] then cost := 0
'//                                 else cost := 1
'//            d[i, j] := minimum(
'//                                 d[i-1, j  ] + 1,     // deletion
'//                                 d[i  , j-1] + 1,     // insertion
'//                                 d[i-1, j-1] + cost   // substitution
'//                             )
'//
'//    return d[lenStr1, lenStr2]

'// Returns the Levenshtein Distance between 2 strings.
Private Function GetLevenshteinDistance(argStr1 As String, argStr2 As String) As Long
    Dim LenStr1 As Long, LenStr2 As Long
    Dim editMatrix() As Long, cost As Long
    Dim str1_i As Integer, str2_j As Integer
    Dim i As Long, j As Long

    LenStr1 = Len(argStr1)
    LenStr2 = Len(argStr2)

    If LenStr1 = 0& Then
        '// The length of Str2 is the minimum number of operations
        '// needed to transform one string into the other
        GetLevenshteinDistance = LenStr2

    ElseIf LenStr2 = 0& Then
        '// The length of Str1 is the minimum number of operations
        '// needed to transform one string into the other
        GetLevenshteinDistance = LenStr1

    Else
        '// editMatrix is a table with lenStr1+1 rows and lenStr2+1 columns
        ReDim editMatrix(LenStr1, LenStr2) As Long

        '// i and j are used to iterate over str1 and str2
        For i = 0& To LenStr1
            editMatrix(i, 0&) = i
        Next
    
        For j = 0& To LenStr2
            editMatrix(0&, j) = j
        Next
    
        For i = 1& To LenStr1
            str1_i = MidI(argStr1, i) 'Mid$(argStr1, i, 1)
            For j = 1& To LenStr2
                str2_j = MidI(argStr2, j) 'Mid$(argStr2, j, 1)

                If str1_i = str2_j Then cost = 0& Else cost = 1&

                '//                     deletion,insertion,substitution
                editMatrix(i, j) = min3(editMatrix(i - 1&, j) + 1&, _
                                        editMatrix(i, j - 1&) + 1&, _
                                        editMatrix(i - 1&, j - 1&) + cost)
            Next j
        Next i
    
        GetLevenshteinDistance = editMatrix(LenStr1, LenStr2)
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Visual Basic is one of the few languages where you
' can't extract a character from or insert a character
' into a string at a given position without creating
' another string.
'
' The following Property fixes that limitation.
'
' Twice as fast as AscW and Mid$ when compiled.
'        iChr = AscW(Mid$(sStr, lPos, 1))
'        iChr = MidI(sStr, lPos)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get MidI(sStr As String, ByVal lPos As Long) As Integer
    CopyMemByR MidI, ByVal StrPtr(sStr) + lPos + lPos - 2&, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Mid$(sStr, lPos, 1) = Chr$(iChr)
'        MidI(sStr, lPos) = iChr
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let MidI(sStr As String, ByVal lPos As Long, ByVal iChr As Integer)
    CopyMemByR ByVal StrPtr(sStr) + lPos + lPos - 2&, iChr, 2&
End Property
