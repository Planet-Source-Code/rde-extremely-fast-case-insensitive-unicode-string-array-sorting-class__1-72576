VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   " Spell Checker Test Form"
   ClientHeight    =   2415
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2115
      HideSelection   =   0   'False
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   90
      Width           =   4665
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpell 
         Caption         =   "&Spell Check..."
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Text1 = "A Comprehensive Spell Checker Revisited. This is a modified version of the spell checker class included in Shelz's COTM submission 'A Comprehensive Spell Checker' at txtCodeId=65992. This class includes the Soundex phonetic algorithm and the incredible Levenshtein Distance algorithm from Wikipedia, the free encyclopedia, which when combined produce a most effective spell checker solution. That submission impressed me with its most effective and concise spell Checking algorithms, a perfect demo project, and the smart way it provided a complete database of words in a small 1 and a half MB download. But like many other spell checkers it had a common limitation. " & _
            "The basic aim of the Soundex algorithm is for words with the same pronunciation to be encoded to the same string so that matching can occur despite minor differences in spelling. Unfortunately, the Soundex code for a word consists of a letter followed by three numbers: the letter is the first letter of the word, and the numbers encode the remaining consonants. Therefore, only words beginning with the same first letter are compared for similar pronunciation using the standard algorithm. For example, one may seek the correct spelling for 'upholstery' and may inadvertently type 'apolstry' but would not retrieve the correct spelling for this word. " & _
            "This version of the Soundex algorithm has been modified to allow the matching of words that start with differing first letters so as not to assume that the first letter is always known. In this version all Soundex's begin with the letter 'S', and the encoding always begins with the first letter of the word. The Levenshtein Distance algo marries perfectly with the results produced by the Soundex algo to identify the correct spelling for the given (mis-spelt) word every time! A search on the word 'apolstry' with the minimum successful Levenshtein Distance setting returns just four words where one of these words is 'upholstery'. Happy coding, Rd :)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmSpellCheck
End Sub

Private Sub mnuSpell_Click()
    frmSpellCheck.Show
    If (Text1.SelText <> vbNullString) Then
        frmSpellCheck.CorrectedWord = Text1.SelText
    End If
    Do While frmSpellCheck.Visible
        DoEvents
    Loop
    Text1.SelText = frmSpellCheck.CorrectedWord
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrHandler
    ' Check if right mouse button was clicked
    If ((Button And vbRightButton) = vbRightButton) Then
        ' Display the Edit menu as a pop-up menu
        PopupMenu frmTest.mnuEdit, &H2&
    End If
ErrHandler:
End Sub

Private Sub mnuCut_Click()
    If (Text1.SelText <> vbNullString) Then
        Clipboard.Clear
        'Clipboard.SetText rtfText.SelRTF, vbCFRTF
        Clipboard.SetText Text1.SelText, vbCFText
        Text1.SelText = vbNullString
    End If
End Sub

Private Sub mnuCopy_Click()
    If (Text1.SelText <> vbNullString) Then
        Clipboard.Clear
        'Clipboard.SetText rtfText.SelRTF, vbCFRTF
        Clipboard.SetText Text1.SelText, vbCFText
    End If
End Sub

Private Sub mnuPaste_Click()
    'If Clipboard.GetFormat(vbCFRTF) Then
    '    rtfText.SelRTF = Clipboard.GetText(vbCFRTF)
    'ElseIf Clipboard.GetFormat(vbCFText) Then
        Text1.SelText = Clipboard.GetText(vbCFText)
    'End If
End Sub
