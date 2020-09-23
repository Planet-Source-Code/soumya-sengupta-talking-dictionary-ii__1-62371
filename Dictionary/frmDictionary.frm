VERSION 5.00
Begin VB.Form frmDictionary 
   Caption         =   "Ventura Word Glossary"
   ClientHeight    =   5865
   ClientLeft      =   150
   ClientTop       =   630
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDictionary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   5520
      Width           =   9255
   End
   Begin VB.CommandButton cmdVoice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.ListBox lstWordInfo 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3630
      Left            =   3240
      TabIndex        =   5
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Frame fraWordInfo 
      Caption         =   "Query Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton optSynonym 
         Caption         =   "&Synonyms"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   390
         Width           =   1455
      End
      Begin VB.OptionButton optRelatedWords 
         Caption         =   "Related &Words"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton optRelatedExpressions 
         Caption         =   "Related &Expressions"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton optAntonym 
         Caption         =   "&Antonyms"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.ListBox lstCorrectionList 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   3630
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Spelling Suggestions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Type in the word below and press ENTER to check for spellings..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   480
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Word Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   210
      Left            =   3360
      TabIndex        =   10
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Press F3 to hear it!"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   825
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xx As New VenturaDictionary.clsDictionary
Dim WithEvents YY As ClsVoiceGenerator 'Check if TTS initialization succeeded.
Attribute YY.VB_VarHelpID = -1
Dim ZZ As ClsAnimation
Dim AA As New ClsRotateText
Dim IconPath As String
Dim CursorPath As String
Dim TTSFailure As Boolean
Dim DiscardKey As Boolean
Dim WIType As WordInfoType



Private Sub cmdVoice_Click()
If txtInput = vbNullString Then txtInput.SetFocus: Exit Sub
    On Error GoTo errhandler
    
    Dim Result As Boolean
    If TTSFailure Then Exit Sub
    Result = YY.TextToSpeech(Me.txtInput)
    If Not Result Then Call DisableSoundButton
    txtInput.SetFocus
    Exit Sub
errhandler:
    Call DisableSoundButton
    
End Sub

'Private Sub Command1_Click()
    'Call frmMail.Show(, Me)
'End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyF3 Then
        cmdVoice = True
    ElseIf KeyCode = vbKeyReturn Then
        If txtInput <> "" Then Call GetSpellSugg(txtInput)
    ElseIf KeyCode = vbKeyEscape Then
        Call Unload(Me)
    End If
        
End Sub

Private Sub Form_Load()

    Dim mstr As String
    If VBA.Right(App.Path, 1) = "\" Then
        mstr = Left(App.Path, (Len(App.Path) - 1))
    Else
        mstr = App.Path
    End If
    IconPath = mstr + "\Icons\"
    CursorPath = mstr + "\Cursors\dog.ani"
    cmdVoice.Picture = LoadPicture(IconPath + "soundon.ico")
    Set YY = New ClsVoiceGenerator
    Set ZZ = New VenturaDictionary.ClsAnimation
    txtInput = vbNullString
    Text1 = vbNullString
    Call Me.Show
    Call AA.GetRotateText(Text1, "Ventura Word Glossary... Conceived, designed and programmed by Soumya Sengupta. Design assistance from Arani Ghosh. Please email your suggestions to soumyas_v@hotmail.com", 0.2)
    'Stop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Verifying Exit
    If MsgBox("Wanna Quit?", vbQuestion + vbYesNo + vbDefaultButton2, "Ventura Glossary") = vbNo Then
        Cancel = True
        Exit Sub
    End If
    'Releasing the objects
    Set xx = Nothing
    Set YY = Nothing
    Set ZZ = Nothing
    End
End Sub

Private Sub lstCorrectionList_Click()
    If lstCorrectionList.ListIndex >= 0 Then
        txtInput = lstCorrectionList.List(lstCorrectionList.ListIndex)
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If DiscardKey Then
        KeyAscii = 0
        DiscardKey = False
        Exit Sub
    End If
    
    If KeyAscii = 13 Then 'Enter is pressed
        If txtInput <> "" Then
            Call GetSpellSugg(txtInput)
        End If
    End If
End Sub
Private Function GetSpellSugg(inputword As String) As Boolean

    Dim SpellSugg As Variant
    Dim wstatus As WordStatus
    Dim i As Integer
    
    
    Screen.MousePointer = vbHourglass
    Call lstCorrectionList.Clear
    Call lstCorrectionList.Refresh
    Call lstCorrectionList.AddItem("Checking spelling..")
    SpellSugg = xx.GetSpellSuggestions(inputword, wstatus)
    Call lstCorrectionList.Clear
    If IsArray(SpellSugg) Then
        For i = LBound(SpellSugg) To UBound(SpellSugg)
                Call lstCorrectionList.AddItem(SpellSugg(i))
                GetSpellSugg = True
            Next
     Else
        If wstatus = vtrCorrect Then
            Call lstCorrectionList.AddItem("Correct Word!")
            GetSpellSugg = True
        Else
            Call lstCorrectionList.AddItem("Sorry, couldn't help..")
            GetSpellSugg = False
        End If
    End If
    Screen.MousePointer = vbDefault
End Function

Private Function CheckWildCard(mstr As String) As Boolean
    If InStr(1, mstr, "?") <> 0 Or InStr(1, mstr, "*") <> 0 Then
        CheckWildCard = True
    End If
End Function

Private Sub lstWordInfo_Click()
If lstWordInfo.ListIndex >= 0 Then
        txtInput = lstWordInfo.List(lstWordInfo.ListIndex)
    End If
End Sub

Private Sub optAntonym_Click()

    WIType = vtrAntonyms
    Call GetWordInformation(txtInput)
    txtInput.SetFocus
    
End Sub

Private Sub optRelatedExpressions_Click()
    WIType = vtrRelatedExpressions
    Call GetWordInformation(txtInput)
    txtInput.SetFocus
End Sub

Private Sub optRelatedWords_Click()

    WIType = vtrRelatedWords
    Call GetWordInformation(txtInput)
    txtInput.SetFocus
    
End Sub

Private Sub optSynonym_Click()

    WIType = vtrSynonyms
    Call GetWordInformation(txtInput)
    txtInput.SetFocus
    
End Sub

Private Sub txtInput_Change()

    If Not TTSFailure Then
        cmdVoice.ToolTipText = "Readout " + txtInput
    End If
    optSynonym.Value = False
    optAntonym.Value = False
    optRelatedWords.Value = False
    optRelatedExpressions.Value = False
    
End Sub

Private Sub YY_TTSFailure(Cause As String)

    Call MsgBox("An error occurred while initializing" + vbCrLf + "the MS-TTS Engine. Check whether the component " + vbCrLf + "is correctly registered. The ReadOut feature will now be disabled." + vbCrLf + "Feel free to email me at soumyas_v@hotmail.com", vbCritical + vbOKOnly, "Ventura TTS Failure")
    Call DisableSoundButton
    
End Sub
Private Sub DisableSoundButton()

    cmdVoice.Picture = LoadPicture(IconPath + "soundoff.ico")
    TTSFailure = True
    cmdVoice.ToolTipText = "Readout feature disabled."
    
End Sub

Private Function GetWordInformation(word As String) As Boolean
   
    Dim Result As Variant
    Dim temp As Variant
    Dim wstatus As WordStatus
    
    Screen.MousePointer = vbHourglass
    Call lstWordInfo.Clear
    Call lstWordInfo.AddItem("Retrieving information for '" + VBA.UCase(word) + "'")
    Call lstWordInfo.Refresh
    'Retrieving the Word Info
    Result = xx.GetWordInfo(word, wstatus, WIType)
    Call lstWordInfo.Clear
    
    If IsArray(Result) Then
        For Each temp In Result
            Call lstWordInfo.AddItem(temp)
        Next
        GetWordInformation = True
    Else
        Call lstWordInfo.AddItem("Sorry, couldn't help..")
    End If
    Dim mstr As String
    If WIType = vtrAntonyms Then
        mstr = "Antonyms for '" + VBA.UCase(word) + "'"
    ElseIf WIType = vtrSynonyms Then
        mstr = "Synonyms for '" + VBA.UCase(word) + "'"
    ElseIf WIType = vtrRelatedWords Then
        mstr = "Related-Words for '" + VBA.UCase(word) + "'"
    Else
        mstr = "Related-Expressions for '" + VBA.UCase(word) + "'"
    End If
    Label2 = mstr
    Screen.MousePointer = vbDefault
    
End Function
