VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum WordStatus
    vtrNotInDictionary
    vtrCapitalizationError
    vtrCorrect
End Enum
Public Enum WordInfoType
    vtrSynonyms
    vtrAntonyms
    vtrRelatedWords
    vtrRelatedExpressions
End Enum
Dim colSpellingSugg As SpellingSuggestions
Dim SuggestionMode As WdSpellingWordType
Dim MyDoc As Word.Document
Dim AgentFailed As Boolean


Public Function GetSpellSuggestions(Word As String, Wstatus As WordStatus) As Variant
    'THIS TAKES A WORD AS INPUT AND RETURNS AN ARRAY OF
    'PROBABLE SUGGESTIONS
    On Error GoTo errhandler:
    Dim ReturnArray() As String
    Dim UserEntry As String
    Dim SpellingSugg As SpellingSuggestion
    Dim X As Dictionary
    Dim ErrType As WdSpellingErrorType
    Dim i As Integer
    
    UserEntry = Word
    Call App.LogEvent("Getting Spelling suggestions for  '" + Word + "'...", 4)
    
    If MyDoc Is Nothing Then
        Call App.LogEvent("Instantiating a new Word Document...", 4)
        Set MyDoc = New Word.Document
    End If
    
    'RESETTING CONTENTS OF THE DOCUMENT
    MyDoc.Range.Text = ""
    MyDoc.Range.Text = UserEntry
    Call App.LogEvent("SuggestionMode: SpellWord", 4)
    SuggestionMode = wdSpellword
    
    'STORING THE SUGGESTIONS ARRAY IN A VARIANT
    Call App.LogEvent("STORING THE SUGGESTIONS ARRAY IN A VARIANT...")
    Set colSpellingSugg = MyDoc.Range.GetSpellingSuggestions(, , , SuggestionMode)
    
    'IN CASE NOTHING IS FOUND
    If colSpellingSugg.Count = 0 Then
        Call App.LogEvent("No suggestions found...", 2)
    End If
    
    'DETERMINING THE SPECIFIC CAUSE OF FALURE
     ErrType = colSpellingSugg.SpellingErrorType
     If ErrType = wdSpellingNotInDictionary Then
         Call App.LogEvent("Not found in Dictionary...", 2)
         Wstatus = vtrNotInDictionary
     ElseIf ErrType = wdSpellingCorrect Then
         Call App.LogEvent("The word is correctly spelled...", 4)
         Wstatus = vtrCorrect
     Else
         Call App.LogEvent("Capitalization Error...", 1)
         Wstatus = vtrCapitalizationError
     End If
     If colSpellingSugg.Count = 0 Then
         Call App.LogEvent("No Spelling suggestions could be retrieved for '" + Word + "'...", 2)
         GetSpellSuggestions = vbNullString
         Exit Function
     End If
    
     i = 1
     Call App.LogEvent("Preparing the return array of suggestions...", 4)
     For Each SpellingSugg In colSpellingSugg
         ReDim Preserve ReturnArray(1 To i) As String
         ReturnArray(i) = SpellingSugg.Name
         i = i + 1
     Next SpellingSugg
     'Call App.LogEvent("Closing the Word Document...", 4)
     'MyDoc.Close wdDoNotSaveChanges
     'Set MyDoc = Nothing
     GetSpellSuggestions = ReturnArray()
    Exit Function
errhandler:
    Call HandleError(Err)
End Function


Public Function GetWordInfo(Word As String, Wstatus As WordStatus, winfotype As WordInfoType) As Variant

    On Error GoTo errhandler:
    
    
    Dim MyInfo As SynonymInfo
    Dim ReturnArray As Variant
    Dim UserEntry As String
    Dim TextEntry As String
    
    
    UserEntry = Word 'The word
    
    
    'Check if MyDoc contains a valid reference to a Document Object
    'else, create one
    If MyDoc Is Nothing Then
        Call App.LogEvent("Instantiating a new Word document...", 4)
        Set MyDoc = New Word.Document
    End If
    
    'Setting the range and Text
    MyDoc.Range.Text = ""
    MyDoc.Range.Text = UserEntry
    Call App.LogEvent("Retrieving SynonymInfo for " + "'" + UCase(Word) + "'...", 4)
    Set MyInfo = MyDoc.Range.SynonymInfo
    If Not MyInfo.Found Then
        Call App.LogEvent("No information found...", 2)
        ReturnArray = vbNullString
        Exit Function
    End If
            
    ReturnArray = vbEmpty
    If winfotype = vtrSynonyms Then
        Call App.LogEvent("Generating meaning list for '" + Word + "'...", 4)
        ReturnArray = MyInfo.MeaningList
    ElseIf winfotype = vtrRelatedWords Then
        If MyInfo.MeaningCount >= 1 Then '
            Call App.LogEvent("Generating Related-Words list for '" + Word + "'...", 4)
            ReturnArray = MyInfo.SynonymList(meaning:=1)
        Else
            Call App.LogEvent("No help was available...", 2)
            ReturnArray = vbNullString
            Exit Function
        End If
    ElseIf winfotype = vtrAntonyms Then
        Call App.LogEvent("Generating Antonym List for '" + Word + "'...", 4)
        ReturnArray = MyInfo.AntonymList
    ElseIf winfotype = vtrRelatedExpressions Then
        Call App.LogEvent("Generating Related Expressions List for '" + Word + "'...", 4)
        ReturnArray = MyInfo.RelatedExpressionList
    End If
    
    
    If UBound(ReturnArray) = 0 Then
        ReturnArray = vbNullString
        Exit Function
    End If
    GetWordInfo = ReturnArray
    Exit Function
    
errhandler:
    Call HandleError(Err)
    

End Function






Private Sub Class_Initialize()

    Dim mStr As String
    Dim ApplicationPath As String
    ApplicationPath = IIf(Right(App.Path, 1) = "\", App.Path, (App.Path + "\"))

    On Local Error Resume Next
    
    Call Kill(ApplicationPath + "dlllog.log")
    If Err.Number <> 0 Then
        Err.Clear
        Call App.StartLogging(ApplicationPath + "dllLog.log", vbLogToFile)
    Else
        Call App.StartLogging(ApplicationPath + "dllLog.log", vbLogOff)
    End If
    mStr = "Initializing Dictionary DLL at " + Str(Time) + " on " + Format(CStr(Date), "dd-mmm-yyyy")
    Call App.LogEvent(mStr, 4)
    
End Sub

Private Sub Class_Terminate()

On Error GoTo errhandler:

    Dim mStr As String
    
    Call App.LogEvent("Terminating Dictionary...", 4)
    If MyDoc Is Nothing Then Exit Sub
    Call App.LogEvent("Closing Word Document...", 4)
    MyDoc.Close wdDoNotSaveChanges
    Call App.LogEvent("Destroying Word Document...", 4)
    Set MyDoc = Nothing
    mStr = "Dictionary DLL terminated at " + Str(Time) + " on " + Format(CStr(Date), "dd-mmm-yyyy")
    Call App.LogEvent(mStr, 4)
    Exit Sub
    
errhandler:
    Call HandleError(Err)
End Sub
