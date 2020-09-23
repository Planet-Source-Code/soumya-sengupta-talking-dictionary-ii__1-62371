VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsVoiceGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim xx As Object
Dim TTSFailed As Boolean
Public Event TTSFailure(Cause As String)


Public Function TextToSpeech(TextToRead As String) As Boolean

On Error Resume Next
    
If TTSFailed Then Exit Function

xx.Speak (TextToRead)
If Err.Number <> 0 Then
    TextToSpeech = False
    Call App.LogEvent("TTS failure occurred...")
    RaiseEvent TTSFailure(Err.Description)
    Err.Clear
Else
    TextToSpeech = True
End If
    
        
End Function

Private Sub Class_Initialize()
    On Error Resume Next
    Call App.LogEvent("Initializing the TTS Engine...", 4)
    Set xx = New ACTIVEVOICEPROJECTLib.DirectSS
    
    If Err.Number <> 0 Then
        TTSFailed = True
        Call App.LogEvent("TTS Failure occurred...", 1)
        RaiseEvent TTSFailure(Err.Description) 'Informing the client that initializatin failed
        Err.Clear
    Else
        Call App.LogEvent("TTS Engine initialized...", 4)
    End If
End Sub


