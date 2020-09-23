VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRotateText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRotateStop As Boolean
Public Event TextRotate(RotateStop As Boolean)
Private lngOldCursor As Long, lngNewCursor As Long


Public Function GetAllowedText(TargetObject As Object) As Integer
    On Error GoTo GetOut
    GetAllowedText = TargetObject.Width / 120
GetOut:
    Exit Function
End Function

Public Function GetRotateText(TargetObject As Object, CaptionText As String, Optional Speed As Single = 0.5) As Boolean
Attribute GetRotateText.VB_MemberFlags = "40"

    Dim NoOfText As Integer
    Dim i As Integer
    Dim t
    Dim LenOfText As Integer
        
    On Error GoTo GetOut
'This validates the target object.
    Call ValidateRotateText(TargetObject)
'mRotateStop is a module level variable which acts as a flag
'deciding when the function ceases to work.
    If mRotateStop Then GoTo GetOut
    
    LenOfText = Len(VBA.Trim(CaptionText)) 'Length of text
    NoOfText = GetAllowedText(TargetObject) 'Max string length within display area of target
    If NoOfText = 0 Then GoTo GetOut
    
RotateText: 'This actually does the trick!
    For i = 1 To NoOfText
        If i <> 1 Then
            'Delay (Speed)
            t = Timer
            While Timer - t <= Speed
                DoEvents
            Wend
        End If
        RaiseEvent TextRotate(mRotateStop)
        'Meant for asynchronous automation server messages
        'checking for a boolean switch..
        If mRotateStop Then GoTo GetOut
        'Wiping out a space and put in an additional letter
        'from the main text to be displayed
        If TypeName(TargetObject) = "TextBox" Then
            TargetObject.Text = VBA.Space(NoOfText - i) & VBA.Mid(CaptionText, 1, i)
        Else
            TargetObject.Caption = VBA.Space(NoOfText - i) & VBA.Mid(CaptionText, 1, i)
        End If
        DoEvents
    Next
    
    If Len(CaptionText) > NoOfText Then
        For i = 1 To (Len(CaptionText) - NoOfText) + 1
            RaiseEvent TextRotate(mRotateStop)
            If mRotateStop Then GoTo GetOut
            If TypeName(TargetObject) = "TextBox" Then
                TargetObject.Text = VBA.Mid(CaptionText, i, NoOfText)
            Else
                TargetObject.Caption = VBA.Mid(CaptionText, i, NoOfText)
            End If
            DoEvents
            'Delay (Speed)
            t = Timer
            While Timer - t <= Speed
                DoEvents
            Wend
        Next
    End If

    For i = 1 To NoOfText
        RaiseEvent TextRotate(mRotateStop)
        If mRotateStop Then GoTo GetOut
        If TypeName(TargetObject) = "TextBox" Then
            TargetObject.Text = VBA.Right(TargetObject.Text, (Len(TargetObject.Text) - 1)) & VBA.Space(LenOfText) & VBA.Mid(CaptionText, i, NoOfText)
        Else
            TargetObject.Caption = VBA.Right(TargetObject.Caption, (Len(TargetObject.Caption) - 1)) & VBA.Space(LenOfText) & VBA.Mid(CaptionText, i, NoOfText)
        End If
        DoEvents
        
        t = Timer
            While Timer - t <= Speed
                DoEvents
            Wend
    Next
    
    GoTo RotateText
    GetRotateText = True
    
GetOut:
    Exit Function
    
End Function

Private Sub Delay(Speed As Integer)
    Dim S_Timer As Long
    
    S_Timer = Timer
    While (Timer - S_Timer) < Speed
        DoEvents
    Wend
End Sub


Private Sub ValidateRotateText(TargetObject As Object)

    If TypeName(TargetObject) = "TextBox" Then
        TargetObject.Alignment = 0
    ElseIf TypeName(TargetObject) = "Label" Then
        TargetObject.BackStyle = 0
    ElseIf TypeName(TargetObject) = "CommandButton" Then
        'Do something here for Command button
    Else
        mRotateStop = True
    End If
    
End Sub


