Attribute VB_Name = "modAPIModule"
Option Explicit

Public Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Const OCR_NORMAL = 32512


Public Sub HandleError(pErr As ErrObject)
    Dim mStr As String
    mStr = "The following error occurred at " + pErr.Source + "..." + vbCrLf
    mStr = mStr + pErr.Description
    Call App.LogEvent(mStr, 1)
    Call MsgBox(mStr, vbOKOnly + vbCritical, "Dictionary Error")
  End Sub
