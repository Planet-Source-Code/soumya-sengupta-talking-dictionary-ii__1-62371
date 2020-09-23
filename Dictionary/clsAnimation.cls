VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private lngOldCursor As Long, lngNewCursor As Long

Public Sub StartAnimatedCursor(AniFilePath As String)

    'Create a copy of the current cursor for Windows NT compatibility
    On Error GoTo errhandler:
        
    Call App.LogEvent("Loading animated cursor from :" + AniFilePath + "...", 4)
    lngOldCursor = CopyIcon(GetCursor())
    
    'Check the passed string, if it containsa solid file path, then load the cursorfrom file. If not, add the App.Path,
    '*then* load cursor...
    If InStr(1, AniFilePath, "\") Then
        lngNewCursor = LoadCursorFromFile(AniFilePath)
    Else
        lngNewCursor = LoadCursorFromFile(App.Path & "\" & AniFilePath)
    End If
    'Activate the cursor
    SetSystemCursor lngNewCursor, OCR_NORMAL
    Exit Sub
errhandler:
    Call HandleError(Err)
End Sub
  
Public Sub RestoreLastCursor()
    On Error GoTo errhandler:
    Call App.LogEvent("Restoring old cursor..", 4)
    Call SetSystemCursor(lngOldCursor, OCR_NORMAL)
    Exit Sub
errhandler:
    Call HandleError(Err)
End Sub

