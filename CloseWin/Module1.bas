Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10

Public Function EnumCallBack(ByVal hWndChild As Long, ByVal lParam As Long) As Long
Dim lngSize As Long
Dim strPadString As String
strPadString = String(255, 0)
lngSize = GetWindowText(hWndChild, strPadString, Len(strPadString))
strPadString = Left$(strPadString, lngSize)
If strPadString <> "" Then
    Form1.List1.AddItem strPadString
End If
EnumCallBack = True
End Function

Public Function CloseApplication(ByVal sAppCaption As String) As Boolean

    Dim lHwnd As Long
    Dim lRetVal As Long
    lHwnd = FindWindow(vbNullString, sAppCaption)


    If lHwnd <> 0 Then
        lRetVal = PostMessage(lHwnd, WM_CLOSE, 0&, 0&)
    End If

End Function

