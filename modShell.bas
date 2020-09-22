Attribute VB_Name = "modShell"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Const SW_SHOWNORMAL = 1

Public Function MailTo(addr As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    MailTo = ShellExecute(Scr_hDC, "Open", addr, vbNullString, "C:\", SW_SHOWNORMAL)
End Function
