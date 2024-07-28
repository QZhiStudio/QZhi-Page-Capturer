Attribute VB_Name = "modMain"
Option Explicit

Global gFileName As String

Sub Main()
    gFileName = String(4096, 0)
    GetModuleFileNameW 0, StrPtr(gFileName), 4096
    gFileName = Replace(gFileName, Chr(0), "")
    
    frmMain.Show
End Sub
