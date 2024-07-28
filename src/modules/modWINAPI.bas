Attribute VB_Name = "modWINAPI"
' Copyright 2024 QZhi Studio
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

Option Explicit

Public Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_SETTEXT = &HC

Public Const PW_CLIENTONLY = 1

Public Declare Function PrintWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Boolean

Public Declare Function IUnknown_GetWindow Lib "shlwapi.dll" (ByVal punk As IUnknown, ByVal phwnd As Long) As Long

Public Declare Function ShellAboutW Lib "shell32.dll" (ByVal hwnd As Long, ByVal szApp As Long, ByVal szOtherStuff As Long, ByVal hIcon As Long) As Long

Public Declare Function GetSaveFileNameA Lib "comdlg32.dll" (ByRef pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
