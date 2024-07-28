Attribute VB_Name = "modFunctions"
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

Public Type DLGRET
    lngStat As Long
    strFileName As String
    blnError As Boolean
End Type

Public Function GetObjectHWnd(ByVal obj As Object) As Long

    IUnknown_GetWindow obj.object, VarPtr(GetObjectHWnd)

End Function

Public Function GetWebBrowserSnapShot(ByVal brwBrowser As WebBrowser, ByVal picOutput As PictureBox) As Boolean

    ' 必然没有加载好
    If brwBrowser.ReadyState <> READYSTATE_COMPLETE Then
        GetWebBrowserSnapShot = False
        Exit Function
    End If
    
    Dim hWndBrowser As Long
    hWndBrowser = GetObjectHWnd(brwBrowser)
    
    Dim docDocument As HTMLDocument
    Set docDocument = brwBrowser.Document
    
    Dim strScroll As String
    Dim strMargin As String
    Dim strBorder As String
    Dim strOverflow As String
    
    strScroll = docDocument.body.Scroll
    brwBrowser.Document.body.Scroll = "no"

    strMargin = docDocument.body.Style.margin
    brwBrowser.Document.body.Style.margin = "0px"

    strBorder = docDocument.body.Style.border
    brwBrowser.Document.body.Style.border = "0px"

    strOverflow = docDocument.body.Style.overflow
    brwBrowser.Document.body.Style.overflow = "hidden"

    brwBrowser.Document.parentWindow.Scroll 0, 0
    
    Dim lngOldScaleMode As Long
    
    With picOutput
    
        lngOldScaleMode = picOutput.Parent.ScaleMode
        .Parent.ScaleMode = vbPixels
    
        .Width = docDocument.body.scrollWidth
        .Height = docDocument.body.scrollHeight
        .AutoRedraw = True
        
        .Parent.ScaleMode = lngOldScaleMode
    
    End With
    
    Dim lngOldWidth As Long
    Dim lngOldHeight As Long
    
    With brwBrowser
        lngOldScaleMode = brwBrowser.Parent.ScaleMode
        .Parent.ScaleMode = vbPixels
        
        lngOldWidth = .Width
        lngOldHeight = .Height
    
        .Width = docDocument.body.scrollWidth
        .Height = docDocument.body.scrollHeight
        
        PrintWindow GetObjectHWnd(brwBrowser), picOutput.hDC, PW_CLIENTONLY
        
        .Document.body.Scroll = strScroll
        .Document.body.Style.margin = strMargin
        .Document.body.Style.border = strBorder
        .Document.body.Style.overflow = strOverflow

        .Width = lngOldWidth
        .Height = lngOldHeight
        
        .Parent.ScaleMode = lngOldScaleMode
        
    End With

End Function

Public Function GetSaveFile(ByVal hwndOwner As Long, ByVal lpstrFilter As String, ByVal lpstrInitialDir As String, ByVal lpstrTitle As String) As DLGRET

    Dim ofnSaveFileName As OPENFILENAME
    Dim dlrReturn As DLGRET
    
    With ofnSaveFileName
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .lpstrFilter = lpstrFilter
        .lpstrFile = String(&HFF, 0)
        .nMaxFile = &HFF
        .lpstrFileTitle = String(&HFF, 0)
        .nMaxFileTitle = &HFF
        .lpstrInitialDir = lpstrInitialDir
        .lpstrTitle = lpstrTitle
        .flags = &H1804
        .lStructSize = Len(ofnSaveFileName)
    End With
    
    dlrReturn.lngStat = GetSaveFileNameA(ofnSaveFileName)
    If dlrReturn.lngStat >= 1 Then
        dlrReturn.strFileName = ofnSaveFileName.lpstrFile
        dlrReturn.blnError = False
    Else
        dlrReturn.strFileName = vbNullString
        dlrReturn.blnError = True
    End If
    
    GetSaveFile = dlrReturn
    
End Function

