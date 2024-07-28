VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6615
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "Capture"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboURL 
      Height          =   300
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   180
      Width           =   5775
   End
   Begin VB.CommandButton cmdGoForward 
      Caption         =   ">>"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "<<"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New window"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseThisWindow 
         Caption         =   "&Close this window"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuNavigate 
      Caption         =   "&Navigate"
      Begin VB.Menu mnuNavigateGoBack 
         Caption         =   "Go &back"
      End
      Begin VB.Menu mnuNavigateGoForward 
         Caption         =   "Go &forward"
      End
      Begin VB.Menu mnuNavigateRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuNavigateBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNavigateCapture 
         Caption         =   "&Capture"
         Shortcut        =   ^{INSERT}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpLicense 
         Caption         =   "Apache License, Version 2.0"
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Page Capturer"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Const DEFAULT_URL = "https://www.0xaa55.com/"

Private Sub brwWebBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    cmdCapture.Enabled = False
End Sub

Private Sub brwWebBrowser_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
    If (Command = CSC_NAVIGATEBACK) Then
        cmdGoBack.Enabled = Enable
        mnuNavigateGoBack.Enabled = Enable
    ElseIf (Command = CSC_NAVIGATEFORWARD) Then
        cmdGoForward.Enabled = Enable
        mnuNavigateGoForward.Enabled = Enable
    End If
End Sub

Private Sub brwWebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    cboURL.Text = brwWebBrowser.LocationURL
    
    If pDisp = brwWebBrowser.object Then
        cmdCapture.Enabled = True
        
        Dim i As Long
        
        For i = 0 To cboURL.ListCount - 1
            If brwWebBrowser.LocationURL = cboURL.List(i) Then Exit Sub
        Next i
        
        cboURL.AddItem brwWebBrowser.LocationURL
    End If
End Sub

Private Sub brwWebBrowser_DownloadBegin()
    brwWebBrowser.Silent = True
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    brwWebBrowser.Silent = True
End Sub

Private Sub brwWebBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Dim frmBrowser As New frmMain
    
    Set ppDisp = frmBrowser.brwWebBrowser.object
    
    frmBrowser.Show
End Sub

Private Sub brwWebBrowser_TitleChange(ByVal Text As String)
    DefWindowProcW Me.hwnd, WM_SETTEXT, 0, StrPtr(Text)
End Sub

Private Sub cboURL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdGo_Click
End Sub

Private Sub cmdCapture_Click()
    Dim frmShot As New frmSnapShot

    frmShot.Caption = "The snapshot of " & brwWebBrowser.LocationURL

    GetWebBrowserSnapShot brwWebBrowser, frmShot.picSnapShot
    
    frmShot.Show
    
End Sub

Private Sub cmdGo_Click()
    brwWebBrowser.Navigate cboURL.Text
End Sub

Private Sub cmdGoBack_Click()
    brwWebBrowser.GoBack
End Sub

Private Sub cmdGoForward_Click()
    brwWebBrowser.GoForward
End Sub

Private Sub cmdRefresh_Click()
    brwWebBrowser.Refresh
End Sub

Private Sub Form_Load()
    If Me Is frmMain Then brwWebBrowser.Navigate DEFAULT_URL
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdCapture.Left = Me.ScaleWidth - cmdCapture.Width - 120 / Screen.TwipsPerPixelX
    cmdRefresh.Left = cmdCapture.Left - cmdRefresh.Width - 120 / Screen.TwipsPerPixelX
    cmdGo.Left = cmdRefresh.Left - cmdGo.Width - 120 / Screen.TwipsPerPixelX
    
    cboURL.Width = cmdGo.Left - cboURL.Left - 120 / Screen.TwipsPerPixelX
    
    brwWebBrowser.Width = Me.ScaleWidth - 240 / Screen.TwipsPerPixelX
    brwWebBrowser.Height = Me.ScaleHeight - brwWebBrowser.Top - 120 / Screen.TwipsPerPixelY
End Sub

Private Sub mnuFileCloseThisWindow_Click()
    Unload Me
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileNew_Click()
    Dim frmNewMain As New frmMain
    
    frmNewMain.brwWebBrowser.Navigate DEFAULT_URL
    frmNewMain.Show
End Sub

Private Sub mnuHelpAbout_Click()
    ShellAboutW Me.hwnd, StrPtr(App.ProductName), 0, Me.Icon
End Sub

Private Sub mnuHelpLicense_Click()
    Dim frmNewLicense As New frmLicense
    frmNewLicense.Show vbModal
End Sub

Private Sub mnuNavigateCapture_Click()
    cmdCapture_Click
End Sub

Private Sub mnuNavigateGoBack_Click()
    cmdGoBack_Click
End Sub

Private Sub mnuNavigateGoForward_Click()
    cmdGoForward_Click
End Sub

Private Sub mnuNavigateRefresh_Click()
    cmdRefresh_Click
End Sub
