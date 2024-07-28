VERSION 5.00
Begin VB.Form frmSnapShot 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9360
   Icon            =   "frmSnapShot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9360
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      Height          =   6375
      Left            =   120
      ScaleHeight     =   421
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   605
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8760
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar hsbScroll 
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   6000
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.VScrollBar vsbScroll 
         Height          =   6015
         Left            =   8760
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picSnapShot 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   1
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseThisWindow 
         Caption         =   "&Close this window"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmSnapShot"
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

Private Sub Form_Load()
    '
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picPanel.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
End Sub

Private Sub hsbScroll_Scroll()
    picSnapShot.Left = -hsbScroll.Value
End Sub

Private Sub mnuFileCloseThisWindow_Click()
    Unload Me
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim drRet As DLGRET
    drRet = GetSaveFile(Me.hwnd, "Bitmap(*.bmp)" & Chr(0) & "*.bmp" & Chr(0) & Chr(0), App.Path, "")
    
    If drRet.blnError = False Then
        drRet.strFileName = Replace(drRet.strFileName, Chr(0), "")
        If LCase(Right(drRet.strFileName, 4)) <> ".bmp" Then drRet.strFileName = drRet.strFileName & ".bmp"
        SavePicture picSnapShot.Image, drRet.strFileName
    End If
End Sub

Private Sub picPanel_Resize()

    vsbScroll.Move picPanel.ScaleWidth - (240 / Screen.TwipsPerPixelX), 0, (240 / Screen.TwipsPerPixelX), picPanel.ScaleHeight - (240 / Screen.TwipsPerPixelY)
    hsbScroll.Move 0, picPanel.ScaleHeight - (240 / Screen.TwipsPerPixelY), picPanel.ScaleWidth - (240 / Screen.TwipsPerPixelX), (240 / Screen.TwipsPerPixelY)
    picMask.Move picPanel.ScaleWidth - (240 / Screen.TwipsPerPixelX), picPanel.ScaleHeight - (240 / Screen.TwipsPerPixelY), (240 / Screen.TwipsPerPixelX), (240 / Screen.TwipsPerPixelY)
    
    If picSnapShot.Width > picPanel.ScaleWidth Then
        hsbScroll.Visible = True
        hsbScroll.Enabled = True
        hsbScroll.Max = picSnapShot.Width - picPanel.ScaleWidth + (240 / Screen.TwipsPerPixelX)
    Else
        hsbScroll.Enabled = False
        hsbScroll.Visible = False
    End If
    
    If picSnapShot.Height > picPanel.ScaleHeight Then
        vsbScroll.Visible = True
        vsbScroll.Enabled = True
        vsbScroll.Max = picSnapShot.Height - picPanel.ScaleHeight + (240 / Screen.TwipsPerPixelY)
    Else
        vsbScroll.Enabled = False
        vsbScroll.Visible = False
    End If
    
    If (hsbScroll.Visible = True) Or (vsbScroll.Visible = True) Then
        picMask.Visible = True
    Else
        picMask.Visible = False
    End If
End Sub

Private Sub picSnapShot_Resize()
    picPanel_Resize
End Sub

Private Sub vsbScroll_Scroll()
    picSnapShot.Top = -vsbScroll.Value
End Sub
