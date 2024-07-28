VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmLicense 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apache License, Version 2.0"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6960
   Icon            =   "frmLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin SHDocVwCtl.WebBrowser brwLicense 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   7435
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
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This product licensed under the Apache License, Version 2.0."
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6300
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub brwWebBrowser_StatusTextChange(ByVal Text As String)

End Sub

Private Sub Form_Load()
    brwLicense.Navigate "res://" & gFileName & "/License.html"
End Sub
