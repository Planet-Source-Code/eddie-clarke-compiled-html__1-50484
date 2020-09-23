VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "My VB Application"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      ExtentX         =   3413
      ExtentY         =   2990
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'###########################################################

'Date: 16 Dec 2003
'Author: Eddie Clarke
'Title: Compiled HTML within VB6 App

'Forget about using the ADD-IN MANAGER to create your
'resource file as it causes lots of headaches for some
'reason. Edit the RC and BAT files with a text editor
'as required then run the BAT to make your RES file.
'This will compile the HTML files. You then need to add
'the RES file to your VB app and compile the EXE before
'running. If done correctly, you can use the "RES://"
'protocol to point a web page contained inside the EXE.

'###########################################################

Private Sub Form_Load()
    
    Dim strHelpFile As String
    strHelpFile = "res://" & App.Path & "\" & App.EXEName & ".exe/index.htm"
    WebBrowser1.Navigate strHelpFile
    
End Sub

Private Sub Form_Resize()
    
    With WebBrowser1
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
End Sub
