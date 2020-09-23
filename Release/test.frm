VERSION 5.00
Begin VB.Form frmTest
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "Test"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6600
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  '¿ÞÂÊ
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Template"
         HelpContextID   =   1234
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenTemplate 
         Caption         =   "&Open Template"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open &Form"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileSaveAsTemplate 
         Caption         =   "Save As Template"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileSaveAsForm 
         Caption         =   "Save As Form"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu mnuFileSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
