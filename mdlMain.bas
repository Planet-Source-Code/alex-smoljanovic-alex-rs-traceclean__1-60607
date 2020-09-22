Attribute VB_Name = "mdlMain"
Option Explicit
'***********************************************************************
'This application was developed for a
'PSC(Planet Source Code) User(s) on request.
'
'If you compile this application, please dont distribute it.
'However, feel free to use any of this code in you're own application(s).
'
'Alex Smoljanovic [Salex] 2005
'salex_software@shaw.ca, alexrs@gmail.com
'***********************************************************************

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Sub Main()
Dim rICc&: rICc = InitCommonControls
    Load frmMain
    frmMain.Show
End Sub
