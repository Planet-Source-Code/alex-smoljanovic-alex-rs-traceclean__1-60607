VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trace Clean"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clean"
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   14
      Top             =   5100
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clean"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   13
      Top             =   4080
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clean"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   12
      Top             =   3120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clean"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clean"
      Height          =   255
      Index           =   1
      Left            =   5460
      TabIndex        =   10
      Top             =   1260
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clean"
      Height          =   255
      Index           =   0
      Left            =   5460
      TabIndex        =   9
      Top             =   360
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Selected"
      Default         =   -1  'True
      Height          =   375
      Left            =   4860
      TabIndex        =   8
      Top             =   5760
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      Picture         =   "frmMain.frx":109A
      ScaleHeight     =   540
      ScaleWidth      =   2370
      TabIndex        =   7
      Top             =   5700
      Width           =   2370
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start All"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   5760
      Width           =   915
   End
   Begin prjTraceClean.TCFolder TCFolder1 
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   5295
      _extentx        =   9340
      _extenty        =   1508
      folder          =   "%History"
      title           =   "History"
   End
   Begin prjTraceClean.TCReg TCReg0 
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5295
      _extentx        =   9340
      _extenty        =   1614
      regkey          =   "Software\Microsoft\Internet Explorer\TypedURLs"
      title           =   "Typed URL's [Internet Explorer]"
   End
   Begin prjTraceClean.TCFolder TCFolder2 
      Height          =   855
      Left            =   60
      TabIndex        =   2
      Top             =   1920
      Width           =   5295
      _extentx        =   9340
      _extenty        =   1508
      folder          =   "%Cookies"
      title           =   "Cookies"
   End
   Begin prjTraceClean.TCReg TCReg3 
      Height          =   915
      Left            =   60
      TabIndex        =   3
      Top             =   2820
      Width           =   5235
      _extentx        =   9234
      _extenty        =   1614
      regkey          =   "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
      title           =   "Recent Documents"
   End
   Begin prjTraceClean.TCReg TCReg5 
      Height          =   915
      Left            =   60
      TabIndex        =   4
      Top             =   4740
      Width           =   5235
      _extentx        =   9234
      _extenty        =   1614
      regkey          =   "Software\Kazaa\Search"
      title           =   "KaZaA Search History"
   End
   Begin prjTraceClean.TCReg TCReg4 
      Height          =   915
      Left            =   60
      TabIndex        =   6
      Top             =   3780
      Width           =   5295
      _extentx        =   9340
      _extenty        =   1614
      regkey          =   "Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List"
      title           =   "MSPaint Recent File List"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub cmdStart_Click()
On Error GoTo errh
    If MsgBox("Are you sure you want to start all Cleaners?" & vbCrLf & "This action is permanent.", vbQuestion + vbYesNo, "Clean All") <> vbYes Then Exit Sub
    
    Dim control As control
    For Each control In Me.Controls
        If VarType(control) = vbObject Then
            control.Start
        End If
    Next control
    Exit Sub
errh:
    MsgBox Err.Description
End Sub

Private Sub Command1_Click()
On Error GoTo errh
    If MsgBox("Are you sure you want to start the selected Cleaners?" & vbCrLf & "This action is permanent.", vbQuestion + vbYesNo, "Clean Selected") <> vbYes Then Exit Sub
    
    Dim control As control, i%
    For Each control In Me.Controls
        If VarType(control) = vbObject Then
            If control.Name Like "*#" Then
                i = CInt(Val(Right(control.Name, 1)))
                If Check1(i).Value = 1 Then
                    control.Start
                End If
            End If
        End If
    Next control
    Exit Sub
errh:
    MsgBox Err.Description
End Sub

