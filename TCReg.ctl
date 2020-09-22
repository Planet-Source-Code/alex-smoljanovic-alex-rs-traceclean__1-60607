VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.UserControl TCReg 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   ScaleHeight     =   900
   ScaleWidth      =   5220
   Begin VB.PictureBox picCheck 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4740
      Picture         =   "TCReg.ctx":0000
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picXMark 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4740
      Picture         =   "TCReg.ctx":0F74
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picNoMark 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4740
      Picture         =   "TCReg.ctx":1EE8
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   20
      Left            =   960
      Picture         =   "TCReg.ctx":2E5C
      ScaleHeight     =   15
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   840
      Width           =   3840
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Not Complete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   540
      TabIndex        =   2
      Top             =   600
      Width           =   4650
   End
End
Attribute VB_Name = "TCReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Const m_def_Title = ""
Const m_def_RegKey = ""
Const m_def_HKey = &H80000001

Dim m_Title As String
Dim m_RegKey As String
Dim m_HKey As Long
Event Done()

Public Enum rHKEY
    CURRENT_USER = &H80000001
    LOCAL_MACHINE = &H80000002
End Enum


Public Property Get HKey() As Long
    HKey = m_HKey
End Property

Public Property Let HKey(ByVal New_HKey As rHKEY)
    m_HKey = New_HKey
    PropertyChanged "HKey"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get RegKey() As String
    RegKey = m_RegKey
End Property

Public Property Let RegKey(ByVal New_RegKey As String)
    If Ambient.UserMode Then Err.Raise 382
    m_RegKey = New_RegKey
    PropertyChanged "RegKey"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Start() As Boolean
    Dim i%
    pb.Max = 2: pb.Value = 1
    i = DelEnumValues(m_HKey, m_RegKey)
    
    picNoMark.Visible = False
    picCheck.Visible = True
    lblStatus = i & " items cleaned"
    pb.Value = pb.Max
End Function

Private Sub UserControl_InitProperties()
    m_RegKey = m_def_RegKey
    m_Title = m_def_Title
    m_HKey = m_def_HKey
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_HKey = PropBag.ReadProperty("HKey", m_def_HKey)
    m_RegKey = PropBag.ReadProperty("RegKey", m_def_RegKey)
    m_Title = PropBag.ReadProperty("Title", m_def_Title)
    lblTitle.Caption = m_Title
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("HKey", m_HKey, m_def_HKey)
    Call PropBag.WriteProperty("RegKey", m_RegKey, m_def_RegKey)
    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal New_Title As String)
    m_Title = New_Title
    lblTitle.Caption = m_Title
    PropertyChanged "Title"
End Property

