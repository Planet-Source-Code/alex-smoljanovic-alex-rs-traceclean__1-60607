VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.UserControl TCFolder 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   ScaleHeight     =   825
   ScaleWidth      =   5265
   Begin VB.PictureBox picNoMark 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4740
      Picture         =   "TCFolder.ctx":0000
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   60
      Width           =   540
   End
   Begin VB.PictureBox picXMark 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4740
      Picture         =   "TCFolder.ctx":0F74
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picCheck 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4740
      Picture         =   "TCFolder.ctx":1EE8
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   20
      Left            =   960
      Picture         =   "TCFolder.ctx":2E5C
      ScaleHeight     =   15
      ScaleWidth      =   3840
      TabIndex        =   4
      Top             =   780
      Width           =   3840
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
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
      TabIndex        =   3
      Top             =   540
      Width           =   4650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   495
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
      TabIndex        =   0
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "TCFolder"
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
Const m_def_Folder = ""

Dim m_Title As String
Dim m_Folder As String

Event Done()


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get Folder() As String
    Folder = m_Folder
End Property

Public Property Let Folder(ByVal New_Folder As String)
    If Ambient.UserMode Then Err.Raise 382
    m_Folder = New_Folder
    PropertyChanged "Folder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Start() As Boolean
On Error Resume Next
    Dim iShell As Shell, iFolder As Folder, iFolderItem As FolderItem
    Dim buffer$, FolderPath$, DeleteCount%
    
    If Left(m_Folder, 1) = "%" Then
        buffer = Mid(m_Folder, 2)
        FolderPath$ = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", buffer)
        Debug.Print FolderPath$
    Else
        FolderPath = m_Folder
    End If
    
    If Not GetAttr(FolderPath) And vbDirectory Then
        lblStatus.Caption = "Can't Find Directory"
        picNoMark.Visible = False
        picXMark.Visible = True
    Else
        Dim i%
        Set iShell = New Shell
        Set iFolder = iShell.NameSpace(FolderPath)
        'Debug.Print iFolder.Title; iFolder.Items.Count
        pb.Max = iFolder.Items.Count
        lblStatus.Caption = "Deleting Items"
        For i = 0 To iFolder.Items.Count - 1
            DoEvents
            'iFolder.Items.Item(i).InvokeVerb ("delete")
            'Debug.Print iFolder.Items.Item(i).Path; ","; iFolder.Items.Item(i).Name, FolderPath
            buffer = IIf(Right(FolderPath, 1) = "\", FolderPath & iFolder.Items.Item(i).Name, FolderPath & "\" & iFolder.Items.Item(i).Name)
            Debug.Print buffer
            If Dir(buffer, vbSystem Or vbNormal Or vbHidden Or vbDirectory) <> "" Then
                Kill buffer
                Debug.Print buffer
            Else
                iFolder.Items.Item(i).InvokeVerb ("delete")
            End If
            pb.Value = i + 1
            DeleteCount% = DeleteCount% + 1
        Next i
        
        lblStatus.Caption = "Deleted " & DeleteCount% & " items..."
        picNoMark.Visible = False
        picCheck.Visible = True
    End If
    
    If Not iFolderItem Is Nothing Then Set iFolderItem = Nothing
    If Not iFolder Is Nothing Then Set iFolder = Nothing
    If Not iShell Is Nothing Then Set iShell = Nothing
    pb.Value = pb.Max
    
End Function



Private Sub UserControl_InitProperties()
    m_Folder = m_def_Folder
    m_Title = m_def_Title
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Folder = PropBag.ReadProperty("Folder", m_def_Folder)
    m_Title = PropBag.ReadProperty("Title", m_def_Title)
     lblTitle.Caption = m_Title
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Folder", m_Folder, m_def_Folder)
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

