VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3495
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   2160
      Top             =   1080
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   2040
      ScaleHeight     =   15
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   22000
      Left            =   720
      Top             =   840
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright: 2003"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail Address: philipnaparan@eudoramail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2040
      MouseIcon       =   "frmSplash.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "E-mail Me !"
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company: Naparan Business Solution"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by: Philip V. Naparan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Java Text Editor version.1.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   3510
      Left            =   0
      Picture         =   "frmSplash.frx":015E
      Top             =   0
      Width           =   5700
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim cTime As Byte
Private Sub Label4_Click()
On Error Resume Next
Dim gotoURL
gotoURL = ShellExecute(Me.hWnd, vbNullString, "mailto:philipnaparan@eudoramail.com", "", vbNullString, 1)
End Sub

Private Sub Timer1_Timer()
Form1.Show
Unload Me
End Sub

Private Sub Timer2_Timer()
cTime = cTime + 1
If cTime = 1 Then Label5.Caption = "Creating Java Vertual Machine..."
If cTime = 2 Then Label5.Caption = "Reading the System Processor..."
If cTime = 3 Then Label5.Caption = "Repairing Bad Sector on Hard Disk..."
If cTime = 4 Then Label5.Caption = "Creating Operating System..."
If cTime = 5 Then Label5.Caption = "Upgrading Your Computer..."
If cTime = 6 Then Label5.Caption = "Upgrading System Memory into 256 Terabyte..."
If cTime = 7 Then Label5.Caption = "Upgrading Hard Disk Space into 950 Terabyte..."
If cTime = 8 Then Label5.Caption = "Upgrading System Processor into 833 Terabyte..."
If cTime = 9 Then Label5.Caption = "Upgrading Video Card into 128 Terabyte..."
If cTime = 10 Then Label5.Caption = "Joke Only !"
End Sub
