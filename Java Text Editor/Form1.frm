VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Untitled-Java Text Editor version.1.1"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6375
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2520
      Top             =   3240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4755
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5821
            MinWidth        =   5821
            Picture         =   "Form1.frx":058A
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   435
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   767
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0B24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuN 
         Caption         =   "&New               "
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOF 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuS 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSA 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu blnk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPS 
         Caption         =   "&Export Preview"
         Shortcut        =   ^P
      End
      Begin VB.Menu blnk9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuE 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEDT 
      Caption         =   "&Edit"
      Begin VB.Menu mnuTnD 
         Caption         =   "&Sheck Spelling"
         Shortcut        =   {F4}
      End
      Begin VB.Menu blnk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFND 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuP 
      Caption         =   "&Project"
      Begin VB.Menu mnuEP 
         Caption         =   "&Execute Application"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEA 
         Caption         =   "Execute &Applet"
         Shortcut        =   {F6}
      End
      Begin VB.Menu blnk10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCJ 
         Caption         =   "&Create Jar"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuJR 
      Caption         =   "&Java Resources"
      Begin VB.Menu mnuOT 
         Caption         =   "[Onlie Tutorials]                                        "
      End
      Begin VB.Menu blnk4 
         Caption         =   "-"
      End
      Begin VB.Menu mnui 
         Caption         =   "http://java.sun.com/j2ee"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnujavaworld 
         Caption         =   "http://www.javaworld.com"
         Checked         =   -1  'True
      End
      Begin VB.Menu blnk7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDWN 
         Caption         =   "[Downloads]"
      End
      Begin VB.Menu blnk8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJS 
         Caption         =   "http://java.sun.com"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnudeitel 
         Caption         =   "http://www.deitel.com"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnumindview 
         Caption         =   "http://www.mindview.net"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuintelinfo 
         Caption         =   "http://www.intelinfo.com"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuRQ 
         Caption         =   "&Requirements"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuAB 
         Caption         =   "&About Java Text Editor..."
         Shortcut        =   {F11}
      End
      Begin VB.Menu blnk5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAW 
         Caption         =   "&Author's Website"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'This all code in this Application was created was
'created by me Philip V. Naparan (Web Developer/
'Programmer) in just 7 Hour and 23 Min. last Aug.
'20,2003. This Application is use in java progra-
'mming to thosejava programers that don't have Java Sun
'ONE Studio CE. This Application can Compile,Run
'Java Applicaiton and Applet, and it also have the
'ability to create Java Executable jar.
'(See Requirement's for more!)
'
'PLS. DON'T FORGET TO VOTE THIS APPLICAITON.
'Thank you! Happy coding and God Bless...
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim curPath As String
Dim nLine As String
Dim authorsName As String
Dim i As Byte
Dim conSave As Boolean
Dim conOPSV As Boolean
Dim nDoc As Boolean
Dim nDoc1 As Boolean
Private Sub Form_Load()
nDoc1 = False
conSave = False
conOPSV = False
nDoc = True
authorsName = "Program created by: Philip V. Naparan"
RichTextBox1.Top = 0
RichTextBox1.Left = 0
CommonDialog1.Filter = "Java File (*.java)|*.java"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reply As Integer
If nDoc1 = True Then
    reply = MsgBox("Do you want to save the code?", vbExclamation + vbYesNo, "Java Text Editor")
    If reply = vbYes Then
        mnuS_Click
    End If
    Exit Sub
End If
If conSave = True Then
    reply = MsgBox("Do you want to save changes in the code?", vbExclamation + vbYesNoCancel, "Java Text Editor")
    If reply = vbYes Then
        Me.MousePointer = vbHourglass
        If curPath = "" Then
            CommonDialog1.CancelError = False
            CommonDialog1.ShowSave
        If CommonDialog1.FileName = "" Then
            Cancel = 1
            Me.MousePointer = vbDefault
            Exit Sub
        End If
            curPath = CommonDialog1.FileName
            RichTextBox1.SaveFile curPath, 1
        Else
            RichTextBox1.SaveFile curPath, 1
        End If
        CommonDialog1.FileName = ""
        Me.MousePointer = vbDefault
        MsgBox "Changes has been saved.", vbInformation, "Java Text Editor"
    End If
    If reply = vbCancel Then
        Cancel = 1
    End If
Else
    reply = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo, "Java Text Editor")
    If reply = vbNo Then
        Cancel = 1
    End If
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
RichTextBox1.Width = Me.Width - 100
RichTextBox1.Height = Me.Height - 1010
End Sub

Private Sub mnuAB_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuAW_Click()
getURL "http://www.philipnaparan.cjb.net", Me.hWnd
End Sub
Private Sub mnudeitel_Click()
getURL "http://www.deitel.com", Me.hWnd
End Sub

Private Sub mnuDFN()
Dim i, ii As Integer
ReDim fileDir(0) As String
Dim slshFinder As String
If curPath = "" Then
    MsgBox "Pls. save the file first!", vbExclamation, "Java Text Editor"
    Exit Sub
End If
For i = 1 To Len(curPath)
    If Not Left(Right(curPath, i), 1) = "\" Then
        slshFinder = Left(Right(curPath, i), 1) & slshFinder
    Else
        Me.Caption = slshFinder & " - Java Text Editor version.1.1"
        Exit For
    End If
Next i
End Sub

Private Sub mnuE_Click()
Unload Me
End Sub

Private Sub mnuEA_Click()
On Error Resume Next
Dim i, ii As Integer
ReDim fileDir(0) As String
Dim slshFinder As String
If curPath = "" Then
    MsgBox "Pls. save the file first before executing it!", vbExclamation, "Java Text Editor"
    Exit Sub
End If
For i = 1 To Len(curPath)
    If Not Left(Right(curPath, i), 1) = "\" Then
        slshFinder = Left(Right(curPath, i), 1) & slshFinder
    Else
        ReDim Preserve fileDir(UBound(fileDir) + 1)
        fileDir(UBound(fileDir) - 1) = slshFinder
        slshFinder = ""
    End If
Next i
Call MakeDIR
'creating HTML File
Open "c:\JavaTextEditor\HTML\AppletViewer.htm" For Output As #3
    Print #3, "<html>"
    Print #3, "<body topmargin='0' leftmargin='0' >"
    Print #3, "<applet code='" & Left(fileDir(0), Len(fileDir(0)) - 5) & ".class" & "' width='500' height='400'>"
    Print #3, "</applet>"
    Print #3, "<html>"
    Print #3, "</body>"
    Print #3, "</html>"
Close #3
'compiling and running
Open "c:\JavaTextEditor\IO\ExecuteApplet.bat" For Output As #3
    Print #3, "@ ECHO OFF"
    Print #3, "ECHO ------------------------------------------------"
    Print #3, "ECHO JAVA TEXT EDITOR (PROGRAMMER: PHILIP V. NAPARAN)"
    Print #3, "ECHO ------------------------------------------------"
    Print #3, "ECHO **************** Compiling Start ***************"
    Print #3, "cd\"
    Print #3, Left(curPath, 1) & ":"
    For ii = UBound(fileDir) - 1 To LBound(fileDir) Step -1
        Print #3, "cd " & fileDir(ii)
    Next ii
    Print #3, "javac " & fileDir(0)
    Print #3, "copy " & Left(fileDir(0), Len(fileDir(0)) - 5) & ".class" & " c:\JavaTextEditor\HTML\" & Left(fileDir(0), Len(fileDir(0)) - 5) & ".class"
    Print #3, "ECHO ***************** Compiling End ****************"
    Print #3, "ECHO **************** Execution Start ***************"
    Print #3, "cd\"
    Print #3, Left(curPath, 1) & ":"
    Print #3, "cd JavaTextEditor"
    Print #3, "cd HTML"
    Print #3, "appletviewer AppletViewer.htm"
    Print #3, "ECHO ***************** Execution End ****************"
    Print #3, "Pause"
    Print #3, "Exit"
    Print #3, "Del " & Left(fileDir(0), Len(fileDir(0)) - 5) & ".class"
    Print #3, "REM This Bat File was generated by Java Text Editor version.1.1 that"
    Print #3, "REM created by Philip V. Naparan (Web Developer/Programmer)"
    Print #3, "REM Visit: www.philipnaparan.cjb.net"
Close #3
Kill Left(curPath, Len(curPath) - Len(fileDir(0))) & Left(fileDir(0), Len(fileDir(0)) - 5) & ".class"
Shell "c:\JavaTextEditor\IO\ExecuteApplet.bat", vbMaximizedFocus
End Sub

Private Sub mnuEP_Click()
On Error Resume Next
Dim i, ii As Integer
ReDim fileDir(0) As String
Dim slshFinder As String
If curPath = "" Then
    MsgBox "Pls. save the file first before executing it!", vbExclamation, "Java Text Editor"
    Exit Sub
End If
For i = 1 To Len(curPath)
    If Not Left(Right(curPath, i), 1) = "\" Then
        slshFinder = Left(Right(curPath, i), 1) & slshFinder
    Else
        ReDim Preserve fileDir(UBound(fileDir) + 1)
        fileDir(UBound(fileDir) - 1) = slshFinder
        slshFinder = ""
    End If
Next i
Call MakeDIR
'compiling and running
Open "c:\JavaTextEditor\IO\Execute.bat" For Output As #3
    Print #3, "@ ECHO OFF"
    Print #3, "ECHO ------------------------------------------------"
    Print #3, "ECHO JAVA TEXT EDITOR (PROGRAMMER: PHILIP V. NAPARAN)"
    Print #3, "ECHO ------------------------------------------------"
    Print #3, "ECHO **************** Compiling Start ***************"
    Print #3, "cd\"
    Print #3, Left(curPath, 1) & ":"
    For ii = UBound(fileDir) - 1 To LBound(fileDir) Step -1
        Print #3, "cd " & fileDir(ii)
    Next ii
    Print #3, "javac " & fileDir(0)
    Print #3, "ECHO ***************** Compiling End ****************"
    Print #3, "ECHO **************** Execution Start ***************"
    Print #3, "java " & Left(fileDir(0), Len(fileDir(0)) - 5)
    Print #3, "ECHO ***************** Execution End ****************"
    Print #3, "Pause"
    Print #3, "Exit"
    Print #3, "REM This Bat File was generated by Java Text Editor version.1.1 that"
    Print #3, "REM created by Philip V. Naparan (Web Developer/Programmer)"
    Print #3, "REM Visit: www.philipnaparan.cjb.net"
Close #3

Kill Left(curPath, Len(curPath) - Len(fileDir(0))) & Left(fileDir(0), Len(fileDir(0)) - 5) & ".class"
Shell "c:\JavaTextEditor\IO\Execute.bat", vbMaximizedFocus
End Sub

Private Sub mnuFND_Click()
On Error Resume Next
nLine = InputBox("Enter the text to search :", "Java Text Editor", "Type Here!")
RichTextBox1.Find (nLine)
End Sub

Private Sub mnui_Click()
getURL "http://java.sun.com/j2ee", Me.hWnd
End Sub

Private Sub mnuintelinfo_Click()
getURL "http://www.intelinfo.com", Me.hWnd
End Sub

Private Sub mnujavaworld_Click()
getURL "http://www.javaworld.com", Me.hWnd
End Sub

Private Sub mnuJS_Click()
getURL "http://java.sun.com", Me.hWnd
End Sub

Private Sub mnumindview_Click()
getURL "http://www.mindview.net", Me.hWnd
End Sub

Private Sub mnuN_Click()
Dim reply As Integer
If conSave = True Then
    reply = MsgBox("Do you want to save changes in the code?", vbExclamation + vbYesNoCancel, "Java Text Editor")
    If reply = vbCancel Then
        Exit Sub
    End If
    If reply = vbYes Then
        Me.MousePointer = vbHourglass
        If curPath = "" Then
            CommonDialog1.CancelError = False
            CommonDialog1.ShowSave
        If CommonDialog1.FileName = "" Then
            Exit Sub
        End If
            curPath = CommonDialog1.FileName
            RichTextBox1.SaveFile curPath, 1
        Else
            RichTextBox1.SaveFile curPath, 1
        End If
        CommonDialog1.FileName = ""
        Me.MousePointer = vbDefault
        MsgBox "Changes has been saved.", vbInformation, "Java Text Editor"
    End If
End If
curPath = ""
RichTextBox1.Text = ""
Me.Caption = "Untitled-Java Text Editor version.1.1"
conSave = False
nDoc = True
End Sub

Private Sub mnuOF_Click()
Dim reply As Integer
If conSave = True Then
    reply = MsgBox("Do you want to save changes in the code?", vbExclamation + vbYesNoCancel, "Java Text Editor")
    If reply = vbCancel Then
        Exit Sub
    End If
    If reply = vbYes Then
        Me.MousePointer = vbHourglass
        If curPath = "" Then
            CommonDialog1.CancelError = False
            CommonDialog1.ShowSave
        If CommonDialog1.FileName = "" Then
            Exit Sub
        End If
            curPath = CommonDialog1.FileName
            RichTextBox1.SaveFile curPath, 1
        Else
            RichTextBox1.SaveFile curPath, 1
        End If
        CommonDialog1.FileName = ""
        Me.MousePointer = vbDefault
        MsgBox "Changes has been saved.", vbInformation, "Java Text Editor"
    End If
End If
Me.MousePointer = vbHourglass
    CommonDialog1.CancelError = False
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then
        conSave = False
        nDoc = False
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    curPath = CommonDialog1.FileName
    RichTextBox1.LoadFile curPath
    CommonDialog1.FileName = ""
    conOPSV = True
    conSave = False
    nDoc = False
    nDoc1 = False
    Call mnuDFN
Me.MousePointer = vbDefault
End Sub

Private Sub mnuPS_Click()
Open "c:\JavaTextEditor\Export\Preview.doc" For Output As #3
    Print #3, RichTextBox1.Text
Close #3
Open "c:\JavaTextEditor\IO\Preview.bat" For Output As #3
    Print #3, "@ ECHO OFF"
    Print #3, "ECHO ------------------------------------------------"
    Print #3, "ECHO JAVA TEXT EDITOR (PROGRAMMER: PHILIP V. NAPARAN)"
    Print #3, "ECHO ------------------------------------------------"
    Print #3, "ECHO Exporting File..."
    Print #3, "c:\JavaTextEditor\Export\Preview.doc"
    Print #3, "Exit"
Close #3
Shell "c:\JavaTextEditor\IO\Preview.bat", vbMaximizedFocus
End Sub

Private Sub mnuRQ_Click()
MsgBox "System Requirements:" & vbCrLf & "J2EE 1.2.1 or higher" & vbCrLf & "Windows 95,98,ME,2000,XP" & vbCrLf & "16MB RAM" & vbCrLf & "133Mhz Processor Speed", vbInformation, "Java Text Editor"
End Sub

Private Sub mnuS_Click()
Me.MousePointer = vbHourglass
If curPath = "" Then
    CommonDialog1.CancelError = False
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    curPath = CommonDialog1.FileName
    RichTextBox1.SaveFile curPath, 1
Else
    RichTextBox1.SaveFile curPath, 1
End If
CommonDialog1.FileName = ""
conSave = False
nDoc = False
nDoc = False
Call mnuDFN
Me.MousePointer = vbDefault
End Sub

Private Sub mnuSA_Click()
Me.MousePointer = vbHourglass
    CommonDialog1.CancelError = False
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    curPath = CommonDialog1.FileName
    
    RichTextBox1.SaveFile curPath, 1
    CommonDialog1.FileName = ""
    conSave = False
    nDoc = False
    nDoc = False
    Call mnuDFN
Me.MousePointer = vbDefault
End Sub

Private Sub mnuTnD_Click()
Dim strCK As String
On Error Resume Next
Me.MousePointer = vbHourglass
strCK = RichTextBox1.SelText

If strCK = "" Then
    MsgBox "Pls. highlight a text to check it's spelling.", vbExclamation, "Java Text Editor"
    Me.MousePointer = vbDefault
    Exit Sub
End If

If chkSplng(strCK) = True Then
    MsgBox "The spelling is correct.", vbInformation, "Java Text Editor"
Else
    MsgBox "The spelling is wrong.", vbCritical, "Java Text Editor"
End If

Me.MousePointer = vbDefault
strCK = ""
End Sub
Private Function chkSplng(sTxt As String) As Boolean
Dim spellChekr As New Word.Application
On Error Resume Next
chkSplng = spellChekr.CheckSpelling(sTxt)
End Function
Public Function getURL(urlADD As String, sourceHWND As String)
On Error Resume Next
Dim gotoURL
gotoURL = ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Function

Private Sub RichTextBox1_Change()
conSave = True
If nDoc = True Then
    nDoc1 = True
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 50
i = i + 1
StatusBar1.Panels(1).Text = Left(authorsName, i)
If i = Len(authorsName) Then
    i = 1
Timer1.Interval = 20000
End If
End Sub
Private Sub MakeDIR()
On Error Resume Next
MkDir ("c:\JavaTextEditor")
MkDir ("c:\JavaTextEditor\IO")
MkDir ("c:\JavaTextEditor\HTML")
MkDir ("c:\JavaTextEditor\Export")
End Sub
Private Sub mnuCJ_Click()
On Error Resume Next
Dim i, ii As Integer
ReDim fileDir(0) As String
Dim slshFinder As String
If curPath = "" Then
    MsgBox "Pls. save the file first before creating Jar !", vbExclamation, "Java Text Editor"
    Exit Sub
End If
For i = 1 To Len(curPath)
    If Not Left(Right(curPath, i), 1) = "\" Then
        slshFinder = Left(Right(curPath, i), 1) & slshFinder
    Else
        ReDim Preserve fileDir(UBound(fileDir) + 1)
        fileDir(UBound(fileDir) - 1) = slshFinder
        slshFinder = ""
    End If
Next i
Call MakeDIR
'Making MFT for Jar
Open Left(curPath, Len(curPath) - Len(fileDir(0))) & Left(fileDir(0), Len(fileDir(0)) - 5) & ".mft" For Output As #3
    Print #3, "Main-Class: " & Left(fileDir(0), Len(fileDir(0)) - 5)
Close #3
'compiling and running
Open "c:\JavaTextEditor\IO\ExecuteJar.bat" For Output As #3
    Print #3, "@ ECHO OFF"
    Print #3, "ECHO ------------------------------------------------"
    Print #3, "ECHO JAVA TEXT EDITOR (PROGRAMMER: PHILIP V. NAPARAN)"
    Print #3, "ECHO ------------------------------------------------"
    Print #3, "ECHO **************** Compiling Start ***************"
    Print #3, "cd\"
    Print #3, Left(curPath, 1) & ":"
    For ii = UBound(fileDir) - 1 To LBound(fileDir) Step -1
        Print #3, "cd " & fileDir(ii)
    Next ii
    Print #3, "javac " & fileDir(0)
    Print #3, "ECHO ***************** Compiling End ****************"
    Print #3, "ECHO ****************** Creating Jar ****************"
    Print #3, "jar cmf  " & Left(fileDir(0), Len(fileDir(0)) - 5) & ".mft " & Left(fileDir(0), Len(fileDir(0)) - 5) & ".jar *.class"
    Print #3, "ECHO ***************** Creating End *****************"
    Print #3, "copy " & Left(fileDir(0), Len(fileDir(0)) - 5) & ".jar" & " " & Left(curPath, Len(curPath) - Len(fileDir(0))) & Left(fileDir(0), Len(fileDir(0)) - 5) & ".jar"
    Print #3, "Del " & Left(fileDir(0), Len(fileDir(0)) - 5) & ".mft"
    Print #3, "Exit"
    Print #3, "REM This Bat File was generated by Java Text Editor version.1.1 that"
    Print #3, "REM created by Philip V. Naparan (Web Developer/Programmer)"
    Print #3, "REM Visit: www.philipnaparan.cjb.net"
Close #3
Kill Left(curPath, Len(curPath) - Len(fileDir(0))) & Left(fileDir(0), Len(fileDir(0)) - 5) & ".class"
Shell "c:\JavaTextEditor\IO\ExecuteJar.bat", vbHide
getURL Left(curPath, Len(curPath) - Len(fileDir(0))), Me.hWnd
End Sub
