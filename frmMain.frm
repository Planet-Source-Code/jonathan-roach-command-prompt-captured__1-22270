VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capture Demo by Jonathan Roach - stormdev@golden.net"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   6120
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   5000
      Left            =   120
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   511
      TabIndex        =   1
      Top             =   720
      Width           =   7730
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Let's Do It"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Click Here to Re-Set Focus on the DOS Window"
      Height          =   195
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hope you like the code, feedback and votes welcome on Planet Source Code, and comments via email !"
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5850
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   6000
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' API Declarations Required
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
' Required Variables
Dim winFind As Long
Dim winParent As Long
Dim foundIt As Boolean

Private Sub Command1_Click()
If Command1.Caption = "Let's Do It" Then
    ' Shell our command prompt window
    ' Note: This demo works best if the command prompt is set to start in a window mode and not
    ' fullscreen mode. If your command prompt window starts up full screen just hit alt+enter
    ' to window it.
    Shell "command.com", vbNormalNoFocus
    Timer1.Enabled = True
    Label3.Visible = True
    Command1.Caption = "E&xit"
Else
    ' Reset the original parent of the dos window and exit
    X% = SetParent(winFind, winParent)
    End
End If
End Sub

Private Sub Form_Load()
foundIt = False
End Sub

Private Sub Label3_Click()
' Use window positioning to re-set focus to the dos window
X% = SetWindowPos(winFind, 1, 10, 10, 500, 300, 0)
End Sub

Private Sub Timer1_Timer()
' Find the window in question, not this can be done with virtually any window, provided you know the class name and/or title.
winFind = FindWindow("tty", "MS-DOS Prompt")
If winFind <> 0 Then
' Found it
    foundIt = True
    Form1.ZOrder
    MsgBox "What has happened so far, a dos prompt window has been opened, once you click the ok button I will transfer the entire dos window into the picturebox on our form...", vbOKOnly, "Info"
    ' Get the parent of the dos window and save it.
    winParent = GetParent(winFind)
    ' Change the parent of the dos prompt to the picturebox.
    X% = SetParent(winFind, Picture1.hwnd)
    ' Set the window size and position inside the picturebox
    X% = SetWindowPos(winFind, 1, 10, 10, 500, 300, 0)
    Picture1.Cls
    Timer1.Enabled = False
End If
End Sub
