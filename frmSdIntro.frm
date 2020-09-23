VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Stormdev Software Development (c), 1999-2000."
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Code Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.Label lblCodeDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSdIntro.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   4065
      End
      Begin VB.Label lblEmailAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "   Email: stormdev@golden.net"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   4065
      End
      Begin VB.Label lblAuthor 
         BackStyle       =   0  'Transparent
         Caption         =   "Author: Jonathan Roach"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   4065
      End
      Begin VB.Label lblDemoTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dos Prompt to Picturebox"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton butOptions 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton butOptions 
      Caption         =   "&Continue"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code is copyright(c) 2000, Stormdev Software Development.
' You are hereby granted rights to use/modify this code as you see fit,
' for commercial or personal use. The only stipulation is that I
' ask for some feedback regarding the code contained herein.
'
' Send feedback to: stormdev@golden.net
'      Code Author: Jonathan Roach
'  Purpose of Code:
'    Level of Code:
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub butOptions_Click(Index As Integer)
' Evaluate which button is being clicked
Select Case Index
    Case 0 ' Continue button code
        Unload Me
        Form1.Show
    Case 1 ' Quit button code
        Unload frmStartup
        End
End Select
End Sub

