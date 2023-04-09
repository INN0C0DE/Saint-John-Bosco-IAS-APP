VERSION 5.00
Begin VB.Form frmsub 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Choose a Subject"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CHOOSE A SUBJECT DO YOU WANT TO ANSWER:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   10575
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "FILIPINO"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ENGLISH"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "MATH"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SCIENCE"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H008080FF&
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5640
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmsub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmeng.Show
frmsub.Hide
End Sub

Private Sub Command2_Click()
frmfil.Show
frmsub.Hide
End Sub

Private Sub Command3_Click()
frmmath.Show
frmsub.Hide
End Sub

Private Sub Command4_Click()
frmsci.Show
frmsub.Hide
End Sub

Private Sub Command5_Click()
Dim frmsub As New frmwelcome
frmwelcome.Show
Me.Hide
End Sub

