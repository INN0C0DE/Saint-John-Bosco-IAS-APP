VERSION 5.00
Begin VB.Form frmhello 
   BackColor       =   &H0080C0FF&
   Caption         =   "Hello World (Ads Edition)"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "CLICK ME"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   4935
   End
   Begin VB.Line Line8 
      BorderWidth     =   5
      X1              =   10800
      X2              =   11400
      Y1              =   3360
      Y2              =   4080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!!!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   11775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You have been chosen to get a FREE"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   7575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "iPhone X"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1455
      Left            =   3720
      TabIndex        =   1
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   720
      X2              =   2640
      Y1              =   3720
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   720
      X2              =   360
      Y1              =   3720
      Y2              =   4560
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   360
      X2              =   2280
      Y1              =   4560
      Y2              =   5280
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   2280
      X2              =   2040
      Y1              =   5280
      Y2              =   5880
   End
   Begin VB.Line Line5 
      BorderWidth     =   5
      X1              =   2880
      X2              =   2640
      Y1              =   3840
      Y2              =   4440
   End
   Begin VB.Line Line6 
      BorderWidth     =   5
      X1              =   2040
      X2              =   3600
      Y1              =   5880
      Y2              =   5280
   End
   Begin VB.Line Line7 
      BorderWidth     =   5
      X1              =   2880
      X2              =   3600
      Y1              =   3840
      Y2              =   5280
   End
   Begin VB.Line Line9 
      BorderWidth     =   5
      X1              =   10800
      X2              =   9480
      Y1              =   3360
      Y2              =   4440
   End
   Begin VB.Line Line10 
      BorderWidth     =   5
      X1              =   11400
      X2              =   10080
      Y1              =   4080
      Y2              =   5160
   End
   Begin VB.Line Line11 
      BorderWidth     =   5
      X1              =   9480
      X2              =   9120
      Y1              =   4440
      Y2              =   3960
   End
   Begin VB.Line Line12 
      BorderWidth     =   5
      X1              =   10440
      X2              =   10080
      Y1              =   5640
      Y2              =   5160
   End
   Begin VB.Line Line13 
      BorderWidth     =   5
      X1              =   9120
      X2              =   8880
      Y1              =   3960
      Y2              =   5520
   End
   Begin VB.Line Line14 
      BorderWidth     =   5
      X1              =   10440
      X2              =   8880
      Y1              =   5640
      Y2              =   5520
   End
End
Attribute VB_Name = "frmhello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "HAHAHA! You have been pranked.", vbInformation, "Its a Trap!"
End Sub

Private Sub Command2_Click()
Dim frmhello As New frmwelcome
frmwelcome.Show
Me.Hide
End Sub
