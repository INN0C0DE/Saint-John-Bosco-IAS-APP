VERSION 5.00
Begin VB.Form frmsci1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Science"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10800
      Top             =   6720
   End
   Begin VB.Frame Frame1 
      Caption         =   "Question # 2"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   10455
      Begin VB.OptionButton opt2 
         Caption         =   "Geology"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   4
         Top             =   3720
         Width           =   2055
      End
      Begin VB.OptionButton opt3 
         Caption         =   "Meteorology"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   4560
         Width           =   2055
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Seismology"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   2
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "It is the study of Earthquake."
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   8895
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CHOOSE"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   7
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Time left:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   855
      Left            =   9480
      Top             =   6360
      Width           =   1095
   End
End
Attribute VB_Name = "frmsci1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
If opt1 = True Then
MsgBox "Correct!", vbInformation, "You got it right!"
frmscore.lblright.Caption = frmscore.lblright.Caption + 1
frmcal.Txtscience.Text = frmcal.Txtscience.Text + 33
Timer1.Enabled = False
Else
MsgBox "Incorrect answer!", vbInformation, "Opss..."
frmscore.lblwrong.Caption = frmscore.lblwrong.Caption + 1
Timer1.Enabled = False
End If
Dim frmsci1 As New frmsci2
frmsci2.Show
Me.Hide
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Label1.Caption - 1
If Label1.Caption = 0 Then
Timer1.Enabled = False
MsgBox "Your time is up!", vbInformation, "STOP!"
frmsci2.Show
Me.Hide
End If
End Sub
