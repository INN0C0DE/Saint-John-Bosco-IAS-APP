VERSION 5.00
Begin VB.Form frmeng2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "English"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10920
      Top             =   7080
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
      TabIndex        =   5
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Question # 3"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      Begin VB.OptionButton opt2 
         Caption         =   "Topic Control"
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
         Left            =   480
         TabIndex        =   3
         Top             =   3840
         Width           =   2055
      End
      Begin VB.OptionButton opt3 
         Caption         =   "Topic Shifting"
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
         Left            =   480
         TabIndex        =   2
         Top             =   4680
         Width           =   2055
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Turn Taking"
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
         Left            =   480
         TabIndex        =   1
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   $"frmeng2.frx":0000
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
         TabIndex        =   4
         Top             =   840
         Width           =   9135
      End
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   855
      Left            =   9600
      Top             =   6720
      Width           =   1095
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
      Left            =   7920
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
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
      Left            =   9480
      TabIndex        =   6
      Top             =   6720
      Width           =   1335
   End
End
Attribute VB_Name = "frmeng2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
If opt3 = True Then
MsgBox "Correct!", vbInformation, "You got it right!"
frmscore.lblright.Caption = frmscore.lblright.Caption + 1
frmcal.Txtenglish.Text = frmcal.Txtenglish.Text + 33
Timer1.Enabled = False
Else
MsgBox "Incorrect answer!", vbInformation, "Opss..."
frmscore.lblwrong.Caption = frmscore.lblwrong.Caption + 1
Timer1.Enabled = False
End If
Dim frmeng2 As New frmscore
frmscore.Show
Me.Hide
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Label1.Caption - 1
If Label1.Caption = 0 Then
Timer1.Enabled = False
MsgBox "Your time is up!", vbInformation, "STOP!"
frmscore.Show
Me.Hide
End If
End Sub
