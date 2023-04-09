VERSION 5.00
Begin VB.Form frmfil 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Filipino"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10560
      Top             =   6960
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
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Question # 1"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      Begin VB.OptionButton opt1 
         Caption         =   "Konotasyon"
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
         Left            =   360
         TabIndex        =   3
         Top             =   3000
         Width           =   2055
      End
      Begin VB.OptionButton opt3 
         Caption         =   "Semantiks"
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
         Left            =   360
         TabIndex        =   2
         Top             =   4680
         Width           =   2055
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Denotasyon"
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
         Left            =   360
         TabIndex        =   1
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Ito ang pokus na kahulugan ng salita, parirala o pangugusap."
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
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   9375
      End
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   855
      Left            =   9240
      Top             =   6600
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
      Left            =   7560
      TabIndex        =   7
      Top             =   6840
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
      Left            =   9120
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "frmfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
If opt3 = True Then
MsgBox "Correct!", vbInformation, "You got it right!"
frmscore.lblright.Caption = frmscore.lblright.Caption + 1
frmcal.txtfil.Text = frmcal.txtfil.Text + 33
Timer1.Enabled = False
Else
MsgBox "Incorrect answer!", vbInformation, "Opss..."
frmscore.lblwrong.Caption = frmscore.lblwrong.Caption + 1
Timer1.Enabled = False
End If
Dim frmfil As New frmfil1
frmfil1.Show
Me.Hide
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Label1.Caption - 1
If Label1.Caption = 0 Then
Timer1.Enabled = False
MsgBox "Your time is up!", vbInformation, "STOP!"
frmfil1.Show
Me.Hide
End If
End Sub

Private Sub Timer2_Timer()

End Sub
