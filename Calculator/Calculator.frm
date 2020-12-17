VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Calculator"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton backspace 
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton div 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton mul 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton sub 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton add 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton ans 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton dot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton num9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton num8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton num7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton num6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton num5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton num4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton num3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton num2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton num1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton num0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox screen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim one As Double
Dim two As Double
Dim sign As String

Private Sub add_Click(Index As Integer)
one = screen.Text
sign = "+"
screen.Text = ""
End Sub

Private Sub ans_Click(Index As Integer)
two = screen.Text

If sign = "+" Then
screen.Text = one + two
Else
If sign = "-" Then
screen.Text = one - two
Else
If sign = "*" Then
screen.Text = one * two
Else
If sign = "/" Then
screen.Text = one / two
End If
End If
End If
End If
End Sub

Private Sub backspace_Click(Index As Integer)
If screen.Text = "" Then
screen.Text = ""
Else
lngLen = Len(screen.Text)
screen.Text = Left$(screen.Text, lngLen - 1)
End If
End Sub

Private Sub clear_Click(Index As Integer)
screen.Text = ""
End Sub

Private Sub div_Click(Index As Integer)
one = screen.Text
sign = "/"
screen.Text = ""
End Sub

Private Sub dot_Click(Index As Integer)
screen.Text = screen.Text & "."
End Sub

Private Sub mul_Click(Index As Integer)
one = screen.Text
sign = "*"
screen.Text = ""
End Sub

Private Sub num0_Click(Index As Integer)
screen.Text = screen.Text & 0
End Sub

Private Sub num1_Click(Index As Integer)
screen.Text = screen.Text & 1
End Sub

Private Sub num2_Click(Index As Integer)
screen.Text = screen.Text & 2
End Sub

Private Sub num3_Click(Index As Integer)
screen.Text = screen.Text & 3
End Sub

Private Sub num4_Click(Index As Integer)
screen.Text = screen.Text & 4
End Sub

Private Sub num5_Click(Index As Integer)
screen.Text = screen.Text & 5
End Sub

Private Sub num6_Click(Index As Integer)
screen.Text = screen.Text & 6
End Sub

Private Sub num7_Click(Index As Integer)
screen.Text = screen.Text & 7
End Sub

Private Sub num8_Click(Index As Integer)
screen.Text = screen.Text & 8
End Sub

Private Sub num9_Click(Index As Integer)
screen.Text = screen.Text & 9
End Sub

Private Sub sub_Click(Index As Integer)
one = screen.Text
sign = "-"
screen.Text = ""
End Sub
