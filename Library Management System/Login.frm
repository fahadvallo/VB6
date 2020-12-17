VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library System"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5131.579
   ScaleMode       =   0  'User
   ScaleWidth      =   9563.452
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton RETURN 
      Caption         =   "Return Books"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton LIBRARIA 
      Caption         =   "Librarian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6720
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton BOOKS 
      Caption         =   "Borrow Books"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label ACCOUNTS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Account"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "              Welcome!               Library System"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   9375
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ACCOUNTS_Click()
AccManage.Show
Unload Me
End Sub

Private Sub BOOKS_Click()
Borrow.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim img As String
img = App.Path & "\library.jpg"
Login.Picture = LoadPicture(img)
End Sub

Private Sub LIBRARIA_Click()
Librarian.Show
Unload Me
End Sub

Private Sub Return_Click()
Return_.Show
Unload Me
End Sub
