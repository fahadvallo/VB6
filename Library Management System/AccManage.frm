VERSION 5.00
Begin VB.Form AccManage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   4335
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "< Back"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "UPDATE EXISTING ACCOUNT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   960
         TabIndex        =   21
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CREATE NEW ACCOUNT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   960
         TabIndex        =   20
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "What you want to do?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   4935
      Begin VB.CommandButton cmdCancer 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   17
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton cmdOKIE 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1200
         TabIndex        =   11
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblPassVis 
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label lblAcc 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "AccManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCan_Click()
Frame2.Visible = True
Frame1.Visible = False
End Sub

Private Sub cmdCancer_Click()
Frame2.Visible = True
Frame1.Visible = False
txtUser.Text = ""
txtPass.Text = ""
End Sub

Private Sub cmdOK_Click()
If lblAcc.Caption = "Create" Then

Call connect
rs.Open "Select * from Account order by userID", con, adOpenStatic, adLockBatchOptimistic
rs.AddNew
rs("userID") = txtID.Text
rs("password") = txtPassword.Text
rs("Full_Name") = txtName.Text
rs("borrowno") = "0"
rs.Update
rs.UpdateBatch
MsgBox "Successfully Created an Account!", , "Success"
txtID.Text = ""
txtPassword.Text = ""
txtName.Text = ""
Frame2.Visible = True
Frame1.Visible = False

ElseIf lblAcc.Caption = "Update" Then

rs("userID") = txtID.Text
rs("password") = txtPassword.Text
rs("Full_Name") = txtName.Text
rs.Update
rs.UpdateBatch
MsgBox "Successfully Update your Account!", , "Success"
txtID.Text = ""
txtPassword.Text = ""
txtName.Text = ""
Frame2.Visible = True
Frame1.Visible = False

End If
End Sub

Private Sub cmdOKIE_Click()
Call connect
rs.Open "Select * from Account where userID='" + txtUser.Text + "' and password='" + txtPass.Text + "'", con, adOpenDynamic, adLockBatchOptimistic

If rs.EOF Then
txtUser.Text = ""
txtPass.Text = ""
MsgBox "Incorrect User/Password!", vbCritical, "Error!"
Else
txtUser.Text = ""
txtPass.Text = ""
txtID.Text = rs("userID")
txtPassword.Text = rs("password")
txtName.Text = rs("Full_Name")
Frame1.Visible = False
lblAcc.Caption = "Update"
End If
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
Frame2.Visible = False
lblAcc.Caption = "Create"
txtID.Text = ""
txtPassword.Text = ""
txtName.Text = ""
End Sub

Private Sub Command2_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Command3_Click()
Unload Me
Login.Show
End Sub

Private Sub lblPassVis_Click()
If txtPass.PasswordChar = "*" Then
txtPass.PasswordChar = ""
lblPassVis.Caption = "Hide"
Else
txtPass.PasswordChar = "*"
lblPassVis.Caption = "Show"
End If
End Sub
