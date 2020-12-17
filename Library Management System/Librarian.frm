VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Librarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Library"
   ClientHeight    =   8055
   ClientLeft      =   3855
   ClientTop       =   2310
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmEditLib 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Change Librarian Password"
      Height          =   3375
      Left            =   3360
      TabIndex        =   20
      Top             =   2280
      Width           =   6015
      Begin VB.CommandButton Command7 
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
         Left            =   3240
         TabIndex        =   26
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtNEW 
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
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   1560
         Width           =   3135
      End
      Begin VB.CommandButton Command6 
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
         Left            =   1560
         TabIndex        =   23
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtOLD 
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
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5400
         TabIndex        =   27
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
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
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
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
         TabIndex        =   22
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Frame frmLogin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Login"
      Height          =   2415
      Left            =   3360
      TabIndex        =   14
      Top             =   2760
      Width           =   6015
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
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   840
         Width           =   3135
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
         Left            =   1080
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
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
         Left            =   3000
         TabIndex        =   15
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         Left            =   600
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblPassVis 
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Librarian Key"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Borrow History"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Users"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Books"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Librarian"
      Height          =   7095
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   11895
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
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
         Left            =   7080
         TabIndex        =   13
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
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
         Left            =   6000
         TabIndex        =   12
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
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
         Left            =   4920
         TabIndex        =   11
         Top             =   6240
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5415
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9551
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
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
         Left            =   3480
         TabIndex        =   5
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
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
         TabIndex        =   3
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
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
         Left            =   240
         TabIndex        =   2
         Top             =   6240
         Width           =   975
      End
   End
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
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Librarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
BOOKS.Show
BOOKS.Caption = "Add Book"
BOOKS.Command1.Caption = "Save"
BOOKS.SetFocus
End Sub

Private Sub cmdBack_Click()
If rs.BOF Then
MsgBox "You already reached the first record", vbInformation
rs.MoveNext
Else
rs.MovePrevious
End If
End Sub

Private Sub cmdCan_Click()
Login.Show
Unload Me
End Sub

Private Sub cmdDel_Click()
If rs.EOF = True Then
MsgBox "No Selected Book!", , "Select"
ElseIf rs.BOF = True Then
MsgBox "No Selected Book!", , "Select"
Else

If MsgBox("Are you want to Delete?", vbQuestion + vbYesNo) = vbYes Then

rs.Delete
rs.Update
rs.UpdateBatch

MsgBox "Record Deleted", vbInformation

Else
MsgBox "Record Not Deleted", vbInformation
End If
End If
End Sub

Private Sub cmdEdit_Click()
If rs.EOF = True Then
MsgBox "No Selected Book!", , "Select"
ElseIf rs.BOF = True Then
MsgBox "No Selected Book!", , "Select"
Else
BOOKS.Show
BOOKS.Caption = "Edit Book"
BOOKS.Command1.Caption = "Update"
BOOKS.SetFocus

BOOKS.Text1.Text = rs("BookNo")
BOOKS.Text2.Text = rs("Title")
BOOKS.Text3.Text = rs("Author")
BOOKS.Text4.Text = rs("Publication_date")
End If
End Sub

Private Sub cmdFirst_Click()
rs.MoveFirst
End Sub

Private Sub cmdLast_Click()
rs.MoveLast
End Sub

Private Sub cmdNext_Click()
If rs.EOF Then
MsgBox "You aleady reached the Last Record", vbInformation, ""
rs.MovePrevious
Else
rs.MoveNext
End If
End Sub

Private Sub cmdOK_Click()
Call connect
rs.Open "Select * from Librarian where masterkey='" + txtPass.Text + "'", con, adOpenDynamic, adLockBatchOptimistic

If rs.EOF Then
txtPass.Text = ""
MsgBox "Incorrect Password!", vbCritical, "Error!"
Else
Frame1.Enabled = True
Call CMDD
Call EDITD
Command3.Caption = "LOGOUT"
frmLogin.Visible = False
Command5.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Call connect
rs.Open "Select * from Book", con, adOpenStatic, adLockBatchOptimistic
Set DataGrid1.DataSource = rs
Call CMDE
Call EDITE
End Sub

Private Sub Command2_Click()
Call connect
rs.Open "Select * from Account", con, adOpenStatic, adLockBatchOptimistic
Set DataGrid1.DataSource = rs
Call CMDE
Call EDITD
End Sub

Private Sub Command3_Click()
If Command3.Caption = "LOGOUT" Then
If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Logout") = vbYes Then
Unload Me
Login.Show
End If
Else
Unload Me
Login.Show
End If
End Sub

Private Sub Command4_Click()
Call connect
rs.Open "Select * from Borrow where Full_Name", con, adOpenStatic, adLockBatchOptimistic
rs.Sort = "Book_Returned"
Set DataGrid1.DataSource = rs
Call CMDE
Call EDITD
End Sub

Private Sub Command5_Click()
frmEditLib.Visible = True
End Sub

Private Sub Command6_Click()
Call connect
rs.Open "Select * from Librarian where masterkey='" + txtOLD.Text + "'", con, adOpenDynamic, adLockBatchOptimistic
If rs.EOF = True Then
MsgBox "Incorrect old password!", , "Update Password"
Else
rs("masterkey") = txtNEW.Text
rs.Update
rs.UpdateBatch
MsgBox "Librarian Password Successfully Updated!"
txtOLD.Text = ""
txtNEW.Text = ""
frmEditLib.Visible = False
End If
End Sub

Private Sub Command7_Click()
frmEditLib.Visible = False
txtOLD.Text = ""
txtNEW.Text = ""
End Sub

Private Sub Form_Load()
Dim img As String
img = App.Path & "\library.jpg"
Librarian.Picture = LoadPicture(img)
Call CMDD
Call EDITD
Frame1.Enabled = False
frmEditLib.Visible = False
Command5.Enabled = False
End Sub

Private Sub CMDD()
cmdFirst.Enabled = False
cmdLast.Enabled = False
cmdBack.Enabled = False
cmdNext.Enabled = False
End Sub

Private Sub CMDE()
cmdFirst.Enabled = True
cmdLast.Enabled = True
cmdBack.Enabled = True
cmdNext.Enabled = True
End Sub

Private Sub EDITD()
cmdAdd.Enabled = False
cmdEdit.Enabled = False
cmdDel.Enabled = False
cmdAdd.Visible = False
cmdEdit.Visible = False
cmdDel.Visible = False
End Sub

Private Sub EDITE()
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cmdDel.Enabled = True
cmdAdd.Visible = True
cmdEdit.Visible = True
cmdDel.Visible = True
End Sub

Private Sub Label4_Click()
If txtNEW.PasswordChar = "*" Then
txtNEW.PasswordChar = ""
Label4.Caption = "Hide"
Else
txtNEW.PasswordChar = "*"
Label4.Caption = "Show"
End If
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
