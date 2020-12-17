VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Return_ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmReturn 
      BackColor       =   &H00C0C0C0&
      Height          =   5415
      Left            =   2760
      TabIndex        =   22
      Top             =   960
      Width           =   6615
      Begin VB.CommandButton cmdB 
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
         Left            =   4440
         TabIndex        =   36
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return Book"
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
         Left            =   2760
         TabIndex        =   23
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label lblDateR 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   35
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label lblReturne 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   34
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label lblDateD 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   33
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label lblDateB 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   32
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   31
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Returned Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   29
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Returned:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   28
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Barrow Due Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Borrow:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   26
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Book Title:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame frmLogin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Login"
      Height          =   3015
      Left            =   3000
      TabIndex        =   11
      Top             =   2040
      Width           =   6015
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
         Top             =   2040
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
         Left            =   1080
         TabIndex        =   14
         Top             =   2040
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
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1200
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
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblPassVis 
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   1200
         Width           =   495
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Left            =   600
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User"
      Height          =   6255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   11055
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
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
         Left            =   9720
         TabIndex        =   10
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
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
         Left            =   8520
         TabIndex        =   9
         Top             =   5040
         Width           =   975
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
         Left            =   3840
         TabIndex        =   8
         Top             =   5040
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
         Left            =   2640
         TabIndex        =   7
         Top             =   5040
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
         Left            =   1440
         TabIndex        =   6
         Top             =   5040
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
         TabIndex        =   5
         Top             =   5040
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Book Borrowed History"
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   10815
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2895
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   5106
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
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Unreturned Books Total:"
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
         Left            =   6960
         TabIndex        =   21
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblBorrowNo 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   9480
         TabIndex        =   20
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblUName 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Return_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dates As String

Private Sub cmdB_Click()
frmReturn.Enabled = False
frmReturn.Visible = False
End Sub

Private Sub cmdCan_Click()
Login.Show
Unload Me
End Sub

Private Sub cmdFirst_Click()
rs.MoveFirst
End Sub

Private Sub cmdLast_Click()
rs.MoveLast
End Sub

Private Sub cmdBack_Click()
If rs.BOF Then
MsgBox "You already reached the first record", vbInformation
rs.MoveNext
Else
rs.MovePrevious
End If
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
rs.Open "Select * from Account where userID='" + txtUser.Text + "' and password='" + txtPass.Text + "'", con, adOpenDynamic, adLockBatchOptimistic


If rs.EOF Then
txtUser.Text = ""
txtPass.Text = ""
MsgBox "Incorrect User/Password!", vbCritical, "Error!"

Else

dates = DateAdd("d", 10, Date)
lblUName.Caption = rs("Full_Name")
lblBorrowNo.Caption = rs("borrowno")

frmLogin.Visible = False
Call eni
Call connect
rs.Open "Select * from Borrow where Full_Name like" & "'%" & lblUName.Caption & "%'", con, adOpenStatic, adLockBatchOptimistic
rs.Sort = "Book_Returned"
Set DataGrid1.DataSource = rs
Command1.Caption = "LOGOUT"

End If
End Sub

Private Sub cmdReturn_Click()
If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Return") = vbYes Then
MsgBox "Successfully Returned Book", , "Return"

rs("Book_Returned") = "Yes"
rs("DateReturned") = DateValue(Now)
rs.Update
rs.UpdateBatch

Dim addd As Integer
Dim adddd As Integer

Call connect
rs.Open "Select * from Account where Full_Name like" & "'%" & lblUName.Caption & "%'", con, adOpenStatic, adLockBatchOptimistic

addd = rs("borrowno")
adddd = addd - 1

rs("borrowno") = adddd
rs.Update
rs.UpdateBatch
lblBorrowNo.Caption = rs("borrowno")

Call connect
rs.Open "Select * from Borrow where Full_Name like" & "'%" & lblUName.Caption & "%'", con, adOpenStatic, adLockBatchOptimistic
Set DataGrid1.DataSource = rs

Call connect
rs.Open "Select * from Borrow where Full_Name like" & "'%" & lblUName.Caption & "%'", con, adOpenStatic, adLockBatchOptimistic
rs.Sort = "Book_Returned"
Set DataGrid1.DataSource = rs

frmReturn.Visible = False
End If
End Sub

Private Sub cmdSelect_Click()
If rs.EOF = True Then

MsgBox "No Selected from the List!", , "Select"

ElseIf rs.BOF = True Then

MsgBox "No Selected from the List!", , "Select"

Else
frmReturn.Enabled = True
frmReturn.Visible = True

lblName.Caption = rs("Full_Name")
lblTitle.Caption = rs("Title")
lblDateB.Caption = rs("DateStart")
lblDateD.Caption = rs("DateEnd")

Dim ret As String
ret = rs("Book_Returned")

lblReturne.Caption = ret
lblDateR.Caption = rs("DateReturned")

If ret = "Yes" Then
cmdReturn.Visible = False
ElseIf ret = "No" Then
cmdReturn.Visible = True
End If

'If IsNull(rs("DateReturened")) = True Then

'lblDateR.Caption = "-"

'Else

'lblDateR.Caption = rs("DateReturned")

'End If

End If
End Sub

Private Sub cmdRefresh_Click()
Call connect
rs.Open "Select * from Borrow where Full_Name like" & "'%" & lblUName.Caption & "%'", con, adOpenStatic, adLockBatchOptimistic
rs.Sort = "Book_Returned"
Set DataGrid1.DataSource = rs
cmdSelect.Enabled = True
End Sub

Private Sub Command1_Click()
If Command1.Caption = "LOGOUT" Then
If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Logout") = vbYes Then
Unload Me
Login.Show
End If
Else
Unload Me
Login.Show
End If

End Sub

Private Sub Form_Load()
Call des
Dim img As String
img = App.Path & "\library.jpg"
Return_.Picture = LoadPicture(img)
End Sub

Private Sub des()
Frame1.Enabled = False
frmReturn.Visible = False
frmReturn.Enabled = False
End Sub

Private Sub eni()
Frame1.Enabled = True
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
