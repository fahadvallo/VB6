VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Borrow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   13935
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   30
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame frmLogin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Login"
      Height          =   3015
      Left            =   3960
      TabIndex        =   15
      Top             =   1440
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblPassVis 
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5040
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borrow"
      Height          =   4575
      Left            =   7920
      TabIndex        =   8
      Top             =   840
      Width           =   5895
      Begin VB.CommandButton cmdBorrow 
         Caption         =   "Borrow Book"
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
         Left            =   2280
         TabIndex        =   9
         Top             =   3840
         Width           =   1455
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
         Left            =   1800
         TabIndex        =   32
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Borrowed:"
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
         Left            =   360
         TabIndex        =   31
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblUName 
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
         Left            =   1800
         TabIndex        =   29
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblBookNo 
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
         Left            =   1800
         TabIndex        =   27
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Book No.:"
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
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblDate 
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
         Left            =   1800
         TabIndex        =   25
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label lblUsers 
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
         Left            =   1800
         TabIndex        =   24
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
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
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTitle 
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
         Left            =   1800
         TabIndex        =   14
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Book Title:"
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
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Book Selection"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7575
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
         Left            =   6360
         TabIndex        =   13
         Top             =   3840
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2655
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4683
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
         Left            =   5280
         TabIndex        =   7
         Top             =   3840
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
         Left            =   3480
         TabIndex        =   6
         Top             =   3840
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
         TabIndex        =   5
         Top             =   3840
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
         TabIndex        =   4
         Top             =   3840
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
         TabIndex        =   3
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Search: Title/Author/Date[MM/DD/YYYY]"
         Height          =   255
         Left            =   2760
         TabIndex        =   33
         Top             =   720
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Borrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dates As String

Private Sub cmdBack_Click()
If rs.BOF Then
MsgBox "You already reached the first record", vbInformation
rs.MoveNext
Else
rs.MovePrevious
End If
End Sub

Private Sub cmdBorrow_Click()
If lblBorrowNo.Caption = "5" Then
MsgBox "You Already have Borrowed 5 Books. Return Atleast 1 Book first before you borrow Again!", , "Greedy Ass"
Else
If lblTitle.Caption = "-" Then
MsgBox "Select Book First!", , "Borrow"
Else
If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Borrow") = vbYes Then
Dim addd As Integer
addd = lblBorrowNo.Caption + 1
Call connect
rs.Open "Select * from Account where Full_Name like" & "'%" & lblUName.Caption & "%'", con, adOpenStatic, adLockBatchOptimistic
rs("borrowno") = addd
rs.Update
rs.UpdateBatch
lblBorrowNo.Caption = rs("borrowno")
MsgBox "Thanks!, Please return the book after 10 days...", , "Success"
Call connect
rs.Open "Select * from Borrow order by Full_Name", con, adOpenStatic, adLockBatchOptimistic
rs.AddNew
rs("Full_Name") = lblUName.Caption
rs("Title") = lblTitle.Caption
rs("DateStart") = lblDate.Caption
rs("DateEnd") = dates
rs("Book_Returned") = "No"
rs("DateReturned") = "-"
rs.Update
rs.UpdateBatch
lblTitle.Caption = "-"
lblBookNo.Caption = "-"
cmdSelect.Enabled = False
End If
End If
End If
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
lblDate.Caption = DateValue(Now)
dates = DateAdd("d", 10, Date)

lblUName.Caption = rs("Full_Name")
lblBorrowNo.Caption = rs("borrowno")
lblUsers.Caption = txtUser.Text
frmLogin.Visible = False
Call ene
Call connect
rs.Open "Select * from Book order by Title", con, adOpenStatic, adLockBatchOptimistic
Set DataGrid1.DataSource = rs
Command1.Caption = "LOGOUT"

End If
End Sub

Private Sub cmdRefresh_Click()
Call connect
rs.Open "Select * from Book order by Title", con, adOpenStatic, adLockBatchOptimistic
Set DataGrid1.DataSource = rs
cmdSelect.Enabled = True
End Sub

Private Sub cmdSearch_Click()
Call connect
rs.Open "Select * from Book where Title = '" & Text1.Text & "' or Author = '" & Text1.Text & "' or Publication_Date = '" & Text1.Text & "'", con, adOpenUnspecified, adLockUnspecified
If rs.EOF = True Then
MsgBox "Not Found :<", , "Search"
Call connect
rs.Open "Select * from Book order by Title", con, adOpenStatic, adLockBatchOptimistic
End If
Set DataGrid1.DataSource = rs
End Sub

Private Sub cmdSelect_Click()
If rs.EOF = True Then
MsgBox "No Selected Book!", , "Select"
ElseIf rs.BOF = True Then
MsgBox "No Selected Book!", , "Select"
Else
lblTitle.Caption = rs("Title")
lblBookNo.Caption = rs("BookNo")
End If
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
Call dis
Dim img As String
img = App.Path & "\library.jpg"
Borrow.Picture = LoadPicture(img)
End Sub

Private Sub dis()
Frame1.Enabled = False
Frame2.Enabled = False
End Sub

Private Sub ene()
Frame1.Enabled = True
Frame2.Enabled = True
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

