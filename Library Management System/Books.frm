VERSION 5.00
Begin VB.Form Books 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Book Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Publication Date: (MM/DD/YYYY)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
End
Attribute VB_Name = "Books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


If Command1.Caption = "Save" Then
If MsgBox("Are you sure you want to Save?", vbQuestion + vbYesNo) = vbYes Then
rs.AddNew
rs("BookNo") = Text1.Text
rs("Title") = Text2.Text
rs("Author") = Text3.Text
rs("Publication_Date") = Text4.Text
rs.Update
rs.UpdateBatch

MsgBox "Record Saved", vbInformation

Unload Me
End If

ElseIf Command1.Caption = "Update" Then
If MsgBox("Are you sure you want to Update Record?", vbQuestion + vbYesNo) = vbYes Then

rs("BookNo") = Text1.Text
rs("Title") = Text2.Text
rs("Author") = Text3.Text
rs("Publication_Date") = Text4.Text
rs.Update
rs.UpdateBatch

MsgBox "Record Updated", vbInformation
Unload Me

End If
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

