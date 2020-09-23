VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogalaxy Search"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AGSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "View Queue"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Audiogalaxy Homepage"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   280
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   100
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Search for:"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   140
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "" Then
MsgBox "Select a search type!", vbOKOnly, "Search error"
Exit Sub
End If

If Combo1.Text = "FTP Search" Then
SearchFTP Form1
Else
SearchMusic Form1
End If
End Sub

Private Sub Command2_Click()
VisitAG Form1
End Sub

Private Sub Command3_Click()
VisitQ Form1
End Sub

Private Sub Form_Load()
Combo1.AddItem "Music Search"
Combo1.AddItem "FTP Search"
End Sub
