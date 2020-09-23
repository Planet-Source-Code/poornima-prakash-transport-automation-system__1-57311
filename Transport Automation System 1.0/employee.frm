VERSION 5.00
Begin VB.Form employee 
   Caption         =   "Employee Information"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "employee.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "employee.frx":08CA
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   " Operations "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1200
      TabIndex        =   15
      Top             =   5160
      Width           =   9615
      Begin VB.CommandButton Command7 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   8160
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Previous"
         Height          =   495
         Left            =   6840
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Next"
         Height          =   495
         Left            =   5520
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Reset"
         Height          =   495
         Left            =   2880
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Update"
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8040
      TabIndex        =   13
      Top             =   3480
      Width           =   3110
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   1380
      Left            =   8040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   2040
      Width           =   3110
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      TabIndex        =   9
      Top             =   3480
      Width           =   3110
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      TabIndex        =   7
      Top             =   3120
      Width           =   3110
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   3110
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   3110
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Width           =   3110
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Mobile Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      TabIndex        =   14
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   5880
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Phone Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   10
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Designation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Last Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   " First Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Employee Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim add As Boolean

Private Sub Command1_Click()
add = True
clear
Text1.SetFocus
End Sub

Private Sub Command2_Click()
If add = True Then
add = False
rs.AddNew
save
rs.Update
MsgBox "The Record has been Saved"
display
Else
save
rs.Update
MsgBox "The Record has been Saved"
display
End If
End Sub

Private Sub Command3_Click()
clear
Text1.SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next
rs.Delete
clear
MsgBox "The Record has been Deleted"
rs.MoveFirst
display
If Err Then
MsgBox Err.Number + Err.Description
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
rs.MoveNext
display
If rs.EOF = True Then
MsgBox "This is the last record"
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
rs.MovePrevious
display
If rs.BOF = True Then
MsgBox "This is the first record"
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TRANSPORT.MDB;Persist Security Info=False"
rs.Open "select * from employee", cn, adOpenDynamic, adLockPessimistic
rs.MoveFirst
display
If Err Then
MsgBox Err.Number + Err.Description
End If
End Sub

Public Sub display()
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(5)
Text6.Text = rs.Fields(4)
Text7.Text = rs.Fields(6)
End Sub

Public Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Public Sub save()
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(5) = Text5.Text
rs.Fields(4) = Text6.Text
rs.Fields(6) = Text7.Text
End Sub

