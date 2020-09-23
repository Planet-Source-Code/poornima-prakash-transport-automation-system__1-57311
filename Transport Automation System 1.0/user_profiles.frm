VERSION 5.00
Begin VB.Form user_profiles 
   Caption         =   "User Profiles"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "user_profiles.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "user_profiles.frx":08CA
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1440
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3960
         PasswordChar    =   "0"
         TabIndex        =   11
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3960
         PasswordChar    =   "0"
         TabIndex        =   10
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   5160
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Update"
         Height          =   495
         Left            =   2160
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Reset"
         Height          =   495
         Left            =   3720
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   8655
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Choose this option to rename the user name"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   5415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Choose this option to change the user password"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Width           =   5295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Proceed"
         Height          =   495
         Left            =   3120
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Profiles"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "user_profiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim u, p As String
Private Sub Command1_Click()

If Option1.Value = True Then
Frame2.Top = 1560
Frame2.Visible = True
Label2.Caption = " Enter Old Password"
Label3.Caption = " Enter New Password"
Text1.SetFocus
End If

If Option2.Value = True Then
Frame2.Top = 1560
Frame2.Visible = True
Label2.Caption = " Enter Password"
Label3.Caption = " Enter New Username"
Text1.SetFocus
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()

If Option1.Value = True Then
    If Text1.Text = p Then
    Open App.Path & "\configtas.dll" For Output As #1
    Write #1, Encrypt(1234, Trim(u))
    Write #1, Encrypt(1234, Text2.Text)
    Close #1
    MsgBox "The Password Has Been Modified"
    Text1.Text = ""
    Text2.Text = ""
    Frame2.Visible = False
    Else
    MsgBox "Invalid Old Password"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    End If
End If

If Option2.Value = True Then
    If Text2.Text = "" Then
    MsgBox "Enter a valid Username"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    End If
    If Text1.Text = p And Text2.Text <> "" Then
    Open App.Path & "\configtas.dll" For Output As #1
    Write #1, Encrypt(1234, Text2.Text)
    Write #1, Encrypt(1234, Text1.Text)
    Close #1
    MsgBox "The Username Has Been Modified"
    Text1.Text = ""
    Text2.Text = ""
    Frame2.Visible = False
    Else
    If Text2.Text <> "" Then
    MsgBox "Invalid Old Password"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    End If
    End If
End If


End Sub

Private Sub Command5_Click()
Frame2.Visible = False
End Sub

Private Sub Form_Load()
    Option1.Value = True
    On Error Resume Next
    Open App.Path & "\configtas.dll" For Input As #1
    Input #1, u
    Input #1, p
    Close #1
    u = Decrypt(1234, Trim(u))
    p = Decrypt(1234, Trim(p))
    If Err Then
    MsgBox Err.Description
    End If
End Sub
