VERSION 5.00
Begin VB.Form time 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Time"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "time.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset All"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Insert Time"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "time.frx":08CA
      Left            =   2880
      List            =   "time.frx":08D4
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "time.frx":08E0
      Left            =   1680
      List            =   "time.frx":0998
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "time.frx":0A8C
      Left            =   480
      List            =   "time.frx":0AB4
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Time and Click Insert Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If p_booking_information.Text4.Locked = False Then
p_booking_information.Text4.Text = time.Combo1.Text + " : " + time.Combo2.Text + "  " + time.Combo3.Text
Unload Me
p_booking_information.Text4.Locked = True
End If
If p_booking_information.Text5.Locked = False Then
p_booking_information.Text5.Text = time.Combo1.Text + " : " + time.Combo2.Text + "  " + time.Combo3.Text
Unload Me
p_booking_information.Text5.Locked = True
End If
End Sub

Private Sub Command2_Click()
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
