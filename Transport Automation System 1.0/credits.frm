VERSION 5.00
Begin VB.Form credits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Credits"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "credits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "credits.frx":08CA
   ScaleHeight     =   2790
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "poorniprakash85@rediffmail.com"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Poornima Prakash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Programmed By"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport Automation System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
