VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form search 
   Caption         =   "Search & Status"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "search.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "search.frx":08CA
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   5640
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2004
      Month           =   10
      Day             =   14
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Prodeed"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   3360
      ScaleHeight     =   5865
      ScaleWidth      =   7905
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton Command4 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   4440
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Search"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3480
         TabIndex        =   9
         Top             =   600
         Width           =   3465
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   3465
      End
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   600
         TabIndex        =   7
         Top             =   2040
         Width           =   6735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "The  Following are the data found  as per requirement"
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
         TabIndex        =   19
         Top             =   1680
         Width           =   6735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
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
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
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
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   6960
         Picture         =   "search.frx":146CC6
         Top             =   600
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Choose "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Contacts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   2415
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Booking "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seats Availability"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search && Status"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn1, cn2, cn3  As ADODB.Connection
Dim rs1, rs2, rs3 As ADODB.Recordset

Private Sub Calendar1_Click()
Text2.Text = Calendar1.Month & "/" & Calendar1.Day & "/" & Calendar1.Year
Calendar1.Visible = False
End Sub

Private Sub Command1_Click()

If Option1.Value = True Then
    Picture1.Visible = True
    Label2.Caption = " Bus Code"
    Label3.Caption = " Transport Date"
    Text2.Locked = True
    Image1.Visible = True
    Text1.SetFocus
End If

If Option2.Value = True Then
    Picture1.Visible = True
    Label2.Caption = " Booking Number"
    Label3.Caption = " Booking Date"
    Text2.Locked = True
    Image1.Visible = True
    Text1.SetFocus
End If

If Option3.Value = True Then
    Picture1.Visible = True
    Label2.Caption = " Employee Number"
    Label3.Caption = " First Name"
    Text1.SetFocus
End If

If Option4.Value = True Then
    Picture1.Visible = True
    Label2.Caption = " Contact Code"
    Label3.Caption = " First Name"
    Text1.SetFocus
End If

End Sub

Private Sub Command2_Click()

If Option1.Value = True Then
On Error Resume Next
    If Check1.Value = 1 Or Check2.Value = 1 Then
        load_booking1
        List1.clear
        l = rs1.RecordCount
        rs1.MoveFirst
        If l <> 0 Then
        List1.AddItem ("Booking No  Bus Code  Seat No")
        While (Not (rs1.EOF))
        If rs1.Fields(9) = Text1.Text Or rs1.Fields(3) = Text2.Text Then
        List1.AddItem (rs1.Fields(0) & "    " & rs1.Fields(9) & "    " & rs1.Fields(10))
        End If
        rs1.MoveNext
        Wend
        End If
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
        load_booking1
        List1.clear
        l = rs1.RecordCount
        rs1.MoveFirst
        If l <> 0 Then
        List1.AddItem ("Booking No  Bus Code  Seat No")
        While (Not (rs1.EOF))
        If rs1.Fields(3) = Text2.Text And rs1.Fields(9) = Text1.Text Then
        List1.AddItem (rs1.Fields(0) & "    " & rs1.Fields(9) & "    " & rs1.Fields(10))
        End If
        rs1.MoveNext
        Wend
        End If
    End If
End If

If Option2.Value = True Then
    If Check1.Value = 1 Or Check2.Value = 1 Then
        load_booking1
        List1.clear
        l = rs1.RecordCount
        rs1.MoveFirst
        If l <> 0 Then
        List1.AddItem ("Booking No   Booking Date  Name")
        While (Not (rs1.EOF))
        If rs1.Fields(0) = Text1.Text Or rs1.Fields(2) = Text2.Text Then
        List1.AddItem (rs1.Fields(0) & "     " & rs1.Fields(2) & "       " & rs1.Fields(6))
        End If
        rs1.MoveNext
        Wend
        End If
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
        load_booking1
        List1.clear
        l = rs1.RecordCount
        rs1.MoveFirst
        If l <> 0 Then
        List1.AddItem ("Booking No  Bus Code  Seat No")
        While (Not (rs1.EOF))
        If rs1.Fields(0) = Text1.Text Or rs1.Fields(2) = Text2.Text Then
        List1.AddItem (rs1.Fields(0) & "     " & rs1.Fields(2) & "       " & rs1.Fields(6))
        End If
        rs1.MoveNext
        Wend
        End If
    End If
End If

If Option3.Value = True Then
    If Check1.Value = 1 Or Check2.Value = 1 Then
        load_employee
        List1.clear
        l = rs2.RecordCount
        rs2.MoveFirst
        If l <> 0 Then
        List1.AddItem ("Emp No       Name        Designation")
        While (Not (rs2.EOF))
        If rs2.Fields(0) = Text1.Text Or rs2.Fields(1) = Text2.Text Then
        List1.AddItem (rs2.Fields(0) & "       " & rs2.Fields(1) & "        " & rs2.Fields(3))
        End If
        rs2.MoveNext
        Wend
        End If
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
        load_employee
        List1.clear
        l = rs2.RecordCount
        rs2.MoveFirst
        If l <> 0 Then
        List1.AddItem ("Emp No       Name        Designation")
        While (Not (rs2.EOF))
        If rs2.Fields(0) = Text1.Text And rs2.Fields(1) = Text2.Text Then
        List1.AddItem (rs2.Fields(0) & "       " & rs2.Fields(1) & "        " & rs2.Fields(3))
        End If
        rs2.MoveNext
        Wend
        End If
    End If
End If

If Option4.Value = True Then
    If Check1.Value = 1 Or Check2.Value = 1 Then
        load_contact
        List1.clear
        rs3.MoveFirst
        l = rs3.RecordCount
        rs3.MoveFirst
        If l <> 0 Then
        List1.AddItem ("Contact Code  Name          Designation")
        While (Not (rs3.EOF))
        If rs3.Fields(0) = Text1.Text Or rs3.Fields(1) = Text2.Text Then
        List1.AddItem (rs3.Fields(0) & "         " & rs3.Fields(1) & "         " & rs3.Fields(3))
        End If
        rs3.MoveNext
        Wend
        End If
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
        load_contact
        List1.clear
        rs3.MoveFirst
        l = rs3.RecordCount
        rs3.MoveFirst
        If l <> 0 Then
        List1.AddItem ("Contact Code  Name          Designation")
        While (Not (rs3.EOF))
        If rs3.Fields(0) = Text1.Text And rs3.Fields(1) = Text2.Text Then
        List1.AddItem (rs3.Fields(0) & "         " & rs3.Fields(1) & "         " & rs3.Fields(3))
        End If
        rs3.MoveNext
        Wend
        End If
    End If
End If

If Err Then
MsgBox Err.Number + Err.Description
End If
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
List1.clear
Text1.SetFocus
If Check2.Value = Checked Then
Check2.Value = Unchecked
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Check1.Value = 1
End Sub

Private Sub Image1_Click()
Calendar1.Visible = True
Calendar1.Top = 2160
Calendar1.Left = 6840
End Sub

Private Sub Option1_Click()
Picture1.Visible = False
Image1.Visible = False
End Sub

Private Sub Option2_Click()
Picture1.Visible = False
Image1.Visible = False
End Sub

Private Sub Option3_Click()
Picture1.Visible = False
Image1.Visible = False
End Sub

Private Sub Option4_Click()
Picture1.Visible = False
Image1.Visible = False
End Sub

Public Sub clear()
Text1.Text = ""
Text2.Text = ""
End Sub

Public Sub load_booking1()
On Error Resume Next
Set cn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
cn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TRANSPORT.MDB;Persist Security Info=False"
rs1.Open "select * from booking", cn1, adOpenDynamic, adLockPessimistic
rs1.MoveFirst
If Err Then
MsgBox Err.Number + Err.Description
End If
End Sub

Public Sub load_employee()
On Error Resume Next
Set cn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TRANSPORT.MDB;Persist Security Info=False"
rs2.Open "select * from employee", cn2, adOpenDynamic, adLockPessimistic
rs2.MoveFirst
If Err Then
MsgBox Err.Number + Err.Description
End If
End Sub

Public Sub load_contact()
On Error Resume Next
Set cn3 = New ADODB.Connection
Set rs3 = New ADODB.Recordset
cn3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TRANSPORT.MDB;Persist Security Info=False"
rs3.Open "select * from contacts", cn3, adOpenDynamic, adLockPessimistic
rs3.MoveFirst
If Err Then
MsgBox Err.Number + Err.Description
End If
End Sub
