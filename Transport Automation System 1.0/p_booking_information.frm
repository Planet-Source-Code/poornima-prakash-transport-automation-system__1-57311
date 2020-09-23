VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form p_booking_information 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Passenger Booking Information"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "p_booking_information.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "p_booking_information.frx":08CA
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   1200
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2004
      Month           =   10
      Day             =   6
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
      Left            =   1080
      TabIndex        =   33
      Top             =   5400
      Width           =   9615
      Begin VB.CommandButton Command7 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   8160
         TabIndex        =   40
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Previous"
         Height          =   495
         Left            =   6840
         TabIndex        =   39
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Next"
         Height          =   495
         Left            =   5520
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   4200
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Reset"
         Height          =   495
         Left            =   2880
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Update"
         Height          =   495
         Left            =   1560
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   495
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8040
      TabIndex        =   31
      Top             =   4560
      Width           =   3110
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8040
      TabIndex        =   29
      Top             =   4200
      Width           =   3110
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8040
      TabIndex        =   27
      Top             =   3840
      Width           =   3110
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8040
      TabIndex        =   25
      Top             =   3480
      Width           =   3110
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   1740
      Left            =   8040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   24
      Top             =   1680
      Width           =   3110
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8040
      TabIndex        =   21
      Top             =   1320
      Width           =   3110
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "p_booking_information.frx":146CC6
      Left            =   2760
      List            =   "p_booking_information.frx":146CC8
      TabIndex        =   19
      Top             =   4560
      Width           =   3135
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "p_booking_information.frx":146CCA
      Left            =   2760
      List            =   "p_booking_information.frx":146CCC
      TabIndex        =   17
      Top             =   4200
      Width           =   3135
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "p_booking_information.frx":146CCE
      Left            =   2760
      List            =   "p_booking_information.frx":146CD0
      TabIndex        =   15
      Top             =   3840
      Width           =   3135
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "p_booking_information.frx":146CD2
      Left            =   2760
      List            =   "p_booking_information.frx":146CD4
      TabIndex        =   13
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   2745
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Width           =   2745
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   2745
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   2745
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "p_booking_information.frx":146CD6
      Left            =   2760
      List            =   "p_booking_information.frx":146CE3
      TabIndex        =   4
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   3110
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Total Charge"
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
      TabIndex        =   32
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Booking Charge"
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
      Left            =   5880
      TabIndex        =   30
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Ticket Charge"
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
      Left            =   5880
      TabIndex        =   28
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label14 
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
      Height          =   420
      Left            =   5880
      TabIndex        =   26
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label13 
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
      Height          =   1860
      Left            =   5880
      TabIndex        =   23
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Name"
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
      Left            =   5880
      TabIndex        =   22
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Seat Number"
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
      TabIndex        =   20
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Bus Code"
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
      TabIndex        =   18
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Ending Place"
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
      TabIndex        =   16
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Starting Place"
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
      TabIndex        =   14
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   5520
      Picture         =   "p_booking_information.frx":146D03
      Top             =   3120
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Reaching Time"
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
      TabIndex        =   12
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   5520
      Picture         =   "p_booking_information.frx":148FF4
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Starting Time"
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
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   5520
      Picture         =   "p_booking_information.frx":14B2E5
      Top             =   2400
      Width           =   300
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Travel Date"
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
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   5520
      Picture         =   "p_booking_information.frx":14B79D
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Booking Date"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Category"
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
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Booking Number"
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
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Information"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "p_booking_information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bd, td As Boolean
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cn1 As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim cn2 As ADODB.Connection
Dim rs2 As ADODB.Recordset
Dim add As Boolean

Private Sub Calendar1_Click()
If bd = True Then
Text2.Text = Calendar1.Month & "/" & Calendar1.Day & "/" & Calendar1.Year
Calendar1.Visible = False
bd = False
End If
If td = True Then
Text3.Text = Calendar1.Month & "/" & Calendar1.Day & "/" & Calendar1.Year
Calendar1.Visible = False
td = False
End If
End Sub

Private Sub Combo4_Click()
On Error Resume Next
l = rs1.RecordCount
If l <> 0 Then
rs1.MoveFirst
While (Not (rs1.EOF))
If rs1.Fields(0) = Combo4.Text Then
t = rs1.Fields(1)
For i = 1 To t
Combo5.AddItem (i)
Next i
End If
rs1.MoveNext
Wend
End If
End Sub

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
Set cn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set cn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TRANSPORT.MDB;Persist Security Info=False"
rs.Open "select * from booking", cn, adOpenDynamic, adLockPessimistic
cn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TRANSPORT.MDB;Persist Security Info=False"
rs1.Open "select * from vehicle", cn, adOpenDynamic, adLockPessimistic
cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TRANSPORT.MDB;Persist Security Info=False"
rs2.Open "select * from services", cn, adOpenDynamic, adLockPessimistic
load_city
load_bus
rs.MoveFirst
display
If Err Then
MsgBox Err.Number + Err.Description
End If
End Sub

Private Sub Image1_Click()
bd = True
td = False
Calendar1.Left = 2760
Calendar1.Top = 2360
Calendar1.Visible = True
End Sub

Private Sub Image2_Click()
td = True
bd = False
Calendar1.Left = 2760
Calendar1.Top = 2720
Calendar1.Visible = True
End Sub

Private Sub Image3_Click()
p_booking_information.Text4.Locked = False
time.Show
End Sub

Private Sub Image4_Click()
p_booking_information.Text5.Locked = False
time.Show
End Sub

Public Sub load_city()
On Error Resume Next
l = rs2.RecordCount
If l <> 0 Then
rs2.MoveFirst
While (Not (rs2.EOF))
Combo2.AddItem (rs2.Fields(1))
Combo3.AddItem (rs2.Fields(1))
rs2.MoveNext
Wend
End If
End Sub

Public Sub load_bus()
On Error Resume Next
l = rs1.RecordCount
If l <> 0 Then
rs1.MoveFirst
While (Not (rs1.EOF))
Combo4.AddItem (rs1.Fields(0))
rs1.MoveNext
Wend
End If
End Sub

Public Sub display()
Text1.Text = rs.Fields(0)
Combo1.Text = rs.Fields(1)
Text2.Text = rs.Fields(2)
Text3.Text = rs.Fields(3)
Text4.Text = rs.Fields(14)
Text5.Text = rs.Fields(15)
Combo2.Text = rs.Fields(4)
Combo3.Text = rs.Fields(5)
Combo4.Text = rs.Fields(9)
Combo5.Text = rs.Fields(10)
Text6.Text = rs.Fields(6)
Text7.Text = rs.Fields(7)
Text8.Text = rs.Fields(8)
Text9.Text = rs.Fields(11)
Text10.Text = rs.Fields(12)
Text11.Text = rs.Fields(13)
End Sub

Public Sub clear()
Text1.Text = ""
Combo1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
End Sub

Public Sub save()
rs.Fields(0) = Text1.Text
rs.Fields(1) = Combo1.Text
rs.Fields(2) = Text2.Text
rs.Fields(3) = Text3.Text
rs.Fields(14) = Text4.Text
rs.Fields(15) = Text5.Text
rs.Fields(4) = Combo2.Text
rs.Fields(5) = Combo3.Text
rs.Fields(9) = Combo4.Text
rs.Fields(10) = Combo5.Text
rs.Fields(6) = Text6.Text
rs.Fields(7) = Text7.Text
rs.Fields(8) = Text8.Text
rs.Fields(11) = Text9.Text
rs.Fields(12) = Text10.Text
rs.Fields(13) = Text11.Text
End Sub

