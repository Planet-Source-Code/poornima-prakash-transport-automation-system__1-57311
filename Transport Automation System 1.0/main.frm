VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm main 
   BackColor       =   &H8000000C&
   Caption         =   " Transport Automation System"
   ClientHeight    =   3195
   ClientLeft      =   210
   ClientTop       =   1215
   ClientWidth     =   4680
   Icon            =   "main.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "booking"
            Object.ToolTipText     =   "Booking Information"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "contact"
            Object.ToolTipText     =   "Contacts Manager"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "vehicle"
            Object.ToolTipText     =   "Vechicle Manager"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "employee"
            Object.ToolTipText     =   "Employee Details"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "service"
            Object.ToolTipText     =   "Service Information"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":145C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2156
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":387A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":46DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":544E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu m_login_manager 
      Caption         =   "&Login Manager"
      Begin VB.Menu m_lock_screen 
         Caption         =   "Lock Screen"
      End
      Begin VB.Menu m_sep_0 
         Caption         =   "-"
      End
      Begin VB.Menu m_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu m_application_manager 
      Caption         =   "&Application Manager"
      Begin VB.Menu m_booking_information 
         Caption         =   "Booking Information"
      End
      Begin VB.Menu m_sep_1 
         Caption         =   "-"
      End
      Begin VB.Menu m_contact_manager 
         Caption         =   "Contact Manager"
      End
      Begin VB.Menu m_employee_information 
         Caption         =   "Employee Information"
      End
      Begin VB.Menu m_sep_2 
         Caption         =   "-"
      End
      Begin VB.Menu m_vehicle_manager 
         Caption         =   "Vehicle Manager"
      End
      Begin VB.Menu m_services_information 
         Caption         =   "Services Information"
      End
   End
   Begin VB.Menu m_security_settings 
      Caption         =   "&Security Settings"
      Begin VB.Menu m_user_profiles 
         Caption         =   "User Profiles"
      End
   End
   Begin VB.Menu m_search_and_status 
      Caption         =   "S&earch && Status"
      Begin VB.Menu m_seats_availability 
         Caption         =   "Seats Availability"
      End
      Begin VB.Menu m_sep_3 
         Caption         =   "-"
      End
      Begin VB.Menu m_search_booking 
         Caption         =   "Search Booking"
      End
      Begin VB.Menu m_search_employee 
         Caption         =   "Search Employee"
      End
      Begin VB.Menu m_search_contact 
         Caption         =   "Search Contact"
      End
   End
   Begin VB.Menu m_reports 
      Caption         =   "&Reports"
      Begin VB.Menu m_booking_list 
         Caption         =   "Booking List"
      End
      Begin VB.Menu m_sep5 
         Caption         =   "-"
      End
      Begin VB.Menu m_contact_list 
         Caption         =   "Contacts List"
      End
      Begin VB.Menu m_employee_list 
         Caption         =   "Employee List"
      End
      Begin VB.Menu m_sep6 
         Caption         =   "-"
      End
      Begin VB.Menu m_vehicle_list 
         Caption         =   "Vehicle List"
      End
      Begin VB.Menu m_services_list 
         Caption         =   "Services List"
      End
   End
   Begin VB.Menu m_about 
      Caption         =   "A&bout"
      Begin VB.Menu m_about_tas 
         Caption         =   "About TAS"
      End
      Begin VB.Menu m_sep_4 
         Caption         =   "-"
      End
      Begin VB.Menu m_credits 
         Caption         =   "Credits"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub m_contacts_manager_Click()
contact_manager.Show
End Sub

Private Sub m_about_tas_Click()
about.Show
End Sub

Private Sub m_booking_information_Click()
p_booking_information.Show
End Sub

Private Sub m_booking_list_Click()
bookingreport.Show
End Sub

Private Sub m_contact_list_Click()
contactsreport.Show
End Sub

Private Sub m_contact_manager_Click()
contact_manager.Show
End Sub

Private Sub m_credits_Click()
credits.Show
End Sub

Private Sub m_employee_information_Click()
employee.Show
End Sub

Private Sub m_employee_list_Click()
employeereport.Show
End Sub

Private Sub m_exit_Click()
End
End Sub

Private Sub mpassenger_booking_Click()
p_booking_information.Show
End Sub

Private Sub m_lock_screen_Click()
main.Hide
authentication.Show
End Sub

Private Sub m_search_booking_Click()
search.Show
search.Option2.Value = True
End Sub

Private Sub m_search_contact_Click()
search.Show
search.Option4.Value = True
End Sub

Private Sub m_search_employee_Click()
search.Show
search.Option3.Value = True
End Sub

Private Sub m_seats_availability_Click()
search.Show
search.Option1.Value = True
End Sub

Private Sub m_services_information_Click()
services.Show
End Sub

Private Sub m_services_list_Click()
servicereport.Show
End Sub

Private Sub m_user_profiles_Click()
user_profiles.Show
End Sub

Private Sub m_vehicle_list_Click()
vehiclereport.Show
End Sub

Private Sub m_vehicle_manager_Click()
vehicle.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
'To Select the required application from the designed toolbar
'The Key of the button in the toolbar is used for calling it
Select Case Button.Key
   Case "booking"
      p_booking_information.Show
   Case "contact"
      contact_manager.Show
   Case "employee"
      employee.Show
    Case "vehicle"
      vehicle.Show
    Case "service"
      services.Show
End Select
End Sub

