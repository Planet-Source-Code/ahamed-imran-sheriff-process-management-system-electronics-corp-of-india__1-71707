VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Durdans Hospitals (Pvt) Ltd. - Hospital Management System"
   ClientHeight    =   10290
   ClientLeft      =   180
   ClientTop       =   750
   ClientWidth     =   14970
   Picture         =   "MDIforREPORTS.frx":0000
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpFile        =   "G:\MyVB\VB Project\Forms\Help.HLP"
   End
   Begin VB.PictureBox picRightNavigation 
      Align           =   3  'Align Left
      Height          =   8895
      Left            =   0
      Picture         =   "MDIforREPORTS.frx":1D78B
      ScaleHeight     =   8835
      ScaleWidth      =   3405
      TabIndex        =   3
      Top             =   1020
      Width           =   3465
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Engine"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   5
         Left            =   1200
         TabIndex        =   19
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label lblUserAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "User Account Panel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   720
         TabIndex        =   16
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label lblRecordExplorer 
         BackStyle       =   0  'Transparent
         Caption         =   "Record Explorer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   840
         TabIndex        =   15
         Top             =   200
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         X1              =   0
         X2              =   3360
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Label lblSolutionsProvider 
         BackStyle       =   0  'Transparent
         Caption         =   "Powered By : UDS Labs Inc.,"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   420
         Left            =   360
         TabIndex        =   14
         Top             =   8520
         Width           =   2775
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Log Off / Exit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   7
         Left            =   1200
         TabIndex        =   13
         Top             =   7500
         Width           =   2055
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   6
         Left            =   1200
         TabIndex        =   12
         Top             =   6600
         Width           =   2055
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Reports Quick Launch"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   4
         Left            =   1200
         TabIndex        =   11
         Top             =   4185
         Width           =   2055
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Payments"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   3
         Left            =   1200
         TabIndex        =   10
         Top             =   3435
         Width           =   1935
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Channeling Services"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   2
         Left            =   1200
         TabIndex        =   9
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Outpatients"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Inpatients"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   1000
         Width           =   2535
      End
      Begin VB.Image imgUserAccount 
         Height          =   3465
         Left            =   0
         Picture         =   "MDIforREPORTS.frx":261BC
         Top             =   5475
         Width           =   3405
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Record Explorer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   555
         Left            =   360
         TabIndex        =   6
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Image imgRecordExplorer 
         Height          =   5460
         Left            =   0
         Picture         =   "MDIforREPORTS.frx":2FF0C
         Top             =   0
         Width           =   3405
      End
   End
   Begin VB.PictureBox picTopNavigation 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      Picture         =   "MDIforREPORTS.frx":3D576
      ScaleHeight     =   960
      ScaleMode       =   0  'User
      ScaleWidth      =   15115.58
      TabIndex        =   1
      Top             =   0
      Width           =   14970
      Begin VB.CommandButton Command5 
         Height          =   975
         Left            =   43523
         Picture         =   "MDIforREPORTS.frx":462E5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   435
         Index           =   9
         Left            =   14040
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image5 
         Height          =   765
         Left            =   13200
         Picture         =   "MDIforREPORTS.frx":495C9
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Log Off"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   435
         Index           =   8
         Left            =   12240
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image4 
         Height          =   735
         Left            =   11280
         Picture         =   "MDIforREPORTS.frx":4F11C
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblDateTime 
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   540
         Width           =   4335
      End
      Begin VB.Label lblDesignation 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   280
         Width           =   2295
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   120
         Picture         =   "MDIforREPORTS.frx":54B16
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar BottomStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9915
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
            Text            =   "Durdans Hospital Management System "
            TextSave        =   "Durdans Hospital Management System "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "AD000001"
            TextSave        =   "AD000001"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Administrator"
            TextSave        =   "Administrator"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuInpatients 
      Caption         =   "&Inpatients"
      Begin VB.Menu mnuPatientAdmission 
         Caption         =   "Patient Admission"
         Begin VB.Menu mnuInpatientsMaintenance 
            Caption         =   "Step 1 - In-patients Maintenance"
         End
         Begin VB.Menu mnuSeparator9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGuardiansMaintenance 
            Caption         =   "Step 2 - Guardians Maintenance"
         End
         Begin VB.Menu mnuSeparator10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRegisterAdmitPatient 
            Caption         =   "Step 3 - Register / Admit Patient"
         End
         Begin VB.Menu mnuSeparator11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewPatientListings 
            Caption         =   "View Patient Listings"
         End
      End
      Begin VB.Menu mnuSeparator12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreatment 
         Caption         =   "Treatments"
         Begin VB.Menu mnuMedicalTreatmentsMaintenance 
            Caption         =   "Medical Treatments Maintenance"
            Begin VB.Menu mnuAddMedicalTreatments 
               Caption         =   "Add Medical Treatments"
            End
            Begin VB.Menu mnuEditMedicalTreatmentRecords 
               Caption         =   "Edit Medical Treatment Records"
            End
         End
         Begin VB.Menu mnuSeparator14 
            Caption         =   "-"
         End
         Begin VB.Menu mnuServiceTreatmentMaintenance 
            Caption         =   "Service Treatments Maintenance"
            Begin VB.Menu mnuAddServiceTreatments 
               Caption         =   "Add Service Treatments"
            End
            Begin VB.Menu mnuEditServiceTreatmentRecord 
               Caption         =   "Edit Service Treatment Records"
            End
         End
      End
      Begin VB.Menu mnuSeparator15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOveralPatientBill 
         Caption         =   "View Overall Patient Bill"
      End
      Begin VB.Menu mnuSeparator16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDischargePatient 
         Caption         =   "Discharge Patient"
      End
   End
   Begin VB.Menu mnuOutpatients 
      Caption         =   "&Outpatients"
      Begin VB.Menu mnuOutpatientsMaintenance 
         Caption         =   "Outpatients Maintenance"
      End
      Begin VB.Menu mnuSeparator17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreatments 
         Caption         =   "Treatments"
         Begin VB.Menu mnuMedicalTreatmentMaintenance 
            Caption         =   "Medical Treatments Maintenance"
            Begin VB.Menu mnuAddMedicalTreatment 
               Caption         =   "Add Medical Treatments"
            End
            Begin VB.Menu mnuEditMedicalTreatment 
               Caption         =   "Edit Medical Treatments"
            End
         End
         Begin VB.Menu mnuSeparator19 
            Caption         =   "-"
         End
         Begin VB.Menu mnuServiceTreatmentsMaintenance 
            Caption         =   "Service Treatments Maintenance"
            Begin VB.Menu mnuAddServiceTreatment 
               Caption         =   "Add Service Treatments"
            End
            Begin VB.Menu mnuEditServiceTreatmentRecords 
               Caption         =   "Edit Service Treatments Records"
            End
         End
      End
      Begin VB.Menu mnuSeparator20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOverallPatientBill 
         Caption         =   "View Overall Patient Bill"
      End
   End
   Begin VB.Menu mnuChanneling 
      Caption         =   "&Channeling"
      Begin VB.Menu mnuViewDoctorAppointments 
         Caption         =   "View Doctor's Appointments"
      End
      Begin VB.Menu mnuSeparator200 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManageAppointments 
         Caption         =   "Manage Appointments"
      End
   End
   Begin VB.Menu mnuPayments 
      Caption         =   "&Payments"
      Begin VB.Menu mnuInpatientTransactions 
         Caption         =   "Inpatients"
         Begin VB.Menu mnuManagePatientBill 
            Caption         =   "Manage Patient Bill"
         End
         Begin VB.Menu mnuSeparator48 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchPayments 
            Caption         =   "Search Payments"
         End
      End
      Begin VB.Menu mnuSeparator21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOutpatientsBilling 
         Caption         =   "Outpatients"
         Begin VB.Menu mnuManagePatientsBill 
            Caption         =   "Manage Patient Bill"
         End
         Begin VB.Menu mnuSeparator49 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOSearchPayments 
            Caption         =   "Search Payments"
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuPatientReports 
         Caption         =   "Patient Reports"
         Begin VB.Menu mnuInpatientsMasterReport 
            Caption         =   "Inpatients Master Report"
         End
         Begin VB.Menu mnuOutpatientsMasterReport 
            Caption         =   "Outpatients Master Report"
         End
         Begin VB.Menu mnuChannelingPatientsMasterReport 
            Caption         =   "Channeling Patients Master Report"
         End
         Begin VB.Menu mnuSeparator22 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPatientAdmissionMasterReport 
            Caption         =   "Patient Admission Master Report"
         End
         Begin VB.Menu mnuPatientDischargeMasterReport 
            Caption         =   "Patient Discharge Master Report"
         End
         Begin VB.Menu mnuSeparator23 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDoctorsVisitsMasterReport 
            Caption         =   "Doctors Visits Master Report"
         End
         Begin VB.Menu mnuPatientsMedicalTreatmentsMasterReport 
            Caption         =   "Patients' Medical Treatments Master Report"
         End
         Begin VB.Menu mnuPatientsMedicalServicesMasterReport 
            Caption         =   "Patients' Medical Services Master Report"
         End
      End
      Begin VB.Menu mnuSeparator24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDoctorReports 
         Caption         =   "Doctor Reports"
         Begin VB.Menu mnuAllDoctors 
            Caption         =   "Doctors Master Report"
         End
         Begin VB.Menu mnuAllDoctorsSchedulesMasterReport 
            Caption         =   "All Doctors' Schedules Master Report"
         End
         Begin VB.Menu mnuIndividualDoctorsScheduleReport 
            Caption         =   "Individual Doctor's Schedule Report"
         End
      End
      Begin VB.Menu mnuSeparator25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorporateReports 
         Caption         =   "Corporate Reports"
         Begin VB.Menu mnuAllCompaniesMasterReport 
            Caption         =   "All Companies Master Report"
         End
         Begin VB.Menu mnuIndividualCompanyPatientsReport 
            Caption         =   "Individual Company's Patients Report"
         End
      End
      Begin VB.Menu mnuSeparator26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHospitalReports 
         Caption         =   "Hospital Reports"
         Begin VB.Menu mnuMedicinesMasterReport 
            Caption         =   "Medicines Master Report"
         End
         Begin VB.Menu mnuMedicalServicesMasterReport 
            Caption         =   "Medical Services Master Report"
         End
         Begin VB.Menu mnuSeparator27 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDepartmentsMasterReport 
            Caption         =   "Departments Master Report"
         End
         Begin VB.Menu mnuRoomsMasterReport 
            Caption         =   "Rooms Master Report"
         End
      End
      Begin VB.Menu mnuSeparator28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminReports 
         Caption         =   "Log Reports"
         Begin VB.Menu mnuLoginDetailsReport 
            Caption         =   "Login Details Report"
         End
      End
      Begin VB.Menu mnSeparator29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRevenueReports 
         Caption         =   "Revenue Reports"
         Begin VB.Menu mnuInpatientsInvoice 
            Caption         =   "Inpatient's Invoice"
         End
         Begin VB.Menu mnuSeparator100 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInpatientsReceipt 
            Caption         =   "Inpatient's Receipt"
         End
         Begin VB.Menu mnuSeparator110 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOutpatientsReceipt 
            Caption         =   "Outpatient's Receipt"
         End
         Begin VB.Menu mnuSeparator111 
            Caption         =   "-"
         End
         Begin VB.Menu mnuChannelingPatientReceipt 
            Caption         =   "Channeling Patient's Receipt"
         End
         Begin VB.Menu mnuSeparator104 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInpatientsRevenueReports 
            Caption         =   "Inpatients Revenue Reports"
            Begin VB.Menu mnuWeeklyInpatientsRevenueReport 
               Caption         =   "Weekly Inpatients Revenue Report"
            End
            Begin VB.Menu mnuMonthlyInpatientsRevenueReport 
               Caption         =   "Monthly Inpatients Revenue Report"
            End
         End
         Begin VB.Menu mnuSeparator102 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOutpatientsRevenueReports 
            Caption         =   "Outpatients Revenue Reports"
            Begin VB.Menu mnuWeeklyOutpatientsRevenueReport 
               Caption         =   "Weekly Outpatients Revenue Report"
            End
            Begin VB.Menu mnuMonthlyOutpatientsRevenueReport 
               Caption         =   "Monthly Outpatients Revenue Report"
            End
         End
         Begin VB.Menu mnuSeparator103 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOverallRevenueReports 
            Caption         =   "Overall Revenue Reports"
            Begin VB.Menu mnuWeeklyOverallRevenueReport 
               Caption         =   "Weekly Overall Revenue Report"
            End
            Begin VB.Menu mnuMonthlyOverallRevenueReport 
               Caption         =   "Monthly Overall Revenue Report"
            End
         End
         Begin VB.Menu mnuSeparator106 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRevenueComparisonReports 
            Caption         =   "Revenue Comparison Reports"
            Begin VB.Menu mnuWeeklyRevenueReport 
               Caption         =   "WeeklyRevenue Report"
            End
            Begin VB.Menu mnuMonthlyRevenueReport 
               Caption         =   "Monthly Revenue Report"
            End
         End
         Begin VB.Menu mnuSeparator107 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDoctorEarningsReports 
            Caption         =   "Doctor's Earnings Reports"
            Begin VB.Menu mnuWeeklyDoctorEarningsReport 
               Caption         =   "Weekly Doctor's Earnings Report"
            End
            Begin VB.Menu mnuMonthlyDoctorEarningsReport 
               Caption         =   "Monthly Doctor's Earnings Report"
            End
         End
         Begin VB.Menu mnuSeparator108 
            Caption         =   "-"
         End
         Begin VB.Menu AgingReport 
            Caption         =   "Aging Report"
         End
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuDoctorMaintenance 
         Caption         =   "Doctor's Maintenance"
         Begin VB.Menu mnuDoctorsMaintenance 
            Caption         =   "Doctor's Maintenance"
         End
         Begin VB.Menu mnuSeparator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDoctorScheduleMaintenance 
            Caption         =   "Doctor's Schedule Maintenance"
         End
         Begin VB.Menu mnuSeparator2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewDoctorListings 
            Caption         =   "View Doctor Listings"
         End
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicinesMaintenance 
         Caption         =   "Medicines Maintenance"
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHopitalServicesMaintenance 
         Caption         =   "Hospital Services Maintenance"
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepartmentsMaintenance 
         Caption         =   "Departments Maintenance"
      End
      Begin VB.Menu mnuSeparator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWardsMaintenance 
         Caption         =   "Wards Maintenance"
      End
      Begin VB.Menu mnuSeparator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRoomsMaintenance 
         Caption         =   "Rooms Maintenance"
         Begin VB.Menu mnuRoomMaintenance 
            Caption         =   "Rooms Maintenance"
         End
         Begin VB.Menu mnuSeparator45 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewRoomListings 
            Caption         =   "View Rooms Listings"
         End
      End
      Begin VB.Menu mnuSeparator8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorporateMaintenance 
         Caption         =   "Corporate Maintenance"
         Begin VB.Menu mnuCompanyMaintenance 
            Caption         =   "Company Maintenance"
         End
         Begin VB.Menu mnuSeparator46 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewCorporateListing 
            Caption         =   "View Corporate Listings"
         End
      End
   End
   Begin VB.Menu mnuUserAccount 
      Caption         =   "&User Account"
      Begin VB.Menu mnuManageUserAccounts 
         Caption         =   "Manage User Accounts"
      End
      Begin VB.Menu mnuSeparator98 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuSeparator97 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuSeparator101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSearchEngine 
         Caption         =   "Search Engine"
      End
      Begin VB.Menu mnuSeparator96 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMicrosoftMagnifier 
         Caption         =   "Microsoft Magnifier"
      End
      Begin VB.Menu mnuSeparator95 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMicrosoftNarrator 
         Caption         =   "Microsoft Narrator"
      End
      Begin VB.Menu mnuSeparator94 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemMediaPlayer 
         Caption         =   "System Media Player"
      End
      Begin VB.Menu mnuSeparator93 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalendar 
         Caption         =   "Calendar"
      End
      Begin VB.Menu mnuSeparator92 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemCalculator 
         Caption         =   "System Calculator"
      End
      Begin VB.Menu mnuSeparator91 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemNotepad 
         Caption         =   "System Notepad"
      End
      Begin VB.Menu mnuSeparator90 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackupDatabase 
         Caption         =   "Backup Database"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuSeparator88 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuSeparator87 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuSeparator86 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSeparator85 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "Credits"
      End
      Begin VB.Menu mnuSeparator84 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpFile 
         Caption         =   "Help File"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

End Sub

Private Sub mnuAllCompaniesMasterReport_Click()
    RptAllCompaniesMaster.Show
End Sub

Private Sub mnuAllDoctors_Click()
    RptDoctorsMaster.Show
End Sub

Private Sub mnuAllDoctorsSchedulesMasterReport_Click()
    RptAllDoctorsShedule.Show
End Sub

Private Sub mnuChannelingPatientsMasterReport_Click()
'frmReportChannelingMaster
End Sub

Private Sub mnuDepartmentsMasterReport_Click()
    RptDepartmentMaster.Show
End Sub

Private Sub mnuDoctorsVisitsMasterReport_Click()
    RptVisitingDoctorsShedule.Show
End Sub

'RECHECK THIS CODE ACCORDING TO NAME CHANGE
Private Sub mnuIndividualDoctorsScheduleReport_Click()
    RptAllDoctorsShedule.Show
End Sub

Private Sub mnuInpatientsInvoice_Click()
    frmReportInpatientInvoice.Show
End Sub

Private Sub mnuInpatientsMasterReport_Click()
    frmReportInpatientMaster.Show
End Sub

Private Sub mnuMedicalServicesMasterReport_Click()
    RptMedicalServicesMaster.Show
End Sub

Private Sub mnuMedicinesMasterReport_Click()
    RptMedicinesMaster.Show
End Sub

Private Sub mnuOutpatientsMasterReport_Click()
    frmReportOutpatientMaster.Show
End Sub

Private Sub mnuPatientAdmissionMasterReport_Click()
    frmReportPatientAdmission.Show
End Sub

Private Sub mnuPatientDischargeMasterReport_Click()
    frmReportPatientDischarge.Show
End Sub

Private Sub mnuRoomsMasterReport_Click()
    RptRoomsMaster.Show
End Sub
