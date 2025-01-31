VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Members 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Members"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17985
   LinkTopic       =   "Form1"
   Picture         =   "Members.frx":0000
   ScaleHeight     =   6525
   ScaleWidth      =   17985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox combog 
      Height          =   315
      ItemData        =   "Members.frx":9860
      Left            =   2760
      List            =   "Members.frx":986A
      TabIndex        =   15
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtgrade 
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtadd 
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox txtmi 
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtln 
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtfn 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Members.frx":987C
      Height          =   3615
      Left            =   5160
      TabIndex        =   2
      Top             =   1800
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "LRN"
         Caption         =   "LRN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "First Name"
         Caption         =   "First Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Last Name"
         Caption         =   "Last Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Middle Name"
         Caption         =   "Middle Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Gender"
         Caption         =   "Gender"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Grade"
         Caption         =   "Grade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Date of Membership"
         Caption         =   "Date of Membership"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc memado 
      Height          =   735
      Left            =   6600
      Top             =   9240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LEAN FILES\LEAN DATABASE\MEMBER.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LEAN FILES\LEAN DATABASE\MEMBER.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "MEMBER"
      Caption         =   "memado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtlrn 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblmemd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date of Membership  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label lblgrade 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grade                           :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label lbladd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Address                        :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblgen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gender                         :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblmi 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Middle Name               :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblln 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Last Name                    :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label lbllrn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LRN                               :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblfn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "First Name                    :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "Members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()

End Sub

Private Sub Text4_Change()

End Sub

Private Sub Command1_Click()
B = MsgBox("Are you sure you want to saved this information", vbInformation + vbYesNo, "Membership Registration")

If B = vbYes Then
With memado.Recordset
.AddNew
.Fields("LRN") = txtlrn.Text
.Fields("First Name") = txtfn.Text
.Fields("Last Name") = txtln.Text
.Fields("Middle Name") = txtmi.Text
.Fields("Gender") = combog.Text
.Fields("Address") = txtadd.Text
.Fields("Grade") = txtgrade.Text
.Fields("Date of Membership") = Date
.Update
memado.Recordset.Update
End With
End If

MsgBox "Record has been successfully saved!", vbInformation, "Membership"
txtlrn.Text = ""
txtfn.Text = ""
txtln.Text = ""
txtmi.Text = ""
combog.Text = ""
txtadd.Text = ""
txtgrade.Text = ""

End Sub

Private Sub Command3_Click()
C = MsgBox("Are you sure you want to delete this information", vbCritical + vbYesNo, "Delete Info")
If C = vbYes Then
memado.Recordset.Delete adAffectCurrent
MsgBox "Record has been successfully deleted!", vbInformation, "Delete Info"
memado.Recordset.Update
memado.Refresh
Else
MsgBox "Record not deleted", vbCritical
End If

End Sub

