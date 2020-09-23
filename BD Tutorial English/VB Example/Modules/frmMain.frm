VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Tutorial and a example by Santiago Favaro"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmThree 
      Caption         =   "Datagrid and non visible AdoDC"
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   7455
      Begin MSDataGridLib.DataGrid dgList 
         Bindings        =   "frmMain.frx":0000
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   13290175
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "BOOKS"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cName"
            Caption         =   "Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cEditorial"
            Caption         =   "Editorial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cPrice"
            Caption         =   "Price"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "$ #0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cDate"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MMM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoNonDC 
         Height          =   375
         Left            =   5400
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Example.mdb"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Example.mdb"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblBooks"
         Caption         =   "Adodc3"
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
      Begin VB.Image imgDel 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   4680
         Picture         =   "frmMain.frx":0017
         Top             =   3720
         Width           =   420
      End
      Begin VB.Image imgNew 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   4200
         Picture         =   "frmMain.frx":071B
         Top             =   3720
         Width           =   420
      End
      Begin VB.Image imgEnd 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   3480
         Picture         =   "frmMain.frx":0E1F
         Top             =   3720
         Width           =   420
      End
      Begin VB.Image imgNext 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   3000
         Picture         =   "frmMain.frx":1523
         Top             =   3720
         Width           =   420
      End
      Begin VB.Image imgPrevious 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   2400
         Picture         =   "frmMain.frx":1C27
         Top             =   3720
         Width           =   420
      End
      Begin VB.Image imgBegin 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   1920
         Picture         =   "frmMain.frx":232B
         Top             =   3720
         Width           =   420
      End
   End
   Begin VB.Frame frmTwo 
      Caption         =   "Find"
      Height          =   3015
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1140
         TabIndex        =   15
         Top             =   675
         Width           =   2295
      End
      Begin VB.ComboBox cmbField 
         Height          =   315
         ItemData        =   "frmMain.frx":2A2F
         Left            =   1155
         List            =   "frmMain.frx":2A3C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   300
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid dgList2 
         Bindings        =   "frmMain.frx":2A5E
         Height          =   1455
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         Caption         =   "Radios"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cRadio"
            Caption         =   "Radio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cFrequency"
            Caption         =   "Freq"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cRanking"
            Caption         =   "P"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1470,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   450,142
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoDCFind 
         Height          =   330
         Left            =   240
         Top             =   2520
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Example.mdb"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Example.mdb"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblRadios"
         Caption         =   "Adodc2"
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
      Begin VB.Label Label6 
         Caption         =   "Find what:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Field name:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frmOne 
      Caption         =   "View, Add and Edit"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdDel 
         Caption         =   "del"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         ToolTipText     =   "Erase the current register"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "cTelefono"
         DataSource      =   "AdoDCEdit"
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1635
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         DataField       =   "cSurname"
         DataSource      =   "AdoDCEdit"
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1170
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc AdoDCEdit 
         Height          =   330
         Left            =   120
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
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
         EOFAction       =   2
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Example.mdb"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Example.mdb"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblPersons"
         Caption         =   "Adodc1"
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
      Begin VB.Label Label4 
         Caption         =   "Telephone:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "here is the name"
         DataField       =   "cName"
         DataSource      =   "AdoDCEdit"
         Height          =   195
         Left            =   1260
         TabIndex        =   7
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   "Surname:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This code is an example of what you can do with the
' Database Tutorial for Visual Basic Step by Step by
' Santiago Favaro
'
' Please if you found this tutorial or example useful
' think in my expend time in it for you and think if
' i really deserve a vote. If you vote me i will still
' writing this useful tutorials for you
'
' If you wanna this tutorial in Spanish please send me
' an e-mail and say what you wanna. my e-mail is
'               smfavaro@hotmail.com




'--------------------------------------------------------'
'#          VIEW, ADD NEW AND EDIT                      #'
'--------------------------------------------------------'

Private Sub AdoDCEdit_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' Change the caption to show more info
    AdoDCEdit.Caption = "Register " & AdoDCEdit.Recordset.AbsolutePosition & _
                       " of " & AdoDCEdit.Recordset.RecordCount
End Sub

Private Sub cmdDel_Click()
    ' Erase present record
    AdoDCEdit.Recordset.Delete
    ' Refresh and Save the database
    AdoDCEdit.Recordset.Save
    AdoDCEdit.Recordset.Requery
End Sub

'--------------------------------------------------------'
'#              FIND (Apply FILTER)                     #'
'--------------------------------------------------------'

Private Sub AdoDCFind_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' Change the caption to show more info
    AdoDCFind.Caption = "Filter Registers: " & AdoDCFind.Recordset.RecordCount
End Sub

Private Sub txtFind_Change()
    ' We will apply filter when:
    '       1- cmbField is NOT empty
    '       2- txtFind is NOT empty
    ' We will take off the filter when:
    '       1- cmbField is empty
    '       2- txtFind is empty
    
    ' With this condition we don't have to have a button for ShowAll
    If cmbField.Text = "" Or txtFind = "" Then
        AdoDCFind.Recordset.Filter = ""
        AdoDCFind.Refresh
        Exit Sub
    End If
    ' Now if all ok we apply the filter
    AdoDCFind.Recordset.Filter = cmbField & " LIKE '*" & txtFind & "*'"
    ' If you wanna filter by date you must use this
    ' AdoDCFind.Recordset.Filter = "fieldNamedFecha >= #23/05/2002# AND fieldNamedFecha <= #23/05/2003#"
    ' This will filter all registers between 23/05/2002 and 23/05/2003
End Sub

'--------------------------------------------------------'
'#              ADO DC NON VISIBLE                      #'
'--------------------------------------------------------'

Private Sub imgBegin_Click()
    ' Goto to begin
    AdoNonDC.Recordset.MoveFirst
End Sub

Private Sub imgPrevious_Click()
    ' Go previous
    AdoNonDC.Recordset.MovePrevious
End Sub

Private Sub imgNext_Click()
    ' Go next
    AdoNonDC.Recordset.MoveNext
End Sub

Private Sub imgEnd_Click()
    ' Go to the end
    AdoNonDC.Recordset.MoveLast
End Sub

Private Sub imgNew_Click()
    ' Add new
    AdoNonDC.Recordset.AddNew
End Sub

Private Sub imgDel_Click()
    ' Delete current register
    AdoNonDC.Recordset.Delete
End Sub

Private Sub dgList_AfterColEdit(ByVal ColIndex As Integer)
    ' Grabo
    AdoNonDC.Recordset.Save
    AdoNonDC.Recordset.Requery
End Sub
