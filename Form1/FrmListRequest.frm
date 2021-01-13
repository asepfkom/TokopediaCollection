VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmListRequest 
   Caption         =   "Request Form"
   ClientHeight    =   8895
   ClientLeft      =   180
   ClientTop       =   555
   ClientWidth     =   13185
   LinkTopic       =   "Form2"
   ScaleHeight     =   8895
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDExit 
      Caption         =   "&Exit"
      Height          =   555
      Left            =   11280
      TabIndex        =   4
      Top             =   8220
      Width           =   1635
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   435
      Left            =   180
      TabIndex        =   3
      Top             =   8280
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Request PUM"
      TabPicture(0)   =   "FrmListRequest.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LvReqPUM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdLoadPUM"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TxtJmlPUM"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Request EC"
      TabPicture(1)   =   "FrmListRequest.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdLoadEC"
      Tab(1).Control(1)=   "TxtJmlEc"
      Tab(1).Control(2)=   "LvReqEC"
      Tab(1).Control(3)=   "Label2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Request BS"
      TabPicture(2)   =   "FrmListRequest.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtJMLBS"
      Tab(2).Control(1)=   "CmdLoadBS"
      Tab(2).Control(2)=   "LvReqBS"
      Tab(2).Control(3)=   "Label3"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Request RS"
      TabPicture(3)   =   "FrmListRequest.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdLoadRS"
      Tab(3).Control(1)=   "TxtJmlRS"
      Tab(3).Control(2)=   "LvReqRS"
      Tab(3).Control(3)=   "Label4"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Request OST"
      TabPicture(4)   =   "FrmListRequest.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TxtJmlOST"
      Tab(4).Control(1)=   "CmdLoadOST"
      Tab(4).Control(2)=   "LvReqOST"
      Tab(4).Control(3)=   "Label5"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Problem..."
      TabPicture(5)   =   "FrmListRequest.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "CmdLoadProblem"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "TxtJmlProblem"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "LvReqProblem"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label6"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      Begin VB.CommandButton CmdLoadProblem 
         Caption         =   "Load data &Problem"
         Height          =   555
         Left            =   -64020
         TabIndex        =   24
         Top             =   7380
         Width           =   1755
      End
      Begin VB.TextBox TxtJmlProblem 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         TabIndex        =   23
         Text            =   "0"
         Top             =   7500
         Width           =   555
      End
      Begin VB.TextBox TxtJmlOST 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         TabIndex        =   20
         Text            =   "0"
         Top             =   7500
         Width           =   555
      End
      Begin VB.CommandButton CmdLoadOST 
         Caption         =   "Load data &OST"
         Height          =   555
         Left            =   -64020
         TabIndex        =   19
         Top             =   7380
         Width           =   1755
      End
      Begin VB.CommandButton CmdLoadRS 
         Caption         =   "Load data &RS"
         Height          =   555
         Left            =   -64080
         TabIndex        =   16
         Top             =   7380
         Width           =   1755
      End
      Begin VB.TextBox TxtJmlRS 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73740
         TabIndex        =   15
         Text            =   "0"
         Top             =   7500
         Width           =   555
      End
      Begin VB.TextBox TxtJMLBS 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         TabIndex        =   12
         Text            =   "0"
         Top             =   7500
         Width           =   555
      End
      Begin VB.CommandButton CmdLoadBS 
         Caption         =   "Load data &BS"
         Height          =   555
         Left            =   -64020
         TabIndex        =   11
         Top             =   7380
         Width           =   1755
      End
      Begin VB.CommandButton CmdLoadEC 
         Caption         =   "Load data &EC"
         Height          =   555
         Left            =   -64080
         TabIndex        =   9
         Top             =   7380
         Width           =   1755
      End
      Begin VB.TextBox TxtJmlEc 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73740
         TabIndex        =   8
         Text            =   "0"
         Top             =   7500
         Width           =   555
      End
      Begin VB.TextBox TxtJmlPUM 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "0"
         Top             =   7500
         Width           =   555
      End
      Begin VB.CommandButton CmdLoadPUM 
         Caption         =   "Load data &PUM"
         Height          =   555
         Left            =   10980
         TabIndex        =   2
         Top             =   7380
         Width           =   1755
      End
      Begin MSComctlLib.ListView LvReqPUM 
         Height          =   6600
         Left            =   120
         TabIndex        =   1
         Top             =   780
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LvReqEC 
         Height          =   6600
         Left            =   -74880
         TabIndex        =   7
         Top             =   780
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LvReqBS 
         Height          =   6600
         Left            =   -74820
         TabIndex        =   13
         Top             =   780
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LvReqRS 
         Height          =   6600
         Left            =   -74880
         TabIndex        =   17
         Top             =   780
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LvReqOST 
         Height          =   6600
         Left            =   -74820
         TabIndex        =   21
         Top             =   780
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LvReqProblem 
         Height          =   6600
         Left            =   -74820
         TabIndex        =   25
         Top             =   780
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label6 
         Caption         =   "Jumlah Data:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   26
         Top             =   7500
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Jumlah Data:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   22
         Top             =   7500
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Jumlah Data:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   18
         Top             =   7500
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah Data:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   14
         Top             =   7500
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Jumlah Data:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   10
         Top             =   7500
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Data:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   7500
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmListRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HeaderPUM()
    LvReqPUM.ColumnHeaders.ADD , , "Id PUM", 500
    LvReqPUM.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvReqPUM.ColumnHeaders.ADD , , "Custid", 1000
    LvReqPUM.ColumnHeaders.ADD , , "Agent", 1000
    LvReqPUM.ColumnHeaders.ADD , , "Amountwo", 1000
    LvReqPUM.ColumnHeaders.ADD , , "Payment Date", 1000
    LvReqPUM.ColumnHeaders.ADD , , "Remarks PUM", 4000
    LvReqPUM.ColumnHeaders.ADD , , "Remarks By Admin", 4000
    LvReqPUM.ColumnHeaders.ADD , , "Status", 3000
End Sub
Private Sub IsiReqPUM()
    Dim listitem As listitem
    Dim m_objrs As ADODB.Recordset
    Dim Cmdsql As String
    
    Cmdsql = "select * from tbl_req_pum order by status asc"
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlPUM.Text = m_objrs.RecordCount
    
    If m_objrs.RecordCount > 0 Then
        PB1.Max = m_objrs.RecordCount
        While Not m_objrs.EOF
            PB1.Value = m_objrs.Bookmark
            Set listitem = LvReqPUM.ListItems.ADD(, , m_objrs("id"))
                listitem.SubItems(1) = Format(m_objrs("tgl_req"), "yyyy-mm-dd")
                listitem.SubItems(2) = m_objrs("custid")
                listitem.SubItems(3) = m_objrs("agent")
                listitem.SubItems(4) = IIf(IsNull(m_objrs("amountwo")), "0", m_objrs("amountwo"))
                listitem.SubItems(5) = IIf(IsNull(m_objrs("payment_date")), "", Format(m_objrs("payment_date"), "yyyy-mm-dd"))
                listitem.SubItems(6) = IIf(IsNull(m_objrs("remarks_agent")), "", m_objrs("remarks_agent"))
                listitem.SubItems(7) = IIf(IsNull(m_objrs("remarks")), "", m_objrs("remarks"))
                listitem.SubItems(8) = IIf(IsNull(m_objrs("status")), "", m_objrs("status"))
            If m_objrs("status") = "0" Then
                listitem.ForeColor = vbRed
                listitem.ListSubItems(1).ForeColor = vbRed
                listitem.ListSubItems(2).ForeColor = vbRed
                listitem.ListSubItems(3).ForeColor = vbRed
                listitem.ListSubItems(4).ForeColor = vbRed
                listitem.ListSubItems(5).ForeColor = vbRed
                listitem.ListSubItems(6).ForeColor = vbRed
                listitem.ListSubItems(7).ForeColor = vbRed
                listitem.ListSubItems(8).ForeColor = vbRed
            Else
                listitem.ForeColor = vbBlue
                listitem.ListSubItems(1).ForeColor = vbBlue
                listitem.ListSubItems(2).ForeColor = vbBlue
                listitem.ListSubItems(3).ForeColor = vbBlue
                listitem.ListSubItems(4).ForeColor = vbBlue
                listitem.ListSubItems(5).ForeColor = vbBlue
                listitem.ListSubItems(6).ForeColor = vbBlue
                listitem.ListSubItems(7).ForeColor = vbBlue
                listitem.ListSubItems(8).ForeColor = vbBlue
            End If
            m_objrs.MoveNext
        Wend
    End If
End Sub


Private Sub HeaderEC()
    LvReqEC.ColumnHeaders.ADD , , "Id EC", 700
    LvReqEC.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvReqEC.ColumnHeaders.ADD , , "Custid", 1000
    LvReqEC.ColumnHeaders.ADD , , "Agent", 1000
    LvReqEC.ColumnHeaders.ADD , , "Nama CH", 1000
    LvReqEC.ColumnHeaders.ADD , , "Remarks EC", 4000
    LvReqEC.ColumnHeaders.ADD , , "Remarks By Admin", 4000
    LvReqEC.ColumnHeaders.ADD , , "Status", 700
End Sub

Private Sub IsiReqEC()
    Dim listitem As listitem
    Dim m_objrs As ADODB.Recordset
    Dim Cmdsql As String
    
    Cmdsql = "select * from tbl_req_ec order by status asc"
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvReqEC.ListItems.CLEAR
    TxtJmlEc.Text = m_objrs.RecordCount
    
    If m_objrs.RecordCount > 0 Then
        PB1.Max = m_objrs.RecordCount
        While Not m_objrs.EOF
            PB1.Value = m_objrs.Bookmark
            Set listitem = LvReqEC.ListItems.ADD(, , m_objrs("id"))
                listitem.SubItems(1) = Format(m_objrs("tgl_req_ec"), "yyyy-mm-dd")
                listitem.SubItems(2) = m_objrs("custid")
                listitem.SubItems(3) = m_objrs("agent")
                listitem.SubItems(4) = IIf(IsNull(m_objrs("nama")), "", m_objrs("nama"))
                listitem.SubItems(5) = IIf(IsNull(m_objrs("remarks_agent")), "", m_objrs("remarks_agent"))
                listitem.SubItems(6) = IIf(IsNull(m_objrs("remarks")), "", m_objrs("remarks"))
                listitem.SubItems(7) = IIf(IsNull(m_objrs("status")), "", m_objrs("status"))
                
            If m_objrs("status") = "0" Then
                listitem.ForeColor = vbRed
                listitem.ListSubItems(1).ForeColor = vbRed
                listitem.ListSubItems(2).ForeColor = vbRed
                listitem.ListSubItems(3).ForeColor = vbRed
                listitem.ListSubItems(4).ForeColor = vbRed
                listitem.ListSubItems(5).ForeColor = vbRed
                listitem.ListSubItems(6).ForeColor = vbRed
                listitem.ListSubItems(7).ForeColor = vbRed
            Else
                listitem.ForeColor = vbBlue
                listitem.ListSubItems(1).ForeColor = vbBlue
                listitem.ListSubItems(2).ForeColor = vbBlue
                listitem.ListSubItems(3).ForeColor = vbBlue
                listitem.ListSubItems(4).ForeColor = vbBlue
                listitem.ListSubItems(5).ForeColor = vbBlue
                listitem.ListSubItems(6).ForeColor = vbBlue
                listitem.ListSubItems(7).ForeColor = vbBlue
            End If
            m_objrs.MoveNext
        Wend
    End If
End Sub

Private Sub HeaderBS()
    LvReqBS.ColumnHeaders.ADD , , "Id BS", 700
    LvReqBS.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvReqBS.ColumnHeaders.ADD , , "Custid", 1000
    LvReqBS.ColumnHeaders.ADD , , "Agent", 1000
    LvReqBS.ColumnHeaders.ADD , , "Nama CH", 1000
    LvReqBS.ColumnHeaders.ADD , , "Month", 1000
    LvReqBS.ColumnHeaders.ADD , , "Year", 1000
    LvReqBS.ColumnHeaders.ADD , , "Remarks BS", 4000
    LvReqBS.ColumnHeaders.ADD , , "Remarks By Admin", 4000
    LvReqBS.ColumnHeaders.ADD , , "Status", 700
End Sub
Private Sub IsiReqBS()
    Dim listitem As listitem
    Dim m_objrs As ADODB.Recordset
    Dim Cmdsql As String
    
    Cmdsql = "select * from tbl_req_bs order by status asc"
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJMLBS.Text = m_objrs.RecordCount
    
    LvReqBS.ListItems.CLEAR
    
    If m_objrs.RecordCount > 0 Then
        PB1.Max = m_objrs.RecordCount
        While Not m_objrs.EOF
            PB1.Value = m_objrs.Bookmark
            Set listitem = LvReqBS.ListItems.ADD(, , m_objrs("id"))
                listitem.SubItems(1) = Format(m_objrs("tgl_req_bs"), "yyyy-mm-dd")
                listitem.SubItems(2) = m_objrs("custid")
                listitem.SubItems(3) = m_objrs("agent")
                listitem.SubItems(4) = IIf(IsNull(m_objrs("nama")), "", m_objrs("nama"))
                listitem.SubItems(5) = IIf(IsNull(m_objrs("month_bs")), "", m_objrs("month_bs"))
                listitem.SubItems(6) = IIf(IsNull(m_objrs("year_bs")), "", m_objrs("year_bs"))
                listitem.SubItems(7) = IIf(IsNull(m_objrs("remarks_agent")), "", m_objrs("remarks_agent"))
                listitem.SubItems(8) = IIf(IsNull(m_objrs("remarks")), "", m_objrs("remarks"))
                listitem.SubItems(9) = IIf(IsNull(m_objrs("status")), "", m_objrs("status"))
            If m_objrs("status") = "0" Then
                listitem.ForeColor = vbRed
                listitem.ListSubItems(1).ForeColor = vbRed
                listitem.ListSubItems(2).ForeColor = vbRed
                listitem.ListSubItems(3).ForeColor = vbRed
                listitem.ListSubItems(4).ForeColor = vbRed
                listitem.ListSubItems(5).ForeColor = vbRed
                listitem.ListSubItems(6).ForeColor = vbRed
                listitem.ListSubItems(7).ForeColor = vbRed
                listitem.ListSubItems(8).ForeColor = vbRed
                listitem.ListSubItems(9).ForeColor = vbRed
            Else
                listitem.ForeColor = vbBlue
                listitem.ListSubItems(1).ForeColor = vbBlue
                listitem.ListSubItems(2).ForeColor = vbBlue
                listitem.ListSubItems(3).ForeColor = vbBlue
                listitem.ListSubItems(4).ForeColor = vbBlue
                listitem.ListSubItems(5).ForeColor = vbBlue
                listitem.ListSubItems(6).ForeColor = vbBlue
                listitem.ListSubItems(7).ForeColor = vbBlue
                listitem.ListSubItems(8).ForeColor = vbBlue
                listitem.ListSubItems(9).ForeColor = vbBlue
            End If
            m_objrs.MoveNext
        Wend
    End If
End Sub


Private Sub HeaderRS()
    LvReqRS.ColumnHeaders.ADD , , "Id RS", 700
    LvReqRS.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvReqRS.ColumnHeaders.ADD , , "Custid", 1000
    LvReqRS.ColumnHeaders.ADD , , "Agent", 1000
    LvReqRS.ColumnHeaders.ADD , , "Total Payment", 1000
    LvReqRS.ColumnHeaders.ADD , , "Installment Period", 1000
    LvReqRS.ColumnHeaders.ADD , , "Remarks BS", 4000
    LvReqRS.ColumnHeaders.ADD , , "Remarks By Admin", 4000
    LvReqRS.ColumnHeaders.ADD , , "Status", 700
End Sub
Private Sub IsiReqRS()
    Dim listitem As listitem
    Dim m_objrs As ADODB.Recordset
    Dim Cmdsql As String
    
    Cmdsql = "select * from tbl_req_rs order by status asc"
    
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlRS.Text = m_objrs.RecordCount
    LvReqRS.ListItems.CLEAR
    If m_objrs.RecordCount > 0 Then
        PB1.Max = m_objrs.RecordCount
        While Not m_objrs.EOF
            PB1.Value = m_objrs.Bookmark
            Set listitem = LvReqRS.ListItems.ADD(, , m_objrs("id"))
                listitem.SubItems(1) = Format(m_objrs("tgl_req_rs"), "yyyy-mm-dd")
                listitem.SubItems(2) = m_objrs("custid")
                listitem.SubItems(3) = m_objrs("agent")
                listitem.SubItems(4) = IIf(IsNull(m_objrs("tot_payment")), "0", m_objrs("tot_payment"))
                listitem.SubItems(5) = IIf(IsNull(m_objrs("installment_period")), "0", m_objrs("installment_period"))
                listitem.SubItems(6) = IIf(IsNull(m_objrs("remarks_agent")), "", m_objrs("remarks_agent"))
                listitem.SubItems(7) = IIf(IsNull(m_objrs("remarks")), "", m_objrs("remarks"))
                listitem.SubItems(8) = IIf(IsNull(m_objrs("status")), "", m_objrs("status"))
            If m_objrs("status") = "0" Then
                listitem.ForeColor = vbRed
                listitem.ListSubItems(1).ForeColor = vbRed
                listitem.ListSubItems(2).ForeColor = vbRed
                listitem.ListSubItems(3).ForeColor = vbRed
                listitem.ListSubItems(4).ForeColor = vbRed
                listitem.ListSubItems(5).ForeColor = vbRed
                listitem.ListSubItems(6).ForeColor = vbRed
                listitem.ListSubItems(7).ForeColor = vbRed
                listitem.ListSubItems(8).ForeColor = vbRed
            Else
                listitem.ForeColor = vbBlue
                listitem.ListSubItems(1).ForeColor = vbBlue
                listitem.ListSubItems(2).ForeColor = vbBlue
                listitem.ListSubItems(3).ForeColor = vbBlue
                listitem.ListSubItems(4).ForeColor = vbBlue
                listitem.ListSubItems(5).ForeColor = vbBlue
                listitem.ListSubItems(6).ForeColor = vbBlue
                listitem.ListSubItems(7).ForeColor = vbBlue
                listitem.ListSubItems(8).ForeColor = vbBlue
            End If
            m_objrs.MoveNext
        Wend
    End If
End Sub


Private Sub HeaderOST()
    LvReqOST.ColumnHeaders.ADD , , "Id OST", 700
    LvReqOST.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvReqOST.ColumnHeaders.ADD , , "Custid", 1000
    LvReqOST.ColumnHeaders.ADD , , "Agent", 1000
    LvReqOST.ColumnHeaders.ADD , , "Address Request", 1000
    LvReqOST.ColumnHeaders.ADD , , "Remarks OST", 4000
    LvReqOST.ColumnHeaders.ADD , , "Remarks By Admin", 4000
    LvReqOST.ColumnHeaders.ADD , , "Status", 700
End Sub
Private Sub IsiReqOST()
    Dim listitem As listitem
    Dim m_objrs As ADODB.Recordset
    Dim Cmdsql As String
    
    Cmdsql = "select * from tbl_req_ost order by status asc"
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvReqOST.ListItems.CLEAR
    
    TxtJmlOST.Text = m_objrs.RecordCount
    If m_objrs.RecordCount > 0 Then
        PB1.Max = m_objrs.RecordCount
        While Not m_objrs.EOF
            PB1.Value = m_objrs.Bookmark
            Set listitem = LvReqOST.ListItems.ADD(, , m_objrs("id"))
                listitem.SubItems(1) = Format(m_objrs("tgl_req_ost"), "yyyy-mm-dd")
                listitem.SubItems(2) = m_objrs("custid")
                listitem.SubItems(3) = m_objrs("agent")
                listitem.SubItems(4) = IIf(IsNull(m_objrs("addr")), "0", m_objrs("addr"))
                listitem.SubItems(5) = IIf(IsNull(m_objrs("remarks_agent")), "", m_objrs("remarks_agent"))
                listitem.SubItems(6) = IIf(IsNull(m_objrs("remarks")), "", m_objrs("remarks"))
                listitem.SubItems(7) = IIf(IsNull(m_objrs("status")), "", m_objrs("status"))
            If m_objrs("status") = "0" Then
                listitem.ForeColor = vbRed
                listitem.ListSubItems(1).ForeColor = vbRed
                listitem.ListSubItems(2).ForeColor = vbRed
                listitem.ListSubItems(3).ForeColor = vbRed
                listitem.ListSubItems(4).ForeColor = vbRed
                listitem.ListSubItems(5).ForeColor = vbRed
                listitem.ListSubItems(6).ForeColor = vbRed
                listitem.ListSubItems(7).ForeColor = vbRed
            Else
                listitem.ForeColor = vbBlue
                listitem.ListSubItems(1).ForeColor = vbBlue
                listitem.ListSubItems(2).ForeColor = vbBlue
                listitem.ListSubItems(3).ForeColor = vbBlue
                listitem.ListSubItems(4).ForeColor = vbBlue
                listitem.ListSubItems(5).ForeColor = vbBlue
                listitem.ListSubItems(6).ForeColor = vbBlue
                listitem.ListSubItems(7).ForeColor = vbBlue
            End If
            m_objrs.MoveNext
        Wend
    End If
End Sub


Private Sub HeaderProblem()
    LvReqProblem.ColumnHeaders.ADD , , "Id Problem", 700
    LvReqProblem.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvReqProblem.ColumnHeaders.ADD , , "Custid", 1000
    LvReqProblem.ColumnHeaders.ADD , , "Agent", 1000
    LvReqProblem.ColumnHeaders.ADD , , "Nama Agent", 1000
    LvReqProblem.ColumnHeaders.ADD , , "Problem", 4000
    LvReqProblem.ColumnHeaders.ADD , , "Solving", 4000
    LvReqProblem.ColumnHeaders.ADD , , "Status", 700
End Sub
Private Sub IsiReqProblem()
    Dim listitem As listitem
    Dim m_objrs As ADODB.Recordset
    Dim Cmdsql As String
    
    Cmdsql = "select * from tbl_req_problem order by status asc"
    
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvReqProblem.ListItems.CLEAR
    
    TxtJmlProblem.Text = m_objrs.RecordCount
    
    If m_objrs.RecordCount > 0 Then
        PB1.Max = m_objrs.RecordCount
        While Not m_objrs.EOF
            PB1.Value = m_objrs.Bookmark
            Set listitem = LvReqProblem.ListItems.ADD(, , m_objrs("id"))
                listitem.SubItems(1) = Format(m_objrs("tgl"), "yyyy-mm-dd")
                listitem.SubItems(2) = m_objrs("custid")
                listitem.SubItems(3) = m_objrs("agent")
                listitem.SubItems(4) = IIf(IsNull(m_objrs("nama_agent")), "0", m_objrs("nama_agent"))
                listitem.SubItems(5) = IIf(IsNull(m_objrs("problem")), "", m_objrs("problem"))
                listitem.SubItems(6) = IIf(IsNull(m_objrs("solve")), "", m_objrs("solve"))
                listitem.SubItems(7) = IIf(IsNull(m_objrs("status")), "", m_objrs("status"))
            If m_objrs("status") = "0" Then
                listitem.ForeColor = vbRed
                listitem.ListSubItems(1).ForeColor = vbRed
                listitem.ListSubItems(2).ForeColor = vbRed
                listitem.ListSubItems(3).ForeColor = vbRed
                listitem.ListSubItems(4).ForeColor = vbRed
                listitem.ListSubItems(5).ForeColor = vbRed
                listitem.ListSubItems(6).ForeColor = vbRed
                listitem.ListSubItems(7).ForeColor = vbRed
            Else
                listitem.ForeColor = vbBlue
                listitem.ListSubItems(1).ForeColor = vbBlue
                listitem.ListSubItems(2).ForeColor = vbBlue
                listitem.ListSubItems(3).ForeColor = vbBlue
                listitem.ListSubItems(4).ForeColor = vbBlue
                listitem.ListSubItems(5).ForeColor = vbBlue
                listitem.ListSubItems(6).ForeColor = vbBlue
                listitem.ListSubItems(7).ForeColor = vbBlue
            End If
            m_objrs.MoveNext
        Wend
    End If
End Sub



Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub CmdLoadBS_Click()
    IsiReqBS
End Sub

Private Sub CmdLoadEC_Click()
    Call IsiReqEC
End Sub

Private Sub CmdLoadOST_Click()
    IsiReqOST
End Sub

Private Sub CmdLoadProblem_Click()
    IsiReqProblem
End Sub

Private Sub CmdLoadPUM_Click()
    LvReqPUM.ListItems.CLEAR
    Call IsiReqPUM
End Sub

Private Sub CmdLoadRS_Click()
    Call IsiReqRS
End Sub

Private Sub Form_Load()
   Call HeaderPUM
   Call HeaderEC
   Call HeaderBS
   Call HeaderRS
   Call HeaderOST
   Call HeaderProblem
End Sub




Private Sub LvReqBS_DblClick()
    If LvReqBS.ListItems.Count = 0 Then
        Exit Sub
    End If
        
    With FrmRemarksRequest
        .TxtForm = "BS"
        .TxtIdForm = LvReqBS.SelectedItem.Text
        .TxtCustid.Text = LvReqBS.SelectedItem.SubItems(2)
        .TxtAgent.Text = LvReqBS.SelectedItem.SubItems(3)
        .Show vbModal
    End With
End Sub

Private Sub LvReqEC_DblClick()
    If LvReqEC.ListItems.Count = 0 Then
        Exit Sub
    End If
        
    With FrmRemarksRequest
        .TxtForm = "EC"
        .TxtIdForm = LvReqEC.SelectedItem.Text
        .TxtCustid.Text = LvReqEC.SelectedItem.SubItems(2)
        .TxtAgent.Text = LvReqEC.SelectedItem.SubItems(3)
        .Show vbModal
    End With
End Sub

Private Sub LvReqOST_DblClick()
     If LvReqOST.ListItems.Count = 0 Then
        Exit Sub
    End If
        
    With FrmRemarksRequest
        .TxtForm = "OST"
        .TxtIdForm = LvReqOST.SelectedItem.Text
        .TxtCustid.Text = LvReqOST.SelectedItem.SubItems(2)
        .TxtAgent.Text = LvReqOST.SelectedItem.SubItems(3)
        .Show vbModal
    End With
End Sub



Private Sub LvReqProblem_DblClick()
     If LvReqProblem.ListItems.Count = 0 Then
        Exit Sub
    End If
        
    With FrmRemarksRequest
        .TxtForm = "PROBLEM"
        .TxtIdForm = LvReqProblem.SelectedItem.Text
        .TxtCustid.Text = LvReqProblem.SelectedItem.SubItems(2)
        .TxtAgent.Text = LvReqProblem.SelectedItem.SubItems(3)
        .Show vbModal
    End With
End Sub

Private Sub LvReqPUM_DblClick()
    If LvReqPUM.ListItems.Count = 0 Then
        Exit Sub
    End If
        
    With FrmRemarksRequest
        .TxtForm = "PUM"
        .TxtIdForm = LvReqPUM.SelectedItem.Text
        .TxtCustid.Text = LvReqPUM.SelectedItem.SubItems(2)
        .TxtAgent.Text = LvReqPUM.SelectedItem.SubItems(3)
        .Show vbModal
    End With
End Sub



Private Sub LvReqRS_DblClick()
     If LvReqRS.ListItems.Count = 0 Then
        Exit Sub
    End If
        
    With FrmRemarksRequest
        .TxtForm = "RS"
        .TxtIdForm = LvReqRS.SelectedItem.Text
        .TxtCustid.Text = LvReqRS.SelectedItem.SubItems(2)
        .TxtAgent.Text = LvReqRS.SelectedItem.SubItems(3)
        .Show vbModal
    End With
End Sub
