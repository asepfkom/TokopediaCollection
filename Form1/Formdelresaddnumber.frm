VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Formdelresaddnumber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Restore Additional Number"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7965
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7965
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Delete Add Number"
      TabPicture(0)   =   "Formdelresaddnumber.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DataGrid1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CommonDialog1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtdataupload"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmddata"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbsheet"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Restore Add Number"
      TabPicture(1)   =   "Formdelresaddnumber.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "Combo1"
      Tab(1).Control(2)=   "DataGrid2"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "Line2"
      Tab(1).Control(5)=   "Label6"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton Command2 
         Caption         =   "Run"
         Height          =   375
         Left            =   -68040
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74280
         TabIndex        =   13
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Run"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Backup"
         Height          =   615
         Left            =   3840
         TabIndex        =   8
         Top             =   3480
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000010&
            Caption         =   "Backupaddnum"
            Height          =   255
            Left            =   600
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackColor       =   &H0080FFFF&
            Caption         =   "Nama"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ComboBox cbsheet 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   795
         Width           =   2415
      End
      Begin VB.CommandButton cmddata 
         Caption         =   "..."
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   435
         Width           =   615
      End
      Begin VB.TextBox txtdataupload 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   465
         Width           =   3495
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3285
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   5794
         _Version        =   393216
         BackColor       =   -2147483636
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   3285
         Left            =   -74640
         TabIndex        =   15
         Top             =   1320
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   5794
         _Version        =   393216
         BackColor       =   -2147483636
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   -71040
         TabIndex        =   16
         Top             =   720
         Width           =   3135
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   -75000
         X2              =   -67080
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
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
         Left            =   -74760
         TabIndex        =   14
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Height          =   735
         Left            =   3840
         TabIndex        =   7
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   7920
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet"
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
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
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
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Formdelresaddnumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_XLSCONN As New ADODB.Connection
Dim rsTemp As ADODB.Recordset
Dim sbatch As String

Private Sub cbsheet_Click()
    Dim OBJRECORD As New ADODB.Recordset
    
        Set rsTemp = New ADODB.Recordset
        rsTemp.CursorLocation = adUseClient
    
        ssql = "SELECT * FROM [" & cbsheet.text & "] "
            rsTemp.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
            Set rsTemp = Nothing
            Set OBJRECORD = New ADODB.Recordset
            OBJRECORD.CursorLocation = adUseClient
            
        ssql = "SELECT * FROM [" & cbsheet.text & "] "
            DoEvents
            OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
            Set DataGrid1.DATASOURCE = OBJRECORD
        
        
        q = "select * from information_schema.columns  where table_name = 'tbldelresaddnumtemp'"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If rs.RecordCount = 0 Then
            qc = "create table tbldelresaddnumtemp ( custid varchar );"
            M_OBJCONN.Execute qc
            
            For i = 1 To OBJRECORD.RecordCount
                scustid = IIf(IsNull(OBJRECORD("custid")), "", OBJRECORD("custid"))
                qi = "insert into tbldelresaddnumtemp values ('" & CustId & "');"
                M_OBJCONN.Execute qi
                OBJRECORD.MoveNext
            Next i
            
            q1 = "select * from mgm where custid in (select custid from tbldelresaddnumtemp)"
            Set rs1 = New ADODB.Recordset
            rs1.CursorLocation = adUseClient
            rs1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        Else
            qd = "drop table tbldelresaddnumtemp;"
            M_OBJCONN.Execute qd
            
            qc = "create table tbldelresaddnumtemp ( custid varchar );"
            M_OBJCONN.Execute qc
            
            For i = 1 To OBJRECORD.RecordCount
                scustid = IIf(IsNull(OBJRECORD("custid")), "", OBJRECORD("custid"))
                qi = "insert into tbldelresaddnumtemp values ('" & scustid & "');"
                M_OBJCONN.Execute qi
                OBJRECORD.MoveNext
            Next i
            
            q1 = "select * from mgm where custid in (select custid from tbldelresaddnumtemp)"
            Set rs1 = New ADODB.Recordset
            rs1.CursorLocation = adUseClient
            rs1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        End If
            Label3.Caption = "Data Excel ada sebanyak " & OBJRECORD.RecordCount & vbCrLf & " Data disource ada sebanyak " & rs1.RecordCount
    
    Command1.Enabled = True

End Sub

Private Sub cmddata_Click()
    With CommonDialog1
            .DialogTitle = "Import From File"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
    End With
        
    txtdataupload.text = CommonDialog1.FileName
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtdataupload.text & ";Extended Properties=Excel 8.0;"
    Set rsTemp = M_XLSCONN.OpenSchema(adSchemaTables)
    cbsheet.clear
    If rsTemp.EOF And rsTemp.BOF Then Exit Sub
        
    While Not rsTemp.EOF
        cbsheet.AddItem IIf(IsNull(rsTemp!table_name), "", rsTemp!table_name)
        rsTemp.MoveNext
    Wend
    
    Set rsTemp = Nothing
End Sub

Private Sub Combo1_Click()
    Dim OBJRECORD As New ADODB.Recordset
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
            
    ssql = "SELECT * FROM " & Combo1.text & " "
    DoEvents
    OBJRECORD.Open ssql, M_OBJCONN, adOpenKeyset, adLockOptimistic
    Set DataGrid2.DATASOURCE = OBJRECORD
    
    Label7.Caption = "Ada Sebanyak " & OBJRECORD.RecordCount & " Data "

End Sub

Private Sub Command1_Click()
'    If Text1.text = "" Then
'        MsgBox "Harap Buat Backup Terlebih Dahulu"
'        Exit Sub
'    End If
    
'    q = "select * from information_schema.columns  where table_name = '" & LCase(Label4.Caption) & "_" & LCase(Text1.text) & "'"
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If rs.RecordCount > 0 Then
'        MsgBox "Nama Backup Sudah Ada Harap Ganti"
'        Exit Sub
'    End If
    
    qwkt = "select to_char(now(),'yyyymmddhhmiss') wkt"
    Set rswkt = New ADODB.Recordset
    rswkt.CursorLocation = adUseClient
    rswkt.Open qwkt, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    qc = "create table " & Label4.Caption & "_" & rswkt!wkt & "_" & MDIForm1.Text1.text & " as " & vbCrLf
    qc = qc + "select * from mgm where custid in (select custid from tbldelresaddnumtemp)"
    M_OBJCONN.Execute qc
    
    qu = "update mgm set HOMENOADD1 = '' ,OFFICENOADD1 = '',MOBILENOADD1 = '' where " & vbCrLf
    qu = qu + " custid in (select custid from tbldelresaddnumtemp)"
    M_OBJCONN.Execute qu
    
    MsgBox "Berhasil di Delete dan Backup"
    Call clear
    Call cmbrestore
End Sub

Private Sub clear()
    txtdataupload.text = ""
    cbsheet.clear
    Set DataGrid1.DATASOURCE = Nothing
    Text1.text = ""
    Label3.Caption = ""
End Sub

Private Sub Command2_Click()
    If Combo1.text = "" Then
        MsgBox "Pilih Data"
        Exit Sub
    End If
    
    qwkt = "select to_char(now(),'yyyymmddhhmiss') wkt"
    Set rswkt = New ADODB.Recordset
    rswkt.CursorLocation = adUseClient
    rswkt.Open qwkt, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    qc = "create table res_" & Label4.Caption & "_" & rswkt!wkt & "_" & MDIForm1.Text1.text & " as " & vbCrLf
    qc = qc + "select * from " & Combo1.text & ";"
    M_OBJCONN.Execute qc
    
    qu = "update mgm set HOMENOADD1 = " & Combo1.text & ".homenoadd1 ,OFFICENOADD1 = " & Combo1.text & ".officenoadd1,MOBILENOADD1 = " & Combo1.text & ".mobilenoadd1 from " & Combo1.text & " where " & vbCrLf
    qu = qu + " " & Combo1.text & ".custid = mgm.custid"
    M_OBJCONN.Execute qu
    
    qd = "Drop table " & Combo1.text & ";"
    M_OBJCONN.Execute qd
    
    MsgBox "Berhasil di Restore"
    'Call clear
    Label7.Caption = ""
    Set DataGrid2.DATASOURCE = Nothing
    Call cmbrestore
End Sub

Private Sub Form_Load()
    Call cmbrestore
End Sub

Private Sub cmbrestore()
    Combo1.clear

    q = "select distinct table_name from information_schema.columns  where table_name ilike '" & LCase(Label4.Caption) & "%' order by 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    For i = 1 To rs.RecordCount
        Combo1.AddItem rs!table_name
        rs.MoveNext
    Next i
End Sub
