VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form formreportsms 
   Caption         =   "REPORT SMS"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   915
   ClientWidth     =   10935
   LinkTopic       =   "Form5"
   ScaleHeight     =   5790
   ScaleWidth      =   10935
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Filter"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton Command3 
         Caption         =   "Inbox"
         Height          =   495
         Left            =   5880
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Height          =   495
         Left            =   7560
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         Height          =   495
         Left            =   9240
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate tgl_mulai1 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   195
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "formreportsms.frx":0000
         Caption         =   "formreportsms.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "formreportsms.frx":0184
         Keys            =   "formreportsms.frx":01A2
         Spin            =   "formreportsms.frx":0200
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd, mmm yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__, ___ ____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tgl_akhir1 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   195
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "formreportsms.frx":0228
         Caption         =   "formreportsms.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "formreportsms.frx":03AC
         Keys            =   "formreportsms.frx":03CA
         Spin            =   "formreportsms.frx":0428
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd, mmm yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__, ___ ____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "DATE   :"
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
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12648447
      TabCaption(0)   =   "REPORT SMS"
      TabPicture(0)   =   "formreportsms.frx":0450
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "REPORT DETAIL SMS"
      TabPicture(1)   =   "formreportsms.frx":046C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView1 
         Height          =   4230
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7461
         View            =   3
         LabelEdit       =   1
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
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4230
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7461
         View            =   3
         LabelEdit       =   1
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
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "formreportsms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q, q1 As String

Private Sub headerreport()
    ListView1.ColumnHeaders.ADD , , "DATE", 2500
    ListView1.ColumnHeaders.ADD , , "SMS TYPE", 1500
    ListView1.ColumnHeaders.ADD , , "UNIT", 1000
    ListView1.ColumnHeaders.ADD , , "SUCCESS", 1000
    ListView1.ColumnHeaders.ADD , , "QUEUE", 1000
    ListView1.ColumnHeaders.ADD , , "FAILED", 1000
    ListView1.ColumnHeaders.ADD , , "OTHERS", 1000
    ListView1.ColumnHeaders.ADD , , "TOTAL SMS", 1500
    
    ListView2.ColumnHeaders.ADD , , "DATE", 2500
    ListView2.ColumnHeaders.ADD , , "CUSTID", 2000
    ListView2.ColumnHeaders.ADD , , "SMS TYPE", 1500
    ListView2.ColumnHeaders.ADD , , "STATUS SMS", 1500
    ListView2.ColumnHeaders.ADD , , "DETAIL SMS", 1500
End Sub

Private Sub Command1_Click()
    Call isilv1
    Call isilv2
End Sub

Private Sub Command2_Click()
    Call export
End Sub

Private Sub export()
    Dim rs As ADODB.Recordset
    Dim Strsql, sSPV As String
    Dim objExcel As Object
    Dim objBook As Object
    Dim objSheet As Object
    Dim M As Integer
    
    Dim i, jmlSheet As Integer
    Dim m_msgbox, TxtKrm, TxtADM, txtTgl As String

    txtTgl = Format(FungsiWaktuServer, "DD.MM.YYYY_HH.NN.SS")
    
    CD_save.FileName = "REPORT SMS" + txtTgl
    CD_save.CancelError = True
    CD_save.ShowSave
       
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    
    objBook.Sheets.ADD , , 2
                
    'sheet1
    i = 1
    a = Format(tgl_mulai1.Value, "yyyy-mm-dd") & " 00:00:00"
    B = Format(tgl_akhir1.Value, "yyyy-mm-dd") & " 23:59:50"
    
    a1 = Format(tgl_mulai1.Value, "yyyy-mm-dd")
    a2 = Format(tgl_akhir1.Value, "yyyy-mm-dd")
    
    If a1 = a2 Then
        a3 = a1
    Else
        a3 = a1 & " To " & a2
    End If
    
    q = "select '" & a3 & "' as hari,a.tipe, 'RIT1' unit,coalesce(success,0) success,coalesce(pending,0) success,coalesce(failed,0) failed,'0' others, coalesce(success,0)+coalesce(pending,0)+coalesce(failed,0)+'0' total  from (" & vbCrLf
    q = q & "select case " & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5' " & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe, count(tipe) as total from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe from sentitems  where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%') a group by 1) a left join " & vbCrLf
    q = q & "(" & vbCrLf
    q = q & "select case " & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5' " & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe, count(tipe) as SUCCESS from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe from sentitems  where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%' and status = 'SendingOKNoReport') a group by 1) b on a.tipe = b.tipe" & vbCrLf
    q = q & "left join " & vbCrLf
    q = q & "(select case " & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5' " & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe, count(tipe) as failed from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe from sentitems  where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%' and status = 'SendingError') a group by 1" & vbCrLf
    q = q & ") c on a.tipe=c.tipe left join" & vbCrLf
    q = q & "(" & vbCrLf
    q = q & "select case " & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5' " & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe, count(tipe) as pending from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe from outbox where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%') a group by 1" & vbCrLf
    q = q & ") d on a.tipe=d.tipe"

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open q, M_OBJCONN1, adOpenDynamic, adLockOptimistic
        
        Set objSheet = objBook.Sheets(1)
        objSheet.Name = "REPORT SMS"
            
        Dim x, Y As Integer
        If rs.state = 1 Then
            x = 0
            Y = rs.fields.Count - 1
            Do Until x > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(rs.fields(x).Name)
                i = i + 1
                x = x + 1
            Loop
        End If
        
        objSheet.Cells.EntireColumn.AutoFit
        
        objSheet.Range("A2").CopyFromRecordset rs
        objSheet.Application.ActiveSheet.EnableSelection = xlUnlockedCells
        
        Set objSheet = Nothing
    
    
        'sheet2
        i = 1
        a = Format(tgl_mulai1.Value, "yyyy-mm-dd") & " 00:00:00"
        B = Format(tgl_akhir1.Value, "yyyy-mm-dd") & " 23:59:50"
        
        a1 = Format(tgl_mulai1.Value, "yyyy-mm-dd")
        a2 = Format(tgl_akhir1.Value, "yyyy-mm-dd")
        
        If a1 = a2 Then
            a3 = a1
        Else
            a3 = a1 & " To " & a2
        End If
        
            ListView2.ListItems.clear
    
        a = Format(tgl_mulai1.Value, "yyyy-mm-dd") & " 00:00:00"
        B = Format(tgl_akhir1.Value, "yyyy-mm-dd") & " 23:59:50"
        
        a1 = Format(tgl_mulai1.Value, "yyyy-mm-dd")
        a2 = Format(tgl_akhir1.Value, "yyyy-mm-dd")
        
        If a1 = a2 Then
            a3 = a1
        Else
            a3 = a1 & " To " & a2
        End If
        
        q = "select ''''||updatedindb as hari,custid, case" & vbCrLf
        q = q & "when tipe = 'AUTO+5' then 'FWO +5'" & vbCrLf
        q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
        q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
        q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
        q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
        q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
        q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
        q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
        q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
        q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
        q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
        q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
        q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
        q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
        q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe,status,textdecoded from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe,status,textdecoded,updatedindb from sentitems  where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%') a order by 3 asc, 4 desc" & vbCrLf
    
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open q, M_OBJCONN1, adOpenDynamic, adLockOptimistic
        
        Set objSheet = objBook.Sheets(2)
        objSheet.Name = "REPORT DETAIL SMS"
            
        If rs.state = 1 Then
            x = 0
            Y = rs.fields.Count - 1
            Do Until x > Y
                DoEvents
                objSheet.Cells(1, i).Value = "'" & CStr(rs.fields(x).Name)
                i = i + 1
                x = x + 1
            Loop
        End If
        
        'objSheet.Cells.EntireColumn.AutoFit
        objSheet.Columns.AutoFit
        
        objSheet.Range("A2").CopyFromRecordset rs
        objSheet.Application.ActiveSheet.EnableSelection = xlUnlockedCells
        
        Set objSheet = Nothing
        
    objBook.SaveAs CD_save.FileName, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set rs = Nothing
    
End Sub

Private Sub Command3_Click()
    frmsmsinbox.Show
End Sub

Private Sub Form_Load()
    Call headerreport
End Sub

Private Sub isilv1()
    ListView1.ListItems.clear

    a = Format(tgl_mulai1.Value, "yyyy-mm-dd") & " 00:00:00"
    B = Format(tgl_akhir1.Value, "yyyy-mm-dd") & " 23:59:50"
    
    a1 = Format(tgl_mulai1.Value, "yyyy-mm-dd")
    a2 = Format(tgl_akhir1.Value, "yyyy-mm-dd")
    
    If a1 = a2 Then
        a3 = a1
    Else
        a3 = a1 & " To " & a2
    End If
    
    q = "select '" & a3 & "' as hari,a.tipe, 'RIT1' unit,coalesce(success,0) success,coalesce(pending,0) pending,coalesce(failed,0) failed,'0' others, coalesce(success,0)+coalesce(pending,0)+coalesce(failed,0)+'0' total  from (" & vbCrLf
    q = q & "select case " & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5' " & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe, count(tipe) as total from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe from sentitems  where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%') a group by 1) a left join " & vbCrLf
    q = q & "(" & vbCrLf
    q = q & "select case " & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5' " & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe, count(tipe) as SUCCESS from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe from sentitems  where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%' and status = 'SendingOKNoReport') a group by 1) b on a.tipe = b.tipe" & vbCrLf
    q = q & "left join " & vbCrLf
    q = q & "(select case " & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5' " & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe, count(tipe) as failed from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe from sentitems  where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%' and status = 'SendingError') a group by 1" & vbCrLf
    q = q & ") c on a.tipe=c.tipe left join" & vbCrLf
    q = q & "(" & vbCrLf
    q = q & "select case " & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5' " & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe, count(tipe) as pending from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe from outbox where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%') a group by 1" & vbCrLf
    q = q & ") d on a.tipe=d.tipe"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open q, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView1.ListItems.ADD(, , cnull(rs(0)))
             listItem.SubItems(1) = cnull(rs(1))
             listItem.SubItems(2) = cnull(rs(2))
             listItem.SubItems(3) = cnull(rs(3))
             listItem.SubItems(4) = cnull(rs(4))
             listItem.SubItems(5) = cnull(rs(5))
             listItem.SubItems(6) = cnull(rs(6))
             listItem.SubItems(7) = cnull(rs(7))
        rs.MoveNext
    Wend
    
End Sub


Private Sub isilv2()
    ListView2.ListItems.clear

    a = Format(tgl_mulai1.Value, "yyyy-mm-dd") & " 00:00:00"
    B = Format(tgl_akhir1.Value, "yyyy-mm-dd") & " 23:59:50"
    
    a1 = Format(tgl_mulai1.Value, "yyyy-mm-dd")
    a2 = Format(tgl_akhir1.Value, "yyyy-mm-dd")
    
    If a1 = a2 Then
        a3 = a1
    Else
        a3 = a1 & " To " & a2
    End If
    
    q = "select updatedindb as hari,custid, case" & vbCrLf
    q = q & "when tipe = 'AUTO+5' then 'FWO +5'" & vbCrLf
    q = q & "when tipe = 'AUTO+20' then 'FWO +20'" & vbCrLf
    q = q & "when tipe = 'AUTO+30' then 'FWO +30'" & vbCrLf
    q = q & "when tipe = 'AUTO+40' then 'FWO +40'" & vbCrLf
    q = q & "when tipe = 'AUTO+53' then 'FWO +53 LP/VLP'" & vbCrLf
    q = q & "when tipe = 'AUTO+75' then 'FWO +75 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+100' then 'FWO +100 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+150' then 'FWO +150 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO+175' then 'FWO +175 VHP/HP/MP'" & vbCrLf
    q = q & "when tipe = 'AUTO4th' then 'FWO/NFWO Regular Payer 4th'" & vbCrLf
    q = q & "when tipe = 'AUTO4ths' then 'FWO/NFWO BP 4ths'" & vbCrLf
    q = q & "when tipe = 'AUTO25th' then 'FWO/NFWO 25ths'" & vbCrLf
    q = q & "when tipe = 'AUTO+175s' then 'NFWO +175s'" & vbCrLf
    q = q & "when tipe = 'AUTO+8s' then 'NFWO +8s'" & vbCrLf
    q = q & "when tipe = 'AUTO+90s' then 'NFWO +90s' end tipe,status,textdecoded from (select split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as tipe,status,textdecoded,updatedindb from sentitems  where updatedindb between '" & a & "' and '" & B & "' and creatorid ilike '%AUTO%') a order by 3 asc, 4 desc" & vbCrLf
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open q, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , cnull(rs(0)))
             listItem.SubItems(1) = cnull(rs(1))
             listItem.SubItems(2) = cnull(rs(2))
             listItem.SubItems(3) = cnull(rs(3))
             listItem.SubItems(4) = cnull(rs(4))
        rs.MoveNext
    Wend
    
    SSTab1.TabCaption(1) = "REPORT DETAIL SMS " & "(" & rs.RecordCount & ")"
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
End Sub

Private Sub ListView2_DblClick()
    If ListView2.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = ListView2.SelectedItem.SubItems(1)
        Me.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If

End Sub

