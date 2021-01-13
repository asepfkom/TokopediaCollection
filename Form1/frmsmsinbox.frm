VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmsmsinbox 
   Caption         =   "REPORT SMS INBOX"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10965
   LinkTopic       =   "Form5"
   ScaleHeight     =   4560
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   3855
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   10935
      Begin MSComctlLib.ListView ListView2 
         Height          =   3030
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5345
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         Height          =   255
         Left            =   9240
         TabIndex        =   9
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Filter"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         Height          =   495
         Left            =   9240
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Height          =   495
         Left            =   7560
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate tgl_mulai1 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   195
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "frmsmsinbox.frx":0000
         Caption         =   "frmsmsinbox.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsmsinbox.frx":0184
         Keys            =   "frmsmsinbox.frx":01A2
         Spin            =   "frmsmsinbox.frx":0200
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
         TabIndex        =   4
         Top             =   195
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "frmsmsinbox.frx":0228
         Caption         =   "frmsmsinbox.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsmsinbox.frx":03AC
         Keys            =   "frmsmsinbox.frx":03CA
         Spin            =   "frmsmsinbox.frx":0428
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
         TabIndex        =   6
         Top             =   240
         Width           =   1695
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
   End
End
Attribute VB_Name = "frmsmsinbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    ListView2.ColumnHeaders.ADD , , "DATE", 1500
    ListView2.ColumnHeaders.ADD , , "PERKIRAAN CUSTID", 2000
    ListView2.ColumnHeaders.ADD , , "TIPE CUSTID", 2000
    ListView2.ColumnHeaders.ADD , , "NO HP", 1500
    ListView2.ColumnHeaders.ADD , , "DETAIL SMS", 6000
End Sub

Private Sub isilv1()
    q = " select distinct table_name  from information_schema.columns  where table_name = 'tbltampinbox' "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs.RecordCount = 0 Then
        qc = "create table tbltampinbox ( updatedindb timestamp without time zone, sendernumber varchar, textdecoded text );"
        M_OBJCONN.Execute qc
    Else
        qd = "delete from tbltampinbox;"
        M_OBJCONN.Execute qd
    End If
    
    ListView2.ListItems.clear

    a = Format(tgl_mulai1.Value, "yyyy-mm-dd") & " 00:00:00"
    B = Format(tgl_akhir1.Value, "yyyy-mm-dd") & " 23:59:50"
    
    q = " select updatedindb, sendernumber, textdecoded from inbox where updatedindb between '" + a + "' and '" + B + "' order by 1 desc "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open q, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs.RecordCount = 0 Then
        MsgBox "SMS tidak ditemukan"
        Exit Sub
    End If
    
    While Not rs.EOF
        Tanggal = Format(rs(0), "yyyy-mm-dd hh:mm:ss")
        qins = "insert into tbltampinbox values ('" & Tanggal & "','" & rs(1) & "','" & rs(2) & "');"
        M_OBJCONN.Execute qins
        rs.MoveNext
    Wend
    
    q = " select updatedindb, custid, type, sendernumber, textdecoded from (" & vbCrLf
    q = q & " select updatedindb,sendernumber, textdecoded, right(sendernumber,8) kananinbox from tbltampinbox" & vbCrLf
    q = q & " ) a left join (" & vbCrLf
    q = q & " select custid," & vbCrLf
    q = q & " case when date(now())-date(b_d)  >= 5 and date(now())-date(b_d)  < 20 then '+5'" & vbCrLf
    q = q & " when date(now())-date(b_d) >= 20 and date(now())-date(b_d) < 30 then '+20'" & vbCrLf
    q = q & " when date(now())-date(b_d) >= 30 and date(now())-date(b_d) < 40 then '+30'" & vbCrLf
    q = q & " when date(now())-date(b_d) >= 40 and date(now())-date(b_d) < 53 then '+40'" & vbCrLf
    q = q & " when date(now())-date(b_d) >= 53 and date(now())-date(b_d) < 75 then '+53'" & vbCrLf
    q = q & " when date(now())-date(b_d) >= 75 and date(now())-date(b_d) < 100 then '+75'" & vbCrLf
    q = q & " when date(now())-date(b_d) >= 100 and date(now())-date(b_d) < 150 then '+100'" & vbCrLf
    q = q & " when date(now())-date(b_d) >= 150 and date(now())-date(b_d) < 175 then '+150'" & vbCrLf
    q = q & "  when date(now())-date(b_d) >= 175 then '+175'" & vbCrLf
    q = q & "end as type,"
    q = q & " right(MOBILENO,8) kananmgm1,right(MOBILENO2,8) kananmgm2,right(MOBILENOADD1,8) kananmgm3, right(MOBILENOADD2,8) kananmgm4 from mgm" & vbCrLf
    q = q & " ) b on a.kananinbox = b.kananmgm1 or a.kananinbox = b.kananmgm2 or a.kananinbox = b.kananmgm3 or a.kananinbox = b.kananmgm4" & vbCrLf
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Label1.Caption = "Total : " & rs.RecordCount
    
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , rs(0))
             listItem.SubItems(1) = cnull(rs(1))
             listItem.SubItems(2) = cnull(rs(2))
             listItem.SubItems(3) = cnull(rs(3))
             listItem.SubItems(4) = cnull(rs(4))
        rs.MoveNext
    Wend
    
    
End Sub

Private Sub Command1_Click()
    Call isilv1
End Sub

Private Sub Command2_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    If ListView2.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView2.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView2.ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView2.ListItems.Count + 1
            For col = 1 To ListView2.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = ListView2.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = "'" + ListView2.ListItems(Row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(Row, col).Value = hasil1
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD_save.ShowOpen
        a = CD_save.FileName
     
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
        MsgBox "No data to export", vbInformation, Me.Caption
    End If

End Sub

Private Sub Form_Load()
    Call header
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
