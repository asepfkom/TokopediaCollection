VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form form_tracking_agent 
   Caption         =   "Tracking Agent"
   ClientHeight    =   10050
   ClientLeft      =   345
   ClientTop       =   270
   ClientWidth     =   15165
   LinkTopic       =   "Form5"
   ScaleHeight     =   10050
   ScaleWidth      =   15165
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Height          =   8535
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   15015
      Begin MSComctlLib.ListView ListView2 
         Height          =   8310
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   14658
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Height          =   375
         Left            =   9240
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   375
         Left            =   9240
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin TDBDate6Ctl.TDBDate StartDate 
         Height          =   315
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "form_tracking_agent.frx":0000
         Caption         =   "form_tracking_agent.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_tracking_agent.frx":0184
         Keys            =   "form_tracking_agent.frx":01A2
         Spin            =   "form_tracking_agent.frx":0200
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         Format          =   "dd/mm/yyyy"
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.12794198814265E-317
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate EndDate 
         Height          =   315
         Left            =   7035
         TabIndex        =   10
         Top             =   240
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "form_tracking_agent.frx":0228
         Caption         =   "form_tracking_agent.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_tracking_agent.frx":03AC
         Keys            =   "form_tracking_agent.frx":03CA
         Spin            =   "form_tracking_agent.frx":0428
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         Format          =   "dd/mm/yyyy"
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.12794198814265E-317
         CenturyMode     =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6600
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "form_tracking_agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
ListView2.ColumnHeaders.CLEAR

    ListView2.ColumnHeaders.ADD 1, , "TL", 10 * 120
    ListView2.ColumnHeaders.ADD 2, , "Agent", 8 * 120
    ListView2.ColumnHeaders.ADD 3, , "BLANK", 8 * 120
    ListView2.ColumnHeaders.ADD 4, , "OS", 8 * 120
    ListView2.ColumnHeaders.ADD 5, , "VL", 8 * 120
    ListView2.ColumnHeaders.ADD 6, , "PR", 8 * 120
    ListView2.ColumnHeaders.ADD 7, , "ON", 8 * 120
    ListView2.ColumnHeaders.ADD 8, , "PTP", 8 * 120
    ListView2.ColumnHeaders.ADD 9, , "BP", 8 * 120
    ListView2.ColumnHeaders.ADD 10, , "POP", 8 * 120
    ListView2.ColumnHeaders.ADD 11, , "PO", 8 * 120
    ListView2.ColumnHeaders.ADD 12, , "SP", 8 * 120
    ListView2.ColumnHeaders.ADD 13, , "CO", 8 * 120
    ListView2.ColumnHeaders.ADD 14, , "Jumlah touch", 8 * 120
    ListView2.ColumnHeaders.ADD 15, , "Jumlah Data", 8 * 120
End Sub

Private Sub Command1_Click()
    Call search
End Sub

Private Sub Command2_Click()
    Call export
End Sub

Private Sub Form_Load()
    Call header
End Sub

Private Sub export()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If ListView2.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView2.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView2.ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView2.ListItems.Count + 1
            For col = 1 To ListView2.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = "'" + ListView2.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = ListView2.ListItems(Row - 1).SubItems(col - 1)
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
zzz:
        MsgBox "No data to export", vbInformation, Me.Caption
    End If
End Sub


Private Sub search()
    ListView2.ListItems.CLEAR
        
    query = " select b.* , a.jml_data, c.team from (" & vbCrLf
    query = query + "  select agent, count(agent) as jml_data from mgm group by 1 ) a" & vbCrLf
    query = query + "  right join" & vbCrLf
    query = query + " (select agent, ""BP-POP"" + ""BP-"" as bp, ""SP-"" as sp, ""OS-"" + ""OS-On"" as os , ""PTP-NE"" + ""PTP"" + ""PTP-PO"" as PTP, ""VL"" + ""VL-"" as vl, ""ON-"" as on, ""CO-"" as co, ""PO-"" as po, ""PR-"" as pr, ""POP"" as pop, ""blank"" , " & vbCrLf
    query = query + " ""BP-POP"" + ""SP-"" + ""OS-"" + ""PTP-NE"" + ""VL"" + ""PTP"" + ""ON-"" + ""BP-"" + ""CO-"" + ""PO-"" + ""OS-On"" + ""PR-"" + ""PTP-PO"" + ""VL-"" + ""POP"" + ""blank"" as jumlah from (" & vbCrLf
    query = query + " select agent, sum(""BP-POP"") ""BP-POP"", sum(""SP-SETTLE PAYMENT"") ""SP-"", sum(""OS-"") ""OS-"", sum(""PTP-NE"") ""PTP-NE"", sum(""VL-VALID"") ""VL"", sum(""PTP"") ""PTP"", sum(""ON-"") ""ON-"", sum(""BP-"") ""BP-""," & vbCrLf
    query = query + " sum(""CO-"") ""CO-"", sum(""PO-"") ""PO-"", sum(""OS-On Process"") ""OS-On"", sum(""PR-"") ""PR-"", sum(""PTP-PO"") ""PTP-PO"", sum(""VL-"") ""VL-"", sum(""POP"") ""POP"", sum(""blank"") ""blank"" from ( " & vbCrLf
    query = query + " select agent, tglcall, custid," & vbCrLf
    query = query + " case when f_cek_new = 'BP-POP' then 1 else 0 end as ""BP-POP""," & vbCrLf
    query = query + " case when f_cek_new = 'SP-SETTLE PAYMENT' then 1 else 0 end as ""SP-SETTLE PAYMENT""," & vbCrLf
    query = query + " case when f_cek_new = 'OS-' then 1 else 0 end as ""OS-""," & vbCrLf
    query = query + " case when f_cek_new = 'PTP-NE' then 1 else 0 end as ""PTP-NE""," & vbCrLf
    query = query + " case when f_cek_new = 'VL-VALID' then 1 else 0 end as ""VL-VALID""," & vbCrLf
    query = query + " case when f_cek_new = 'PTP' then 1 else 0 end as ""PTP""," & vbCrLf
    query = query + " case when f_cek_new = 'ON-' then 1 else 0 end as ""ON-""," & vbCrLf
    query = query + " case when f_cek_new = 'BP-' then 1 else 0 end as ""BP-""," & vbCrLf
    query = query + " case when f_cek_new = 'CO-' then 1 else 0 end as ""CO-""," & vbCrLf
    query = query + " case when f_cek_new = 'PO-' then 1 else 0 end as ""PO-""," & vbCrLf
    query = query + " case when f_cek_new = 'OS-On Process' then 1 else 0 end as ""OS-On Process"", " & vbCrLf
    query = query + " case when f_cek_new = 'PR-' then 1 else 0 end as ""PR-""," & vbCrLf
    query = query + " case when f_cek_new = 'PTP-PO' then 1 else 0 end as ""PTP-PO""," & vbCrLf
    query = query + " case when f_cek_new = 'VL-' then 1 else 0 end as ""VL-""," & vbCrLf
    query = query + " case when f_cek_new = 'POP' then 1 else 0 end as ""POP""," & vbCrLf
    query = query + " case when f_cek_new = '' or f_cek_new is null then 1 else 0 end as ""blank""" & vbCrLf
    If StartDate.text = "__/__/____" And EndDate.text = "__/__/____" Then
        query = query + "  from mgm where date(tglcall) = date(now()) order by 1) a group by 1 " & vbCrLf
    Else
        a = Format(StartDate.Value, "yyyy-mm-dd")
        B = Format(StartDate.Value, "yyyy-mm-dd")
        
        query = query + "  from mgm where date(tglcall) between '" + a + "' and '" + B + "' order by 1) a group by 1 " & vbCrLf
    End If
    query = query + "  ) zzz" & vbCrLf
    query = query + " ) b on a.agent = b.agent " & vbCrLf
    query = query + " join (select userid, substring(team, 3, 8)::integer as team from usertbl where userid ilike 'D%' and userid <> 'DODDY' and userid <> 'DECEASE' and userid <> 'DESSY' order by 2) c on a.agent = c.userid" & vbCrLf
    
    If Text1.text <> "" And Text2.text <> "" Then
        a = Mid(Text1.text, 3, 8)
        B = Mid(Text2.text, 3, 8)
        query = query + " and ( team >= '" & a & "' and team <= '" & B & "' ) "
    ElseIf Text1.text <> "" And Text2.text = "" Then
        Text2.text = Text1.text
        a = Mid(Text1.text, 3, 8)
        B = Mid(Text2.text, 3, 8)
        query = query + " and ( team >= '" & a & "' and team <= '" & B & "' ) "
    ElseIf Text1.text = "" And Text2.text <> "" Then
        Text1.text = Text2.text
        a = Mid(Text1.text, 3, 8)
        B = Mid(Text2.text, 3, 8)
        query = query + " and ( team >= '" & a & "' and team <= '" & B & "' ) "
    End If
    
    If Text3.text <> "" And Text4.text <> "" Then
        query = query + " and ( a.agent >= '" & Text3.text & "' and a.agent <= '" & Text4.text & "' ) "
    ElseIf Text3.text <> "" And Text4.text = "" Then
        Text4.text = Text3.text
        query = query + " and ( a.agent >= '" & Text3.text & "' and a.agent <= '" & Text4.text & "' ) "
    ElseIf Text3.text = "" And Text4.text <> "" Then
        Text4.text = Text3.text
        query = query + " and ( a.agent >= '" & Text3.text & "' and a.agent <= '" & Text4.text & "' ) "
    End If
    
    query = query + " order by team, agent" & vbCrLf
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , "TL" & cnull(rs("team")))
             listItem.SubItems(1) = cnull(rs("agent"))
             listItem.SubItems(2) = cnull(rs("blank"))
             listItem.SubItems(3) = cnull(rs("os"))
             listItem.SubItems(4) = cnull(rs("vl"))
             listItem.SubItems(5) = cnull(rs("pr"))
             listItem.SubItems(6) = cnull(rs("on"))
             listItem.SubItems(7) = cnull(rs("ptp"))
             listItem.SubItems(8) = cnull(rs("bp"))
             listItem.SubItems(9) = cnull(rs("pop"))
             listItem.SubItems(10) = cnull(rs("po"))
             listItem.SubItems(11) = cnull(rs("sp"))
             listItem.SubItems(12) = cnull(rs("co"))
             listItem.SubItems(13) = cnull(rs("jumlah"))
             listItem.SubItems(14) = cnull(rs("jml_data"))
        rs.MoveNext
    Wend

End Sub
