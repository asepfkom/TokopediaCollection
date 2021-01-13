VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_report_call 
   Caption         =   "Report Call "
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15540
   LinkTopic       =   "Form3"
   ScaleHeight     =   4815
   ScaleWidth      =   15540
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Caption         =   "Dashboard"
      Height          =   4800
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   15540
      Begin VB.Frame Frame4 
         Caption         =   "Search"
         Height          =   4455
         Left            =   12180
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frm_report_call.frx":0000
            Left            =   120
            List            =   "frm_report_call.frx":0002
            TabIndex        =   5
            Top             =   1560
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Search"
            Height          =   375
            Left            =   1680
            TabIndex        =   4
            Top             =   3360
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Export"
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Touch per Custid per Agent"
            Height          =   435
            Left            =   1680
            TabIndex        =   2
            Top             =   2880
            Visible         =   0   'False
            Width           =   1455
         End
         Begin TDBDate6Ctl.TDBDate TDBDate3 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   503
            Calendar        =   "frm_report_call.frx":0004
            Caption         =   "frm_report_call.frx":011C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frm_report_call.frx":0188
            Keys            =   "frm_report_call.frx":01A6
            Spin            =   "frm_report_call.frx":0204
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648447
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate4 
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Top             =   720
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   503
            Calendar        =   "frm_report_call.frx":022C
            Caption         =   "frm_report_call.frx":0344
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frm_report_call.frx":03B0
            Keys            =   "frm_report_call.frx":03CE
            Spin            =   "frm_report_call.frx":042C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648447
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin MSComDlg.CommonDialog CD_save 
            Left            =   2670
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label8 
            Caption         =   "Client"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
      End
      Begin MSComctlLib.ListView LvAgent 
         Height          =   4455
         Left            =   45
         TabIndex        =   9
         Top             =   240
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frm_report_call"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dashboard()
   
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = " select kdnoprodpresented as stts from contacteddesc where status = '1' order by nmnoprodpresented"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LVAgent.ColumnHeaders.clear
    LVAgent.ColumnHeaders.ADD 1, , "No", 10 * 120
    LVAgent.ColumnHeaders.ADD 2, , "Agent", 10 * 120
    z = 3
    While Not M_objrs.EOF
        LVAgent.ColumnHeaders.ADD z, , "" & M_objrs!stts & "", 7 * 120
        M_objrs.MoveNext
        z = z + 1
    Wend
    LVAgent.ColumnHeaders.ADD z, , "TOTAL", 10 * 120
    
    'isi
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = " select nmnoprodpresented as stts from contacteddesc where status = '1'  order by 1"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    a = ""
    b = ""
    c = ""
    
    While Not M_objrs.EOF
        a = a + " ,case when kodeds = '" & "" & M_objrs!stts & "" & "' then 1 else 0 end as """ & "" & M_objrs!stts & """"
        b = b + """" & "" & M_objrs!stts & """+"
        c = c + " ,sum( """ & "" & M_objrs!stts & """" & " ) as """ & "" & M_objrs!stts & """"
        M_objrs.MoveNext
    Wend
        b = Left(b, Len(b) - 1)
        c = c
    
    '=========asep19/01/2020======'
    q = " select agent" & "" & c & ", sum(total) as total from ("
    q = q + "select *," & "" & b & " as Total from ("
    q = q & "select agent " & "" & a & ""
    q = q & "from (select agent, custid, kodeds from mgm_hst"
    q = q & " where tgl between '" & Format(TDBDate3.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(TDBDate4.Value, "yyyy-mm-dd") & " 23:59:59' "

'    If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Then
'        q = q & " and in (select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.Text1.text & "' or userid = '" & MDIForm1.Text1.text & "' )) "
'    End If
'
    If Combo1.text = "RUPIAH PLUS" Then
        q = q & " ) hst "
    ElseIf Combo1.text = "UANGEXPRESS" Then
        q = q & " ) hst "
    ElseIf Combo1.text = "GLOBALINDO" Then
        q = q & " ) hst "
    Else
        q = q & " ) hst "
    End If


    q = q & " ) abc "
    q = q & " ) a group by agent "
    Set M_objrs = New ADODB.Recordset
    
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LVAgent.ListItems.clear
    While Not M_objrs.EOF
        Set ListItem = LVAgent.ListItems.ADD(, , M_objrs.Bookmark)
        For i = 1 To z - 1
            ListItem.SubItems(i) = IIf(IsNull(M_objrs(i - 1)), "", M_objrs(i - 1))
        Next i
        M_objrs.MoveNext
    Wend

End Sub

Private Sub Command6_Click()
    Call dashboard
End Sub
Private Sub Command7_Click()
Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet
Dim col, row As Integer
Dim a As String
If LVAgent.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To LVAgent.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = LVAgent.ColumnHeaders(col)
    Next
 
    For row = 2 To LVAgent.ListItems.Count + 1
        For col = 1 To LVAgent.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(row, col).Value = LVAgent.ListItems(row - 1).text
        Else
            '" 'cararandy 29032016 "
            Dim hasil1 As String
                hasil1 = "'" + LVAgent.ListItems(row - 1).SubItems(col - 1)
                objExcelSheet.Cells(row, col).Value = hasil1
            End If
        Next
    Next
 
    objExcelSheet.Columns.AutoFit
    Cd_save.ShowOpen
    a = Cd_save.FileName
 
    objExcelSheet.SaveAs a & ".xls"
    MsgBox "Export Completed", vbInformation, Me.Caption
 
    objExcel.Workbooks.Open a & ".xls"
    objExcel.Visible = True
Else
    MsgBox "No data to export", vbInformation, Me.Caption
End If
End Sub
