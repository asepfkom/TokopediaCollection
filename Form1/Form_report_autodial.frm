VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_Report_AutoDial 
   Caption         =   "Form1"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3525
      Left            =   0
      ScaleHeight     =   3465
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   1950
         Left            =   75
         TabIndex        =   1
         Top             =   -15
         Width           =   5715
         Begin VB.CommandButton cmdProses 
            BackColor       =   &H80000004&
            Caption         =   "Proses"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4155
            TabIndex        =   3
            Top             =   1425
            Width           =   975
         End
         Begin VB.TextBox TxtPath 
            Height          =   285
            Left            =   1185
            TabIndex        =   2
            Top             =   1500
            Visible         =   0   'False
            Width           =   1515
         End
         Begin TDBDate6Ctl.TDBDate TdTglCall1 
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   1035
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            Calendar        =   "Form_report_autodial.frx":0000
            Caption         =   "Form_report_autodial.frx":0118
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_report_autodial.frx":0184
            Keys            =   "Form_report_autodial.frx":01A2
            Spin            =   "Form_report_autodial.frx":0200
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd-mmm-yyyy"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   0
            Format          =   "dd-mm-yyyy"
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
            Text            =   "__-__-____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   37468
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TdTglCall2 
            Height          =   285
            Left            =   3570
            TabIndex        =   5
            Top             =   1035
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            Calendar        =   "Form_report_autodial.frx":0228
            Caption         =   "Form_report_autodial.frx":0340
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_report_autodial.frx":03AC
            Keys            =   "Form_report_autodial.frx":03CA
            Spin            =   "Form_report_autodial.frx":0428
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd-mmm-yyyy"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   0
            Format          =   "dd-mm-yyyy"
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
            Text            =   "__-__-____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   37468
            CenturyMode     =   0
         End
         Begin MSComDlg.CommonDialog CD_Save 
            Left            =   135
            Top             =   1395
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5460
            Y1              =   795
            Y2              =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2925
            TabIndex        =   8
            Top             =   1035
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   7
            Top             =   1035
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Report Agent Activity"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   450
            TabIndex        =   6
            Top             =   165
            Width           =   4440
         End
      End
   End
End
Attribute VB_Name = "Form_Report_AutoDial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProses_Click()
    Dim objExcel As Excel.Application
    Dim objBook As Excel.Workbook
    Dim objSheet As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    sWhere = ("date(date_break) between '" & Format(TdTglCall1.Value, "YYYY-MM-DD") & "' and '" & Format(TdTglCall2.Value, "YYYY-MM-DD") & "'")
    
    sQuerySelect = "select agent,  sum(""ManualDial Duration"")as ""ManualDial Duration"", sum(""Start Auto Dial Duration"") as ""Start Auto Dial Duration"", sum(""Form Break Show Duration"") as ""Form Break Show Duration"","
    sQuerySelect = sQuerySelect + vbCrLf + "sum(""Lunch Duration"") as ""Lunch Duration"", sum(""Meeting Duration"") as ""Meeting Duration"", sum(""Pray Duration"") as ""Pray Duration"" from"
    sQuerySelect = sQuerySelect + vbCrLf + " ( select agent,"
    sQuerySelect = sQuerySelect + vbCrLf + " CASE WHEN status_break::text = 'ManualDial'::text THEN durasi else '0'::interval END AS ""ManualDial Duration"","
    sQuerySelect = sQuerySelect + vbCrLf + "CASE WHEN status_break::text = 'start_autodialer'::text THEN durasi else  '0'::interval END AS ""Start Auto Dial Duration"","
    sQuerySelect = sQuerySelect + vbCrLf + "CASE WHEN status_break::text = 'form break show'::text THEN durasi else  '0'::interval END AS ""Form Break Show Duration"","
    sQuerySelect = sQuerySelect + vbCrLf + "CASE WHEN status_break::text = 'Lunch'::text THEN durasi else  '0'::interval END AS ""Lunch Duration"","
    sQuerySelect = sQuerySelect + vbCrLf + "CASE WHEN status_break::text = 'Meeting'::text THEN durasi else  '0'::interval END AS ""Meeting Duration"","
    sQuerySelect = sQuerySelect + vbCrLf + "CASE WHEN status_break::text = 'Pray'::text THEN durasi else  '0'::interval END AS ""Pray Duration"" from ("
    sQuerySelect = sQuerySelect + vbCrLf + "select agent, status_break ,sum(durasi) as durasi from tbl_autodialer_agent_break where status_break <> '' and " & sWhere & ""
    sQuerySelect = sQuerySelect + vbCrLf + "group by agent, status_break)a ) b group by agent"
    
    i = 1
    Set aRS = New ADODB.Recordset
    aRS.CursorLocation = adUseClient
    aRS.Open sQuerySelect, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
form_save:
    
    Cd_save.ShowSave
    TxtPath.text = Cd_save.FileName
    
    If TxtPath.text = Empty Then
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        If m_msgbox = vbYes Then
            MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then
            GoTo form_save
        End If
    End If
    
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
    
    On Error GoTo SALAH
    Dim x, Y    As Integer
    If aRS.state = 1 Then
        x = 0
        Y = aRS.fields.Count - 1
        Do Until x > Y
            DoEvents
            objSheet.Cells(1, i).Value = CStr(aRS.fields(x).Name)
            i = i + 1
            x = x + 1
        Loop
    End If
    objSheet.Range("A2").CopyFromRecordset aRS
    objBook.SaveAs TxtPath.text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set aRS = Nothing
    Exit Sub
SALAH:
    MsgBox err.Description
    Exit Sub
End Sub

