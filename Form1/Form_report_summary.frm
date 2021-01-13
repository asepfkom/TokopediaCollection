VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_report_summary 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Report Summary Status Call"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14415
   LinkTopic       =   "Form3"
   ScaleHeight     =   8670
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   6360
      Left            =   30
      TabIndex        =   11
      Top             =   2085
      Width           =   14250
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6150
         Left            =   0
         TabIndex        =   12
         Top             =   150
         Width           =   14160
         _ExtentX        =   24977
         _ExtentY        =   10848
         _Version        =   393216
         AllowUpdate     =   0   'False
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria Report"
      Height          =   2100
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   14280
      Begin VB.TextBox txtlead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   13245
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   915
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   13275
         TabIndex        =   8
         Top             =   -930
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cbocampaign 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   270
         Width           =   4035
      End
      Begin VB.CommandButton SSCommand1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel"
         Height          =   405
         Left            =   5655
         Picture         =   "Form_report_summary.frx":0000
         TabIndex        =   6
         Top             =   1125
         Width           =   1590
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Show"
         Height          =   360
         Left            =   5655
         Picture         =   "Form_report_summary.frx":0766
         TabIndex        =   5
         Top             =   255
         Width           =   1605
      End
      Begin VB.CommandButton SSCommand2 
         BackColor       =   &H00F1E5DB&
         Caption         =   "Batal"
         Height          =   375
         Left            =   5655
         Picture         =   "Form_report_summary.frx":0D54
         TabIndex        =   4
         Top             =   690
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E87211&
         Caption         =   "Show Phone Number"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   13995
         TabIndex        =   3
         Top             =   -1290
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ComboBox cboagentname 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   615
         Width           =   1635
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3225
         TabIndex        =   1
         Top             =   615
         Width           =   2370
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   13755
         Top             =   -570
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.xls"
      End
      Begin Crystal.CrystalReport RPT 
         Left            =   13275
         Top             =   -570
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Lead"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12240
         TabIndex        =   14
         Top             =   1695
         Width           =   915
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Campaign"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Telesales "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   9
         Top             =   615
         Width           =   1425
      End
   End
End
Attribute VB_Name = "Form_report_summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sGetSPV As String
Dim querygwa As String
Private Sub cboagentname_Click()
    cboagentname_LostFocus
End Sub

Private Sub cboagentname_DropDown()
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

Strsql = "select * from usertbl where usertype='1'"

M_objrs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
cboagentname.clear
    While Not M_objrs.EOF
        cboagentname.AddItem IIf(IsNull(M_objrs!Userid), "", M_objrs!Userid)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub cboagentname_LostFocus()
Dim MOBJR As New ADODB.Recordset
Set MOBJR = New ADODB.Recordset
   MOBJR.CursorLocation = adUseClient
   
Strsql = "select * from usertbl where userid='" + cboagentname.text + "'"
MOBJR.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If Not MOBJR.EOF Then
Combo3.text = IIf(IsNull(MOBJR!agent), "", MOBJR!agent)
End If
Set MOBJR = Nothing

End Sub

Private Sub cbocampaign_DropDown()
sStrsql = "select * from datasourcetbl where status ='1' "

'    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
'        sStrsql = sStrsql & " and KODEDS in (select distinct recsource from mgm where agent in (select userid from usertbl where spvcode = '" & MDIForm1.txtUserName.text & "' or userid = '" & MDIForm1.txtUserName.text & "'))"
'    End If

Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql & " order by  kodeds ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    cbocampaign.clear
    While Not M_objrs.EOF
        cbocampaign.AddItem IIf(IsNull(M_objrs!KODEDS), "", M_objrs!KODEDS)
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing
End Sub
'Public Sub load_spv()
'    If MDIForm1.Text2.text = "Supervisor" Then
'        sStrsql = "select userid , agent from usertbl where  userid = '" + MDIForm1.txtUserName.text + "' and  aktif ='1'"
'    Else
'        sStrsql = "select userid , agent  from  usertbl  where  aktif ='1' and  kdlevel ='2'"
'    End If
'
'    Set M_objrs = New ADODB.Recordset
'        M_objrs.CursorLocation = adUseClient
'        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'        CBOTEAMNAME.clear
'        While Not M_objrs.EOF
'                CBOTEAMNAME.AddItem IIf(IsNull(M_objrs!Userid), "", M_objrs!Userid) & "!" & IIf(IsNull(M_objrs!agent), "", M_objrs!agent)
'                M_objrs.MoveNext
'        Wend
'
'    Set M_objrs = Nothing
'End Sub

Private Sub CmdCari_Click()
    Dim mobj As ADODB.Recordset
    Dim jml As Double
    Dim getSpvcode As String
    Dim getSpv_name As String
    Dim getUserid As String
    Dim getCampaign_code As String
    Dim getCampaign_name As String
    Dim strsql1 As String
    If cboagentname.text <> "" Then
        intvrl = InStr(1, cboagentname.text, "!", vbTextCompare)
        If intvrl <> 0 Then
            ArrayString = Split(cboagentname.text, "!", 2, vbTextCompare)
            getUserid = ArrayString(0)
            getUser_name = ArrayString(1)
        End If
    End If
    
    Strsql = " select * from vw_report_summary2"
    'strsql = " select * from report_summarytracking"
    mwhere = " WHERE 1=1 "
    
    If cbocampaign.text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where ""Campaign"" = '" & cbocampaign.text & "'"
        Else
            mwhere = mwhere + " and ""Campaign"" = '" & cbocampaign.text & "'"
        End If
    End If
    
    If cboagentname.text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where ""Agent"" ='" + cboagentname.text + "'"
        Else
            mwhere = mwhere + " and  ""Agent"" ='" + cboagentname.text + "'"
        End If
    End If

    Set mobj = New ADODB.Recordset
    mobj.CursorLocation = adUseClient
    
    mobj.Open Strsql + mwhere, M_OBJCONN, adOpenKeyset, adLockOptimistic
    
    txtlead.text = mobj.RecordCount
    Set DataGrid1.DATASOURCE = mobj
    cmdCari.Enabled = True
End Sub

Private Sub Combo3_Click()
Combo3_LostFocus
End Sub

Private Sub Combo3_DropDown()
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

Strsql = "select * from usertbl where aktif='1'  and usertype='1'"

M_objrs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
Combo3.clear
    While Not M_objrs.EOF
      Combo3.AddItem IIf(IsNull(M_objrs!agent), "", M_objrs!agent)
        M_objrs.MoveNext
    Wend
 Set M_objrs = Nothing

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo3_LostFocus()
Dim MOBJR As New ADODB.Recordset
Set MOBJR = New ADODB.Recordset
   MOBJR.CursorLocation = adUseClient
   
Strsql = "select * from usertbl where agent='" + Combo3.text + "'"
MOBJR.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If Not MOBJR.EOF Then
    cboagentname.text = IIf(IsNull(MOBJR!Userid), "", MOBJR!Userid)
End If
Set MOBJR = Nothing

End Sub
Public Sub load_spv1()
    Dim listv As ListItem
    If MDIForm1.Text2.text = "Supervisor" Then
        sStrsql = " select userid , agent  from usertbl where  userid = '" + MDIForm1.Text1.text + "' and  aktif ='1'"
    Else
        sStrsql = "select userid , agent  from usertbl  where  aktif ='1' and  level_name ='Supervisor'"
    End If
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        List_Supervisor.ListItems.clear
        While Not M_objrs.EOF
                Set listv = List_Supervisor.ListItems.ADD(, , IIf(IsNull(M_objrs!Userid), "", M_objrs!Userid))
                listv.SubItems(1) = IIf(IsNull(M_objrs!agent), "", M_objrs!agent)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

Private Sub Check_all1_Click()
    Dim i As Integer
    i = 0
    If Check_all1.Value = 1 Then
        For i = 1 To List_Supervisor.ListItems.Count
            List_Supervisor.ListItems(i).Checked = True
        Next i
    ElseIf Check_all1.Value = 0 Then
        For i = 1 To List_Supervisor.ListItems.Count
            List_Supervisor.ListItems(i).Checked = False
        Next i
    End If
    Call GetSPVs
End Sub
Private Sub GetSPVs()
    Dim sWhere As String
    sWhere = ""
    sWhere = GETSPV
    If sWhere <> "" Then
        sGetSPV = ""
        sGetSPV = sWhere
        Exit Sub
    End If
End Sub
Public Function GETSPV() As Variant
    Dim row As Double
    row = 1
    Strsql = ""
    For i = 1 To List_Supervisor.ListItems.Count
       If List_Supervisor.ListItems(i).Checked = True Then
            If row = 1 Then
                Strsql = "'" + List_Supervisor.ListItems(i).text + "'"
            Else
                Strsql = Strsql + ",'" + List_Supervisor.ListItems(i).text + "'"
            End If
            row = row + 1
      End If
    Next i
    GETSPV = Strsql
End Function

Private Sub List_Supervisor_Click()
    Call GetSPVs
End Sub

Private Sub SSCommand1_Click()
    export_data
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub
Public Sub export_data()
    Dim mobj As ADODB.Recordset
    Dim jml As Double
    Dim getSpvcode As String
    Dim getSpv_name As String
    Dim getUserid As String
    Dim getCampaign_code As String
    Dim getCampaign_name As String
    Dim strsql1 As String
    Dim Strsql As String
    If cboagentname.text <> "" Then
        intvrl = InStr(1, cboagentname.text, "!", vbTextCompare)
        If intvrl <> 0 Then
            ArrayString = Split(cboagentname.text, "!", 2, vbTextCompare)
            getUserid = ArrayString(0)
            getUser_name = ArrayString(1)
        End If
    End If
    
    Strsql = " select * from vw_report_summary2"
    'strsql = " select * from report_summarytracking"
    mwhere = " WHERE 1=1 "
    
    If cbocampaign.text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where ""Campaign"" = '" & cbocampaign.text & "'"
        Else
            mwhere = mwhere + " and ""Campaign"" = '" & cbocampaign.text & "'"
        End If
    End If
    
    If cboagentname.text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where ""Agent"" ='" + cboagentname.text + "'"
        Else
            mwhere = mwhere + " and  ""Agent"" ='" + cboagentname.text + "'"
        End If
    End If
    
    isi_data (Strsql + mwhere)

End Sub
Private Sub isi_data(Strsql As String)
On Error GoTo SALAH
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    
    i = 1
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
  
form_save:
    Cd_save.ShowSave
    TxtPath.text = Cd_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo form_save        '-> maka goto form_save
        End If
    End If

    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    On Error GoTo SALAH
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Integer
        If M_objrs.state = 1 Then
            x = 0
            Y = M_objrs.fields().Count - 1
            Do Until x > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_objrs.fields(x).Name)
                i = i + 1
                x = x + 1
            Loop
        End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing

    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
SALAH:
    Exit Sub
End Sub

Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
End Sub
