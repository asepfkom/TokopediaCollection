Attribute VB_Name = "Module"
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
 
 
Private Const MF_BYPOSITION = &H400&

Public signtimer2 As Boolean
Public signtimer7 As Boolean
Public signtimes As String

Option Base 0
Public bcekptp As Boolean
Public vrcekamont As String
Public strStatusCpa As String
'VARIABEL NENTUIN OBELISK APA ORANGE
Public Obelisk As Boolean
Public waktu_iddel As String

Public f_must_open As Boolean

Public uniqpublic As String
'-----------------------------------
Global regnego As Boolean
Global Const CB_ERR = -1
Global Const CB_FINDSTRING = &H14C
Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
    
Private Const EM_GETLINECOUNT = &HBA

Public M_OBJCONN As New ADODB.Connection
Public HELPER_OBJCONN As New ADODB.Connection
Public Declare Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public list As ListItem

'untuk bill
Global COST, zon, tot, totcost, tarif, rounding As Integer
Global detik1, menit1, jam1, Cnt, SEC, SECA As Integer
Global AWALTELP, akhirtelp As Date
Global jammulai As Date
Global telpno, idcust As String
Global hari As String
Global TGLSET As ADODB.Recordset
Public CMDSQLOPEN1 As String

Public glexp As String
Public bRenderrecord As Boolean
Public M_RPTCONN As New ADODB.Connection
Public CMDSQLOPEN As String
Public Addmgm As Boolean
Public StsmgmSchedule As Boolean
Public search_ok As Boolean
Public Flag_mgm As Boolean
Public statusclaim As Boolean
Public ADD_CUST As Boolean
Public Const TXT = 120
Public IPServer As String
Public mbIgnoreListClick As Boolean
Public fso As FileSystemObject
Public KET As String
Public reff_View As Boolean
Public reff_Duplikasi As Boolean
Public reff_Duplikasi1 As Boolean
Public TodayList As Boolean
Public POD As Boolean
Public Const SW_SHOWNORMAL = 1
Public F_LOCK As Boolean
'  updata listview after saving customer database==> value 1 untuk form prescreen,,,,  2 untuk form view_mgmdata
Public Status_Form As Integer

Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public b_pindah As Boolean
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public rsTemporary As New ADODB.Recordset
Public Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Public iListitem As ListItem
Public shedulePTP As Boolean
Public shedulePTP_Show As Boolean

Public cek_aksesall As String
Public waktu_start As Date
Public waktu_finish As Date
Public waktu_mulai_ngitung As Date
Public waktu_selesai_ngitung As Date


'Declare Function FindWindow Lib "user32" Alias _
'"FindWindowA" (ByVal lpClassName As String, _
'ByVal lpWindowName As String) As Long

Declare Function GetWindow Lib "user32" (ByVal hwnd _
As Long, ByVal wCmd As Long) As Long

Declare Function OpenIcon Lib "user32" (ByVal hwnd _
As Long) As Long

Declare Function SetForegroundWindow Lib "user32" _
(ByVal hwnd As Long) As Long
 Public M_OBJCONN1 As New ADODB.Connection
        
Public Const GW_HWNDPREV = 3
'@@ 5/04/2011 Buat nandain FrmCPA dipanggil dari Frmcc_collection atau FrmCC2_Collection
Public StatusCPA As String


'Buat bikin direktori
'Setelah Anda menjalankan program ini, pilih direktori 'yang Anda inginkan pada kotak dialog tersebut. Anda 'akan melihat sebuah kotak pesan yang menampilkan
'nama direktori yang Anda pilih tadi.

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Public Type BrowseInfo
  hwndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

'@@08102012, Buat HangUp X-Lite
Public THandle As Long


Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName _
As Any) As Long

'@@15012013 Tambahan buat ambil Nilai TLnya
Public UseridTL As String

'@@11022013
Public AksesAllAcc As String

' ## 08 April 2013
Public i_monitoring_activity As Integer
Public i_monitoring_activity_2 As Integer
Public main_timer_activity As Integer
Public b_cmdhangup As Boolean

Public sConnstring As String
' Api Functions
' For screen resolutions
Private Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
' Cursor function
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long

' -- LAST UPDATE 19 April 2013 untuk fungsi 5x blok
Public sPhone_Agent As String
Public sPhone_CustID As String
Public sPhone_TelpNo As String
'--------------------------------------------------

Public bReminder_agent As Boolean
Public sReminder_CUST_ID As String
Public bAktif_form_customer As Boolean
Public bAktif_Cust_Review As Boolean
Public open_sms As Boolean

Public count_timer_detik As Integer

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public f_lockAccount_enabled As Boolean

Public lemparformcc As String
Public strategi As Boolean
Public id_strategi As String
Public nm_strategi As String
Public query_strategi As String
Public exit_klik As Boolean
Public custid_autodial As String
Public custid_autodial_not_in As String

Public cti_get As String
Public c_rs_global As New ADODB.Recordset
Public Session_login As String
Public Session_ManualDial As String
Public Session_AutoDial As String
Public Session_Break As String
Public break_time As Boolean

Public bcp As Boolean
Public WsckCti_initiated, WsckCti_connected, WsckCti_busy, WsckCti_hangup As String

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub
    
Sub Main()
    'On Error GoTo SqlConnErr
    AutoDialerHangup = True
    AutoDialerBreak = False
    AutoDialerCall = False
    
    UseridTL = ""
    
    AksesAllAcc = ""
        
    bcp = False
    
    ' Server Development
    
    'CMDSQLOPEN = "Driver={PostgreSQL ANSI}; Server=192.168.20.2; PORT=5432; Database=tokopedia; UID=userappl; PWD=Monyong!"
    'BUAT LEPTOP VPN'
    'CMDSQLOPEN = "Driver={PostgreSQL ANSI}; Server=10.8.0.241; PORT=5432; Database=tokopedia; UID=userappl; PWD=Monyong!"
    
    'local
    CMDSQLOPEN = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=tokopedia; UID=postgres; PWD=123"

    ' Koneksi Utama
    
    M_OBJCONN.Open CMDSQLOPEN
    ' SMS
    'asli M_OBJCONN1.Open CMDSQLOPEN1
    
    ' Report
    On Error GoTo AccessConnErr
    'M_RPTCONN.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=Admin;Data Source=TINS_RITCARD"

    'On Error GoTo AccessConnErr
    'M_RPTCONN.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=Admin;Data Source=TINS_RITCARD"

    'M_RPTCONN.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=Admin;Data Source=TINS_RITPIL"
    'M_RPTCONN.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=Admin;Data Source=TINS_AWARNESS"
    'frmAbout1.Show
    frmlogin.Show
    'FrmMonitoringHeadset.Show
    Exit Sub
AccessConnErr:
    MsgBox err.Description
    'Set M_RPTCONN = Nothing
    Exit Sub
    
'SqlConnErr:
'    MsgBox err.Description
'    Set M_OBJCONN = Nothing
'    Exit Sub
End Sub

Public Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function

Public Sub offsesilogin(abc As String)
Dim query As String
    
    query = "UPDATE usertbl SET f_status_login=null,last_logout='" & waktu_server_sekarang & "' where userid = '" + abc + "'"
    M_OBJCONN.execute query
    
'    '---- update tbluser_log status logout
'    M_OBJCONN.execute " update usertbl_log set waktu_logout = now()::timestamp , durasi = now() - substring(waktu_login,1,19)::timestamp where session_login= '" + Session_login + "' and username='" + MDIForm1.Text1.text + "'"
'
'    'update session manual dialer
'    If Session_ManualDial <> "" Then
'    M_OBJCONN.execute " update tbl_autodialer_agent_break set waktu_end =now(), durasi = now() - waktu_start where sessionid= '" + Session_ManualDial + "' and agent='" + MDIForm1.Text1.text + "'"
'    End If
'    'update session autodialer
'    If Session_AutoDial <> "" Then
'    M_OBJCONN.execute " update tbl_autodialer_agent_break set waktu_end = '" & waktu_server_sekarang & "', durasi = '" & waktu_server_sekarang & "'::timestamp - '" + Session_AutoDial + "' where sessionid= '" + Session_AutoDial + "' and agent='" + MDIForm1.Text1.text + "'"
'    End If
'    'update session break
'    If Session_Break <> "" Then
'    M_OBJCONN.execute " update tbl_autodialer_agent_break set waktu_end = now(), durasi = now() - waktu_start where sessionid= '" + Session_Break + "' and agent='" + MDIForm1.Text1.text + "'"
'    End If
End Sub
Public Sub offsesilogin_new(abc As String)
Dim query As String
    
    query = "UPDATE usertbl SET f_status_login=null,last_logout='" & waktu_server_sekarang & "' where userid = '" + abc + "'"
    M_OBJCONN.execute query
    
    '---- update tbluser_log status logout
    M_OBJCONN.execute " update usertbl_log set waktu_logout = now()::timestamp , durasi = now() - substring(waktu_login,1,19)::timestamp where session_login= '" + Session_login + "' and username='" + abc + "' and coalesce(waktu_logout,'')=''"
    
    'update session manual dialer
    If Session_ManualDial <> "" Then
    M_OBJCONN.execute " update tbl_autodialer_agent_break set waktu_end =now(), durasi = now() - waktu_start where sessionid= '" + Session_ManualDial + "' and agent='" + abc + "'"
    End If
    'update session autodialer
    If Session_AutoDial <> "" Then
    M_OBJCONN.execute " update tbl_autodialer_agent_break set waktu_end = '" & waktu_server_sekarang & "', durasi = '" & waktu_server_sekarang & "'::timestamp - '" + Session_AutoDial + "' where sessionid= '" + Session_AutoDial + "' and agent='" + abc + "'"
    End If
    'update session break
    If Session_Break <> "" Then
    M_OBJCONN.execute " update tbl_autodialer_agent_break set waktu_end = now(), durasi = now() - waktu_start where sessionid= '" + Session_Break + "' and agent='" + MDIForm1.Text1.text + "'"
    End If
End Sub




Public Sub DisableCloseBtn(ByVal frm As Form)
    Dim h As Long
    
    h = GetSystemMenu(frm.hwnd, 0)
    RemoveMenu h, 6, &H400
    RemoveMenu h, 5, &H400
    
End Sub

Public Sub logwktcti(pesan As String)
    Dim ifilenumber As Integer
    Static iErrCtr As Integer
    
    
    iErrCtr = iErrCtr + 1
    
    ifilenumber = FreeFile
    Open "C:\LogCTI.txt" For Append As #ifilenumber
    
    
        Write #ifilenumber, pesan
    
    Close #ifilenumber
End Sub


Public Function CUSTNOMOR(M_OBJCONN As ADODB.Connection, VARNAME As String) As String
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim NOMOR1 As String
    Dim TGLNOMOR As String
    CMDSQL = "SELECT VARVALUE FROM COMMONCFG"
    CMDSQL = CMDSQL + " WHERE VARNAME = '" + VARNAME + "'"
    On Error GoTo ERRORA
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount <> 0 Then
        If Len(M_objrs("VARVALUE")) < 4 Then
            NOMOR1 = CStr(M_objrs("VARVALUE"))
            NOMOR1 = (Left("0000", 4 - Len(NOMOR1)) + NOMOR1)
        Else
            NOMOR1 = CStr(M_objrs("VARVALUE"))
        End If
        TGLNOMOR = CStr(Format(MDIForm1.TDBDate1.Value, "yyyymmdd"))
        CUSTNOMOR = TGLNOMOR & NOMOR1
        NOMOR1 = CStr((CCur(NOMOR1) + 1))
            CMDSQL = "UPDATE COMMONCFG SET VARVALUE = '" + NOMOR1 + "' "
            CMDSQL = CMDSQL + " WHERE VARNAME = '" + VARNAME + "'"
            M_OBJCONN.Open CMDSQLOPEN
            M_OBJCONN.execute CMDSQL
            Set M_OBJCONN = Nothing
    End If
    Set M_objrs = Nothing
    Exit Function
ERRORA:
    Set M_objrs = Nothing
End Function

Public Function UBAH_QUOTE(KATAUBAH As String)
    UBAH_QUOTE = Replace(KATAUBAH, "'", "`")
End Function

Public Function UBAH_STRIP(KATAUBAH As String)
    UBAH_STRIP = Replace(KATAUBAH, "- -", "-")
End Function

Public Function JADI_QUOTE(KATAJADI As String)
    JADI_QUOTE = Replace(KATAJADI, "`", "'")
End Function
Public Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function

Public Function UBAH_SEPARATOR(SEPARATOR As String)
    UBAH_SEPARATOR = Replace(SEPARATOR, ".", ",")
End Function

' END OPEN HELP
Function GetNamaNoSpace(ByVal Nama As String) As String
    Dim TXT As String ' general use
    Dim a As Long ' General use
    ' we loop through string removing bad values
    For a = 1 To Len(Nama)
        Select Case Asc(Mid(Nama, a, 1))
        Case 32
        Case Else  ' * or ,
            TXT = TXT + Mid(Nama, a, 1) 'add to txt
        End Select
    Next a
    GetNamaNoSpace = TXT
End Function

Sub WaitSecs(Seconds As Single)
    Dim a As Long
    Seconds = Seconds + Timer
    While Seconds > Timer
        a = DoEvents
    Wend
End Sub

Function GetNumber(ByVal NumberTXT As String) As String
    Dim TXT As String ' general use
    Dim a As Integer ' General use
    For a = 1 To Len(NumberTXT)
        Select Case Asc(Mid(NumberTXT, a, 1))
        Case 48 To 57 ' numbers
            TXT = TXT + Mid(NumberTXT, a, 1) 'add to txt
        Case 32, 44, 35  ' * or ,
            TXT = TXT + Mid(NumberTXT, a, 1) 'add to txt
        Case 120, 88
            a = Len(NumberTXT)
        Case Else
        End Select
    Next a
    GetNumber = TXT
End Function

Function GET_EXT(ByVal number As String) As String
    Dim TXT As String
    Dim a As Integer
    For a = 1 To Len(number)
    Next a
End Function

Function DELETE(filespec)
    Dim fso, F
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set F = fso.GetFile(filespec)
    F.DELETE
End Function

Public Sub SAVE_FILE_KONEKSI(M_FILE As String, M_database As String)
    Dim t As TextStream
    Set t = fso.OpenTextFile(App.Path & "\" & M_FILE, ForWriting, True)
    'Set t = fso.OpenTextFile(M_FILE, ForWriting, True)
    t.Write M_database
    t.Close
End Sub


Public Function BUKA_FILE_KONEKSI(M_FILE As String) As String
    Dim F As String
    Dim t As TextStream
    On Error GoTo HELL
    'Set t = fso.OpenTextFile(App.Path & "\" & M_FILE, ForReading)
    Set t = fso.OpenTextFile(App.Path & "\" & M_FILE, ForReading)
    BUKA_FILE_KONEKSI = t.ReadAll
    t.Close
    Exit Function
HELL:
        BUKA_FILE_KONEKSI = ""
    '    MsgBox Err.Description
End Function

Public Function StartMeUp(F As String)
    Dim i As Integer
    Dim d As String
    i = InStrRev(F, "\")
    If i > 0 Then
        d = Left(F, i - 1)
    Else
        d = App.Path
    End If
    StartMeUp = ShellExecute(MDIForm1.hwnd, "open", F, vbNullString, d, SW_SHOWNORMAL)
End Function

Public Sub cari_zone()
    Dim prs As ADODB.Recordset
    Dim rsrate As ADODB.Recordset
    Dim x As Integer
    Dim CMDSQL As String
    'Dim zon, tarif, rounding As Integer
    Dim n As String
    Dim awal, akhir As Date
    x = 8
    rounding = 0
    
    Do While x >= 1
        n = Left(telpno, x)
        Set prs = New ADODB.Recordset
        prs.CursorLocation = adUseClient
        CMDSQL = "select * from bill_countryprefix where prefix='" & n & "'"
        prs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If prs.BOF And prs.EOF Then
            x = x - 1
        Else
            zon = prs!bill_zone_id
            Set prs = Nothing
            Exit Do
        'x = x - 1
        End If
    Loop

    Set rsrate = New ADODB.Recordset
    rsrate.CursorLocation = adUseClient
    If UCase(hari) = "MINGGU" And (zon = 3 Or zon = 4) Then
        CMDSQL = "select timebandstart, timebandstop, cost, rounding from Bill_tarifrate WHERE Bill_Zone_id='" & zon & "' and holiday='t'"
    Else
        CMDSQL = "select timebandstart, timebandstop, cost, rounding from Bill_tarifrate WHERE Bill_Zone_id='" & zon & "'"
    End If
    rsrate.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If rsrate.BOF And rsrate.EOF Then
    Else
        Do While Not rsrate.EOF
            If Format(rsrate!timebandstart, "hh:mm:ss") <= jammulai Then
                If Format(rsrate!timebandstop, "hh:mm:ss") >= jammulai Then
                    tarif = rsrate!COST
                    rounding = rsrate!rounding
                    totcost = tarif
                    Exit Do
                End If
            End If
            rsrate.MoveNext
        Loop
    End If
    'tot = tarif
    'COST = COST - tarif
    FBILL.Label6.Caption = "Tarif : Rp." & tarif & "/" & rounding & " detik"
    FBILL.Label3.Caption = "Cost : " & tarif
    Set rsrate = Nothing
End Sub

Public Sub savecall()
    Dim hiscall As ADODB.Recordset
    Dim CMDSQL, durasi As String
    Dim AKHIRTELPON As String
    Dim AWALTELPON As String
    
    AWALTELPON = Format(AWALTELP, "yyyy-mm-dd hh:mm:ss")
    durasi = jam1 & ":" & menit1 & ":" & detik1
    Set hiscall = New ADODB.Recordset
    hiscall.CursorLocation = adUseClient
    CMDSQL = "Insert into callhistory (custid,agent,notelp,mulaitelp,"
    CMDSQL = CMDSQL + "durasi,detik,cost) values ("
    CMDSQL = CMDSQL + "'" & idcust & "',"
    CMDSQL = CMDSQL + "'" & MDIForm1.Text1.text & "',"
    CMDSQL = CMDSQL + "'" & telpno & "',"
    CMDSQL = CMDSQL + "'" & AWALTELPON & "',"
    CMDSQL = CMDSQL + "'" & durasi & "',"
    CMDSQL = CMDSQL + "'" & Cnt & "',"
    CMDSQL = CMDSQL + "'" & totcost & "')"
    M_OBJCONN.execute CMDSQL
    Set hiscall = Nothing
End Sub

Public Sub tglhost()
    Set TGLSET = New ADODB.Recordset
    TGLSET.CursorLocation = adUseClient
    TGLSET.Open "select tglsystem from vwcallcfg1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Select Case CStr(Weekday(Format(TGLSET!tglsystem, "yyyy-mm-dd")))
    Case 1:
        hari = "Minggu"
    Case 2:
        hari = "Senin"
    Case 3:
        hari = "Selasa"
    Case 4:
        hari = "rabu"
    Case 5:
        hari = "kamis"
    Case 6:
        hari = "Jumat"
    Case 7:
        hari = "Sabtu"
    End Select
    MDIForm1.Label2.Caption = TGLSET!tglsystem
    SEC = DateDiff("S", Now(), MDIForm1.Label2.Caption)
    Set TGLSET = Nothing
End Sub

Public Sub CheckSoftware(x As Form)
    On Error GoTo pesan
    Dim SaveTitle$
    If App.PrevInstance Then
        SaveTitle$ = App.Title
        MsgBox "Program Aplikasi yang dijalankan tidak bisa dua kali dijalankan!", _
               vbCritical, "Sedang Dijalankan"
        App.Title = ""
        x.Caption = ""
        AppActivate SaveTitle$
        Sendkeys "%{ENTER}", True
        End
    End If
    Exit Sub
pesan:
    End
    Exit Sub
End Sub

Public Function ENCRIPY(x As Boolean, strText As String) As String
    Dim str As String
    Dim i As Integer
    'JIKA VARIABEL X = FALSE MAKA DATA DI ENKRIPSI
    If x = False Then
        str = ""
        For i = 1 To Len(strText)
            str = str + Chr(Asc(Mid(strText, i, 1)) + 27)
        Next i
       ENCRIPY = str
    
    Else
    'JIKA X=TRUE MAKA DEKRIPSI
      str = ""
        For i = 1 To Len(strText)
            str = str + Chr(Asc(Mid(strText, i, 1)) - 27)
        Next i
       ENCRIPY = str
    End If
End Function

Public Function TulisJalan(Hitung As Integer, _
    strKalimat As String, PANJANG As Integer)

    If Hitung = Len(strKalimat) + PANJANG Then
       Hitung = 0
    ElseIf Hitung > Len(strKalimat) Then
       TulisJalan = strKalimat & Space(Hitung - _
                    Len(strKalimat))
    Else
       TulisJalan = Mid(strKalimat, 1, Hitung)
    End If
End Function

Public Sub load_reminder()
    Dim ListItem As ListView
    Dim M_objrs As New ADODB.Recordset
    Dim my_strline As String
    Dim ifilenumber As Integer
    Static iErrCtr As Integer
    Dim cmdsql3 As String
    
    If Dir("C:\reminder.txt") = "reminder.txt" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile ("C:\reminder.txt")
    End If
    
    my_strline = ""
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    cmdsql3 = "select CUSTID, NAME, NEXTACTDATE from mgm where NEXTACTDATE BETWEEN '" + Format((Now), "yyyy-mm-dd") & " 00:00" + "' and '" + Format((Now), "yyyy-mm-dd") & " 23:59" + "' and agent ='" + MDIForm1.Text1.text + "'"
    M_objrs.Open cmdsql3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    iErrCtr = iErrCtr + 1

    ifilenumber = FreeFile
    Open "C:\reminder.txt" For Append As #ifilenumber
    
    If M_objrs.RecordCount <> 0 Then
        While Not M_objrs.EOF
            Write #ifilenumber, IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID")) & "|" & IIf(IsNull(M_objrs("NAME")), "", M_objrs("NAME")) & "|" & Format(IIf(IsNull(M_objrs("NEXTACTDATE")), "", M_objrs("NEXTACTDATE")), "hh:nn")
            M_objrs.MoveNext
        Wend
    End If
    
    Close #ifilenumber
End Sub

Public Function EncodeUTF8(S)
    Dim i
    Dim c
    
    i = 1
    Do While i <= Len(S)
        c = Asc(Mid(S, i, 1))
        If c >= &H80 Then
          S = Left(S, i - 1) + Chr(&HC2 + ((c And &H40) / &H40)) + Chr(c And &HBF) + Mid(S, i + 1)
          i = i + 1
        End If
        i = i + 1
    Loop
    EncodeUTF8 = S
End Function

Public Function cnull(ByVal Nilai As Variant) As Variant
    If IsNumeric(Nilai) Then
        Nilai = IIf(IsNull(Nilai), 0, Nilai)
    ElseIf IsDate(Format(Nilai, "yyyy-mm-dd")) Then
        Nilai = IIf(IsNull(Nilai), Null, Format(Nilai, "yyyy-mm-dd"))
    Else
        Nilai = IIf(IsNull(Nilai), "", Nilai)
    End If
    cnull = Nilai
End Function

Public Sub set_count_ol(Optional xKet As String)
    If UCase(MDIForm1.Text2.text) <> "SUPERVISOR" Then
        M_OBJCONN.execute "UPDATE tblabsen_aplikasi SET hours=hours+left((" & (Val(MDIForm1.Label_OL_count.Caption) & "/60::float/60::float)::varchar,6)::float WHERE userid='" & Trim(MDIForm1.Text1.text) & "' AND date(tanggal)=date(now()) ")
        M_OBJCONN.execute "INSERT INTO tbl_count_block(agent,ket) values('" & MDIForm1.Text1.text & "','" & xKet & "')"
    End If
End Sub

Public Sub ConvertToExcel(M_objrs As ADODB.Recordset, TxtPath As String)
    Dim ListItem        As ListItem
    Dim cmdsql_update   As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Double
    Dim m_msgbox As String
    
    i = 1
  
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath = Empty Then
        MsgBox "Nama file tidak boleh kosong, download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Double
    If M_objrs.state = 1 Then
        x = 0
        Y = M_objrs.fields().Count - 1
        Do Until x > Y
            DoEvents
            objSheet.Cells(1, i).Value = UCase(Replace(CStr(M_objrs.fields(x).Name), "_", " "))
            i = i + 1
            x = x + 1
        Loop
    End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    
    objBook.SaveAs TxtPath, xlWorkbookNormal
    objExcel.Quit
    
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
End Sub

Public Function waktu_server_sekarang() As String
    Dim m_objrs_waktu As ADODB.Recordset
    
    Set m_objrs_waktu = New ADODB.Recordset
    
    m_objrs_waktu.Open "SELECT now() as wkt_server", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    waktu_server_sekarang = Format(m_objrs_waktu!wkt_server, "yyyy-mm-dd hh:mm:ss")
    
    Set m_objrs_waktu = Nothing
End Function
'jejaktian30052016 untuk tanggal call PTP
Public Function tanggal_server_sekarang() As String
    Dim m_objrs_waktu As ADODB.Recordset
    
    Set m_objrs_waktu = New ADODB.Recordset
    
    m_objrs_waktu.Open "SELECT now() as wkt_server", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    tanggal_server_sekarang = Format(m_objrs_waktu!wkt_server, "mm-dd")
    
    Set m_objrs_waktu = Nothing
End Function

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
      
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim spath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
      
    'get the resulting string path
     If lpIDList Then
        spath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, spath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(spath, vbNullChar)
        If iNull Then spath = Left$(spath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolder = spath

End Function

Public Function FolderExists(sFullPath As String) As Boolean
    Dim myFSO As Object
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = myFSO.FolderExists(sFullPath)
End Function
'jejaktian08092016
'Public Sub insertlogcti(STATUS As String, nomor_telepon As String)
'    Dim query As String
'    WaitSecs 0.5
'    query = "INSERT INTO log_status_cti (agent,status,tanggal,nomor_telepon) values ('" + MDIForm1.Text1.Text + "','" + STATUS + "',now(), '" & nomor_telepon & "')"
'    M_OBJCONN.Execute (query)
'End Sub

Public Function create_duration(x As Long) As String
    Dim dur         As Long
    Dim dur_hour    As Long
    Dim dur_minute  As Long
    Dim dur_second  As Long
    Dim jm          As String
    
    If x > 0 Then
        dur = x
        dur_hour = dur \ 3600
        dur_minute = (dur - (dur_hour * 3600)) \ 60
        dur_second = dur - (dur_hour * 3600) - (dur_minute * 60)
        
        jm = IIf(dur_hour < 1, "00:", Right("00" & dur_hour, 2) & ":") & IIf(dur_minute < 1, "00:", Right("00" & dur_minute, 2) & ":") & IIf(dur_second < 1, "00:", Right("00" & dur_second, 2))
    Else
        jm = "00:00:00"
    End If
    create_duration = jm
End Function

Public Sub Warna_Row_Listview(frm As Form, LST1 As ListView, ByVal BackColorOne As OLE_COLOR, ByVal BackColorTwo As OLE_COLOR)
    Dim XNIL      As Long
    Dim XBYTE     As Byte
    Dim picTMP  As PictureBox
    With LST1
        If .VIEW = lvwReport And .ListItems.Count Then
            Set picTMP = frm.Controls.ADD("VB.PictureBox", "picTMP")
            XBYTE = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            XNIL = .ListItems(1).Height
            With picTMP
                .BackColor = BackColorOne
                .AutoRedraw = True
                .Height = XNIL * 2
                .BorderStyle = 0
                .Width = 10 * Screen.TwipsPerPixelX
                picTMP.Line (0, XNIL)-(.ScaleWidth, XNIL * 2), BackColorTwo, BF
                Set LST1.Picture = .Image
            End With
            Set picTMP = Nothing
            frm.Controls.Remove "picTMP"
            LST1.Parent.ScaleMode = XBYTE
        End If
    End With
End Sub


Public Sub MakeTopMost(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub


Public Function FungsiWaktuServer(Optional IntervalDay As Integer) As String
awal:
    Dim sQuerySelect As String
    Dim mObjrs_new As ADODB.Recordset
    
    Set mObjrs_new = New ADODB.Recordset
    '------------------------------------------------------------------------------
    'Fungsi Untuk mengambil waktu dan tanggal di server database
    '------------------------------------------------------------------------------
    If IntervalDay = 0 Then
        sQuerySelect = "select now() as waktu"
    Else
        sQuerySelect = "SELECT now() + interval '" & IntervalDay & " day'"
    End If
    
'    lastQuery = sQuerySelect
'    Call errLog(lastQuery, "LOG FUNGSI SERVER")
    mObjrs_new.CursorLocation = adUseClient
    
    mObjrs_new.Open sQuerySelect, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
     
    If mObjrs_new(0) = "" Then
        GoTo awal
    Else
        FungsiWaktuServer = Format(mObjrs_new(0), "yyyy-mm-dd hh:mm:ss")
    End If
    
    Set mObjrs_new = Nothing
    '------------------------------------------------------------------------------
End Function

