Attribute VB_Name = "autodialer"
Public AutoDialerON As Boolean
Public AutoDialerOFF As Boolean
Public AutoDialerCall As Boolean
Public AutoDialerHangup As Boolean
Public AutoDialerStatusFeedBack As String
Public AutodialerPhoneNumber As String
Public AutodialerCustomerID As String
Public AutoDialerBreak As Boolean
Public DoubleClick_ListViewMGM As Boolean
Public FirstLogin  As Boolean
Public Sub Autdialer_CekON(agent As String)
Dim rs As ADODB.Recordset
Dim strsring As String
Dim statusAutodialerAgent As String
strstring = "select * from usertbl where agent = '" + agent + "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strstring, M_OBJCONN, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    statusAutodialerAgent = IIf(IsNull(rs!autodialer_status), "", rs!autodialer_status)
    If statusAutodialerAgent = "ON" Then
        AutoDialerON = True
     Else
        AutoDialerON = False
    End If
    
End If


End Sub

Public Function Autodialer_Calling(agent As String)
Dim rs As ADODB.Recordset
Dim strsring As String
Dim statusAutodialerAgent As String
Dim StrPhone As String
Dim StrCustID As String

strstring = "select * from tbl_autodialer_runningcall where agent = '" + agent + "' order by id limit 1"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strstring, M_OBJCONN, adOpenDynamic, adLockOptimistic
AutodialerCustomerID = ""
AutodialerPhoneNumber = ""
If rs.RecordCount > 0 Then
     AutodialerPhoneNumber = IIf(IsNull(rs!phone), "", rs!phone)
     AutodialerCustomerID = IIf(IsNull(rs!customerid), "", Trim(rs!customerid))
     'MDIForm1.ActionCTI ("DIAL|" & GetNumber(CStr(Replace(AutodialerPhoneNumber, " ", ""))) & "|" & Trim(AutodialerCustomerID))
    WaitSecs (0.5)
    AutoDialerCall = True
    AutoDialerHangup = False
    AutoDialerBreak = False
    MDIForm1.TimerAutoDialer.Enabled = False
    VIEW_MGMDATA.Command1_Click (0)
    FrmCC_Colection.Show
    
    Else
    MsgBox "Data di bucket Autodialer sudah habis"
End If


End Function

Public Sub Autodialer_update_phone_log(no_telp As String, customerid As String, result_status As String, agent As String, StatusCTI As String)
Dim rsTelp As ADODB.Recordset
Dim cmdsqlphone As String

cmdsqlphone = "select * from tbl_autodialer_runningcall_log where phone='" + no_telp + "' "
Set rsTelp = New ADODB.Recordset
rsTelp.CursorLocation = adUseClient
rsTelp.Open cmdsqlphone, M_OBJCONN, adOpenDynamic, adLockOptimistic

If rsTelp.RecordCount = 0 Then
cmdsqlphone = "insert into tbl_autodialer_runningcall_log(insert_date,customerid,phone,retrycall,last_call_date,result_status,agent)"
cmdsqlphone = cmdsqlphone + " values(now()::timestamp(0),'" + customerid + "','" + no_telp + "','1',"
cmdsqlphone = cmdsqlphone + " now()::timestamp(0),'" + result_status + "','" + agent + "')"
Else
cmdsqlphone = "update tbl_autodialer_runningcall_log set last_call_date=now()::timestamp(0), result_status='" + result_status + "', retrycall = retrycall::int + 1  where phone='" + no_telp + "'"
End If
M_OBJCONN.execute cmdsqlphone
M_OBJCONN.execute "delete from tbl_autodialer_runningcall where phone='" + no_telp + "'"
End Sub

'Public Sub Autodialer_Stop(agent As String, reason_stop As String, durasi As Double)
Public Sub Autodialer_Stop(agent As String, reason_stop As String, start_tm As String, session_id As String, local_ip As String)
Dim rsTelp As ADODB.Recordset
Dim cmdsqlphone As String

'cmdsqlphone = " insert into  tbl_autodialer_agent_break(agent,status_break,durasi) values"
'cmdsqlphone = cmdsqlphone + "('" + agent + "','" + reason_stop + "','" + CStr(durasi) + "')"
''M_OBJCONN.execute cmdsqlphone

If break_time = True Then
    cmdsqlphone = " insert into tbl_autodialer_agent_break(agent,status_break,waktu_start,sessionid,ip_login) values"
    cmdsqlphone = cmdsqlphone + "('" & agent & "','" & reason_stop & "','" & start_tm & "', '" & session_id & "','" & local_ip & "')"
    M_OBJCONN.execute cmdsqlphone
Else
    AutoDialerON = False
End If
        
AutoDialerBreak = True
cmdsqlphone = "update usertbl set autodialer_status='OFF' where agent= '" + agent + "'"
M_OBJCONN.execute cmdsqlphone


End Sub

'Public Sub Autodialer_Start(agent As String, reason_stop As String, durasi As Double)
Public Sub Autodialer_Start(agent As String, reason_stop As String, durasi As String, end_tm As String)

Dim cmdsqlphone As String

'cmdsqlphone = " insert into  tbl_autodialer_agent_break(agent,status_break,durasi) values"
'cmdsqlphone = cmdsqlphone + "('" + agent + "','" + reason_stop + "','" + CStr(durasi) + "')"
'M_OBJCONN.execute cmdsqlphone
If break_time = True Then
    cmdsqlphone = "update tbl_autodialer_agent_break set durasi = '" & durasi & "', waktu_end = '" & end_tm & "' where id in "
    cmdsqlphone = cmdsqlphone + "(select max(id) from tbl_autodialer_agent_break where agent = '" & agent & "' and status_break not in ('ManualDial','start_autodialer','AutoDial','form break show'))"
    M_OBJCONN.execute cmdsqlphone
End If

AutoDialerBreak = False
AutoDialerON = True
cmdsqlphone = "update usertbl set autodialer_status='ON' where agent= '" + agent + "'"
M_OBJCONN.execute cmdsqlphone
MDIForm1.TimerAutoDialer.Enabled = True

End Sub

