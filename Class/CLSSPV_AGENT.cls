VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSSPV_AGENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean

Public Function QUERY_KOMISI(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim M_Objrs As ADODB.Recordset

cmdsql = "SELECT * FROM KOMISI"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
    
cmdsql = cmdsql + " ORDER BY KODE"
    
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_KOMISI = M_Objrs
Set M_Objrs = Nothing
End Function


Public Function UPDATE_KOMISI(M_OBJCONN As ADODB.Connection, KODE As String, BTSAWAL As String, BTSAKHR As String, KOMISI As String)
Dim cmdsql As String
On Error GoTo add_error
cmdsql = " UPDATE KOMISI"
cmdsql = cmdsql + " SET"
cmdsql = cmdsql + " BTSAWAL = " + CStr(CCur(BTSAWAL)) + ","
cmdsql = cmdsql + " BTSAKHR = " + CStr(CCur(BTSAKHR)) + ","
cmdsql = cmdsql + " KOMISI = " + CStr(CCur(KOMISI)) + ""
cmdsql = cmdsql + " WHERE"
cmdsql = cmdsql + " KODE = '" + UBAH_QUOTE(KODE) + "'"
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    If err.number = 13 Then
        MsgBox "Angka yang dimasukan tidak valid"
    End If
    ADD_OK = False
End Function


Public Function query(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim M_Objrs As ADODB.Recordset

cmdsql = "SELECT * FROM SPVTBL"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
    
cmdsql = cmdsql + " ORDER BY SPVCODE"
    
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set query = M_Objrs
Set M_Objrs = Nothing

End Function

Public Function DELETE(M_OBJCONN As ADODB.Connection, SPVCODE As String)
Dim cmdsql As String

cmdsql = "DELETE from SPVTBL"
cmdsql = cmdsql + " WHERE SPVCODE = '" + UBAH_QUOTE(SPVCODE) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
  
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function ADD(M_OBJCONN As ADODB.Connection, SPVCODE As String, SPVNAME As String, TEAM As String, UNIT As String, TARGET As String, JABATAN As String)
Dim cmdsql As String

cmdsql = " INSERT INTO SPVTBL"
cmdsql = cmdsql + " (SPVCODE,"
cmdsql = cmdsql + " TEAM,"
cmdsql = cmdsql + " UNIT,"
'CMDSQL = CMDSQL + " TARGET,"
cmdsql = cmdsql + " SPVNAME)"
cmdsql = cmdsql + " VALUES"
cmdsql = cmdsql + " ('" + UBAH_QUOTE(SPVCODE) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(TEAM) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(UNIT) + "',"
'CMDSQL = CMDSQL + " " + CStr(CCur(TARGET)) + ","
cmdsql = cmdsql + " '" + UBAH_QUOTE(SPVNAME) + "')"
    On Error GoTo add_error
     M_OBJCONN.Execute cmdsql
    'Set M_OBJCONN = Nothing
    Call ADD_SPV_DITBLAGENT(M_OBJCONN, SPVCODE, SPVNAME, UNIT, TEAM, JABATAN)
    'Call ADD_CLIENTTBL(M_OBJCONN, SPVCODE)
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Private Function ADD_SPV_DITBLAGENT(M_OBJCONN As ADODB.Connection, Userid As String, agent As String, UNIT As String, TEAM As String, JABATAN As String)
Dim cmdsql As String
    cmdsql = " INSERT INTO usertbl"
    cmdsql = cmdsql + " (USERID,"
    cmdsql = cmdsql + " AGENT,"
    cmdsql = cmdsql + " UNIT,"
    cmdsql = cmdsql + " SPVCODE,"
    cmdsql = cmdsql + " PWD,"
    cmdsql = cmdsql + " ACCREC,"
    cmdsql = cmdsql + " TEAM,"
    cmdsql = cmdsql + " USERTYPE)"
    cmdsql = cmdsql + " VALUES"
    cmdsql = cmdsql + " ('" + UBAH_QUOTE(Userid) + "',"
    cmdsql = cmdsql + " '" + UBAH_QUOTE(agent) + "',"
    cmdsql = cmdsql + " '" + UBAH_QUOTE(UNIT) + "',"
    cmdsql = cmdsql + " '" + UBAH_QUOTE(Userid) + "',"
    cmdsql = cmdsql + " '" + Encrypt(Len(Userid), "PASS12345") + "',"
    cmdsql = cmdsql + " '" + Encrypt(Len(Userid), "PASS12345") + "',"
    cmdsql = cmdsql + " '" + UBAH_QUOTE(TEAM) + "',"
    If JABATAN = "AM" Then
        cmdsql = cmdsql + " 25)"
    Else
        cmdsql = cmdsql + " 20)"
    End If
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function update(M_OBJCONN As ADODB.Connection, SPVCODE As String, SPVNAME As String, TEAM As String, UNIT As String, TARGET As String)
Dim cmdsql As String
cmdsql = " UPDATE SPVTBL"
cmdsql = cmdsql + " SET"
cmdsql = cmdsql + " SPVNAME = '" + UBAH_QUOTE(SPVNAME) + "',"
cmdsql = cmdsql + " UNIT = '" + UBAH_QUOTE(UNIT) + "',"
'CMDSQL = CMDSQL + " TARGET = " + CStr(CCur(TARGET)) + ","
cmdsql = cmdsql + " TEAM = '" + UBAH_QUOTE(TEAM) + "'"
cmdsql = cmdsql + " WHERE"
cmdsql = cmdsql + " SPVCODE ='" + UBAH_QUOTE(SPVCODE) + "'"
    On Error GoTo add_error
  
    M_OBJCONN.Execute cmdsql
  
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function QUERY_AGENT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim M_Objrs As ADODB.Recordset


'    cmdsql = "SELECT usertbl.*, SPVTBL.SPVNAME"
'    cmdsql = cmdsql + " FROM usertbl INNER JOIN"
'    cmdsql = cmdsql + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"

    cmdsql = "select *,teamleader from "
    cmdsql = cmdsql + " (SELECT usertbl.* FROM usertbl ORDER BY USERID) chris "
    cmdsql = cmdsql + " Left Join "
    cmdsql = cmdsql + " (select agent as teamleader,spvcode from usertbl where userid like '%TL%') tian "
    cmdsql = cmdsql + " on chris.spvcode = tian.spvcode "

    If Len(M_WHERE) <> 0 Then
        cmdsql = cmdsql + " WHERE " + M_WHERE
    End If
    
    cmdsql = cmdsql + " ORDER BY USERID"
    
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT = M_Objrs
Set M_Objrs = Nothing

End Function

Public Function DELETE_AGENT(M_OBJCONN As ADODB.Connection, Userid As String)
Dim cmdsql As String
cmdsql = "DELETE from usertbl"
cmdsql = cmdsql + " WHERE USERID = '" + UBAH_QUOTE(Userid) + "'"
    On Error GoTo add_error
      M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_ordering(M_OBJCONN As ADODB.Connection, Userid As String)
Dim cmdsql As String
cmdsql = "DELETE from tbloffering"
cmdsql = cmdsql + " WHERE id_offering  = " + CStr(Userid) + ""
    On Error GoTo add_error
      M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function ADD_AGENT(M_OBJCONN As ADODB.Connection, Userid As String, agent As String, SPVCODE As String, SALARYBSC As String, TEAM As String, AKTIF As String, LVL As String, UNIT As String, AM As String, sId As Integer, usertype As Integer)
Dim cmdsql As String
    If Len(SALARYBSC) = 0 Then
        SALARYBSC = "0"
    End If
    cmdsql = " INSERT INTO usertbl"
    cmdsql = cmdsql + " (id,USERID,"
    cmdsql = cmdsql + " AGENT,"
    cmdsql = cmdsql + " SALARYBSC,"
    cmdsql = cmdsql + " TEAM,"
    cmdsql = cmdsql + " AKTIF,"
    cmdsql = cmdsql + " LVL,"
    cmdsql = cmdsql + " USERTYPE,"
    cmdsql = cmdsql + " ACCREC,"
    cmdsql = cmdsql + " UNIT,"
    cmdsql = cmdsql + " AM,"
    cmdsql = cmdsql + " SPVCODE)"
    cmdsql = cmdsql + " VALUES"
    cmdsql = cmdsql + " (" & sId & ",'" + UBAH_QUOTE(Userid) + "',"
    cmdsql = cmdsql + " '" + UBAH_QUOTE(agent) + "',"
    cmdsql = cmdsql + " " + CStr(CCur(SALARYBSC)) + ","
    cmdsql = cmdsql + " '" + TEAM + "',"
    cmdsql = cmdsql + " " + CStr(CCur(AKTIF)) + ","
    cmdsql = cmdsql + " '" + LVL + "',"
    cmdsql = cmdsql + " " & usertype & ","
    cmdsql = cmdsql + " '" & Encrypt(Len(Userid), "PASS12345") & "',"
    cmdsql = cmdsql + " '" + UNIT + "',"
    cmdsql = cmdsql + " '" + AM + "',"
    cmdsql = cmdsql + " '" + UBAH_QUOTE(SPVCODE) + "')"
    On Error GoTo add_error
    
    M_OBJCONN.Execute cmdsql
    
    If usertype = "6" Then
        Dim a As String
        a = "REVIEW" + Right(UBAH_QUOTE(Userid), 2)
    
        cmdsql = " INSERT INTO usertbl"
        cmdsql = cmdsql + " (id,USERID,"
        cmdsql = cmdsql + " AGENT,"
        cmdsql = cmdsql + " SALARYBSC,"
        cmdsql = cmdsql + " TEAM,"
        cmdsql = cmdsql + " AKTIF,"
        cmdsql = cmdsql + " LVL,"
        cmdsql = cmdsql + " USERTYPE,"
        cmdsql = cmdsql + " ACCREC,"
        cmdsql = cmdsql + " UNIT,"
        cmdsql = cmdsql + " AM,"
        cmdsql = cmdsql + " SPVCODE)"
        cmdsql = cmdsql + " VALUES"
        cmdsql = cmdsql + " (" & sId & ",'" + a + "',"
        cmdsql = cmdsql + " '" + a + "',"
        cmdsql = cmdsql + " " + CStr(CCur(SALARYBSC)) + ","
        cmdsql = cmdsql + " '" + TEAM + "',"
        cmdsql = cmdsql + " " + CStr(CCur(AKTIF)) + ","
        cmdsql = cmdsql + " 'Trainee',"
        cmdsql = cmdsql + " 1,"
        cmdsql = cmdsql + " '" & Encrypt(Len(Userid), "PASS12345") & "',"
        cmdsql = cmdsql + " '" + UNIT + "',"
        cmdsql = cmdsql + " '" + AM + "',"
        cmdsql = cmdsql + " 'RESERVED')"
        M_OBJCONN.Execute cmdsql
    End If
    
    Call ADD_CLIENTTBL(M_OBJCONN, Userid)
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function ADD_offering(M_OBJCONN As ADODB.Connection, keterangan As String, fld As String, discount As String, Remarks As String, operator As String, keyrms As String, existpersentase As String)
Dim cmdsql As String
  
    cmdsql = " INSERT INTO tbloffering"
    cmdsql = cmdsql + " (keterangan,"
    cmdsql = cmdsql + " fldrms ,"
    cmdsql = cmdsql + " persentase,"
    cmdsql = cmdsql + " remarks,operand,idkey,exispersentase)"
    cmdsql = cmdsql + " VALUES"
    cmdsql = cmdsql + " ('" + keterangan + "',"
    cmdsql = cmdsql + " '" + fld + "',"
    cmdsql = cmdsql + " " + CStr(discount) + ","
    cmdsql = cmdsql + " '" + frminputoreder.TxtRemarks.Text + "','" + operator + "','" + keyrms + "','" + existpersentase + "')"
    On Error GoTo add_error
    
    M_OBJCONN.Execute cmdsql
    
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function UPDATE_AGENT(M_OBJCONN As ADODB.Connection, Userid As String, agent As String, SPVCODE As String, SALARYBSC As String, TEAM As String, AKTIF As String, LVL As String, UNIT As String, AM As String)
Dim cmdsql As String

    If Len(SALARYBSC) = 0 Then
    SALARYBSC = 0
    End If

    cmdsql = " UPDATE usertbl"
    cmdsql = cmdsql + " SET"
    cmdsql = cmdsql + " AGENT ='" + UBAH_QUOTE(agent) + "' ,"
    cmdsql = cmdsql + " SALARYBSC = " + CStr(CCur(SALARYBSC)) + ","
    cmdsql = cmdsql + " TEAM = '" + TEAM + "',"
    cmdsql = cmdsql + " LVL = '" + LVL + "',"
    cmdsql = cmdsql + " UNIT = '" + UNIT + "',"
    cmdsql = cmdsql + " AM = '" + AM + "',"
    cmdsql = cmdsql + " AKTIF = " + CStr(CCur(AKTIF)) + ","
    cmdsql = cmdsql + " SPVCODE = '" + UBAH_QUOTE(SPVCODE) + "'"
    cmdsql = cmdsql + " WHERE"
    cmdsql = cmdsql + " USERID = '" + UBAH_QUOTE(Userid) + "'"
    
    On Error GoTo add_error
    
    M_OBJCONN.Execute cmdsql
    
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function
Public Function UPDATE_ordering(M_OBJCONN As ADODB.Connection, id_offer As String, keterangan As String, fld As String, discount As String, Remarks As String, operator As String, keyrms As String, exispersentase As String)
Dim cmdsql As String
    cmdsql = " UPDATE tbloffering "
    cmdsql = cmdsql + " SET keterangan ='" + keterangan + "' ,"
    cmdsql = cmdsql + " idkey ='" + keyrms + "' ,"
    cmdsql = cmdsql + " fldrms ='" + fld + "' ,"
     cmdsql = cmdsql + "  exispersentase ='" + exispersentase + "' ,"
    cmdsql = cmdsql + " persentase = " + discount + ","
     cmdsql = cmdsql + " operand  = '" + operator + "',"
    cmdsql = cmdsql + " remarks = '" + frmheaderoffeer.TxtRemarks.Text + "'"
    cmdsql = cmdsql + " WHERE "
    cmdsql = cmdsql + "  id_offering = " + CStr(id_offer) + ""
    
    On Error GoTo add_error
    
    M_OBJCONN.Execute cmdsql
    
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function COMBO_SPV(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim M_Objrs As ADODB.Recordset
Dim cmdsql As String
    
        cmdsql = " SELECT * FROM usertbl where userid like 'TL%' and spvcode like 'SPV%'"
    If Len(M_WHERE) <> 0 Then
        cmdsql = cmdsql + " AND " + M_WHERE
    End If
    cmdsql = cmdsql + " ORDER BY SPVCODE"
   
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set COMBO_SPV = M_Objrs
Set M_Objrs = Nothing

End Function

Public Function QUERY_PRODUCT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim M_Objrs As ADODB.Recordset

cmdsql = "SELECT * FROM PRODUCTTBL"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
    
cmdsql = cmdsql + " ORDER BY CODE"
    
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_PRODUCT = M_Objrs
Set M_Objrs = Nothing
End Function

Public Function ADD_PRODUCT(M_OBJCONN As ADODB.Connection, CODE As String, STATUS1 As String, PRODUCT As String, UNIT As String)
Dim cmdsql As String
cmdsql = " INSERT INTO PRODUCTTBL"
cmdsql = cmdsql + " (CODE,"
cmdsql = cmdsql + " STATUS,"
cmdsql = cmdsql + " UNIT,"
cmdsql = cmdsql + " PRODUCT)"
cmdsql = cmdsql + " VALUES"
cmdsql = cmdsql + " ('" + UBAH_QUOTE(CODE) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(STATUS1) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(UNIT) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(PRODUCT) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_PRODUCT(M_OBJCONN As ADODB.Connection, CODE As String)
Dim cmdsql As String
cmdsql = "DELETE From PRODUCTTBL"
cmdsql = cmdsql + " WHERE CODE = '" + CODE + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function UPDATE_PRODUCT(M_OBJCONN As ADODB.Connection, CODE As String, PRODUCT As String, STATUS1 As String, UNIT As String)
Dim cmdsql As String
cmdsql = " UPDATE PRODUCTTBL"
cmdsql = cmdsql + " SET"
cmdsql = cmdsql + " PRODUCT = '" + UBAH_QUOTE(PRODUCT) + "',"
cmdsql = cmdsql + " UNIT = '" + UBAH_QUOTE(UNIT) + "',"
cmdsql = cmdsql + " STATUS = '" + UBAH_QUOTE(STATUS1) + "'"
cmdsql = cmdsql + " WHERE"
cmdsql = cmdsql + " CODE ='" + UBAH_QUOTE(CODE) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function


Public Function UPDATE_Password(M_OBJCONN As ADODB.Connection, Userid As String, ACCREC As String, AUTH As String)
Dim cmdsql As String
cmdsql = " UPDATE usertbl"
cmdsql = cmdsql + " SET"
cmdsql = cmdsql + " ACCREC = '" + UBAH_QUOTE(ACCREC) + "',"
cmdsql = cmdsql + " AUTH = '" + UBAH_QUOTE(AUTH) + "'"
cmdsql = cmdsql + " WHERE"
cmdsql = cmdsql + " USERID ='" + UBAH_QUOTE(Userid) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function QUERY_SET_PWDAGENT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim M_Objrs As ADODB.Recordset

    cmdsql = "SELECT * FROM usertbl"
    cmdsql = cmdsql + " WHERE USERID <> 'ADMINISTRATOR'"
    If Len(M_WHERE) <> 0 Then
        cmdsql = cmdsql + " AND " + M_WHERE
    End If
    
    cmdsql = cmdsql + " ORDER BY USERID"
    
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SET_PWDAGENT = M_Objrs
Set M_Objrs = Nothing
End Function

Public Function QUERY_CLOSSING(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim M_Objrs As ADODB.Recordset

cmdsql = "SELECT * FROM CLSREASONTBL"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
   
cmdsql = cmdsql + " ORDER BY KDCLS"
    
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_CLOSSING = M_Objrs
Set M_Objrs = Nothing
End Function

Public Function ADD_CLOSSING(M_OBJCONN As ADODB.Connection, KDCLS As String, KETCLS As String, jenis As String)
Dim cmdsql As String

cmdsql = " INSERT INTO CLSREASONTBL"
cmdsql = cmdsql + " (KDCLS,"
cmdsql = cmdsql + " jenis,"
cmdsql = cmdsql + " KETCLS)"
cmdsql = cmdsql + " VALUES"
cmdsql = cmdsql + " ('" + UBAH_QUOTE(KDCLS) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(jenis) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(KETCLS) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function UPDATE_CLOSSING(M_OBJCONN As ADODB.Connection, KDCLS As String, KETCLS As String, jenis As String)
Dim cmdsql As String
cmdsql = " UPDATE CLSREASONTBL"
cmdsql = cmdsql + " SET"
cmdsql = cmdsql + " KETCLS = '" + UBAH_QUOTE(KETCLS) + "',"
cmdsql = cmdsql + " jenis = '" + UBAH_QUOTE(jenis) + "'"
cmdsql = cmdsql + " WHERE"
cmdsql = cmdsql + " KDCLS ='" + UBAH_QUOTE(KDCLS) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_CLOSSING(M_OBJCONN As ADODB.Connection, KDCLS As String)
Dim cmdsql As String
cmdsql = "DELETE from CLSREASONTBL"
cmdsql = cmdsql + " WHERE KDCLS = '" + UBAH_QUOTE(KDCLS) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function ADD_DATASOURCE(M_OBJCONN As ADODB.Connection, KODEDS As String, STATUS As String, keterangan As String, KDPROGRAM As String)
Dim cmdsql As String

cmdsql = " INSERT INTO DATASOURCETBL"
cmdsql = cmdsql + " (KODEDS,"
cmdsql = cmdsql + " STATUS,"
cmdsql = cmdsql + " KDPROGRAM,"
cmdsql = cmdsql + " KETERANGAN)"
cmdsql = cmdsql + " VALUES"
cmdsql = cmdsql + " ('" + UBAH_QUOTE(KODEDS) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(STATUS) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(KDPROGRAM) + "',"
cmdsql = cmdsql + " '" + UBAH_QUOTE(keterangan) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_DATASOURCE(M_OBJCONN As ADODB.Connection, KODEDS As String)
Dim cmdsql As String
cmdsql = "DELETE from DATASOURCETBL"
cmdsql = cmdsql + " WHERE KODEDS = '" + UBAH_QUOTE(KODEDS) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function


Public Function QUERY_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim M_Objrs As ADODB.Recordset

cmdsql = "SELECT * FROM DATASOURCETBL"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
    
cmdsql = cmdsql + " ORDER BY KODEDS"
    
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_DATASOURCE = M_Objrs
Set M_Objrs = Nothing
End Function

Public Function UPDATE_DATASOURCE(M_OBJCONN As ADODB.Connection, KODEDS As String, STATUS As String, keterangan As String, KDPROGRAM As String)
Dim cmdsql As String
cmdsql = " UPDATE DATASOURCETBL"
cmdsql = cmdsql + " SET"
cmdsql = cmdsql + " KETERANGAN = '" + UBAH_QUOTE(keterangan) + "', "
cmdsql = cmdsql + " KDPROGRAM = '" + UBAH_QUOTE(KDPROGRAM) + "', "
cmdsql = cmdsql + " STATUS =  '" + UBAH_QUOTE(STATUS) + "' "
cmdsql = cmdsql + " WHERE"
cmdsql = cmdsql + " KODEDS = '" + KODEDS + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    'MsgBox Err.Description
End Function

Public Function ADD_CLIENTTBL(M_OBJCONN As ADODB.Connection, agent As String)
Dim cmdsql As String
cmdsql = " INSERT INTO CLIENTTBL"
cmdsql = cmdsql + " (AGENT)"
cmdsql = cmdsql + " VALUES"
cmdsql = cmdsql + " ('" + UBAH_QUOTE(agent) + "')"
    On Error GoTo add_error
     M_OBJCONN.Execute cmdsql
        ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function QUERY_AOC(M_OBJCONN As ADODB.Connection) As Object
Dim cmdsql As String
Dim M_Objrs As ADODB.Recordset


    cmdsql = "SELECT * FROM usertbl WHERE USERID != AGENT ORDER BY USERID"
    
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AOC = M_Objrs
Set M_Objrs = Nothing

End Function

Public Function QUERY_AOC_CHECK(M_OBJCONN As ADODB.Connection) As Boolean
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim cuserid As String
    Dim cagent As String
    Dim caoc As String
    
    cuserid = FormAOC.txtuserid.Text
    cagent = FormAOC.lblnama.Caption
    caoc = FormAOC.txtaoc.Text
    cmdsql = "SELECT USERID FROM tblaoc WHERE USERID != AGENT AND USERID = '" + cagent + "'"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        QUERY_AOC_CHECK = True
    Else
        QUERY_AOC_CHECK = False
    End If
    Set M_Objrs = Nothing
End Function

Public Function SHOW_AOC(M_OBJCONN As ADODB.Connection)
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    cmdsql = "SELECT * FROM tblaoc"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Set SHOW_AOC = M_Objrs
    Set M_Objrs = Nothing
End Function

Public Function UPDATE_AOC(M_OBJCONN As ADODB.Connection)
    Dim cmdsql As String
    Dim cuserid As String
    Dim cagent As String
    Dim caoc As String

    cuserid = FormAOC.txtuserid.Text
    cagent = FormAOC.lblnama.Caption
    caoc = FormAOC.txtaoc.Text
        cmdsql = "UPDATE TBLAOC SET agent = '" + cuserid + "' , aoc = '" + caoc + "' where userid = '" + cagent + "'"
    
    M_OBJCONN.Execute cmdsql
        ADD_OK = True
        Exit Function
add_error:
        ADD_OK = False
        FormAOC.txtuserid.SetFocus
End Function

Public Function ADD_AOC(M_OBJCONN As ADODB.Connection)
    Dim cmdsql As String
    Dim cuserid As String
    Dim cagent As String
    Dim caoc As String
    
    cuserid = FormAOC.txtuserid.Text
    cagent = FormAOC.lblnama.Caption
    caoc = FormAOC.txtaoc.Text
        cmdsql = "INSERT INTO tblaoc values ('" + cagent + "' , '" + cuserid + "' , '" + caoc + "')"
    M_OBJCONN.Execute cmdsql
        ADD_OK = True
        Exit Function
add_error:
        ADD_OK = False
        FormAOC.txtuserid.SetFocus
End Function

Public Function DELETE_AOC(M_OBJCONN As ADODB.Connection)
Dim cmdsql As String
Dim cuserid As String
cuserid = FormAOC.lblnama.Caption
cmdsql = "DELETE from tblaoc where userid = '" + cuserid + "'"
      M_OBJCONN.Execute cmdsql
End Function
