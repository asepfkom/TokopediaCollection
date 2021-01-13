VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_FRMSEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Function QUERY_SEARCH_UPLOAD(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM TEMPCC_CUSTTBL"

 If NamaAgent <> Empty Then
    SYARAT = SYARAT + " WHERE " + NamaAgent
 End If

 If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
CMDSQL = CMDSQL + " " + SYARAT + " ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_UPLOAD = M_objrs
Set M_objrs = Nothing
End Function
Public Function QUERY_SEARCH(M_OBJCONN As ADODB.Connection, M_WHERE As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

Select Case UCase(TIPE_PRODUK)
Case "KTA"
    CMDSQL = "SELECT * FROM KTA_CUSTTBL"
Case "CREDIT CARD"
    CMDSQL = "SELECT * FROM CC_CUSTTBL"
Case "KTA - CROSS SELL"
    CMDSQL = "SELECT * FROM CS_CUSTTBL"
Case "CC - CROSS SELL"
    CMDSQL = "SELECT * FROM CCCS_CUSTTBL"
Case Else
    CMDSQL = "SELECT * FROM CC_CUSTTBL"
End Select
If Len(M_WHERE) <> 0 Then
   CMDSQL = CMDSQL + " WHERE " + M_WHERE
End If
CMDSQL = CMDSQL + " ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_SEARCH_CONDITION(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String, Lcustid1 As String, KETHSLKERJA As String, lLastCallDate1 As String, lStatusCek1 As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT * FROM CC_CUSTTBL"
    
 If Lcustid1 <> Empty Then
    SYARAT = SYARAT + " WHERE " + Lcustid1
 End If
 If NamaAgent <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " and " + NamaAgent
    Else
        SYARAT = SYARAT + " WHERE " + NamaAgent
    End If
 End If
 If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
    
     If lLastCallDate1 <> Empty Then
        If Len(SYARAT) > 0 Then
            SYARAT = SYARAT + " AND " + lLastCallDate1
        Else
            SYARAT = SYARAT + " WHERE " + lLastCallDate1
        End If
     End If
 If lStatusCek1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lStatusCek1
    Else
        SYARAT = SYARAT + " WHERE " + lStatusCek1
    End If
 End If
 
 If KETHSLKERJA <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + KETHSLKERJA
    Else
        SYARAT = SYARAT + " WHERE " + KETHSLKERJA
    End If
  End If

 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
 
 
 
CMDSQL = CMDSQL + " " + SYARAT + " ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_CONDITION = M_objrs
Set M_objrs = Nothing

End Function



Public Function QUERY_SEARCH_nonDist(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset


 CMDSQL = "SELECT * FROM tempcc_custtbl"

 If NamaAgent <> Empty Then
    SYARAT = SYARAT + " WHERE " + NamaAgent
 End If

 If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
CMDSQL = CMDSQL + " " + SYARAT + " ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_nonDist = M_objrs
Set M_objrs = Nothing
End Function

Public Sub strategicheck()
    Dim rs As New ADODB.Recordset
    Dim qsel, cek As String
    Dim a, b, c, d, e, F, g, h, i, J, K, l, M, n, o As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    qsel = "select * from strategi_detail where strategi = " & id_strategi & " and nm_strategi = '" & nm_strategi & "' ;"
    rs.Open qsel, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs.RecordCount > 0 Then
        a = rs!sts_pop
        b = rs!sts_vl
        c = rs!sts_os
        d = rs!sts_ptp
        e = rs!sts_on
        F = rs!sts_bp
        g = rs!sts_co
        h = rs!sts_po
        i = rs!sts_pr
        
        J = cnull(rs!balance_min)
        K = cnull(rs!balance_max)
        
        l = cnull(rs!lpd_min)
        M = cnull(rs!lpd_max)
        
        n = cnull(rs!wo_min)
        o = cnull(rs!wo_max)
    
        cek = " and ("
        If a = 1 Then
            cek = cek & "f_cek_new ilike '%pop%' or"
        End If
        If b = 1 Then
            cek = cek & " f_cek_new ilike '%vl%' or"
        End If
        If c = 1 Then
            cek = cek & " f_cek_new ilike '%os%' or"
        End If
        If d = 1 Then
            cek = cek & " f_cek_new ilike '%ptp%' or"
        End If
        If e = 1 Then
            cek = cek & " f_cek_new ilike '%on%' or"
        End If
        If F = 1 Then
            cek = cek & " f_cek_new ilike '%bp%' or"
        End If
        If g = 1 Then
            cek = cek & " f_cek_new ilike '%co%' or"
        End If
        If h = 1 Then
            cek = cek & " f_cek_new ilike '%po%' or"
        End If
        If i = 1 Then
            cek = cek & " f_cek_new ilike '%pr%' or"
        End If
        If Len(cek) > 6 Then
            cek = Left(cek, Len(cek) - 3)
            cek = cek & ")" & vbCrLf
        Else
            cek = Left(cek, Len(cek) - 6)
        End If
        
        
        If J <> "" And K <> "" Then
            cek = cek & " and curbal between '" & J & "' and '" & K & "' " & vbCrLf
        End If
        If l <> "" And M <> "" Then
            cek = cek & " and to_char(pay_dt,'yyyy-mm') between '" & Format(l, "yyyy-mm") & "' and '" & Format(M, "yyyy-mm") & "' " & vbCrLf
        End If
        If n <> "" And o <> "" Then
            cek = cek & " and to_char(b_d,'yyyy-mm') between '" & Format(n, "yyyy-mm") & "' and '" & Format(o, "yyyy-mm") & "' " & vbCrLf
        End If
    
        query_strategi = cek
    End If
End Sub

Public Function QUERY_SEARCH_CONDITION_mgm(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String, _
                                        AHOMENO As String, AHOMENO2 As String, AOFFICENO As String, AOFFICENO2 As String, extoffice As String, extoffice2 As String, homenoadd1 As String, ahomenoadd1 As String, homenoadd2 As String, ahomenoadd2 As String, officenoadd1 As String, aofficenoadd2 As String, officenoadd2 As String, mobilenoadd1 As String, _
                                        mobilenoadd2 As String, ec_telp As String, alamatrumah As String, alamatkantor As String, alamatec As String, _
                                        Lcustid1 As String, F_CEK As String, lLastCallDate1 As String, lStatusCek1 As String, f_Pending As String, FCEKSTATUS As String, WO_DATE As String, strverify As String, strapprovel As String, strapprovelyet As String, strReject As String, strSukes As String, strBlocked As String, strBlokentry As String, strblokedmarkup As String, nmagentprev As String, LOCKTL As String, BlokedPayment, BlokedPTPNoPayment, Optional page As Long, Optional slimitpage As Double, Optional segments As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim iOffset As Integer
Dim iLimit As String
Dim M_objrs As ADODB.Recordset
'    Cmdsql = "SELECT CUSTID,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek,F_cek_NEW,kethslkerja_new, AGENT,  Principal,DOB,stscpa ,intapprovel,intverify,resultcpa,tglstscpa ,flaglead,"
'    Cmdsql = Cmdsql + "AmountWo, OpenDate, Pay_dt,B_D,TtlPTP, RECSOURCE,TGLSOURCE,TGLSTATUS, Kethslkerja, TGLINCOMING, AmountWo,f_Pending,statusprior,StatusCall,TGLCALL,F_LOCK,tglinsertfrmcpa,curbal,ktpno,dob,lastpay,pay_dt FROM mgm"
    ' UPDATE 12 SEPT 2013 UNTUK tgl cek claim acc
    
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then

        CMDSQL = "SELECT CUSTID,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek,F_cek_NEW,kethslkerja_new,"
        If bcp = False Then
            CMDSQL = CMDSQL + " segment ,"
        End If
        CMDSQL = CMDSQL + " agent , Principal, DOB, stscpa, intapprovel, intverify, resultcpa, tglstscpa, flaglead, "
        CMDSQL = CMDSQL + "ahomeno, ahomeno2, aofficeno, aofficeno2, extoffice, extoffice2, homenoadd1, ahomenoadd1,homenoadd2 , ahomenoadd2, officenoadd1, aofficenoadd2, officenoadd2, mobilenoadd1, mobilenoadd2,ec_telp,addrnow,ecaddr,addrpt,"
        CMDSQL = CMDSQL + "AmountWo, OpenDate, Pay_dt,B_D,TtlPTP, RECSOURCE,TGLSOURCE,TGLSTATUS, Kethslkerja, TGLINCOMING, AmountWo,f_Pending,statusprior,StatusCall,TGLCALL,F_LOCK,tglinsertfrmcpa,curbal,ktpno,dob,lastpay,pay_dt,app_claim,agent_asli, tgl_exp_claim FROM mgm"
        SYARAT = " where 1 = 1 "
        '====asep25022020====='
        'SYARAT = " where 1 = 1 and (status_ptp <>'BP' or coalesce(status_ptp,'') = '')"
        '====================='
    ElseIf UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        CMDSQL = "SELECT CUSTID,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek,F_cek_NEW,kethslkerja_new,"
        If bcp = False Then
            CMDSQL = CMDSQL + " segment ,"
        End If
        CMDSQL = CMDSQL + " agent , Principal, DOB, stscpa, intapprovel, intverify, resultcpa, tglstscpa, flaglead, "
        CMDSQL = CMDSQL + "ahomeno, ahomeno2, aofficeno, aofficeno2, extoffice, extoffice2, homenoadd1, ahomenoadd1,homenoadd2 , ahomenoadd2, officenoadd1, aofficenoadd2, officenoadd2, mobilenoadd1, mobilenoadd2,ec_telp,addrnow,ecaddr,addrpt,"
        CMDSQL = CMDSQL + "AmountWo, OpenDate, Pay_dt,B_D,TtlPTP, RECSOURCE,TGLSOURCE,TGLSTATUS, Kethslkerja, TGLINCOMING, AmountWo,f_Pending,statusprior,StatusCall,TGLCALL,F_LOCK,tglinsertfrmcpa,curbal,ktpno,dob,lastpay,pay_dt,app_claim,agent_asli, tgl_exp_claim FROM mgm"
        'SYARAT = " where 1 = 1 "
        '====asep25022020====='
        SYARAT = " where 1 = 1 and (status_ptp <>'KP' or coalesce(status_ptp,'') = '') and date(tglsource) >= date(now())"
        '====================='
    End If
If strategi = True Then
    strategicheck
End If

If strategi = True Then
    SYARAT = SYARAT + query_strategi
End If

If Lcustid1 <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + Lcustid1
    Else
        SYARAT = SYARAT + " WHERE " + Lcustid1
    End If
End If

If NamaAgent <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + NamaAgent
    Else
       SYARAT = SYARAT + " WHERE " + NamaAgent
    End If
End If
   
   
If nmagentprev <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + nmagentprev
    Else
       SYARAT = SYARAT + " WHERE " + nmagentprev
    End If
End If

If LOCKTL <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + LOCKTL
    Else
       SYARAT = SYARAT + " WHERE " + LOCKTL
    End If
        
End If

If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If

If FCEKSTATUS <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + FCEKSTATUS
    Else
        SYARAT = SYARAT + " WHERE " + FCEKSTATUS
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
  
  If lLastCallDate1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lLastCallDate1
    Else
        SYARAT = SYARAT + " WHERE " + lLastCallDate1
    End If
 End If
 If lStatusCek1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lStatusCek1
    Else
        SYARAT = SYARAT + " WHERE " + lStatusCek1
    End If
 End If
 
  
  If WO_DATE <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + WO_DATE
    Else
        SYARAT = SYARAT + " WHERE " + WO_DATE
    End If
  
  End If
  
 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ""
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ""
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If AHOMENO <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + AHOMENO
    Else
        SYARAT = SYARAT + " WHERE (" + AHOMENO
    End If
 End If
 If AHOMENO2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + AHOMENO2
    Else
        SYARAT = SYARAT + " WHERE (" + AHOMENO2
    End If
 End If
 If AOFFICENO <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + AOFFICENO
    Else
        SYARAT = SYARAT + " WHERE (" + AOFFICENO
    End If
 End If
 If AOFFICENO2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + AOFFICENO2
    Else
        SYARAT = SYARAT + " WHERE (" + AOFFICENO2
    End If
 End If
 If extoffice <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + extoffice
    Else
        SYARAT = SYARAT + " WHERE (" + extoffice
    End If
 End If
 If extoffice2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + extoffice2
    Else
        SYARAT = SYARAT + " WHERE (" + extoffice2
    End If
 End If
 If homenoadd1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + homenoadd1
    Else
        SYARAT = SYARAT + " WHERE (" + homenoadd1
    End If
 End If
 If ahomenoadd1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + ahomenoadd1
    Else
        SYARAT = SYARAT + " WHERE (" + ahomenoadd1
    End If
 End If
 If homenoadd2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + homenoadd2
    Else
        SYARAT = SYARAT + " WHERE (" + homenoadd2
    End If
 End If
 If ahomenoadd2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + ahomenoadd2
    Else
        SYARAT = SYARAT + " WHERE (" + ahomenoadd2
    End If
 End If
 If officenoadd1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + officenoadd1
    Else
        SYARAT = SYARAT + " WHERE (" + officenoadd1
    End If
 End If
  If aofficenoadd2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + aofficenoadd2
    Else
        SYARAT = SYARAT + " WHERE (" + aofficenoadd2
    End If
 End If
   If aofficenoadd2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + aofficenoadd2
    Else
        SYARAT = SYARAT + " WHERE (" + aofficenoadd2
    End If
 End If
 If officenoadd2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + officenoadd2
    Else
        SYARAT = SYARAT + " WHERE (" + officenoadd2
    End If
 End If
  If mobilenoadd1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + mobilenoadd1
    Else
        SYARAT = SYARAT + " WHERE (" + mobilenoadd1
    End If
 End If
   If mobilenoadd2 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + mobilenoadd2
    Else
        SYARAT = SYARAT + " WHERE (" + mobilenoadd2
    End If
 End If
    If ec_telp <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + ec_telp
    Else
        SYARAT = SYARAT + " WHERE (" + ec_telp
    End If
 End If
         
 
 If strverify <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strverify
    Else
        SYARAT = SYARAT + " WHERE " + strverify
    End If
 End If
 
 If strapprovel <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strapprovel
    Else
        SYARAT = SYARAT + " WHERE " + strapprovel
    End If
 End If
 
 If strapprovelyet <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strapprovelyet
    Else
        SYARAT = SYARAT + " WHERE " + strapprovelyet
    End If
 End If
 
 If strReject <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strReject
    Else
        SYARAT = SYARAT + " WHERE " + strReject
    End If
 End If
 
 If strSukes <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strSukes
    Else
        SYARAT = SYARAT + " WHERE " + strSukes
    End If
 End If
 
 If f_Pending <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + f_Pending
    Else
        SYARAT = SYARAT + " WHERE " + f_Pending
    End If
  End If
  
  If segments <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND segment = '" & segments & "' "
    Else
        SYARAT = SYARAT + " WHERE segment = '" & segments & "' "
    End If
  End If
  

 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ""
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
 
If F_LOCK = True Then
    SYARAT = SYARAT + " AND F_LOCK='Y'"
 End If
 
 '@@201010 Ini buat gabungin blok data
 If strBlokentry <> Empty Or F_CEK <> Empty Or strblokedmarkup <> Empty Or BlokedPayment <> Empty Then
    SYARAT = SYARAT + " AND ( "
    Dim SyaratBloked As String

    SyaratBloked = ""

    If strBlokentry <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or " + strBlokentry
        Else
            SyaratBloked = SyaratBloked + strBlokentry
        End If
    End If

    If F_CEK <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or " + F_CEK
        Else
            SyaratBloked = SyaratBloked + F_CEK
        End If
    End If
    
    '@@15 Agustus 2011 Blok Payment Request Gaby
    If BlokedPayment <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or (" + BlokedPayment + ") "
        Else
            SyaratBloked = SyaratBloked + BlokedPayment
        End If
    End If
    
    'BlokedPTPNoPayment
    If BlokedPTPNoPayment <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or (" + BlokedPTPNoPayment + ") "
        Else
            SyaratBloked = SyaratBloked + BlokedPTPNoPayment
        End If
    End If
    
    If strblokedmarkup <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or exclude in (" + Replace(strblokedmarkup, "@", "'") + ")"
        Else
            SyaratBloked = SyaratBloked + " exclude in (" + Replace(strblokedmarkup, "@", "'") + ")"
        End If
    End If
    
    SYARAT = SYARAT + SyaratBloked + ")"
 End If
  
 '@@140710 Tambahan Blok Entry
' If strBlokentry <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + strBlokentry
'    Else
'        SYARAT = SYARAT + " WHERE " + strBlokentry
'    End If
' End If
 
' If F_CEK <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + "(" + F_CEK
'    Else
'        SYARAT = SYARAT + " WHERE " + "(" + F_CEK
'    End If
'
'  End If
'
'
'
' If strblokedmarkup <> Empty Then
'    If F_CEK <> Empty Then ' cek kalo f_cek tidak kosong,  pakai tutup kurung  "'") + ")) "
'        If Len(SYARAT) > 0 Then
'            SYARAT = SYARAT + " or  exclude in (" + Replace(strblokedmarkup, "@", "'") + ")) "
'            Else
'            SYARAT = SYARAT + " WHERE exclude in (" + Replace(strblokedmarkup, "@", "'") + ") "
'        End If
'    Else
'            If Len(SYARAT) > 0 Then
'                     SYARAT = SYARAT + " or  exclude in (" + Replace(strblokedmarkup, "@", "'") + ") "
'                 Else
'                     SYARAT = SYARAT + " WHERE exclude in (" + Replace(strblokedmarkup, "@", "'") + ") "
'             End If
'    End If
'   Else
'      If F_CEK <> Empty Then
'      SYARAT = SYARAT + ")"
'      End If
' End If

 'jejaktian searchbyalamat
 If alamatrumah <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + alamatrumah
    Else
        SYARAT = SYARAT + " WHERE " + alamatrumah
    End If
 End If
 If alamatkantor <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + alamatkantor
    Else
        SYARAT = SYARAT + " WHERE (" + alamatkantor
    End If
 End If
 If alamatec <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + alamatec
    Else
        SYARAT = SYARAT + " WHERE (" + alamatec
    End If
 End If
 
'===================================
' '===asep25022020=='
'  If Status_PTP = "KP" Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + alamatec
'    Else
'        SYARAT = SYARAT + " WHERE (" + alamatec
'    End If
' End If
 '==================='
    CMDSQL = CMDSQL + " " + SYARAT + ""
  
  
    CMDSQL = CMDSQL + " ORDER BY CUSTID "
    slimitpage = 100
    If page > 0 Then
    iOffset = (Val(page) - 1) * slimitpage
    iLimit = " Limit " + CStr(slimitpage) + " offset " + CStr(iOffset)
Else
    If Val(slimitpage) > 0 Then

        iOffset = 0
       iLimit = " Limit " + CStr(slimitpage) + " offset " + CStr(iOffset)
    End If
End If
'    If iLimit > 0 Then cmdsql = cmdsql & " LIMIT " & iLimit
'
'    If iOffset > 0 Then cmdsql = cmdsql & " OFFSET " & iOffset
    CMDSQL = CMDSQL + iLimit
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient

M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
Set QUERY_SEARCH_CONDITION_mgm = M_objrs

Set M_objrs = Nothing
End Function
Public Function QUERY_SEARCH_mgm(M_OBJCONN As ADODB.Connection, M_WHERE As String, TIPE_PRODUK As String, F_CEK As String, f_Pending As String, strBlok As String, strBlokentry As String, strblokedmarkup As String, WO_DATE As String, STSLOCKTL As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset

  
    CMDSQL = "SELECT * from mgm"
    'CMDSQL = "SELECT CUSTID,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek, AGENT,  Principal, "
    'CMDSQL = CMDSQL + "AmountWo, OpenDate, TtlPTP, RECSOURCE, TGLSTATUS, Kethslkerja, TGLINCOMING,  AmountWo FROM mgm"
If Len(M_WHERE) <> 0 Then
   CMDSQL = CMDSQL + " WHERE " + M_WHERE
End If

 If F_CEK <> Empty Then
    
        CMDSQL = CMDSQL + " AND " + F_CEK
'    Else
'        SYARAT = SYARAT + " WHERE " + F_CEK
    End If
 
 If f_Pending <> Empty Then
        CMDSQL = CMDSQL + " AND " + f_Pending
 End If
  
 '@@Start 19 Juli 2010
  If strBlok <> Empty Then
    If Len(CMDSQL) > 0 Then
        CMDSQL = CMDSQL + " AND " + strBlok
    Else
        CMDSQL = CMDSQL + " WHERE " + strBlok
    End If
 End If
 If strBlokentry <> Empty Then
    If Len(CMDSQL) > 0 Then
        CMDSQL = CMDSQL + " AND " + strBlokentry
    Else
        CMDSQL = CMDSQL + " WHERE " + strBlokentry
    End If
 End If
 
 
 If STSLOCKTL <> Empty Then
 If Len(CMDSQL) > 0 Then
        CMDSQL = CMDSQL + " AND " + STSLOCKTL
    Else
        CMDSQL = CMDSQL + " WHERE " + STSLOCKTL
    End If
 
 End If
 
  If strblokedmarkup <> Empty Then
    If Len(SYARAT) > 0 Then
        CMDSQL = CMDSQL + " AND  exclude in (" + Replace(strblokedmarkup, "@", "'") + ") "
    Else
        CMDSQL = CMDSQL + " and exclude in (" + Replace(strblokedmarkup, "@", "'") + ") "
    End If
 End If

If WO_DATE <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + WO_DATE
    Else
        SYARAT = SYARAT + " WHERE " + WO_DATE
    End If
  
  End If
 
 '--------------- @@End 19 Juli 2010
  
CMDSQL = CMDSQL
'ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_mgm = M_objrs
Set M_objrs = Nothing

End Function
Public Function QUERY_SEARCH_jmlrow(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String, Lcustid1 As String, F_CEK As String, lLastCallDate1 As String, lStatusCek1 As String, f_Pending As String, FCEKSTATUS As String, WO_DATE As String, strverify As String, strapprovel As String, strapprovelyet As String, strReject As String, strSukes As String, strBlocked As String, strBlokentry As String, strblokedmarkup As String, nmagentprev As String, LOCKTL As String, BlokedPayment, BlokedPTPNoPayment, Optional page As Long, Optional slimitpage As Double) As Object

Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset

  
    CMDSQL = "SELECT custid from mgm"
SYARAT = ""
If Lcustid1 <> Empty Then
   SYARAT = SYARAT + " WHERE " + Lcustid1
End If

If NamaAgent <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + NamaAgent
    Else
       SYARAT = SYARAT + " WHERE " + NamaAgent
    End If
End If
   
   
If nmagentprev <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + nmagentprev
    Else
       SYARAT = SYARAT + " WHERE " + nmagentprev
    End If
End If

If LOCKTL <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + LOCKTL
    Else
       SYARAT = SYARAT + " WHERE " + LOCKTL
    End If
        
End If

If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If

If FCEKSTATUS <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + FCEKSTATUS
    Else
        SYARAT = SYARAT + " WHERE " + FCEKSTATUS
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
  
  If lLastCallDate1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lLastCallDate1
    Else
        SYARAT = SYARAT + " WHERE " + lLastCallDate1
    End If
 End If
 If lStatusCek1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lStatusCek1
    Else
        SYARAT = SYARAT + " WHERE " + lStatusCek1
    End If
 End If
 
  
  If WO_DATE <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + WO_DATE
    Else
        SYARAT = SYARAT + " WHERE " + WO_DATE
    End If
  
  End If
  
 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 
 If strverify <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strverify
    Else
        SYARAT = SYARAT + " WHERE " + strverify
    End If
 End If
 
 If strapprovel <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strapprovel
    Else
        SYARAT = SYARAT + " WHERE " + strapprovel
    End If
 End If
 
 If strapprovelyet <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strapprovelyet
    Else
        SYARAT = SYARAT + " WHERE " + strapprovelyet
    End If
 End If
 
 If strReject <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strReject
    Else
        SYARAT = SYARAT + " WHERE " + strReject
    End If
 End If
 
 If strSukes <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strSukes
    Else
        SYARAT = SYARAT + " WHERE " + strSukes
    End If
 End If
 
 If f_Pending <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + f_Pending
    Else
        SYARAT = SYARAT + " WHERE " + f_Pending
    End If
  End If
  

 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
 
If F_LOCK = True Then
    SYARAT = SYARAT + " AND F_LOCK='Y'"
 End If
 
 '@@201010 Ini buat gabungin blok data
 If strBlokentry <> Empty Or F_CEK <> Empty Or strblokedmarkup <> Empty Or BlokedPayment <> Empty Then
    SYARAT = SYARAT + " AND ( "
    Dim SyaratBloked As String

    SyaratBloked = ""

    If strBlokentry <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or " + strBlokentry
        Else
            SyaratBloked = SyaratBloked + strBlokentry
        End If
    End If

    If F_CEK <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or " + F_CEK
        Else
            SyaratBloked = SyaratBloked + F_CEK
        End If
    End If
    
    '@@15 Agustus 2011 Blok Payment Request Gaby
    If BlokedPayment <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or (" + BlokedPayment + ") "
        Else
            SyaratBloked = SyaratBloked + BlokedPayment
        End If
    End If
    
    'BlokedPTPNoPayment
    If BlokedPTPNoPayment <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or (" + BlokedPTPNoPayment + ") "
        Else
            SyaratBloked = SyaratBloked + BlokedPTPNoPayment
        End If
    End If
    
    If strblokedmarkup <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or exclude in (" + Replace(strblokedmarkup, "@", "'") + ")"
        Else
            SyaratBloked = SyaratBloked + " exclude in (" + Replace(strblokedmarkup, "@", "'") + ")"
        End If
    End If
    
    SYARAT = SYARAT + SyaratBloked + ")"
 End If
 
 '--------------- @@End 19 Juli 2010
  
CMDSQL = CMDSQL + SYARAT
'ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
DoEvents
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
Set QUERY_SEARCH_jmlrow = M_objrs




Set M_objrs = Nothing

End Function
Public Function QUERY_SEARCH_CALLINITIATED(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String, Lcustid1 As String, F_CEK As String, lLastCallDate1 As String, lStatusCek1 As String) As Object
Dim SYARAT As String
Dim CMDSQL2 As String
Dim M_Objrs2 As ADODB.Recordset

    
    CMDSQL2 = "SELECT CUSTID, COUNT(CUSTID) AS CALL FROM mgm_hst WHERE CUSTID IN "
    CMDSQL2 = CMDSQL2 + " (SELECT CUSTID FROM mgm"
'CMDSQL = "SELECT * FROM SEARCH_mgm"

If Lcustid1 <> Empty Then
   SYARAT = SYARAT + " WHERE " + Lcustid1
End If
If NamaAgent <> Empty Then
   If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + NamaAgent
   Else
       SYARAT = SYARAT + " WHERE " + NamaAgent
   End If
End If
If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
  
  If lLastCallDate1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lLastCallDate1
    Else
        SYARAT = SYARAT + " WHERE " + lLastCallDate1
    End If
 End If
 If lStatusCek1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lStatusCek1
    Else
        SYARAT = SYARAT + " WHERE " + lStatusCek1
    End If
 End If
 
 If F_CEK <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + F_CEK
    Else
        SYARAT = SYARAT + " WHERE " + F_CEK
    End If
  End If

 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
CMDSQL2 = CMDSQL2 + " " + SYARAT + ") GROUP BY CUSTID "
CMDSQL2 = "SELECT * FROM ( " + CMDSQL2 + ")A ORDER BY CUSTID "
Set M_Objrs2 = New ADODB.Recordset
M_Objrs2.CursorLocation = adUseClient
M_Objrs2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_CALLINITIATED = M_Objrs2
Set M_Objrs2 = Nothing

End Function
Public Function QUERY_SEARCH_VISIT(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, Lcustid1 As String, STS As String) As Object
Dim SYARAT As String
Dim CMDSQL2 As String
Dim M_Objrs2 As ADODB.Recordset


    CMDSQL2 = "SELECT TblVisit.*, mgm.NAME, mgm.AGENT AS AGENT"
    CMDSQL2 = CMDSQL2 + " FROM  TblVisit INNER JOIN "
    CMDSQL2 = CMDSQL2 + " mgm ON TblVisit.CUSTID = mgm.CUSTID  "

If Lcustid1 <> Empty Then
   SYARAT = SYARAT + " WHERE " + Lcustid1
End If
If NamaAgent <> Empty Then
   If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + NamaAgent
   Else
       SYARAT = SYARAT + " WHERE " + NamaAgent
   End If
End If
If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 
 If STS <> Empty Then
 If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + STS
    Else
        SYARAT = SYARAT + " WHERE " + STS
    End If
 End If
 
 CMDSQL2 = CMDSQL2 + " " + SYARAT + "  ORDER BY tblvisit.CUSTID"
Set M_Objrs2 = New ADODB.Recordset
M_Objrs2.CursorLocation = adUseClient
M_Objrs2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_VISIT = M_Objrs2
Set M_Objrs2 = Nothing
End Function


Public Function QUERY_AGENT_JADWAL(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "SELECT usertbl.*"
    CMDSQL = CMDSQL + " FROM usertbl "
    If Len(M_WHERE) <> 0 Then
       CMDSQL = CMDSQL + " WHERE " + M_WHERE
    Else
       'cmdsql = cmdsql + " WHERE AKTIF = 0"
       
    End If
    CMDSQL = CMDSQL + " ORDER BY USERID"
        
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT_JADWAL = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_AGENT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT usertbl.*, SPVTBL.SPVNAME"
    CMDSQL = CMDSQL + " FROM usertbl INNER JOIN"
    CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
    CMDSQL = CMDSQL + " WHERE USERTYPE ='1'"
    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
    
    CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT = M_objrs
Set M_objrs = Nothing

End Function
Public Function QUERY_FIELD(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT usertbl.*, SPVTBL.SPVNAME"
    CMDSQL = CMDSQL + " FROM usertbl INNER JOIN"
    CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
    CMDSQL = CMDSQL + " WHERE USERTYPE ='2'"
    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
    
    CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_FIELD = M_objrs
Set M_objrs = Nothing

End Function

Public Function QUERY_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


CMDSQL = "SELECT * FROM datasourcetbl"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY KODEDS"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_DATASOURCE = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_AGENT_xl(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT usertbl.*, SPVTBL.SPVNAME"
    CMDSQL = CMDSQL + " FROM usertbl INNER JOIN"
    CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
'    CMDSQL = CMDSQL + " WHERE USERTYPE ='1'"
    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " where " + M_WHERE
    End If
    
    CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT_xl = M_objrs
Set M_objrs = Nothing

End Function

Public Function QUERY_SEARCH_CONDITION_mgm2(M_OBJCONN As ADODB.Connection, Reff_Num As String, No_Laporan As String, tgl_laporan As String, Idi_User As String, Debitur_Name As String, Born_Place As String, Born_Place2 As String, Born_Date As String, Din As String, NPWP As String, NPWP2 As String, KTP_NUM As String, KTP_NUM2 As String, Passport_Num As String, Passport_Num2 As String, HTML_SOURCE As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset
'    CMDSQL = "SELECT CUSTID,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek, AGENT,  Principal,DOB "
'    CMDSQL = CMDSQL + "AmountWo, OpenDate, TtlPTP, RECSOURCE, TGLSTATUS, Kethslkerja, TGLINCOMING, AmountWo,f_Pending FROM mgm"
    CMDSQL = "SELECT * From MAIN_DEBITUR_INFO "
    
'If Lcustid1 <> Empty Then
'   SYARAT = SYARAT + " WHERE " + Lcustid1
'End If
If Reff_Num <> Empty Then
   If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + Reff_Num
   Else
       SYARAT = SYARAT + " WHERE " + Reff_Num
   End If
End If
If Debitur_Name <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + Debitur_Name
    Else
        SYARAT = SYARAT + " WHERE " + Debitur_Name
    End If
 End If
' If DATASOURCE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + DATASOURCE
'    Else
'        SYARAT = SYARAT + " WHERE " + DATASOURCE
'    End If
' End If
' If TGLLAHIR <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + TGLLAHIR
'    Else
'        SYARAT = SYARAT + " WHERE " + TGLLAHIR
'    End If
' End If
'
'  If lLastCallDate1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + lLastCallDate1
'    Else
'        SYARAT = SYARAT + " WHERE " + lLastCallDate1
'    End If
' End If
' If lStatusCek1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + lStatusCek1
'    Else
'        SYARAT = SYARAT + " WHERE " + lStatusCek1
'    End If
' End If
'
' If F_CEK <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + F_CEK
'    Else
'        SYARAT = SYARAT + " WHERE " + F_CEK
'    End If
'  End If
'
' If OFFPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND (" + OFFPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + OFFPHONE
'    End If
' End If
' If OFFPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
' End If
' If HOMEPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR (" + HOMEPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
'    End If
' End If
' If HOMEPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
' End If
' If MOBILEPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND (" + MOBILEPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
'    End If
' End If
' If f_Pending <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + f_Pending
'    Else
'        SYARAT = SYARAT + " WHERE " + f_Pending
'    End If
'  End If
'
' If MOBILEPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
' End If
' If FAXPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR (" + FAXPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + FAXPHONE
'    End If
' End If
' If FAXPHONE2 <> Empty Then
'     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
' End If
CMDSQL = CMDSQL + " " + SYARAT + ""

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient


M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_CONDITION_mgm2 = M_objrs
Set M_objrs = Nothing


End Function


