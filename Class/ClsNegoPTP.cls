VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNegoPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ADD_OK As Boolean
Public Function ADD_NegoPTP(M_OBJCONN As ADODB.Connection, CustId As String, PromiseDate As String, PromisePay As String, inputdate As String, Jns As String)
Dim cmdsql As String
PromiseDate = Format(PromiseDate, "yyyy-mm-dd")
inputdate = Format(inputdate, "yyyy-mm-dd")
cmdsql = "INSERT INTO TblNegoPTP "
cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
cmdsql = cmdsql + "VALUES "
cmdsql = cmdsql + "('" + CustId + "', "
cmdsql = cmdsql + "'" + PromiseDate + "', "
cmdsql = cmdsql + "" + PromisePay + " , "
cmdsql = cmdsql + "'" + inputdate + "', "
cmdsql = cmdsql + "'" + Jns + "')"
On Error GoTo add_error
M_OBJCONN.Execute cmdsql
ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description

End Function

Public Function UPDATE_NegoPTP(M_OBJCONN As ADODB.Connection, CustId As String, PromiseDate As String, PromisePay As String, ID As String)
Dim cmdsql As String
PromiseDate = Format(PromiseDate, "yyyy-mm-dd")
inputdate = Format(inputdate, "yyyy-mm-dd")
cmdsql = "UPDATE tblnegoPTP SET "
cmdsql = cmdsql + "Promisedate = '" + PromiseDate + "', "
cmdsql = cmdsql + "Promisepay = " + PromisePay + " "
cmdsql = cmdsql + "WHERE id = '" + ID + "'"
On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    
    '@@ 29 Des 2011 Update Status PTP ke tabel mgm
    cmdsql = "update mgm set dateptp='"
    cmdsql = cmdsql + PromiseDate + "' where custid='"
    cmdsql = cmdsql + CustId + "'"
    M_OBJCONN.Execute cmdsql
    
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function


Public Function DELETE_Nego_PTP(M_OBJCONN As ADODB.Connection, ID As String)
    Dim cmdsql As String
    Dim M_Objrs_Log As ADODB.Recordset
    
    '@@ 11042012 Catet yang menghapus data PTP
    cmdsql = "select * from tblnegoptp where id='"
    cmdsql = cmdsql + CStr(ID) + "'"
    Set M_Objrs_Log = New ADODB.Recordset
    M_Objrs_Log.CursorLocation = adUseClient
    M_Objrs_Log.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Log.RecordCount > 0 Then
        cmdsql = "insert into tblnegoptp_logdel (custid,promisedate,promisepay,inputdate,"
        cmdsql = cmdsql + "user_del,tgl_del) values ('"
        cmdsql = cmdsql + M_Objrs_Log("custid") + "','"
        cmdsql = cmdsql + CStr(Format(M_Objrs_Log("promisedate"), "yyyy-mm-dd")) + "','"
        cmdsql = cmdsql + CStr(M_Objrs_Log("promisepay")) + "','"
        cmdsql = cmdsql + CStr(Format(M_Objrs_Log("inputdate"), "yyyy-mm-dd")) + "','"
        cmdsql = cmdsql + MDIForm1.Text1.Text + "',"
        cmdsql = cmdsql + "now())"
        M_OBJCONN.Execute cmdsql
    End If
    Set M_Objrs_Log = Nothing
    
    
    cmdsql = "delete from tblnegoptp WHERE "
    cmdsql = cmdsql + "ID = '" + ID + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    
  
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_Nego_Reserve(M_OBJCONN As ADODB.Connection, ID As String)
Dim cmdsql As String
cmdsql = "DELETE from tblreserve WHERE "
cmdsql = cmdsql + "ID = '" + ID + "'"
On Error GoTo add_error
M_OBJCONN.Execute cmdsql
ADD_OK = True
Exit Function
add_error:
ADD_OK = False
End Function

