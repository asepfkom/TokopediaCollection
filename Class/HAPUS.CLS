VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HAPUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean
Public Function DELETED_USERLOG(M_OBJCONN As ADODB.Connection, agent As String, TANGGAL As String)
Dim CMDSQL As String
On Error GoTo HELL
CMDSQL = " DELETE USERLOG "
CMDSQL = CMDSQL + " WHERE AGENT = '" + agent + "'"
CMDSQL = CMDSQL + " AND DATETIME = '" + TANGGAL + "'"

M_OBJCONN.Execute CMDSQL

Exit Function
HELL:

End Function

Public Function DELETED_CLIENTTBL(M_OBJCONN As ADODB.Connection, agent As String)
Dim CMDSQL As String
On Error GoTo HELL
CMDSQL = " DELETE CLIENTTBL "
CMDSQL = CMDSQL + " WHERE AGENT = '" + agent + "'"

M_OBJCONN.Execute CMDSQL

Exit Function
HELL:

End Function


Public Function UPDATE_CLIENT_CLOSE(M_OBJCONN As ADODB.Connection, WSSTATIONIP As String, agent As String, WSSTATION As String)
Dim CMDSQL As String
CMDSQL = " UPDATE CLIENTTBL"
CMDSQL = CMDSQL + " SET WSSTATIONIP ='" + WSSTATIONIP + "',"
CMDSQL = CMDSQL + " WSSTATION ='" + WSSTATION + "'"
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " AGENT ='" + agent + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function UPDATE_LOGOUT(M_OBJCONN As ADODB.Connection, TANGGAL As String, WSTATION As String, ACTIVITY As String, agent As String)
Dim CMDSQL As String
CMDSQL = " UPDATE USERLOG"
CMDSQL = CMDSQL + " SET WSTATION ='" + WSTATION + "',"
CMDSQL = CMDSQL + " ACTIVITY='" + ACTIVITY + "' "
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " AGENT ='" + agent + "' "
CMDSQL = CMDSQL + " AND DATETIME = '" + TANGGAL + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox Err.Description
End Function


