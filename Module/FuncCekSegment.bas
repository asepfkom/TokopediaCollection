Attribute VB_Name = "FuncCekSegment"
Public Function FuncCekSegmen(notlp As String) As Double
    Dim sQuery As String
    Dim Rs_Segmen As ADODB.Recordset
    
    FuncCekSegmen = 0
    
    sQuery = "SELECT * FROM tbl_temp_segment_call "
    sQuery = sQuery + " WHERE no_telpon = '" & notlp & "' AND date(tgl_call) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "' "
    Set Rs_Segmen = New ADODB.Recordset
    Rs_Segmen.CursorLocation = adUseClient
    Rs_Segmen.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Rs_Segmen.RecordCount > 0 Then
        FuncCekSegmen = Rs_Segmen!jumlah_call
    Else
        FuncCekSegmen = 0
    End If
End Function


Public Function FuncCekReview(notlp As String, CustId As String) As Double
    Dim sQuery As String
    Dim Rs_Jumlah_Call As ADODB.Recordset

    'CustId = Trim(FrmCC_Colection.lblCustId.Caption)
    
    FuncCekReview = 0

    sQuery = "SELECT * FROM tbl_temp_telfon_review "
    sQuery = sQuery + " WHERE no_telfon = '" & notlp & "' AND date(tanggal_telfon) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "' AND custId = '" & CustId & "'" 'UPDATETIAN23FEBRUARI2016
    Set Rs_Jumlah_Call = New ADODB.Recordset
    Rs_Jumlah_Call.CursorLocation = adUseClient
    Rs_Jumlah_Call.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If Rs_Jumlah_Call.RecordCount > 0 Then
        FuncCekReview = Rs_Jumlah_Call!jumlah_call
    Else
        FuncCekReview = 0
    End If
End Function
