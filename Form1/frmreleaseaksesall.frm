VERSION 5.00
Begin VB.Form frmreleaseaksesall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancel AksesAll"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   990
   ClientWidth     =   5655
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5655
   Begin VB.Frame Frame1 
      Caption         =   "Running AksesAll"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton Command1 
         Caption         =   "Release"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   280
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Batch AksesAll"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmreleaseaksesall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_DropDown()
    Combo1.clear
        cmdsql = "select kd_profile from tbl_profile_aksesall where waktu_awal < now() and waktu_akhir > now()"
        Set M_ObjrsCekStatus = New ADODB.Recordset
        M_ObjrsCekStatus.CursorLocation = adUseClient
        M_ObjrsCekStatus.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_ObjrsCekStatus.RecordCount > 0 Then
            For i = 1 To M_ObjrsCekStatus.RecordCount
                Combo1.AddItem M_ObjrsCekStatus!kd_profile
                M_ObjrsCekStatus.MoveNext
            Next i
        End If
End Sub

Private Sub Command1_Click()
    qs = "select kd_profile from tbl_profile_aksesall where kd_profile ='" + Combo1.text + "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs.RecordCount > 0 Then
        qu = "UPDATE usertbl SET profile_akses_all=null,f_akses_all_acc=null,f_pesanresetauto=null WHERE profile_akses_all in (SELECT kd_profile FROM tbl_profile_aksesall WHERE kd_profile = '" & Combo1.text & "');"
        qu = qu + "update tbl_profile_aksesall set waktu_akhir = now() where kd_profile = '" + Combo1.text + "' ;"
        M_OBJCONN.Execute qu
        
        cmdsql = "UPDATE mgm SET agent=agent_asli, monitor_akses = null, waktu_akses = null WHERE " & _
                " agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
                " a.kd_profile=b.kd_profile and a.kd_profile = '" + Combo1.text + "')"
        M_OBJCONN.Execute cmdsql
        
        cmdsql = "DELETE FROM tbl_cust_aksesall "
        cmdsql = cmdsql & " WHERE kd_profile = '" + Combo1.text + "' "
        M_OBJCONN.Execute cmdsql
        
    Else
        MsgBox "Tidak Ditemukan"
    End If
End Sub
