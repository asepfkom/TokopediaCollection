VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form formsegment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segmen Export"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FFC0C0&
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Export"
         Height          =   375
         Left            =   8160
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000010&
         Caption         =   "Search"
         Height          =   375
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filter"
         Height          =   1095
         Left            =   7920
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "formsegment.frx":0000
            Left            =   120
            List            =   "formsegment.frx":0019
            TabIndex        =   4
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Segment"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lvsegment 
         Height          =   5550
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   9790
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "formsegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    lvsegment.ColumnHeaders.ADD , , "CUSTID", 2000
    lvsegment.ColumnHeaders.ADD , , "NAMA CH", 2500
    lvsegment.ColumnHeaders.ADD , , "SEGMENT", 1000
    lvsegment.ColumnHeaders.ADD , , "AGENT", 1000
    lvsegment.ColumnHeaders.ADD , , "TL", 1000
End Sub

Private Sub Command1_Click()
    Call isi
End Sub

Private Sub Command2_Click()
Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet
Dim col, Row As Double
Dim a As String
If lvsegment.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To lvsegment.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = lvsegment.ColumnHeaders(col)
    Next
 
    For Row = 2 To lvsegment.ListItems.Count + 1
        For col = 1 To lvsegment.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(Row, col).Value = "'" & lvsegment.ListItems(Row - 1).text
        Else
            '" 'cararandy 29032016 "
            Dim hasil1 As String
                hasil1 = "'" + lvsegment.ListItems(Row - 1).SubItems(col - 1)
                objExcelSheet.Cells(Row, col).Value = hasil1
            End If
        Next
    Next
 
    objExcelSheet.Columns.AutoFit
    CD_save.ShowOpen
    a = CD_save.FileName
 
    objExcelSheet.SaveAs a & ".xls"
    MsgBox "Export Completed", vbInformation, Me.Caption
 
    objExcel.Workbooks.Open a & ".xls"
    objExcel.Visible = True
Else
    MsgBox "No data to export", vbInformation, Me.Caption
End If

End Sub

Private Sub Form_Load()
    Call header
End Sub


Private Sub isi()
    
    If UCase(Combo1.text) = "ALL" Then
        sQuery = "select * from ("
        sQuery = sQuery & " select a.*,b.team from (select custid,name,segment,agent from mgm " & vbCrLf
        'sQuery = sQuery & " where coalesce(segment,'') = '" & Combo1.text & "' " & vbCrLf
        sQuery = sQuery & " ) a left join usertbl b on a.agent = b.userid" & vbCrLf
        sQuery = sQuery & " ) b " & vbCrLf
        
        If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
            sQuery = sQuery & " where team = '" & MDIForm1.Text1.text & "'"
        End If
        
        Set RS_Lv = New ADODB.Recordset
        RS_Lv.CursorLocation = adUseClient
        RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Else
        sQuery = "select * from ("
        sQuery = sQuery & "select a.*,b.team from (select custid,name,segment,agent from mgm " & vbCrLf
        sQuery = sQuery & " where coalesce(segment,'') = '" & Combo1.text & "' " & vbCrLf
        sQuery = sQuery & " ) a left join usertbl b on a.agent = b.userid" & vbCrLf
        sQuery = sQuery & " ) b " & vbCrLf
        
        If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
            sQuery = sQuery & " where team = '" & MDIForm1.Text1.text & "'"
        End If
        
        Set RS_Lv = New ADODB.Recordset
        RS_Lv.CursorLocation = adUseClient
        RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    End If
    lvsegment.ListItems.clear
    If RS_Lv.RecordCount > 0 Then
        Do Until RS_Lv.EOF
            Set listItem = lvsegment.ListItems.ADD(, , Trim(cnull(RS_Lv("custid"))))
            listItem.SubItems(1) = Trim(cnull(RS_Lv("name")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("segment")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("agent")))
            listItem.SubItems(4) = Trim(cnull(RS_Lv("team")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

Private Sub lvsegment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvsegment.SortKey = ColumnHeader.Index - 1
    lvsegment.Sorted = True
End Sub
