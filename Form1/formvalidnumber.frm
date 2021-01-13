VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form formvalidnumber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valid Number"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3780
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Download Valid Number"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvvalid 
      Height          =   510
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   900
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
End
Attribute VB_Name = "formvalidnumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call header
    Call isi
    
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Double
    Dim a As String
    If lvvalid.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
        For col = 1 To lvvalid.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = lvvalid.ColumnHeaders(col)
        Next
     
        For Row = 2 To lvvalid.ListItems.Count + 1
            For col = 1 To lvvalid.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = "'" & lvvalid.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = "'" + lvvalid.ListItems(Row - 1).SubItems(col - 1)
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
End Sub



Private Sub header()
    lvvalid.ColumnHeaders.clear
    lvvalid.ColumnHeaders.ADD , , "CUSTID", 2000
    lvvalid.ColumnHeaders.ADD , , "VALID NUMBER", 2500
End Sub

Private Sub isi()
    sQuery = "select custid,validsms from mgm where coalesce(validsms,'') <> ''"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lvvalid.ListItems.clear
    If RS_Lv.RecordCount > 0 Then
        Do Until RS_Lv.EOF
            Set listItem = lvvalid.ListItems.ADD(, , Trim(cnull(RS_Lv("custid"))))
            listItem.SubItems(1) = Trim(cnull(RS_Lv("validsms")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub
