Imports DevExpress.DataAccess
Imports DevExpress.DataAccess.Excel

Imports Microsoft.Office.Interop
Imports System.Data.Odbc
Imports System.IO
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid

Public Class FrmImportVCT
    Dim FileName As String
    Dim Code As String = ""
    Dim FName As String = ""
    Dim tabel As String = ""
    Private Sub FrmImportVCT_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        SplitContainerControl1.SplitterPosition = Me.Width / 2
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        'FILE OPEN
        On Error Resume Next
        OpenFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()

        FileName = OpenFileDialog1.FileName.ToString
        TextEdit1.Text = FileName
    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        'LOAD
        If FileName = "" Or ComboBoxEdit1.Text = "" Then
            MsgBox("Data Tidak Lengkap", vbInformation, "Process Data")
        Else
            ShowData(FileName, ComboBoxEdit1.Text)
        End If
    End Sub
    Private Sub ShowData(ByVal nFileName As String, ByVal nSheet As String)
        On Error Resume Next
        If nFileName Is Nothing Then Exit Sub
        If nSheet Is Nothing Then Exit Sub

        If TextEdit1.Text = "" Then MsgBox("File Belum di pilih", vbInformation, "Excel File ")

        Dim myExcelSource As New DevExpress.DataAccess.Excel.ExcelDataSource()
        myExcelSource.FileName = nFileName

        Dim worksheetSettings As New ExcelWorksheetSettings(nSheet, "A1:L$")
        myExcelSource.SourceOptions = New ExcelSourceOptions(worksheetSettings)
        myExcelSource.SourceOptions.SkipEmptyRows = True
        myExcelSource.SourceOptions.UseFirstRowAsHeader = True

        GridControl1.DataSource = myExcelSource
        myExcelSource.Fill()
        LabelControl4.Text = "Imports Data : Record Count. " & GridControl1.MainView.RowCount.ToString
    End Sub

    Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click
        'PROSES ISI BARU
        SimpleButton10.Enabled = False
        SimpleButton1.Enabled = False : SimpleButton2.Enabled = False : SimpleButton3.Enabled = False
        SaveGrid(GridView1)
        SimpleButton1.Enabled = True : SimpleButton2.Enabled = True : SimpleButton3.Enabled = True
        SimpleButton10.Enabled = True
    End Sub
    Private Sub SaveGrid(ByRef grid As GridView)

        ' Initializing progress bar properties
        ProgressBarControl1.Properties.Step = 1
        ProgressBarControl1.Properties.PercentView = True
        ProgressBarControl1.Properties.Minimum = 0
        ProgressBarControl1.Properties.Maximum = grid.RowCount
        ProgressBarControl1.EditValue = 0
        'INSERT M_VCT

        LabelControl1.Text = "Progres On "
        Dim i As Integer
        For i = 0 To grid.RowCount - 1

            ProgressBarControl1.PerformStep()
            ProgressBarControl1.Update()

            Dim Tcode As String = CType((grid.GetRowCellValue(i, Code)), String)
            Dim TNama As String = CType((grid.GetRowCellValue(i, FName)), String)
            Dim TAddress As String = CType((grid.GetRowCellValue(i, "ADDRESS")), String)
            Dim Tcity As String = CType((grid.GetRowCellValue(i, "CITY")), String)
            Dim Tstate As String = CType((grid.GetRowCellValue(i, "STATE")), String)
            Dim TPhone As String = CType((grid.GetRowCellValue(i, "PHONE")), String)

            SQL = ""
            If ComboBoxEdit1.Text = "VENDOR" Then
                SQL = "SELECT VENDOR_C,VENDOR_N,ADDRESS,CITY,STATE,PHONE FROM M_VENDOR WHERE VENDOR_C='" & Tcode & "'"
                If CheckRecord(SQL) > 0 Then
                    'UPDATE
                    SQL = "UPDATE M_VENDOR SET VENDOR_N='" & TNama & "',ADDRESS='" & TAddress & "',CITY='" & Tcity & "',PHONE='" & TPhone & "' WHERE VENDOR_C='" & Tcode & "'"
                Else
                    'INSERT
                    SQL = "INSERT INTO M_VENDOR (VENDOR_C,VENDOR_N,ADDRESS,CITY,STATE,PHONE) VALUES " +
                        " ('" & Tcode & "','" & TNama & "' ,'" & TAddress & "','" & Tcity & "', '" & Tstate & "','" & TPhone & "')"
                End If
                ExecuteNonQuery(SQL)
            ElseIf ComboBoxEdit1.Text = "CUSTOMER" Then
                SQL = "SELECT CUSTOMER_C,CUSTOMER_N,ADDRESS,CITY,STATE,PHONE FROM M_CUSTOMER WHERE CUSTOMER_C='" & Tcode & "' "
                If CheckRecord(SQL) > 0 Then
                    'UPDATE
                    SQL = "UPDATE M_CUSTOMER SET CUSTOMER_N='" & TNama & "',ADDRESS='" & TAddress & "',CITY='" & Tcity & "',PHONE='" & TPhone & "' WHERE CUSTOMER_C='" & Tcode & "'"
                Else
                    'INSERT
                    SQL = "INSERT INTO M_CUSTOMER (CUSTOMER_C,CUSTOMER_N,ADDRESS,CITY,STATE,PHONE) VALUES " +
                        " ('" & Tcode & "','" & TNama & "' ,'" & TAddress & "','" & Tcity & "', '" & Tstate & "','" & TPhone & "')"
                End If
                ExecuteNonQuery(SQL)
            ElseIf ComboBoxEdit1.Text = "TRANSPORTER" Then
                SQL = "Select TRANSPORTER_C,TRANSPORTER_N,ADDRESS,CITY,STATE,PHONE FROM M_TRANSPORTER WHERE TRANSPORTER_C='" & Tcode & "' "
                If CheckRecord(SQL) > 0 Then
                    'UPDATE
                    SQL = "UPDATE M_TRANSPORTER SET TRANSPORTER_N='" & TNama & "',ADDRESS='" & TAddress & "',CITY='" & Tcity & "',PHONE='" & TPhone & "' WHERE TRANSPORTER_C='" & Tcode & "'"
                Else
                    'INSERT
                    SQL = "INSERT INTO M_TRANSPORTER (TRANSPORTER_C,TRANSPORTER_N,ADDRESS,CITY,STATE,PHONE) VALUES " +
                        " ('" & Tcode & "','" & TNama & "' ,'" & TAddress & "','" & Tcity & "', '" & Tstate & "','" & TPhone & "')"
                End If
                ExecuteNonQuery(SQL)
            End If


        Next

        LoadDataExsisting()
        MsgBox("PROGRESS COMPLATE", vbInformation, "UPLOAD DATA")
    End Sub

    Private Sub FrmImportVCT_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "IMPORT VENDOR,CUSTOMER,TRANSPORTER"
        'LOAD
        TextEdit1.Text = ""
        ComboBoxEdit1.Text = ""

        ComboBoxEdit1.SelectedIndex = 0

        'CREATE HEADER
        CreateHeader()
        CreateHeaderEx()
        LoadDataExsisting()

    End Sub
    Private Sub LoadDataExsisting()
        'LOAD GRID 2
        SQL = ""
        If ComboBoxEdit1.Text = "VENDOR" Then
            SQL = "SELECT VENDOR_C,VENDOR_N,ADDRESS,CITY,STATE,PHONE FROM M_VENDOR "
        ElseIf ComboBoxEdit1.Text = "CUSTOMER" Then
            SQL = "SELECT CUSTOMER_C,CUSTOMER_N,ADDRESS,CITY,STATE,PHONE FROM M_CUSTOMER "
        ElseIf ComboBoxEdit1.Text = "TRANSPORTER" Then
            SQL = "SELECT TRANSPORTER_C,TRANSPORTER_N,ADDRESS,CITY,STATE,PHONE FROM M_TRANSPORTER "
        End If

        FILLGridView(SQL, GridControl2)
        LabelControl5.Text = "Exsisting Data : Record Count. " & GridControl2.MainView.RowCount.ToString

    End Sub
    Private Sub CreateHeader()
        Dim View As ColumnView = CType(GridControl1.MainView, ColumnView)
        If ComboBoxEdit1.Text = "VENDOR" Then
            Code = "VENDOR_C"
            FName = "VENDOR_N"
        ElseIf ComboBoxEdit1.Text = "CUSTOMER" Then
            Code = "CUSTOMER_C"
            FName = "CUSTOMER_N"
        ElseIf ComboBoxEdit1.Text = "TRANSPORTER" Then
            Code = "TRANSPORTER_C"
            FName = "TRANSPORTER_N"
        End If

        Dim FieldNames() As String = New String() {Code, FName, "ADDRESS", "CITY", "STATE", "PHONE"}
        Dim I As Integer
        Dim Column As DevExpress.XtraGrid.Columns.GridColumn

        View.Columns.Clear()
        For I = 0 To FieldNames.Length - 1
            Column = View.Columns.AddField(FieldNames(I))
            Column.VisibleIndex = I
        Next
        GridView1.BestFitColumns()
    End Sub
    Private Sub CreateHeaderEx()
        Dim View As ColumnView = CType(GridControl1.MainView, ColumnView)
        If ComboBoxEdit1.Text = "VENDOR" Then
            Code = "VENDOR_C"
            FName = "VENDOR_N"
        ElseIf ComboBoxEdit1.Text = "CUSTOMER" Then
            Code = "CUSTOMER_C"
            FName = "CUSTOMER_N"
        ElseIf ComboBoxEdit1.Text = "TRANSPORTER" Then
            Code = "TRANSPORTER_C"
            FName = "TRANSPORTER_N"
        End If

        Dim FieldNames() As String = New String() {Code, FName, "ADDRESS", "CITY", "STATE", "PHONE"}
        Dim I As Integer
        Dim Column As DevExpress.XtraGrid.Columns.GridColumn

        Dim View2 As ColumnView = CType(GridControl2.MainView, ColumnView)
        View2.Columns.Clear()
        For I = 0 To FieldNames.Length - 1
            Column = View2.Columns.AddField(FieldNames(I))
            Column.VisibleIndex = I
        Next
        GridView2.BestFitColumns()
    End Sub

    Private Sub ComboBoxEdit1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit1.SelectedIndexChanged
        CreateHeader()
        CreateHeaderEx()
        If ComboBoxEdit1.Text <> "" Then
            ShowData(FileName, ComboBoxEdit1.Text)
        End If
        LoadDataExsisting()
    End Sub

    Private Sub SimpleButton10_Click(sender As Object, e As EventArgs) Handles SimpleButton10.Click
        'CLOSE
        Me.Close()
    End Sub

    Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click
        GridView1.Columns.Clear()
        GridView2.Columns.Clear()

        CreateHeader()
        CreateHeaderEx()

        TextEdit1.Text = ""

        LabelControl1.Text = "Progres On"
    End Sub
End Class
