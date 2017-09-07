Imports DevExpress.DataAccess
Imports DevExpress.DataAccess.Excel

Imports Microsoft.Office.Interop
Imports System.Data.Odbc
Imports System.IO
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid

Public Class FrmImportItem
    Dim FileName As String
    Private Sub FrmImportItem_Resize(sender As Object, e As EventArgs) Handles Me.Resize
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
        Dim KdItem As String
        Dim Item As String
        Dim Jenis As String
        Dim Unit As String
        Dim i As Integer

        'HAPUS M_ITEM_TEMP
        SQL = "DELETE FROM M_ITEM_TEMP WHERE AKTIF='Y' AND IDUSER='" & USERNAME & "' AND JENIS='" & ComboBoxEdit1.Text & "'"
        ExecuteNonQuery(SQL)

        ' Initializing progress bar properties
        ProgressBarControl1.Properties.Step = 1
        ProgressBarControl1.Properties.PercentView = True
        ProgressBarControl1.Properties.Minimum = 0
        ProgressBarControl1.Properties.Maximum = grid.RowCount
        ProgressBarControl1.EditValue = 0
        'INSERT M_ITEM_TEMP

        LabelControl1.Text = "Progres On M_ITEM_TEMP"
        For i = 0 To grid.RowCount - 1


            KdItem = CType((grid.GetRowCellValue(i, "ITEM_C")), String)
            Item = CType((grid.GetRowCellValue(i, "ITEM")), String)
            Jenis = ComboBoxEdit1.Text
            Unit = CType((grid.GetRowCellValue(i, "UNIT")), String)

            SQL = "INSERT M_ITEM_TEMP (ITEM_C,ITEM,JENIS,UNIT,AKTIF,IDUSER) " +
                  "VALUES('" & KdItem & "','" & Item & "','" & Jenis & "','" & Unit & "','Y','" & USERNAME & "')"
            ExecuteNonQuery(SQL)

            ProgressBarControl1.PerformStep()
            ProgressBarControl1.Update()


        Next
        ProgressBarControl1.Properties.PercentView = True
        ProgressBarControl1.Properties.Minimum = 0
        ProgressBarControl1.Properties.Maximum = grid.RowCount
        ProgressBarControl1.Properties.Step = 1
        ProgressBarControl1.EditValue = 0
        'INSERT M_ITEM
        LabelControl1.Text = "Progres On M_ITEM"
        For i = 0 To grid.RowCount - 1

            ProgressBarControl1.PerformStep()
            ProgressBarControl1.Update()

            KdItem = CType((grid.GetRowCellValue(i, "ITEM_C")), String)
            Item = CType((grid.GetRowCellValue(i, "ITEM")), String)
            Jenis = ComboBoxEdit1.Text
            Unit = CType((grid.GetRowCellValue(i, "UNIT")), String)

            SQL = "SELECT * FROM M_ITEM WHERE ITEM_C='" & KdItem & "' AND JENIS='" & Jenis & "'"
            If CheckRecord(SQL) > 0 Then
                'UPDATE
                SQL = " Update M_ITEM " +
                " Set ITEM='" & KdItem & "',JENIS='" & Jenis & "',UNIT='" & Unit & "'" +
                " WHERE ITEM_C ='" & KdItem & "'  AND JENIS='" & Jenis & "' "

                ExecuteNonQuery(SQL)
            Else
                'INSERT
                SQL = " INSERT M_ITEM (ITEM_C,ITEM,JENIS,UNIT) " +
                      " VALUES('" & KdItem & "','" & Item & "','" & Jenis & "','" & Unit & "')"
                ExecuteNonQuery(SQL)
            End If
        Next

        SQL = "SELECT ITEM_C,ITEM,JENIS,UNIT FROM M_ITEM WHERE JENIS='" & ComboBoxEdit1.Text & "' "
        GridControl2.DataSource = ExecuteQuery(SQL)
        LabelControl5.Text = "Exsisting Data : Record Count. " & GridControl2.MainView.RowCount.ToString
        ProgressBarControl1.EditValue = 0
        LabelControl1.Text = "Progres On"
        MsgBox("PROGRESS COMPLATE", vbInformation, "UPLOAD DATA")

    End Sub

    Private Sub FrmImportItem_Load(sender As Object, e As EventArgs) Handles Me.Load
        'LOAD
        TextEdit1.Text = ""
        ComboBoxEdit1.Text = ""

        'ISI COMBO JENIS
        SQL = "SELECT JENIS,JENIS_DESC FROM M_JENIS"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit1, False)

        'CREATE HEADER
        CreateHeader()
        CreateHeaderEx()
        LoadDataExsisting()

    End Sub
    Private Sub LoadDataExsisting()
        'LOAD GRID 2
        SQL = "SELECT ITEM_C,ITEM,JENIS,UNIT FROM M_ITEM WHERE JENIS LIKE '%" & ComboBoxEdit1.Text & "%'"
        FILLGridView(SQL, GridControl2)
        LabelControl5.Text = "Exsisting Data : Record Count. " & GridControl2.MainView.RowCount.ToString

    End Sub
    Private Sub CreateHeader()
        Dim View As ColumnView = CType(GridControl1.MainView, ColumnView)
        Dim FieldNames() As String = New String() {"ITEM_C", "ITEM", "UNIT"}
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
        Dim View As ColumnView = CType(GridControl2.MainView, ColumnView)
        Dim FieldNames() As String = New String() {"ITEM_C", "ITEM", "JENIS", "UNIT"}
        Dim I As Integer
        Dim Column As DevExpress.XtraGrid.Columns.GridColumn

        View.Columns.Clear()
        For I = 0 To FieldNames.Length - 1
            Column = View.Columns.AddField(FieldNames(I))
            Column.VisibleIndex = I
        Next
        GridView2.BestFitColumns()
    End Sub

    Private Sub ComboBoxEdit1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit1.SelectedIndexChanged
        'SQL = "SELECT ITEM_C,ITEM,JENIS,UNIT FROM M_ITEM WHERE JENIS='" & ComboBoxEdit1.Text & "'"
        'FILLGridView(SQL, GridControl2)
        'LabelControl5.Text = "Exsisting Data : Record Count. " & GridControl2.MainView.RowCount.ToString
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
Public Class ExcelDataBaseHelper
    Private Shared Function OpenFile(ByVal fileName As String) As Object
        Dim fullFileName = String.Format("{0}\{1}", Directory.GetCurrentDirectory(), fileName)
        If (Not File.Exists(fullFileName)) Then
            System.Windows.Forms.MessageBox.Show("File Not found")
            Return Nothing
        End If
        Dim connectionString As String = String.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fullFileName)
        Dim adapter = New OdbcDataAdapter("Select * from [Sheet1$]", connectionString)
        Dim ds = New DataSet()
        Dim tableName As String = "excelData"
        adapter.Fill(ds, tableName)
        Dim data As DataTable = ds.Tables(tableName)
        Return data
    End Function
End Class