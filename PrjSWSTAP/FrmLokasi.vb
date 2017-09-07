
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraEditors.Repository

Imports Devart.Data
Imports Devart.Data.Oracle
Imports Devart.Common
Public Class FrmLokasi
    Private Sub FrmLokasi_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "LOKASI ITEM"
        SQL = "SELECT WH_C FROM M_WH"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit1, False)
        GridHeader()
        LoadView()
        ClearInput()
    End Sub
    Private Sub ClearInput()
        TextEdit1.Text = ""
        TextEdit2.Text = ""
        ComboBoxEdit1.Text = ""
        TextEdit4.Text = ""
        TextEdit5.Text = ""
        TextEdit6.Text = ""

        SimpleButton1.Enabled = True 'add
        SimpleButton2.Enabled = False 'save
        SimpleButton3.Enabled = False 'delete
    End Sub
    Private Sub GridHeader()
        Dim view As ColumnView = CType(GridControl1.MainView, ColumnView)
        Dim fieldNames() As String = New String() {"LOKASI", "LOKASI_DESC"}
        Dim I As Integer
        Dim Column As DevExpress.XtraGrid.Columns.GridColumn

        view.Columns.Clear()
        For I = 0 To fieldNames.Length - 1
            Column = view.Columns.AddField(fieldNames(I))
            Column.VisibleIndex = I
        Next
    End Sub

    Private Sub SimpleButton5_Click(sender As Object, e As EventArgs) Handles SimpleButton5.Click
        'close
        Me.Close()
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        'add
        UnlockAll()
        ClearInput()
        TextEdit2.Select()
        SimpleButton1.Enabled = False 'add
        SimpleButton2.Enabled = True 'Save
        SimpleButton3.Enabled = False 'delete
    End Sub

    Private Sub UnlockAll()
        TextEdit1.Enabled = True
        TextEdit2.Enabled = True
        ComboBoxEdit1.Enabled = True
        TextEdit4.Enabled = True
        TextEdit5.Enabled = True
        TextEdit6.Enabled = True
    End Sub

    Private Sub LockAll()
        TextEdit1.Enabled = False
        TextEdit2.Enabled = False
        ComboBoxEdit1.Enabled = False
        TextEdit4.Enabled = False
        TextEdit5.Enabled = False
        TextEdit6.Enabled = False
    End Sub

    Private Sub LoadView()
        SQL = "select LOKASI ,LOKASI_DESC from M_LOKASI "
        FILLGridView(SQL, GridControl1)
    End Sub
    Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click
        'del
        SQL = "DELETE FROM M_LOKASI WHERE LOKASI='" & TextEdit1.Text & "'"
        ExecuteNonQuery(SQL)
        LoadView()
        MsgBox("DELETE SUCCESSFUL", vbInformation, "LOKASI")
    End Sub

    Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click
        'CANCEL
        ClearInput()
        SimpleButton2.Text = "Save" 'save
    End Sub

    Private Sub FrmMaterialType_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        PanelControl1.Height = Me.Height - 165
    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click

        If Not IsEmptyCombo({ComboBoxEdit1}) = True Then
            If Not IsEmptyText({TextEdit2, TextEdit4, TextEdit5, TextEdit6}) = True Then
                TextEdit1.Text = Trim(ComboBoxEdit1.Text) & Trim(TextEdit4.Text) & Trim(TextEdit5.Text) & Trim(TextEdit6.Text)
                Dim TCode As String = TextEdit1.Text
                Dim TDesc As String = TextEdit2.Text
                SQL = "SELECT * FROM M_LOKASI WHERE LOKASI='" & TCode & "'"
                If CheckRecord(SQL) = 0 Then
                    'INSERT
                    SQL = "INSERT INTO M_LOKASI (LOKASI,LOKASI_DESC) VALUES ('" & TCode & "','" & TDesc & "')"
                    ExecuteNonQuery(SQL)
                    LoadView()
                    MsgBox("SAVE SUCCESSFUL", vbInformation, "LOKASI")
                Else
                    'UPDATE
                    If UCase(SimpleButton2.Text) = "UPDATE" Then
                        SQL = "UPDATE M_LOKASI SET LOKASI_DESC='" & TDesc & "' WHERE LOKASI='" & TCode & "'"
                        ExecuteNonQuery(SQL)
                        LoadView()
                        MsgBox("UPDATE SUCCESSFUL", vbInformation, "LOKASI")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub GridView1_RowCellClick(sender As Object, e As RowCellClickEventArgs) Handles GridView1.RowCellClick
        If e.RowHandle < 0 Then
            SimpleButton1.Enabled = True 'add
            SimpleButton2.Enabled = False 'save
            SimpleButton3.Enabled = False 'delete
        Else
            SimpleButton1.Enabled = False 'add
            SimpleButton2.Enabled = True 'save
            SimpleButton3.Enabled = True 'delete

            SimpleButton2.Text = "Update" 'save
            TextEdit1.Text = GridView1.GetRowCellValue(e.RowHandle, "LOKASI").ToString() 'MATERIAL TYPE CODE
            TextEdit2.Text = GridView1.GetRowCellValue(e.RowHandle, "LOKASI_DESC").ToString() 'MATERIAL TYPE

            TextEdit1.Enabled = False

        End If
    End Sub
End Class