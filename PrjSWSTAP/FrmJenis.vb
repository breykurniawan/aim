
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraEditors.Repository

Imports Devart.Data
Imports Devart.Data.Oracle
Imports Devart.Common
Public Class FrmJenis
    Private Sub FrmJenis_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "JENIS ITEM"
        GridHeader()
        LoadView()
        ClearInput()
    End Sub
    Private Sub ClearInput()
        TextEdit1.Text = ""
        TextEdit2.Text = ""

        SimpleButton1.Enabled = True 'add
        SimpleButton2.Enabled = False 'save
        SimpleButton3.Enabled = False 'delete
    End Sub
    Private Sub GridHeader()
        Dim view As ColumnView = CType(GridControl1.MainView, ColumnView)
        Dim fieldNames() As String = New String() {"JENIS", "JENIS_DESC"}
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
        TextEdit1.Select()

        SimpleButton1.Enabled = False 'add
        SimpleButton2.Enabled = True 'Save
        SimpleButton3.Enabled = False 'delete
    End Sub

    Private Sub UnlockAll()
        TextEdit1.Enabled = True
        TextEdit2.Enabled = True

    End Sub

    Private Sub LockAll()
        TextEdit1.Enabled = False
        TextEdit2.Enabled = False
    End Sub

    Private Sub LoadView()
        SQL = "select JENIS ,JENIS_DESC from M_JENIS "
        FILLGridView(SQL, GridControl1)
    End Sub
    Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click
        'del
        SQL = "DELETE FROM M_JENIS WHERE JENIS='" & TextEdit1.Text & "'"
        ExecuteNonQuery(SQL)
        LoadView()
        MsgBox("DELETE SUCCESSFUL", vbInformation, "JENIS")
    End Sub

    Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click
        'CANCEL
        ClearInput()
        SimpleButton2.Text = "Save" 'save
    End Sub

    Private Sub FrmMaterialType_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        PanelControl1.Height = Me.Height - 150
    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        If Not IsEmptyText({TextEdit1, TextEdit2}) = True Then
            Dim TCode As String = TextEdit1.Text
            Dim TDesc As String = TextEdit2.Text
            SQL = "SELECT * FROM M_JENIS WHERE JENIS='" & TCode & "'"
            If CheckRecord(SQL) = 0 Then
                'INSERT
                SQL = "INSERT INTO M_JENIS (JENIS,JENIS_DESC) VALUES ('" & TCode & "','" & TDesc & "')"
                ExecuteNonQuery(SQL)
                LoadView()
                MsgBox("SAVE SUCCESSFUL", vbInformation, "JENIS")
            Else
                'UPDATE
                If UCase(SimpleButton2.Text) = "UPDATE" Then
                    SQL = "UPDATE M_JENIS SET JENIS_DESC='" & TDesc & "' WHERE JENIS='" & TCode & "'"
                    ExecuteNonQuery(SQL)
                    LoadView()
                    MsgBox("UPDATE SUCCESSFUL", vbInformation, "JENIS")
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
            TextEdit1.Text = GridView1.GetRowCellValue(e.RowHandle, "JENIS").ToString() 'MATERIAL TYPE CODE
            TextEdit2.Text = GridView1.GetRowCellValue(e.RowHandle, "JENIS_DESC").ToString() 'MATERIAL TYPE

            TextEdit1.Enabled = False

        End If
    End Sub
End Class