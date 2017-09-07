Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraEditors.Repository

Imports Devart.Data
Imports Devart.Data.Oracle
Imports Devart.Common
Public Class FrmStock
    Dim FPeriode As String = ""
    Dim FJenis As String = ""
    Dim FItem As String = ""

    Private Sub FrmStock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "STOCK ITEM"
        ComboBoxEdit3.Text = Format(Now, "yyyy/MM")
        LoadPeriode()
        LoadJenis()
        LoadItem(ComboBoxEdit1.Text)
        GridHeader()
        LoadView()
        'ClearInput()
    End Sub
    Private Sub LoadJenis()
        SQL = "Select JENIS from m_jenis"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit1, False)
        ComboBoxEdit1.SelectedIndex = 0
    End Sub
    Private Sub LoadPeriode()
        SQL = "select PERIODE from T_CLOSING ORDER BY PERIODE DESC"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit3, False)

    End Sub
    Private Sub LoadItem(ByVal JENIS As String)
        SQL = "select ITEM from m_item where jenis LIKE '%" & JENIS & "%'"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit2, False)
    End Sub


    Private Sub ClearInput()
        ComboBoxEdit3.Text = ""
        ComboBoxEdit1.Text = ""
        ComboBoxEdit2.Text = ""
        TextEdit1.Text = ""
        TextEdit2.Text = ""
        TextEdit3.Text = ""

        SimpleButton1.Enabled = True 'add
        SimpleButton2.Enabled = False 'save
        SimpleButton3.Enabled = False 'delete
    End Sub
    Private Sub GridHeader()
        Dim view As ColumnView = CType(GridControl1.MainView, ColumnView)
        Dim fieldNames() As String = New String() {"PERIODE", "JENIS", "ITEM", "STOCK_AWAL", "ADJUST", "MIN"}
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
        LoadView()
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
        FPeriode = ComboBoxEdit3.Text
        FJenis = ComboBoxEdit1.Text
        FItem = ComboBoxEdit2.Text

        SQL = "SELECT PERIODE,B.JENIS,B.ITEM,STOCK_AWAL,ADJUST,MIN  " +
            " From T_STOCK A" +
            " LEFT JOIN M_ITEM B ON A.ITEM_C=B.ITEM_C AND A.JENIS=B.JENIS" +
            " Where PERIODE LIKE '%" & FPeriode & "%'" +
            " AND A.JENIS LIKE '%" & FJenis & "%'" +
            " AND ITEM LIKE '%" & FItem & "%'"
        FILLGridView(SQL, GridControl1)
    End Sub

    Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click
        'CANCEL
        ClearInput()
        SimpleButton2.Text = "Save" 'save
    End Sub

    Private Sub FrmMaterialType_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        PanelControl1.Height = Me.Height - 230
    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        If Not IsEmptyText({TextEdit1, TextEdit2}) = True Then
            Dim TPeriode As String = ComboBoxEdit3.Text
            Dim TJenis As String = ComboBoxEdit1.Text
            Dim TItem As String = ComboBoxEdit2.Text
            Dim TStockAwal As String = TextEdit1.Text
            Dim TAdjust As String = TextEdit2.Text
            Dim TMin As String = TextEdit3.Text

            SQL = "SELECT * From T_STOCK " +
            " Where PERIODE='" & TPeriode & "'" +
            " AND JENIS ='" & TJenis & "'" +
            " AND ITEM LIKE '%" & TItem & "%'"

            If CheckRecord(SQL) > 0 Then
                'UPDATE
                If UCase(SimpleButton2.Text) = "UPDATE" Then
                    SQL = "UPDATE T_STOCK SET STOCK_AWAL='" & TStockAwal & "',ADJUST='" & TAdjust & "',MIN='" & TMin & "' " +
                        " Where PERIODE='" & TPeriode & "'" +
                        " AND JENIS ='" & TJenis & "'" +
                        " AND ITEM LIKE '%" & TItem & "%'"
                    ExecuteNonQuery(SQL)
                    LoadView()
                    MsgBox("UPDATE SUCCESSFUL", vbInformation, "UNIT")
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
            ComboBoxEdit3.Text = GridView1.GetRowCellValue(e.RowHandle, "PERIODE").ToString() 'CITY
            ComboBoxEdit1.Text = GridView1.GetRowCellValue(e.RowHandle, "JENIS").ToString() 'STATE
            ComboBoxEdit1.Text = GridView1.GetRowCellValue(e.RowHandle, "ITEM").ToString() 'PHONE
            TextEdit1.Text = GridView1.GetRowCellValue(e.RowHandle, "STOCK_AWAL").ToString() 'CODE
            TextEdit2.Text = GridView1.GetRowCellValue(e.RowHandle, "ADJUST").ToString() 'NAME
            TextEdit3.Text = GridView1.GetRowCellValue(e.RowHandle, "MIN").ToString() 'ADDRESS

            TextEdit1.Enabled = False

        End If
    End Sub

    Private Sub ComboBoxEdit1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit1.SelectedIndexChanged
        LoadItem(ComboBoxEdit1.Text)
    End Sub
End Class