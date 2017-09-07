Imports DevExpress.XtraGrid.Views.Grid
Public Class FrmPenerimaanBarang
    Private Sub FrmPenerimaanBarang_Load(sender As Object, e As EventArgs) Handles Me.Load
        LockAll()
        LockAllDetail()
        SimpleButton6.Enabled = False
        SimpleButton7.Enabled = False
        SimpleButton8.Enabled = False
        Dim Item = FrmMain.NavBarControl1.Cursor.ToString
        If Item = "PENERIMAAN BARANG" Then
            Me.Text = "PENERIMAAN BARANG"

        End If
        loadview()

    End Sub
    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        UnlockAll()
        SimpleButton6.Enabled = True
        SimpleButton7.Enabled = True
        SimpleButton8.Enabled = True
        LockAllDetail()
        Dim nKode As String
        nKode = "MP"
        TextEdit1.Text = GetCode(nKode)
        TextEdit1.Enabled = False
    End Sub

    Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click
        LockAll()
        LockAllDetail()
        SimpleButton6.Enabled = False
        SimpleButton7.Enabled = False
        SimpleButton8.Enabled = False
    End Sub

    Private Sub SimpleButton5_Click(sender As Object, e As EventArgs) Handles SimpleButton5.Click
        Me.Close()
    End Sub

    Private Sub loadview()
        SQL = "select TRSMG_C,TRSMG_T,REFF,NO_SJ,NO_FAK,NO_PO,TRSMG_DATE,TRANSPORTER_C,VENDOR_C,CUSTOMER_C,NO_VEHICLE,PENERIMA,PENGEIRIM,DEP_C from trsmg"
        FILLGridView(SQL, GridControl1)
    End Sub
    Private Sub LoadView1()

        'SQL = "select ITEM_C,QTY,UNIT,LOKASI,BATCH,TAGNO from trsmg_detail "
        'FILLGridView(SQL, GridControl2)


    End Sub
    Private Sub LockAll()
        TextEdit1.Text = ""
        TextEdit2.Text = ""
        TextEdit3.Text = ""
        TextEdit4.Text = ""
        TextEdit5.Text = ""
        TextEdit6.Text = ""
        TextEdit7.Text = ""
        TextEdit8.Text = ""
        TextEdit9.Text = ""
        TextEdit10.Text = ""
        TextEdit11.Text = ""
        TextEdit11.Text = ""
        TextEdit12.Text = ""
        TextEdit13.Text = ""
        TextEdit14.Text = ""
        TextEdit15.Text = ""
        TextEdit16.Text = ""
        TextEdit17.Text = ""
        TextEdit18.Text = ""
        TextEdit19.Text = ""
        TextEdit21.Text = ""
        TextEdit1.Enabled = False
        TextEdit2.Enabled = False
        TextEdit3.Enabled = False
        TextEdit4.Enabled = False
        TextEdit5.Enabled = False
        TextEdit6.Enabled = False
        TextEdit7.Enabled = False
        TextEdit8.Enabled = False
        TextEdit9.Enabled = False
        TextEdit10.Enabled = False
        TextEdit11.Enabled = False
        TextEdit11.Enabled = False
        TextEdit12.Enabled = False
        TextEdit13.Enabled = False
        TextEdit14.Enabled = False

        SimpleButton1.Enabled = True 'add
        SimpleButton2.Enabled = False 'save
        SimpleButton3.Enabled = False 'del


    End Sub

    Private Sub LockAllDetail()
        TextEdit15.Enabled = False
        TextEdit16.Enabled = False
        TextEdit17.Enabled = False
        TextEdit18.Enabled = False
        TextEdit19.Enabled = False
        TextEdit21.Enabled = False

    End Sub
    Private Sub UnlockAll()
        TextEdit1.Enabled = True
        TextEdit2.Enabled = True
        TextEdit3.Enabled = True
        TextEdit4.Enabled = True
        TextEdit5.Enabled = True
        TextEdit6.Enabled = True
        TextEdit7.Enabled = True
        TextEdit8.Enabled = True
        TextEdit9.Enabled = True
        TextEdit10.Enabled = True
        TextEdit11.Enabled = True
        TextEdit11.Enabled = True
        TextEdit12.Enabled = True
        TextEdit13.Enabled = True
        TextEdit14.Enabled = True

        SimpleButton1.Enabled = True 'add
        SimpleButton2.Enabled = True 'save
        SimpleButton3.Enabled = True 'del


    End Sub

    Private Sub UnlockAllDetail()
        TextEdit15.Enabled = True
        TextEdit16.Enabled = True
        TextEdit17.Enabled = True
        TextEdit18.Enabled = True
        TextEdit19.Enabled = True
        TextEdit21.Enabled = True
        SimpleButton6.Enabled = True 'cancel detail
        SimpleButton7.Enabled = True 'save detail
        SimpleButton8.Enabled = True 'add detail

    End Sub

    Private Sub SimpleButton8_Click(sender As Object, e As EventArgs) Handles SimpleButton8.Click
        UnlockAllDetail()
    End Sub

    Private Sub SimpleButton6_Click(sender As Object, e As EventArgs) Handles SimpleButton6.Click
        TextEdit15.Text = ""
        TextEdit16.Text = ""
        TextEdit17.Text = ""
        TextEdit18.Text = ""
        TextEdit19.Text = ""
        TextEdit21.Text = ""
        SimpleButton7.Text = "Save Detail"
        GridView2.ClearDocument()
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
            TextEdit1.Text = GridView1.GetRowCellValue(e.RowHandle, "TRSMG_C").ToString() 'CODE
            TextEdit2.Text = GridView1.GetRowCellValue(e.RowHandle, "TRSMG_T").ToString() 'NAME
            TextEdit3.Text = GridView1.GetRowCellValue(e.RowHandle, "REFF").ToString() 'REFF
            TextEdit4.Text = GridView1.GetRowCellValue(e.RowHandle, "NO_SJ").ToString() 'SURATJALAN
            TextEdit5.Text = GridView1.GetRowCellValue(e.RowHandle, "NO_FAK").ToString() 'NO FAKTUR
            TextEdit6.Text = GridView1.GetRowCellValue(e.RowHandle, "NO_PO").ToString() 'NO PO
            TextEdit7.Text = GridView1.GetRowCellValue(e.RowHandle, "TRSMG_DATE").ToString() 'TGL MASUK
            TextEdit8.Text = GridView1.GetRowCellValue(e.RowHandle, "TRANSPORTER_C").ToString() 'TRANSPORTER CODE
            TextEdit9.Text = GridView1.GetRowCellValue(e.RowHandle, "VENDOR_C").ToString() 'VENDOR CODE
            TextEdit10.Text = GridView1.GetRowCellValue(e.RowHandle, "CUSTOMER_C").ToString() 'CUSTOMER CODE
            TextEdit11.Text = GridView1.GetRowCellValue(e.RowHandle, "NO_VEHICLE").ToString() 'NO KENDARAAN
            TextEdit12.Text = GridView1.GetRowCellValue(e.RowHandle, "PENERIMA").ToString() 'PENERIMA
            TextEdit13.Text = GridView1.GetRowCellValue(e.RowHandle, "PENGEIRIM").ToString() 'PENGIRIM
            TextEdit14.Text = GridView1.GetRowCellValue(e.RowHandle, "DEP_C").ToString() 'DEP CODE
            TextEdit1.Enabled = False
            UnlockAll()
            Dim TRSMG_C As String
            TRSMG_C = TextEdit1.Text
            SQL = "select ITEM_C,QTY,UNIT,LOKASI,BATCH,TAGNO from trsmg_detail WHERE TRSMG_C = '" & TRSMG_C & "' "
            FILLGridView(SQL, GridControl2)

        End If
    End Sub
    Private Sub GridView2_RowCellClick(sender As Object, e As RowCellClickEventArgs) Handles GridView2.RowCellClick
        If e.RowHandle < 0 Then
            SimpleButton8.Enabled = True 'add
            SimpleButton7.Enabled = False 'save
            SimpleButton6.Enabled = False 'delete
        Else
            SimpleButton8.Enabled = False 'add
            SimpleButton7.Enabled = True 'save
            SimpleButton6.Enabled = True 'delete



            SimpleButton7.Text = "Update Detail" 'save
            TextEdit15.Text = GridView2.GetRowCellValue(e.RowHandle, "ITEM_C").ToString() 'CODE
            TextEdit16.Text = GridView2.GetRowCellValue(e.RowHandle, "QTY").ToString() 'NAME
            TextEdit17.Text = GridView2.GetRowCellValue(e.RowHandle, "UNIT").ToString() 'REFF
            TextEdit18.Text = GridView2.GetRowCellValue(e.RowHandle, "BATCH").ToString() 'SURATJALAN
            TextEdit19.Text = GridView2.GetRowCellValue(e.RowHandle, "TAGNO").ToString() 'NO FAKTUR
            TextEdit21.Text = GridView2.GetRowCellValue(e.RowHandle, "LOKASI").ToString() 'TGL MASUK

            UnlockAllDetail()
        End If

    End Sub
End Class
