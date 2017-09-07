Imports Microsoft.VisualBasic
Public Class FrmClosingPeriode
    Dim periode As String = PeriodeJalan()
    Private Sub FrmClosingPeriode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim date1 As Date = Now
        Me.Text = "CLOSING PERIODE & OPNAME"
        LoadPeriode()
        LabelControl3.Text = PeriodeJalan()
        If date1.Month > CInt(Microsoft.VisualBasic.Right(PeriodeJalan, 2)) Then
            MsgBox("PERIODE SAAT INI ANDA BELUM MELAKUKAN CLOSING", vbCritical, "Closing Bulanan")
        End If
        ProgressBarControl1.Properties.Maximum = 100
        ProgressBarControl1.EditValue = 0
    End Sub


    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        Opname()
    End Sub

    Private Sub Opname()
        ' Initializing progress bar properties
        Dim x As Integer = 100
        ProgressBarControl1.Properties.Step = 1
        ProgressBarControl1.Properties.PercentView = True
        ProgressBarControl1.Properties.Minimum = 0
        ProgressBarControl1.Properties.Maximum = x
        ProgressBarControl1.EditValue = 0
        'CREATE DATA BARU DI T_STOCK 
        'PER ITEM JENIS BARANG
        'LOOP PER JENIS

        SQL = "SELECT JENIS FROM M_JENIS"
        Dim DTJenis As DataTable = ExecuteQuery(SQL)
        Dim n As Integer = DTJenis.Rows.Count
        Dim i As Integer = 0
        If n = 0 Then
            Exit Sub  'JIKA TIDAK ADA JENIS BARANG KELUAR SUB
        Else
            For i = 0 To n - 1 'LOOP PER JENIS AMBIL ITEM
                Dim nJenis As String = DTJenis.Rows(i).Item("JENIS").ToString()
                'SQL = "SELECT ITEM_C ,ITEM,JENIS,UNIT FROM M_ITEM WHERE JENIS='" & nJenis & "' "
                SQL = "SELECT A.ITEM_C,A.JENIS ,B.ITEM_C " +
                    " FROM M_ITEM A " +
                    " LEFT JOIN T_STOCK B On A.ITEM_C=B.ITEM_C And A.JENIS=B.JENIS " +
                    " WHERE B.ITEM_C Is NULL And A.JENIS='" & nJenis & "' "
                Dim DTSS As DataTable = ExecuteQuery(SQL)
                Dim ii As Integer
                Dim m As Integer = DTSS.Rows.Count
                If m > 0 Then
                    ProgressBarControl1.Properties.Maximum = m
                    ProgressBarControl1.Properties.Minimum = 0
                    ProgressBarControl1.EditValue = 0
                    For ii = 0 To m - 1
                        ProgressBarControl1.PerformStep()
                        ProgressBarControl1.Update()

                        Dim nItem As String = DTSS.Rows(ii).Item("ITEM_C").ToString
                        Dim STAwal As String = 0 'AMBIL DARI BLN SEBELUM NYA 
                        Dim Adjs As String = 0
                        Dim nMin As String = 0
                        'ISI DATA STOCK BULAN BERJALAN
                        SQL = "SELECT * FROM T_STOCK WHERE ITEM_C ='" & nItem & "' AND PERIODE='" & periode & "'"
                        If CheckRecord(SQL) > 0 Then
                            'BUAT BARU UNTUK BLN KE DEPAN
                            'Dim Bln As String = Microsoft.VisualBasic.Right(PeriodeJalan, 2)
                            'If Bln = "12" Then Bln = "01"

                            'Dim Th As String = Microsoft.VisualBasic.Left(PeriodeJalan, 4)
                            'periode = Th & Bln

                            'SQL = "INSERT INTO T_STOCK (PERIODE,JENIS,ITEM_C,STOCK_AWAL,ADJUST,MIN,CLOSING) VALUES " +
                            '   " ('" & periode & "','" & nJenis & "','" & nItem & "','" & STAwal & "','" & Adjs & "','" & nMin & "','N')"
                            'ExecuteNonQuery(SQL)
                            'ambil Stok bln sebelumnya
                            'update di bln BERJALAN
                        Else
                            'INSERT DI PERIODE BLN BERJALAN UNTUK MENAMBAL DATA YANG ADA DI PERTENGAHAN BLN
                            SQL = "INSERT INTO T_STOCK (PERIODE,JENIS,ITEM_C,STOCK_AWAL,ADJUST,MIN,CLOSING) VALUES " +
                                " ('" & periode & "','" & nJenis & "','" & nItem & "','" & STAwal & "','" & Adjs & "','" & nMin & "','N')"
                            ExecuteNonQuery(SQL)

                        End If

                    Next
                End If
            Next
            ProgressBarControl1.EditValue = 0
            ProgressBarControl1.PerformStep()
            ProgressBarControl1.Update()
            MsgBox("Progress Complated", vbInformation, "OPNAME")
        End If

    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        ClosePeriode()
        'OPEN PERIODE
        'OPNAME
        GetMaxTag()
        MsgBox(GetTAG("AIM"))
    End Sub
    Private Sub LoadPeriode()
        SQL = "select periode from t_closing order by periode desc limit 10"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit1, False)
        ComboBoxEdit1.SelectedIndex = 0
    End Sub

    Private Sub OpenPeriode(ByVal Period As String)
        SQL = "UPDATE T_CLOSING SET CLOSING='N' ,PROSES=CURDATE() WHERE PERIODE='" & Period & "' AND CLOSING='Y'"
        ExecuteNonQuery(SQL)
        LabelControl3.Text = PeriodeJalan()
    End Sub
    Private Sub ClosePeriode()

        SQL = "UPDATE T_CLOSING SET CLOSING='Y' ,PROSES=CURDATE() WHERE PERIODE='" & periode & "' AND CLOSING='N'"
        ExecuteNonQuery(SQL)
        LabelControl3.Text = PeriodeJalan()
        MsgBox("PERIODE BERHASIL DI CLOSE...!!!", vbInformation, "PERIODE BERJALAN")
    End Sub

    Private Sub SimpleButton8_Click(sender As Object, e As EventArgs) Handles SimpleButton8.Click
        'close
        Me.Close()
    End Sub

    Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click
        'open Periode
        If Not IsEmptyCombo({ComboBoxEdit1}) Then
            OpenPeriode(ComboBoxEdit1.Text)
            LabelControl3.Text = PeriodeJalan()
            MsgBox("PERIODE BERHASIL DI BUKA...!!!", vbInformation, "PERIODE BERJALAN")

        End If
    End Sub
End Class