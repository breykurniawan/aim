Imports System.IO

Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraEditors.Repository

Imports Devart.Data
Imports Devart.Data.Oracle
Imports Devart.Common

Imports System.Data.Odbc
Public Class FrmUserProfile
    Dim imagename As String

    Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click
        'DELETE
        SQL = "UPDATE T_USERPROFILE SET AKTIF='N' WHERE USERID='" & TextEdit1.Text & "'"
        ExecuteNonQuery(SQL)
        LoadUser()
        MsgBox("Delete Successful", vbInformation, "USER ROLE")
    End Sub
    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        'SAVE
        Dim UID As String = TextEdit1.Text
        Dim UNAME As String = TextEdit2.Text
        Dim EMAIL As String = TextEdit3.Text
        Dim ROLEID As String = GetCodeRole(ComboBoxEdit1.Text)
        Dim INPUT_BY As String = USERNAME
        Dim INPUT_DATE As String = Now
        Dim UPDATE_BY As String = USERNAME
        Dim UPDATE_DATE As String = Now
        Dim PASSWD As String = TextEdit4.Text
        'Try
        SQL = "SELECT * FROM T_USERPROFILE WHERE AKTIF='Y' AND USERID='" & TextEdit1.Text & "'"
        If CheckRecord(SQL) = 0 Then
            Try
                SQL = "insert into T_USERPROFILE(USERID,USERNAME,PASSWD,EMAIL,ROLEID,INPUT_BY,AKTIF)" +
                "VALUES ('" & UID & "','" & UNAME & "','" & PASSWD & "','" & EMAIL & "','" & ROLEID & "','" & INPUT_BY & "','Y' )"
                ExecuteNonQuery(SQL)
                SQL = "SELECT * FROM T_USERPROFILE WHERE AKTIF='Y' AND USERID='" & TextEdit1.Text & "'"
                If CheckRecord(SQL) > 0 Then
                    MsgBox("Save Successful", vbInformation, "User Profile")
                    UpdateCode("EM")
                    LoadUser()
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            Try
                Dim fls As FileStream
                fls = New FileStream(imagename, FileMode.Open, FileAccess.Read)
                Dim blob As Byte() = New Byte(fls.Length - 1) {}
                fls.Read(blob, 0, System.Convert.ToInt32(fls.Length))
                fls.Close()

                SQL = "UPDATE T_USERPROFILE SET USERNAME= '" & UNAME & "',EMAIL='" & EMAIL & "',ROLEID='" & ROLEID & "',UPDATE_BY='" & UPDATE_BY & "' " +
                " WHERE USERID='" & TextEdit1.Text & "'"
                ExecuteNonQuery(SQL)
                LoadUser()
                MsgBox("Update Successful", vbInformation, "User Profile")
            Catch EX As Exception
                MessageBox.Show(EX.Message)
            End Try
        End If
    End Sub
    Private Sub SimpleButton5_Click(sender As Object, e As EventArgs) Handles SimpleButton5.Click
        'CLOSE
        Me.Close()
    End Sub
    Private Sub GridHeader()
        Dim View As ColumnView = CType(GridControl1.MainView, ColumnView)
        Dim FieldNames() As String = New String() {"USERID", "USERNAME", "EMAIL", "ROLENAME"}
        Dim I As Integer
        Dim Column As DevExpress.XtraGrid.Columns.GridColumn

        View.Columns.Clear()
        For I = 0 To FieldNames.Length - 1
            Column = View.Columns.AddField(FieldNames(I))
            Column.VisibleIndex = I
        Next

        'GROUPING
        Dim GridView As GridView = CType(GridControl1.FocusedView, GridView)
        GridView.SortInfo.ClearAndAddRange(New GridColumnSortInfo() {
        New GridColumnSortInfo(GridView.Columns("ROLENAME"), DevExpress.Data.ColumnSortOrder.Ascending)}, 1)
        GridView.BestFitColumns()
        GridView.ExpandAllGroups()
    End Sub

    Private Sub FrmUserProfile_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "USER PROFILE"
        LoadRole()
        GridHeader()
        LoadUser()
        LockAll()
    End Sub
    Private Sub LoadRole()
        SQL = "select ROLENAME from t_role WHERE AKTIF='Y' ORDER BY ROLEID"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit1, False)
    End Sub

    Private Sub LoadUser()
        SQL = "SELECT USERID,USERNAME,EMAIL,ROLENAME,IMAGE " +
        "FROM T_USERPROFILE A " +
        "LEFT JOIN T_ROLE B On A.ROLEID=B.ROLEID And B.AKTIF='Y' " +
        "WHERE A.AKTIF='Y' " +
        "ORDER BY USERID"
        FILLGridView(SQL, GridControl1)
        GridControl1.DataSource = ExecuteQuery(SQL)
        Dim GridView As GridView = CType(GridControl1.FocusedView, GridView)
        GridView.ExpandAllGroups()
    End Sub
    Private Sub LockAll()
        TextEdit1.Enabled = False
        TextEdit2.Enabled = False
        TextEdit3.Enabled = False
        ComboBoxEdit1.Enabled = False
        SimpleButton1.Enabled = True 'add
        SimpleButton2.Enabled = False 'save
        SimpleButton3.Enabled = False 'del
    End Sub
    Private Sub UnLockAll()
        TextEdit1.Enabled = True
        TextEdit2.Enabled = True
        TextEdit3.Enabled = True
        ComboBoxEdit1.Enabled = True
        TextEdit2.Select()
        SimpleButton1.Enabled = False 'add
        SimpleButton2.Enabled = True 'save
        SimpleButton3.Enabled = True 'del
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        'ADD
        UnLockAll()
        'TextEdit1.Text = Val(Strings.Right(GetCode("EM"), 2))
        TextEdit1.Text = GetUserMaxID()
        TextEdit1.Enabled = False
    End Sub

    Private Sub FrmUserProfile_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        PanelControl1.Height = Me.Height - 250
    End Sub

    Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click
        'CANCEL
        TextEdit1.Text = ""
        TextEdit2.Text = ""
        TextEdit3.Text = ""
        ComboBoxEdit1.Text = ""
        LockAll()
        SimpleButton2.Text = "Save" 'SAVE
    End Sub

    Private Sub GridView1_RowCellClick(sender As Object, e As RowCellClickEventArgs) Handles GridView1.RowCellClick
        If e.RowHandle < 0 Then
            SimpleButton1.Enabled = True 'add
            SimpleButton2.Enabled = False 'save
            SimpleButton3.Enabled = False 'del
        Else
            SimpleButton1.Enabled = False 'ADD
            SimpleButton2.Enabled = True 'SAVE
            SimpleButton3.Enabled = True 'DEL

            SimpleButton2.Text = "Update" 'SAVE

            TextEdit1.Text = GridView1.GetRowCellValue(e.RowHandle, "USERID").ToString()  'ID
            TextEdit2.Text = GridView1.GetRowCellValue(e.RowHandle, "USERNAME").ToString() 'NAME
            TextEdit3.Text = GridView1.GetRowCellValue(e.RowHandle, "EMAIL").ToString() 'NAME
            ComboBoxEdit1.Text = GridView1.GetRowCellValue(e.RowHandle, "ROLENAME").ToString() 'NAME
            UnLockAll()
        End If
    End Sub


End Class