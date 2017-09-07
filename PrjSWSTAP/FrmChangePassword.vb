Public Class FrmChangePassword
    Private Sub FrmChangePassword_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "CHANGE PASSWORD"
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        'save
        If Not IsEmptyText({TextEdit1, TextEdit2, TextEdit3}) Then
            If Len(TextEdit2.Text) < 7 Or Len(TextEdit3.Text) < 7 Then
                MsgBox("Long password min 8 Char", vbInformation, "Change Password")
                Exit Sub

            Else
                If CheckLogin(USERNAME, TextEdit1.Text) = False Then MsgBox("Old Password Wrong, Please Try Again.", vbInformation, "Change Password") : Exit Sub

                If TextEdit1.Text = TextEdit2.Text Then
                    MsgBox("New password must be different", vbInformation, "Change Password")
                Else
                    If TextEdit2.Text <> TextEdit3.Text Then

                        MsgBox("The new password / confirmation must be the same", vbInformation, "Change Password")
                    Else
                        'update password
                        SQL = "UPDATE  T_USERPROFILE  SET PASSWD ='" & TextEdit3.Text & "' WHERE USERNAME='" & USERNAME & "'"
                        ExecuteNonQuery(SQL)
                        TextEdit1.Text = ""
                        TextEdit2.Text = ""
                        TextEdit3.Text = ""
                        MsgBox("Successfully Changed", vbInformation, "Change Password")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        'close
        Me.Close()
    End Sub
End Class