Public Class frmLogin

#Region "Entrar"
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        'If UsernameTextBox.Text <> "" And PasswordTextBox.Text <> "" Then
        '    Try
        '        Conectar()
        '        Iniciar()
        '        Comandar("SELECT USUARIO FROM TB_USUARIOS WHERE USUARIO = '" & UsernameTextBox.Text & "' AND SENHA = '" & PasswordTextBox.Text & "'")
        '        Ler()
        '        If leitura.HasRows Then
        frmInicial.Show()
        Me.Hide()
        '        Fechar()
        '    Else
        '        Fechar()
        '        MessageBox.Show("Usuário ou senha incorretos!", "Login", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
        '        UsernameTextBox.Focus()
        '        UsernameTextBox.SelectAll()
        '        PasswordTextBox.Clear()
        '    End If
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        '    End Try
        'End If
    End Sub
#End Region

#Region "Fechar/Cancelar"
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        If MessageBox.Show("Deseja sair do sistema?", "Sair", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes = True Then
            Me.Close()
        Else
            UsernameTextBox.Focus()
            UsernameTextBox.SelectAll()
        End If
    End Sub
#End Region

End Class
