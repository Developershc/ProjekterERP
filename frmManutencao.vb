Public Class frmManutencao

    Dim visualizatabelan, incluitabelan, modificatabelan, excluitabelan, visualizamenun, incluimenun, modificamenun, excluimenun As Boolean

#Region "Alteração de Senha"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox2.Text <> TextBox3.Text Then
            MessageBox.Show("As senhas não coincidem!", "Manutenção de usuários", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox2.Focus()
            TextBox2.SelectAll()
            TextBox3.Clear()
        Else
            If TextBox1.Text = TextBox2.Text Then
                MessageBox.Show("A nova senha é igual à atual! Digite uma senha diferente!", "Manutenção de usuários", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Focus()
                TextBox1.SelectAll()
                TextBox2.Clear()
                TextBox3.Clear()
            Else
                Try
                    Conectar()
                    Iniciar()
                    Comandar("UPDATE TB_USUARIOS SET SENHA = '" & TextBox3.Text & "' WHERE USUARIO = '" & TextBox9.Text & "'")
                    Executar()
                    Fechar()
                    MessageBox.Show("Senha alterada com sucesso!", "Manutenção de usuários", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    TextBox1.Clear()
                    TextBox2.Clear()
                    TextBox3.Clear()
                    TextBox9.Clear()
                    TextBox9.Focus()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub
#End Region

#Region "Carregar ListBoxes"
    Private Sub frmManutencao_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEGRUPO FROM TB_GRUPOS")
            Ler()
            If leitura.HasRows Then
                Try
                    Conectar()
                    Iniciar()
                    Comandar("SELECT NOMEGRUPO FROM TB_GRUPOS ORDER BY NOMEGRUPO")
                    Ler()
                    While leitura.Read
                        ListBox1.Items.Add(leitura("NOMEGRUPO"))
                    End While
                    Fechar()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                ListBox1.Items.Add("Não há grupos cadastrados!")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("")
        'Ler()
        'While leitura.Read
        'list2, tabelas banco, verificar
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try

        'Try
        'list3, menus, verificar
        'Catch ex As Exception
        'End Try

        Try
            Conectar()
            Iniciar()
            Comandar("SELECT USUARIO FROM TB_USUARIOS")
            Ler()
            If leitura.HasRows Then
                Try
                    Conectar()
                    Iniciar()
                    Comandar("SELECT USUARIO FROM TB_USUARIOS ORDER BY USUARIO")
                    Ler()
                    While leitura.Read
                        ListBox4.Items.Add(leitura("USUARIO"))
                    End While
                    Fechar()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                ListBox4.Items.Add("Não há usuários cadastrados!")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Sub
#End Region 'VERIFICAR TABELAS E MENUS

#Region "Ao Selecionar Tabelas"
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        'If ListBox2.SelectedIndex = "tbN" Then

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT VISUALIZATABELAN, INCLUITABELAN, MODIFICATABELAN, EXCLUITABELAN FROM TBGRUPOS WHERE NOME = '" & ListBox1.SelectedIndex & "'")
        'Ler()
        'visualizatabelan = leitura("VISUALIZATABELAN")
        'incluitabelan = leitura("INCLUITABELAN")
        'modificatabelan = leitura("MODIFICATABELAN")
        'excluitabelan = leitura("EXCLUITABELAN")
        'If visualizatabelan = True Then
        'CheckBox1.Checked = True
        'Else
        'CheckBox1.Checked = False
        'End If
        'If incluitabelan = True Then
        'CheckBox2.Checked = True
        'Else
        'CheckBox2.Checked = False
        'End If
        'If modificatabelan = True Then
        'CheckBox3.Checked = True
        'Else
        'CheckBox3.Checked = False
        'End If
        'If excluitabelan = True Then
        'CheckBox4.Checked = True
        'Else
        'CheckBox4.Checked = False
        'End If
        'Fechar()
        'Catch ex As Exception

        'End Try
        'End If
    End Sub
#End Region 'VERIFICAR E TERMINAR CÓDIGO DE BANCO

#Region "Ao Selecionar Menus"
    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        'If ListBox3.SelectedIndex = "MenuN" Then
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT VISUALIZAMENUN, INCLUIMENUN, MODIFICAMENUN, EXCLUIMENUN FROM TBGRUPOS WHERE NOME = '" & ListBox1.SelectedIndex & "'")
        'Ler()
        'visualizamenun = leitura("VISUALIZAMENUN")
        'incluimenun = leitura("INCLUIMENUN")
        'modificamenun = leitura("MODIFICAMENUN")
        'excluimenun = leitura("EXCLUIMENUN")
        'If visualizamenun = True Then
        'CheckBox5.Checked = True
        'Else
        'CheckBox5.Checked = False
        'End If
        'If incluimenun = True Then
        'CheckBox6.Checked = True
        'Else
        'CheckBox6.Checked = False
        'End If
        'If modificamenun = True Then
        'CheckBox7.Checked = True
        'Else
        'CheckBox7.Checked = False
        'End If
        'If excluimenun = True Then
        'CheckBox8.Checked = True
        'Else
        'CheckBox8.Checked = False
        'End If
        'Fechar()
        'Catch ex As Exception

        'End Try
        'End If
    End Sub
#End Region 'VERIFICAR E TERMINAR CÓDIGO DE BANCO

#Region "Alteração de Usuários"
    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        'TextBox5.ReadOnly = True
        'TextBox6.ReadOnly = True
        'TextBox7.ReadOnly = True
        'TextBox8.ReadOnly = True
        'Try
        '    Conectar()
        '    Iniciar()
        '    Comandar("SELECT NOMEGRUPO, USUARIO, SENHA, OBSERVACAO FROM TB_USUARIOS WHERE USUARIO = '" & ListBox4.SelectedIndex & "'")
        '    Ler()
        '    While leitura.Read
        '        ComboBox1.Text = leitura("NOMEGRUPO")
        '        TextBox5.Text = leitura("USUARIO")
        '        TextBox6.Text = leitura("SENHA")
        '        TextBox7.Text = leitura("SENHA")
        '        TextBox8.Text = leitura("OBSERVACAO")
        '    End While
        '    Fechar()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
        'RemoveHandler Button10.Click, AddressOf Button10_Click
        'AddHandler Button10.Click, AddressOf AlterarDados
    End Sub
#End Region 'VERIFICAR

#Region "Alterar Dados"
    Private Sub AlterarDados()
        'Try
        '    Conectar()
        '    Iniciar() '
        '    Comandar("UPDATE TB_USUARIOS SET USUARIO = '" & ListBox4.SelectedIndex & "' WHERE USUARIO = '" & ListBox4.SelectedIndex & "'")
        '    Executar()
        '    Fechar()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'VERIFICAR

#Region "Botão Verde - Usuários"
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If TextBox6.Text <> TextBox7.Text Then
            MessageBox.Show("As senhas não coincidem!", "Usuários", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox6.Focus()
            TextBox6.SelectAll()
            TextBox7.Clear()
        Else
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT USUARIO FROM TB_USUARIOS WHERE USUARIO = '" & TextBox5.Text & "'")
                Ler()
                If leitura.HasRows Then
                    MessageBox.Show("Usuário já cadastrado!", "Manutenção de usuários", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                    TextBox5.Focus()
                    TextBox5.SelectAll()
                    TextBox6.Clear()
                    TextBox7.Clear()
                    TextBox8.Clear()
                Else
                    Try
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_USUARIOS (NOMEGRUPO, USUARIO, SENHA, OBSERVACAO) VALUES ('" & ComboBox1.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Cadastrado com sucesso!", "Usuários", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        For Each ctrl In Me.Controls
                            If TypeOf ctrl Is TextBox Then
                                ctrl.Clear()
                            End If
                        Next
                        ComboBox1.Text = ""
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub
#End Region

#Region "TextBoxes editáveis"
    'Private Sub TextBox5_Click(sender As Object, e As EventArgs) Handles TextBox5.Click
    '    TextBox5.ReadOnly = False
    'End Sub

    'Private Sub TextBox6_Click(sender As Object, e As EventArgs) Handles TextBox6.Click
    '    TextBox6.ReadOnly = False
    'End Sub

    'Private Sub TextBox7_Click(sender As Object, e As EventArgs) Handles TextBox7.Click
    '    TextBox7.ReadOnly = False
    'End Sub

    'Private Sub TextBox8_Click(sender As Object, e As EventArgs) Handles TextBox8.Click
    '    TextBox8.ReadOnly = False
    'End Sub
#End Region 'VERIFICAR

#Region "Botão Vermelho - Usuários"
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'frmMessageBox.Text = "Usuários"
        'frmMessageBox.Button1.Text = "Excluir"
        'frmMessageBox.Button2.Text = "Cancelar"
        'frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        'frmMessageBox.Show()
    End Sub
#End Region 'VERIFICAR

End Class