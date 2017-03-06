Public Class frmSM

#Region "Carregar Listas"
    Private Sub frmSM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Me.Text = "Seção" Then
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMESECAO FROM TB_SECOES")
                Ler()
                Dim resultado As Boolean = leitura.HasRows
                If resultado = True Then
                    While leitura.Read
                        ListBox2.Items.Add(leitura("NOMESECAO"))
                    End While
                Else
                    'MessageBox.Show("Não há seções cadastradas!","Seção",MessageBoxButtons.OK,MessageBoxIcon.Information)
                    ListBox2.Items.Add("Não há seções cadastradas!")
                End If
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Fechar()
            End Try
        ElseIf Me.Text = "Maquinário" Then
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEMAQUINARIO FROM TB_MAQUINARIOS")
                Ler()
                Dim resultado As Boolean = leitura.HasRows
                If resultado = True Then
                    While leitura.Read
                        ListBox2.Items.Add(leitura("NOMEMAQUINARIO"))
                    End While
                Else
                    MessageBox.Show("Não há maquinários cadastrados!", "Maquinário", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ListBox2.Items.Add("Não há maquinários cadastrados!")
                End If
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Fechar()
            End Try
        End If
    End Sub
#End Region

#Region "Alteração de Dados"
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        If Me.Text = "Seção" Then
            'TextBox8.ReadOnly = True
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("SELECT NOMESECAO, IMAGEM FROM TBSECOES WHERE NOMESECAO = '" & ListBox2.SelectedIndex & "'")
            'Ler()
            'While leitura.Read
            'TextBox8.Text = leitura("NOMESECAO")
            'PictureBox1.Image = leitura("IMAGEM")
            'End While
            'Fechar()
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
            'RemoveHandler Button4.Click, AddressOf Button4_Click
            'AddHandler Button4.Click, AddressOf AlterarDadosSecao
        ElseIf Me.Text = "Maquinário" Then
            'TextBox8.ReadOnly = True
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("SELECT NOME, IMAGEM, CNC FROM TBMAQUINARIOS WHERE NOME = '" & ListBox2.SelectedIndex & "'")
            'Ler()
            'While leitura.Read
            'TextBox8.Text = leitura("NOMESECAO")
            'PictureBox1.Image = leitura("IMAGEM")
            'resultado = leitura("CNC")
            'If resultado = True Then
            'CheckBox1.Checked = True
            'Else
            'CheckBox1.Checked = False
            'End If
            'End While
            'Fechar()
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
            'RemoveHandler Button4.Click, AddressOf Button4_Click
            'AddHandler Button4.Click, AddressOf AlterarDadosMaquinario
        End If
    End Sub
#End Region 'VERIFICAR

#Region "Botão Vermelho"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.Text = "Seção" Then
            frmMessageBox.Text = "Seção"
            frmMessageBox.Button1.Text = "Excluir"
            frmMessageBox.Button2.Text = "Cancelar"
            frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
            frmMessageBox.Show()
        ElseIf Me.Text = "Maquinário" Then
            frmMessageBox.Text = "Maquinário"
            frmMessageBox.Button1.Text = "Excluir"
            frmMessageBox.Button2.Text = "Cancelar"
            frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
            frmMessageBox.Show()
        End If
    End Sub
#End Region 'VERIFICAR

#Region "Botão Verde"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Me.Text = "Seção" Then
            Try
                Conectar()
                Iniciar()
                Comandar("INSERT INTO TB_SECOES (NOMESECAO, CAMINHOIMGSECAO) VALUES ('" & TextBox8.Text & "','" & OpenFileDialog1.FileName & "')")
                Executar()
                Fechar()
                MessageBox.Show("Inserido com sucesso!", "Seção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox8.Clear()
                PictureBox1.Image = Nothing
                Try
                    Conectar()
                    Comandar("SELECT NOMESECAO FROM TB_SECOES")
                    Ler()
                    While leitura.Read
                        ListBox2.Items.Add(leitura("NOMESECAO"))
                    End While
                    Fechar()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf Me.Text = "Maquinário" Then
            If CheckBox1.Checked = True Then
                Try
                    Conectar()
                    Iniciar()
                    Comandar("INSERT INTO TB_MAQUINARIOS (NOMEMAQUINARIO, CAMINHOIMGMAQUINARIO, CNC) VALUES ('" & TextBox8.Text & "','" & OpenFileDialog1.FileName & "', 1)")
                    Executar()
                    Fechar()
                    MessageBox.Show("Inserido com sucesso!", "Maquinário", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    TextBox8.Clear()
                    PictureBox1.Image = Nothing
                    CheckBox1.Checked = False
                    Try
                        Conectar()
                        Iniciar()
                        Comandar("SELECT NOMEMAQUINARIO FROM TB_MAQUINARIOS")
                        Ler()
                        While leitura.Read
                            ListBox2.Items.Add(leitura("NOMEMAQUINARIO"))
                        End While
                        Fechar()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            ElseIf CheckBox1.Checked = False Then
                Try
                    Conectar()
                    Iniciar()
                    Comandar("INSERT INTO TB_MAQUINARIOS (NOMEMAQUINARIO, CAMINHOIMGMAQUINARIO, CNC) VALUES ('" & TextBox8.Text & "','" & OpenFileDialog1.FileName & "', 0)")
                    Executar()
                    Fechar()
                    MessageBox.Show("Inserido com sucesso!", "Maquinário", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    TextBox8.Clear()
                    PictureBox1.Image = Nothing
                    Try
                        Conectar()
                        Iniciar()
                        Comandar("SELECT NOMEMAQUINARIO FROM TB_MAQUINARIOS")
                        Ler()
                        While leitura.Read
                            ListBox2.Items.Add(leitura("NOMEMAQUINARIO"))
                        End While
                        Fechar()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub
#End Region

#Region "Exibir Arquivo"
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
        End If
    End Sub
#End Region

#Region "TextBox Editável"
    Private Sub TextBox8_Click(sender As Object, e As EventArgs) Handles TextBox8.Click
        'TextBox8.ReadOnly = False
    End Sub
#End Region 'VERIFICAR

#Region "Alterar Dados - Seção"
    Public Sub AlterarDadosSecao()
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("UPDATE TBSECOES SET NOMESECAO = '" & TextBox8.Text & "' AND IMAGEM = '" & OpenFileDialog1.FileName & "' WHERE NOMESECAO = '" & ListBox2.SelectedIndex & "'")
        'Executar()
        'Fechar()
        MessageBox.Show("Alterado com sucesso!", "Seção", MessageBoxButtons.OK, MessageBoxIcon.Information)
        TextBox8.Clear()
        PictureBox1.Image = Nothing
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'VERIFICAR

#Region "Alterar Dados - Maquinário"
    Public Sub AlterarDadosMaquinario()
        If CheckBox1.Checked = True Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("UPDATE TBMAQUINARIOS SET NOME = '" & TextBox8.Text & "' AND IMAGEM = '" & OpenFileDialog1.FileName & "' AND CNC = 1 WHERE NOME = '" & ListBox2.SelectedIndex & "'")
            'Executar()
            'Fechar()
            MessageBox.Show("Alterado com sucesso!", "Seção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox8.Clear()
            PictureBox1.Image = Nothing
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf CheckBox1.Checked = False Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("UPDATE TBMAQUINARIOS SET NOME = '" & TextBox8.Text & "' AND IMAGEM = '" & OpenFileDialog1.FileName & "' AND CNC = 0 WHERE NOME = '" & ListBox2.SelectedIndex & "'")
            'Executar()
            'Fechar()
            MessageBox.Show("Alterado com sucesso!", "Seção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox8.Clear()
            PictureBox1.Image = Nothing
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        End If
    End Sub
#End Region 'VERIFICAR

End Class