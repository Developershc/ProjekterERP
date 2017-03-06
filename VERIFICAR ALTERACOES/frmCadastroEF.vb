Public Class frmCadastroEF

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Text = "Cadastro da empresa" Then
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT RAZAOSOCIAL FROM TB_EMPRESAS WHERE RAZAOSOCIAL = '" & TextBox1.Text & "'")
                Ler()
                If leitura.HasRows Then
                    MessageBox.Show("Empresa já cadastrada!", "Cadastro da empresa", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    TextBox1.Focus()
                    TextBox1.SelectAll()
                    TextBox2.Clear()
                    TextBox3.Clear()
                    TextBox4.Clear()
                    TextBox6.Clear()
                    TextBox5.Clear()
                    TextBox7.Clear()
                    Fechar()
                Else
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_EMPRESAS (RAZAOSOCIAL, ENDERECO, CIDADE, ESTADO, TELEFONE, CEP, CNPJ) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox6.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Inserido com sucesso!", "Cadastro da empresa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        For Each ctrl In Me.Controls
                            If TypeOf ctrl Is TextBox Then
                                ctrl.Clear()
                            End If
                        Next
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf Me.Text = "Fornecedor" Then
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT RAZAOSOCIAL FROM TB_FORNECEDORES WHERE RAZAOSOCIAL = '" & TextBox1.Text & "'")
                Ler()
                If leitura.HasRows Then
                    MessageBox.Show("Fornecedor já cadastrado!", "Fornecedor", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    TextBox1.Focus()
                    TextBox1.SelectAll()
                    TextBox2.Clear()
                    TextBox3.Clear()
                    TextBox4.Clear()
                    TextBox6.Clear()
                    TextBox5.Clear()
                    TextBox7.Clear()
                    Fechar()
                Else
                    Try
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_FORNECEDORES (RAZAOSOCIAL, ENDERECO, CIDADE, ESTADO, TELEFONE, CEP, CNPJ) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox6.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Inserido com sucesso!", "Fornecedor", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        For Each ctrl In Me.Controls
                            If TypeOf ctrl Is TextBox Then
                                ctrl.Clear()
                            End If
                        Next
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

#Region "Botão Vermelho"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Me.Text = "Cadastro de empresa" Then
            frmMessageBox.Text = "Cadastro de empresa"
            frmMessageBox.Button1.Text = "Excluir"
            frmMessageBox.Button2.Text = "Cancelar"
            frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
            frmMessageBox.Show()
        ElseIf Me.Text = "Fornecedor" Then
            frmMessageBox.Text = "Fornecedor"
            frmMessageBox.Button1.Text = "Excluir"
            frmMessageBox.Button2.Text = "Cancelar"
            frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
            frmMessageBox.Show()
        End If
    End Sub
#End Region 'VERIFICAR

#Region "Alterar Dados"
    Private Sub AlterarDados()
        If Me.Text = "Cadastro da empresa" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("UPDATE TBEMPRESAS SET RAZAOSOCIAL = '" & frmCadastroEF.TextBox1.Text & "', ENDERECO = '" & frmCadastroEF.TextBox2.Text & "', CIDADE = '" & frmCadastroEF.TextBox3.Text & "', ESTADO = '" & frmCadastroEF.TextBox4.Text & "', TELEFONE = '" & frmCadastroEF.TextBox6.Text & "', CEP = '" & frmCadastroEF.TextBox5.Text & "', CNPJ = '" & frmCadastroEF.TextBox7.Text & "' WHERE RAZAOSOCIAL = '" & frmCadastroEF.TextBox1.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Alterado com sucesso!", "Entrada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Fornecedor" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("UPDATE TBFORNECEDORES SET RAZAOSOCIAL = '" & frmCadastroEF.TextBox1.Text & "', ENDERECO = '" & frmCadastroEF.TextBox2.Text & "', CIDADE = '" & frmCadastroEF.TextBox3.Text & "', ESTADO = '" & frmCadastroEF.TextBox4.Text & "', TELEFONE = '" & frmCadastroEF.TextBox6.Text & "', CEP = '" & frmCadastroEF.TextBox5.Text & "', CNPJ = '" & frmCadastroEF.TextBox7.Text & "' WHERE RAZAOSOCIAL = '" & frmCadastroEF.TextBox1.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Alterado com sucesso!", "Entrada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        End If
    End Sub
#End Region 'VERIFICAR

#Region "Exibição de Dados"
    Private Sub frmCadastroEF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Me.Text = "Cadastro da empresa" Then
            Try
                Conectar()
                Comandar("SELECT RAZAOSOCIAL, ENDERECO, CIDADE, ESTADO, TELEFONE, CEP, CNPJ FROM TB_EMPRESAS")
                Ler()
                While leitura.Read
                    TextBox1.Text = leitura("RAZAOSOCIAL")
                    TextBox2.Text = leitura("ENDERECO")
                    TextBox3.Text = leitura("CIDADE")
                    TextBox4.Text = leitura("ESTADO")
                    TextBox6.Text = leitura("TELEFONE")
                    TextBox5.Text = leitura("CEP")
                    TextBox7.Text = leitura("CNPJ")
                End While
                Fechar()
                TextBox1.ReadOnly = True
                TextBox2.ReadOnly = True
                TextBox3.ReadOnly = True
                TextBox4.ReadOnly = True
                TextBox6.ReadOnly = True
                TextBox5.ReadOnly = True
                TextBox7.ReadOnly = True
            Catch ex As Exception
                MessageBox.Show("Não há dados para exibir! Cadastre a empresa!", "Cadastro de empresa", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Focus()
            End Try
        End If
    End Sub
#End Region

End Class