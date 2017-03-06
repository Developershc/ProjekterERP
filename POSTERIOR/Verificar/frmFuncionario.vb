Public Class frmFuncionario

#Region "Exibir Quadro de Funcionários"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frmQuadro.Show()
        frmQuadro.DataGridView1.Columns.Add("noCracha", "Nº do Crachá")
        frmQuadro.DataGridView1.Columns.Add("Nome", "Nome")
        frmQuadro.DataGridView1.Columns.Add("Salario", "Salário")

        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOCRACHA, NOME, SALARIO FROM TB_FUNCIONARIOS")
            Ler()
            While leitura.Read
                frmQuadro.DataGridView1.Rows.Add(leitura("NoCracha"), leitura("Nome"), leitura("Salario"))
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Fechar()
        End Try
    End Sub
#End Region

#Region "ToolTip"
    Private Sub Button2_MouseHover(sender As Object, e As EventArgs) Handles Button2.MouseHover
        ToolTip1.SetToolTip(Button2, "Acessar quadro de funcionários")
    End Sub
#End Region

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Conectar()
            Iniciar()
            Comandar("INSERT INTO TB_FUNCIONARIOS (NOCRACHA, NOME, CPF, ENDERECO, BAIRRO, CIDADE, ESTADO, CEP, DDD, TELEFONE, BANCO, AGENCIA, CONTA, FUNCAO, SALARIO, HORASSEMANA, VALORHORA, DATAADMISSAO, CTPS, PISPASEP, DATANASC) VALUES ('" & TextBox26.Text & "','" & TextBox25.Text & "','" & TextBox24.Text & "','" & TextBox23.Text & "','" & TextBox22.Text & "','" & TextBox21.Text & "','" & TextBox20.Text & "','" & TextBox19.Text & "','" & TextBox18.Text & "','" & TextBox17.Text & "','" & TextBox16.Text & "','" & TextBox15.Text & "','" & TextBox14.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')")
            Executar()
            Fechar()
            MessageBox.Show("Inserido com sucesso!", "Funcionário", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Clear()
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Fechar()
        End Try
    End Sub
#End Region

End Class