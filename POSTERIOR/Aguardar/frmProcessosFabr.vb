Public Class frmProcessosFabr

#Region "Localizar"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frmLocalizar.Show()
        TextBox4.ReadOnly = True
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
    End Sub
#End Region 'OK

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("INSERT INTO TBPRODUCOES (NOMEPF, NOMESECAO, NOME, CUSTOHORA, FUNCIONAMENTO, MEDIA) VALUES ('" & TextBox4.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')")
        'Executar()
        'Fechar()
        MessageBox.Show("Inserido com sucesso!", "Produção", MessageBoxButtons.OK, MessageBoxIcon.Information)
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Clear()
            End If
        Next
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'OK

#Region "Botão Vermelho"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmMessageBox.Text = "Produção"
        frmMessageBox.Show()
    End Sub
#End Region 'OK

#Region "TextBoxes Editáveis"
    Private Sub TextBox4_Click(sender As Object, e As EventArgs) Handles TextBox4.Click
        TextBox4.ReadOnly = False
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        TextBox1.ReadOnly = False
    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        TextBox2.ReadOnly = False
    End Sub

    Private Sub TextBox3_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        TextBox3.ReadOnly = False
    End Sub
#End Region 'OK

#Region "Carregar Combos"
    Private Sub frmProcessosFabr_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Conectar()
            Comandar("SELECT NOMESECAO FROM TB_SECOES")
            Ler()
            Dim resultado As Boolean = leitura.HasRows
            If resultado = True Then
                While leitura.Read
                    ComboBox1.Items.Add(leitura("NOMESECAO"))
                End While
            Else
                MessageBox.Show("Não há seções cadastradas!", "Processos de Fabricação", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Try
            Conectar()
            Comandar("SELECT NOMEMAQUINARIO FROM TB_MAQUINARIOS")
            Ler()
            Dim resultado As Boolean = leitura.HasRows
            If resultado = True Then
                While leitura.Read
                    ComboBox2.Items.Add(leitura("NOMEMAQUINARIO"))
                End While
            Else
                MessageBox.Show("Não há maquinários cadastrados!", "Processos de Fabricação", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

End Class