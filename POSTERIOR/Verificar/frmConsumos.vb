Public Class frmConsumos

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT * FROM TBCONSUMOS")
        'Ler()
        'resultado = leitura.HasRows
        'If resultado = True Then
        'If MessageBox.Show("Os dados serão alterados. Deseja continuar?", "Consumos", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes = True Then
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("UPDATE TBCONSUMOS SET MESAGUA, VALORAGUA, MEDIAAGUA, MESLUZ, VALORLUZ, MEDIALUZ")
        'Executar()
        'Fechar()
        'MessageBox.Show("Alterado com sucesso!", "Consumos", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'For Each ctrl In Me.Controls
        'If TypeOf ctrl Is TextBox Then
        'ctrl.Clear()
        'End If
        'Next
        'ComboBox1.Text = ""
        'ComboBox2.Text = ""
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
        'Else
        'TextBox1.Focus()
        'TextBox1.SelectAll()
        'End If
        'Else



        Try
            Conectar()
            Iniciar()
            Comandar("INSERT INTO TB_CONSUMOS (MESAGUA, VALORAGUA, MEDIAAGUA, MESLUZ, VALORLUZ, MEDIALUZ) VALUES ('" & ComboBox1.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')")
            Executar()
            Fechar()
            MessageBox.Show("Inserido com sucesso!", "Consumos", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Clear()
                End If
            Next
            ComboBox1.Text = ""
            ComboBox2.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



        'End If
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'VERIFICAR

#Region "Formatação de Moeda"
    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        TextBox1.Text = FormatCurrency(TextBox1.Text, 2)
    End Sub

    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        TextBox2.Text = FormatCurrency(TextBox2.Text, 2)
    End Sub

    Private Sub TextBox3_Leave(sender As Object, e As EventArgs) Handles TextBox3.Leave
        TextBox3.Text = FormatCurrency(TextBox3.Text, 2)
    End Sub

    Private Sub TextBox4_Leave(sender As Object, e As EventArgs) Handles TextBox4.Leave
        TextBox4.Text = FormatCurrency(TextBox4.Text, 2)
    End Sub
#End Region

End Class