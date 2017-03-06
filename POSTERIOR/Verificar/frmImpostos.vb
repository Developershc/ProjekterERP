Public Class frmImpostos

#Region "ToolTips"
    Private Sub Label1_MouseHover(sender As Object, e As EventArgs) Handles Label1.MouseHover
        ToolTip1.SetToolTip(Label1, "Insira a porcentagem do imposto")
    End Sub

    Private Sub Label5_MouseHover(sender As Object, e As EventArgs) Handles Label5.MouseHover
        ToolTip1.SetToolTip(Label5, "Insira a porcentagem do imposto")
    End Sub

    Private Sub Label3_MouseHover(sender As Object, e As EventArgs) Handles Label3.MouseHover
        ToolTip1.SetToolTip(Label3, "Insira a porcentagem do imposto")
    End Sub

    Private Sub Label2_MouseHover(sender As Object, e As EventArgs) Handles Label2.MouseHover
        ToolTip1.SetToolTip(Label2, "Insira a porcentagem do imposto")
    End Sub

    Private Sub Label4_MouseHover(sender As Object, e As EventArgs) Handles Label4.MouseHover
        ToolTip1.SetToolTip(Label4, "Insira a porcentagem do imposto")
    End Sub
#End Region

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT * FROM TB_IMPOSTOS")
        'Ler()
        'resultado = leitura.HasRows
        'If resultado = True Then
        'If MessageBox.Show("Os dados serão alterados. Deseja continuar?", "Impostos", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes = True Then
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("UPDATE TBIMPOSTOS SET ICMS = '" & TextBox1.Text & "', PIS = '" & TextBox2.Text & "', CONFINS = '" & TextBox3.Text & "', ISS = '" & TextBox4.Text & "', IPI = '" & TextBox5.Text & "', IPTU = '" & TextBox7.Text & "', ALUGUELIMOVEL = '" & TextBox6.Text & "', FINANCIMOVEL = '" & TextBox8.Text & "'")
        'Executar()
        'Fechar()
        'MessageBox.Show("Alterado com sucesso!", "Impostos", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'For Each ctrl In Me.Controls
        'If TypeOf ctrl Is TextBox Then
        'ctrl.Clear()
        'End If
        'Next
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
        'Else
        'TextBox1.Focus()
        'TextBox1.SelectAll()
        'End If
        'Else

        '--------------------- OK A PARTIR DAQUI ---------------------

        Try
            Conectar()
            Iniciar()
            Comandar("INSERT INTO TB_IMPOSTOS (ICMS, PIS, CONFINS, ISS, IPI, IPTU, ALUGUEL, FINANC) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "','" & TextBox6.Text & "','" & TextBox8.Text & "')")
            Executar()
            Fechar()
            MessageBox.Show("Inserido com sucesso!", "Impostos", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Clear()
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        '--------------------- OK ANTES DAQUI ---------------------

        'End If
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'VERIFICAR

End Class