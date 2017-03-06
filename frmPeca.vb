
Public Class frmPeca

#Region "Escolher Arquivos"
    Private Sub LinkLabel1_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        OpenFileDialog1.ShowDialog()
        TextBox12.Text = OpenFileDialog1.FileName
    End Sub
    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        OpenFileDialog1.ShowDialog()
        TextBox14.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        OpenFileDialog1.ShowDialog()
        TextBox13.Text = OpenFileDialog1.FileName
    End Sub
#End Region 'OK

#Region "Acessar Arquivos"
    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        Process.Start(TextBox12.Text)
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Process.Start(TextBox13.Text)
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        Process.Start(TextBox14.Text)
    End Sub
#End Region 'OK

#Region "Localizar"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frmLocalizar.Show()
        TextBox1.ReadOnly = True
        TextBox11.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox5.ReadOnly = True
        TextBox7.ReadOnly = True
        TextBox8.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox12.ReadOnly = True
        TextBox13.ReadOnly = True
        TextBox14.ReadOnly = True
        RemoveHandler Button1.Click, AddressOf Button1_Click
        AddHandler Button1.Click, AddressOf AlterarDados
    End Sub
#End Region 'OK

#Region "Alterar Dados"
    Private Sub AlterarDados()
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("UPDATE TBPECAS SET CODIGOP = '" & TextBox1.Text & "', REVISAO = '" & TextBox11.Text & "', TITULOP = '" & TextBox2.Text & "', AREA = '" & TextBox3.Text & "', MASSA = '" & TextBox4.Text & "', RENDIMENTO = '" & TextBox5.Text & "', TRATACAB = '" & ComboBox4.Text & "', COR = '" & TextBox7.Text & "', NOMEGRUPOMAT = '" & ComboBox1.Text & "', NOMEQ = '" & ComboBox2.Text & "', NOMET = '" & ComboBox3.Text & "', NOMEMAT = '" & ComboBox5.Text & "', CUSTOMAT = '" & TextBox8.Text & "', GARANTIA = '" & TextBox6.Text & "', CUSTOTRATACAB = '" & TextBox9.Text & "', CUSTOMO = '" & TextBox10.Text & "', CAD = '" & TextBox12.Text & "', DESENHO = '" & TextBox13.Text & "', 3D = '" & TextBox14.Text & "'")
        'Executar()
        'Fechar()
        MessageBox.Show("Alterado com sucesso!", "Peça", MessageBoxButtons.OK, MessageBoxIcon.Information)
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Clear()
            End If
        Next
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'OK

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("INSERT INTO TBPECAS (CODIGOP, REVISAO, TITULOP, AREA, MASSA, RENDIMENTO, TRATACAB, COR, NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, CUSTOMAT, GARANTIA, CUSTOTRATACAB, CUSTOMO, CAD, DESENHO, 3D) VALUES ('" & TextBox1.Text & "', '" & TextBox11.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "', '" & ComboBox4.Text & "', '" & TextBox7.Text & "', '" & ComboBox1.Text & "', '" & ComboBox2.Text & "', '" & ComboBox3.Text & "', '" & ComboBox5.Text & "', '" & TextBox8.Text & "', '" & TextBox6.Text & "', '" & TextBox9.Text & "', '" & TextBox10.Text & "', '" & TextBox12.Text & "', '" & TextBox13.Text & "', '" & TextBox14.Text & "')")
        'Executar()
        'Fechar()
        MessageBox.Show("Inserido com sucesso!", "Peça", MessageBoxButtons.OK, MessageBoxIcon.Information)
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Clear()
            End If
        Next
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'OK

#Region "Botão Vermelho"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmMessageBox.Text = "Peça"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region 'OK

#Region "TextBoxes Editáveis"
    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        TextBox1.ReadOnly = False
    End Sub

    Private Sub TextBox11_Click(sender As Object, e As EventArgs) Handles TextBox11.Click
        TextBox11.ReadOnly = False
    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        TextBox2.ReadOnly = False
    End Sub

    Private Sub TextBox3_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        TextBox3.ReadOnly = False
    End Sub

    Private Sub TextBox4_Click(sender As Object, e As EventArgs) Handles TextBox4.Click
        TextBox4.ReadOnly = False
    End Sub

    Private Sub TextBox5_Click(sender As Object, e As EventArgs) Handles TextBox5.Click
        TextBox5.ReadOnly = False
    End Sub

    Private Sub TextBox7_Click(sender As Object, e As EventArgs) Handles TextBox7.Click
        TextBox7.ReadOnly = False
    End Sub

    Private Sub TextBox8_Click(sender As Object, e As EventArgs) Handles TextBox8.Click
        TextBox8.ReadOnly = False
    End Sub

    Private Sub TextBox6_Click(sender As Object, e As EventArgs) Handles TextBox6.Click
        TextBox6.ReadOnly = False
    End Sub

    Private Sub TextBox9_Click(sender As Object, e As EventArgs) Handles TextBox9.Click
        TextBox9.ReadOnly = False
    End Sub

    Private Sub TextBox10_Click(sender As Object, e As EventArgs) Handles TextBox10.Click
        TextBox10.ReadOnly = False
    End Sub

    Private Sub TextBox12_Click(sender As Object, e As EventArgs) Handles TextBox12.Click
        TextBox12.ReadOnly = False
    End Sub

    Private Sub TextBox13_Click(sender As Object, e As EventArgs) Handles TextBox13.Click
        TextBox13.ReadOnly = False
    End Sub

    Private Sub TextBox14_Click(sender As Object, e As EventArgs) Handles TextBox14.Click
        TextBox14.ReadOnly = False
    End Sub
#End Region 'OK

End Class

'DataGrid de Processo de Fabricação exibe o que é referente à peça, podendo ser alterado diretamente.
'Custo de Material é calculado a partir da multiplicação do valor do material e a massa.
'Custo de Tratamento/Acabamento é calculado a partir da multiplicação do serviço pela área ou pela massa. O que
'determina é a unidade.
'O valor para cálculo é o mais caro de todos os fornecedores.
