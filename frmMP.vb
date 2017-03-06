Public Class frmMP

#Region "Escolher Arquivo"
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        OpenFileDialog1.ShowDialog()
        TextBox13.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        OpenFileDialog1.ShowDialog()
        TextBox14.Text = OpenFileDialog1.FileName
    End Sub
#End Region 'OK

#Region "Abrir Arquivo"
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
        TextBox7.ReadOnly = True
        TextBox8.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox9.ReadOnly = True
        TextBox10.ReadOnly = True
        TextBox13.ReadOnly = True
        TextBox14.ReadOnly = True
        RemoveHandler Button1.Click, AddressOf Button1_Click
        AddHandler Button1.Click, AddressOf AlterarDados
    End Sub
#End Region 'OK

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Text = "Montagem" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("INSERT INTO TBMONTAGENS (CODIGOM, REVISAO, TITULOM, VOLUME, MASSA, TRATACAB, COR, CUSTOMAT, GARANTIA, CUSTOTRATACAB, CUSTOMO, DESENHO, 3D) VALUES ('" & TextBox1.Text & "','" & TextBox11.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox6.Text & "','" & TextBox9.Text & "','" & TextBox10.Text & "','" & TextBox13.Text & "','" & TextBox14.Text & "')")
            'Executar()
            'Fechar()
            MessageBox.Show("Inserido com sucesso!", "Montagem", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Clear()
                End If
            Next
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Produto" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("INSERT INTO TBPRODUTOS (CODIGOPROD, REVISAO, TITULOPROD, VOLUME, MASSA, TRATACAB, COR, CUSTOMAT, GARANTIA, CUSTOTRATACAB, CUSTOMO, DESENHO, 3D) VALUES ('" & TextBox1.Text & "','" & TextBox11.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox6.Text & "','" & TextBox9.Text & "','" & TextBox10.Text & "','" & TextBox13.Text & "','" & TextBox14.Text & "')")
            'Executar()
            'Fechar()
            MessageBox.Show("Inserido com sucesso!", "Produto", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Clear()
                End If
            Next
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        End If
    End Sub
#End Region 'OK

#Region "Alterar Dados"
    Private Sub AlterarDados()
        If Me.Text = "Montagem" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("UPDATE TBMONTAGENS SET CODIGOM = '" & TextBox1.Text & "' AND REVISAO = '" & TextBox11.Text & "' AND TITULOM = '" & TextBox2.Text & "' AND VOLUME = '" & TextBox3.Text & "' AND MASSA = '" & TextBox4.Text & "' AND TRATACAB = '" & ComboBox1.Text & "' AND COR = '" & TextBox7.Text & "' AND CUSTOMAT = '" & TextBox8.Text & "' AND GARANTIA = '" & TextBox6.Text & "' AND CUSTOTRATACAB = '" & TextBox9.Text & "' AND CUSTOMO = '" & TextBox10.Text & "' AND DESENHO = '" & TextBox13.Text & "' AND 3D = '" & TextBox14.Text & "'")
            'Executar()
            'Fechar()
            MessageBox.Show("Alterado com sucesso!", "Montagem", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Clear()
                End If
            Next
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Produto" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("UPDATE TBPRODUTOS SET CODIGOPROD = '" & TextBox1.Text & "' AND REVISAO = '" & TextBox11.Text & "' AND TITULOPROD = '" & TextBox2.Text & "' AND VOLUME = '" & TextBox3.Text & "' AND MASSA = '" & TextBox4.Text & "' AND TRATACAB = '" & ComboBox1.Text & "' AND COR = '" & TextBox7.Text & "' AND CUSTOMAT = '" & TextBox8.Text & "' AND GARANTIA = '" & TextBox6.Text & "' AND CUSTOTRATACAB = '" & TextBox9.Text & "' AND CUSTOMO = '" & TextBox10.Text & "' AND DESENHO = '" & TextBox13.Text & "' AND 3D = '" & TextBox14.Text & "'")
            'Executar()
            'Fechar()
            MessageBox.Show("Alterado com sucesso!", "Produto", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Clear()
                End If
            Next
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        End If
    End Sub
#End Region 'OK

#Region "Botão Vermelho"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Me.Text = "Montagem" Then
            frmMessageBox.Text = "Montagem"
            frmMessageBox.Button1.Text = "Excluir"
            frmMessageBox.Button2.Text = "Cancelar"
            frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
            frmMessageBox.Show()
        ElseIf Me.Text = "Produto" Then
            frmMessageBox.Text = "Produto"
            frmMessageBox.Button1.Text = "Excluir"
            frmMessageBox.Button2.Text = "Cancelar"
            frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
            frmMessageBox.Show()
        End If
    End Sub
#End Region 'OK

End Class

'DataGrids de Materiais e Componentes exibem os referentes ao selecionado, podendo ser alterados.

'Custos de Material e Tratamento/Acabamento é a soma dos referidos valores exibidos no DataGrid.

'O DataGrid deve ter código, título e quantidade. O último poderá ser editável e, ao apertar ENTER, deve ser
'possível adicionar outro componente. Os dados devem ser exibidos a partir da digitação do código, que "puxará"
'o título para que a quantidade seja colocada. O título é referente à peça.