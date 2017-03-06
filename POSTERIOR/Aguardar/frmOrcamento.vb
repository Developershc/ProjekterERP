Public Class frmOrcamento

#Region "Escolher Arquivo"
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        OpenFileDialog1.ShowDialog()
        TextBox10.Text = OpenFileDialog1.FileName
    End Sub
#End Region 'OK

#Region "ToolTip"
    Private Sub Button2_MouseHover(sender As Object, e As EventArgs) Handles Button2.MouseHover
        ToolTip1.SetToolTip(Button2, "Gerar pedido de venda")
    End Sub
#End Region 'OK

#Region "Abrir Arquivo"
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start(TextBox10.Text)
    End Sub
#End Region 'OK

#Region "Configurações"
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Produto" Then
            GroupBox1.Text = "Informações dos produtos:"
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("SELECT (O QUE FOR INSERIDO NO DATAGRID) FROM (VERIFICAR)")
            'Ler()
            'While leitura.Read
            ''---- CÓDIGO PARA INSERIR NO DATAGRID ----
            'End While
            'Fechar()
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf ComboBox1.Text = "Montagem" Then
            GroupBox1.Text = "Informações das montagens:"
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("SELECT (O QUE FOR INSERIDO NO DATAGRID) FROM TBMONTAGENS")
            'Ler()
            'While leitura.Read
            ''---- CÓDIGO PARA INSERIR NO DATAGRID ----
            'End While
            'Fechar()
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf ComboBox1.Text = "Peça" Then
            GroupBox1.Text = "Informações das peças:"
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("SELECT (O QUE FOR INSERIDO NO DATAGRID) FROM TBPECAS")
            'Ler()
            'While leitura.Read
            ''---- CÓDIGO PARA INSERIR NO DATAGRID ----
            'End While
            'Fechar()
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        End If
    End Sub
#End Region 'VERIFICAR

#Region "Localizar"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frmLocalizar.Show()
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox5.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox11.ReadOnly = True
        TextBox8.ReadOnly = True
        TextBox9.ReadOnly = True
        TextBox7.ReadOnly = True
        TextBox10.ReadOnly = True
    End Sub
#End Region 'OK

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("INSERT INTO TBORCAMENTOS (CODIGO, CLIENTE, REVISAO, CUSTOINTERNO, IMPOSTO, LUCRO, DESCONTO, DATAINCLUSAO, DATAALTERACAO, DATAAPROVACAO, DESENHO) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox11.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "','" & TextBox7.Text & "','" & TextBox10.Text & "')")
        'Executar()
        'Fechar()
        MessageBox.Show("Inserido com sucesso!", "Orçamento", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
        frmMessageBox.Text = "Orçamento"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region 'OK

End Class

'Alteração de dados será feita direto no DataGrid.