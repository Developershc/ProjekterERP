Public Class frmConsumivel

#Region "Localizar Produto"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frmLocalizar.Show()
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox6.ReadOnly = True
        RemoveHandler Button1.Click, AddressOf Button1_Click
        AddHandler Button1.Click, AddressOf AlterarDados
    End Sub
#End Region 'OK

#Region "Alterar Dados"
    Private Sub AlterarDados()
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("UPDATE TBCONSUMIVEIS SET CODINTERNO = '" & frmConsumivel.TextBox1.Text & "', PRODUTO = '" & frmConsumivel.TextBox2.Text & "', SECAO = '" & frmConsumivel.ComboBox1.Text & "', PRAZOVALIDADE = '" & frmConsumivel.TextBox4.Text & "', GARANTIA = '" & frmConsumivel.TextBox6.Text & "'")
        'Executar()
        'Fechar()
        Me.Close()
        MessageBox.Show("Inserido com sucesso!", "Impostos", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'VERIFICAR CÓDIGO DE BANCO

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Try
        '    Conectar()
        '    Iniciar()
        '    Comandar("INSERT INTO TB_CONSUMIVEIS (CODINTERNO, PRODUTO, SECAO, PRAZOVALIDADE, GARANTIA) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox4.Text & "','" & TextBox6.Text & "')")
        '    Executar()
        '    Fechar()
        '    Me.Close()
        '    MessageBox.Show("Inserido com sucesso!", "Impostos", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'OK

#Region "Botão Vermelho"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmMessageBox.Text = "Consumos"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region 'OK

#Region "Carregar ComboBox"
    Private Sub frmConsumivel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT NOMESECAO FROM TBSECOES")
        'Ler()
        'While leitura.Read
        'ComboBox1.Items.Add(leitura("NOMESECAO"))
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'OK

End Class