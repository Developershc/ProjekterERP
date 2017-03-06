Public Class frmComparados

#Region "Carregar Combos"
    Private Sub frmComparados_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT NOMEGRUPOMAT FROM TBGRUPOSMAT")
        'Ler()
        'While leitura.Read
        'ComboBox1.Items.Add(leitura("NOME"))
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT NOMET FROM TBTIPOS")
        'Ler()
        'While leitura.Read
        'ComboBox2.Items.Add(leitura("NOME"))
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT NOMEQ FROM TBQUALIDADES")
        'Ler()
        'While leitura.Read
        'ComboBox3.Items.Add(leitura("NOME"))
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT NOMEMAT FROM TBMATERIAIS")
        'Ler()
        'While leitura.Read
        'ComboBox4.Items.Add(leitura("NOME"))
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT NOMESECAO FROM TBSECOES")
        'Ler()
        'While leitura.Read
        'ComboBox5.Items.Add(leitura("NOME"))
        'End While
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT PRODUTO FROM TBCONSUMIVEIS")
        'Ler()
        'While leitura.Read
        'ComboBox6.Items.Add(leitura("PRODUTO"))
        'End While
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try

        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT NOMES FROM TBSERVICOS")
        'Ler()
        'While leitura.Read
        'ComboBox8.Items.Add(leitura("NOME"))
        'End While
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try

    End Sub
#End Region 'OK

#Region "Alterar Grupo de Material"
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT (O QUE VAI SER SELECIONADO) FROM TBFORNECEDORES WHERE NOMEGRUPOMAT = '" & ComboBox1.Text & "'")
        'Ler()
        'While leitura.Read

        ''---- CÓDIGO PARA DATAGRID RECEBER VALORES ----

        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'TERMINAR CÓDIGO

#Region "Alterar Qualidade"
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT (O QUE FOR SELECIONADO) FROM TBFORNECEDORES WHERE NOMEQ = '" & ComboBox3.Text & "'")
        'Ler()
        'While leitura.Read

        ''---- CÓDIGO PARA DATAGRID RECEBER VALORES ----

        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'TERMINAR CÓDIGO

#Region "Alterar Tipo"
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT (O QUE FOR SELECIONADO) FROM TBFORNECEDORES WHERE NOMET = '" & ComboBox2.Text & "'")
        'Ler()
        'While leitura.Read

        ''---- CÓDIGO PARA DATAGRID RECEBER VALORES ----

        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'TERMINAR CÓDIGO

#Region "Alterar Material"
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT (O QUE FOR SELECIONADO) FROM TBFORNECEDORES WHERE NOMEMAT = '" & ComboBox4.Text & "'")
        'Ler()
        'While leitura.Read

        ''---- CÓDIGO PARA DATAGRID RECEBER VALORES ----

        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'TERMINAR CÓDIGO

#Region "Alterar Seção"
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT (O QUE FOR SELECIONADO) FROM TBFORNECEDORES WHERE NOMESECAO = '" & ComboBox5.Text & "'")
        'Ler()
        'While leitura.Read
        ''---- CÓDIGO PARA DATAGRID RECEBER VALORES ----
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'TERMINAR CÓDIGO

#Region "Alterar Consumível"
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT (O QUE FOR SELECIONADO) FROM TBFORNECEDORES WHERE PRODUTO = '" & ComboBox6.Text & "'")
        'Ler()
        'While leitura.Read
        ''---- CÓDIGO PARA DATAGRID RECEBER VALORES ----
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'TERMINAR CÓDIGO

#Region "Alterar Serviço"
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT (O QUE FOR SELECIONADO) FROM TBFORNECEDORES WHERE NOMES = '" & ComboBox8.Text & "'")
        'Ler()
        'While leitura.Read
        ''---- CÓDIGO PARA DATAGRID RECEBER VALORES ----
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'TERMINAR CÓDIGO

End Class

'DataGrid vai receber informações de todos os fornecedores ao abrir a tela?
'À medida que os dados alteram nas Combos, os dados nele também serão alterados, de acordo com o que é selecionado?