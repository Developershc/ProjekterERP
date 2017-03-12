Imports System.Data.SqlClient

Public Class frmLocalizar

#Region "Carregar DataGridView"
    Private Sub frmLocalizar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Me.Text = "Localizar Fornecedor" Then
            Try
                Conectar()
                Iniciar()
                Dim da As New SqlDataAdapter("SELECT RAZAOSOCIAL, ENDERECO, CIDADE, ESTADO, TELEFONE, CEP, CNPJ FROM TB_FORNECEDORES", conexao)
                da.Fill(ds)
                Me.DataGridView1.DataSource = ds.Tables(0).DefaultView
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub
#End Region

#Region "Selecionar dados para carregar em outra form"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmCadastroEF.TextBox1.ReadOnly = True
        frmCadastroEF.TextBox2.ReadOnly = True
        frmCadastroEF.TextBox3.ReadOnly = True
        frmCadastroEF.TextBox4.ReadOnly = True
        frmCadastroEF.TextBox6.ReadOnly = True
        frmCadastroEF.TextBox5.ReadOnly = True
        If Me.Text = "Localizar Fornecedor" Then
            frmCadastroEF.TextBox1.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString()
        End If
        Me.Close()
        RemoveHandler frmCadastroEF.Button1.Click, AddressOf frmCadastroEF.Button1_Click
        AddHandler frmCadastroEF.Button1.Click, AddressOf frmCadastroEF.AlterarDados
    End Sub
#End Region 'VERIFICAR

End Class

'Botão Verde seleciona os dados para aparecerem na tela que chamar esta.