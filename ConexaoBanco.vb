Imports System.Data.SqlClient

Module ConexaoBanco

    Public conexao As New SqlConnection
    Public comando As New SqlCommand
    Public leitura As SqlDataReader
    Public ds As New DataSet

    Public Sub Executar()
        comando.ExecuteNonQuery()
    End Sub

    Public Sub Iniciar()
        conexao.BeginTransaction.Commit()
    End Sub

    Public Sub Conectar()
        conexao.ConnectionString = My.Settings.con
        conexao.Open()
    End Sub

    Public Sub Ler()
        leitura = comando.ExecuteReader
    End Sub

    Public Sub Comandar(ByVal com As String)
        comando.Connection = conexao
        comando.CommandText = (com)
        comando.CommandType = CommandType.Text
    End Sub

    Public Sub Fechar()
        conexao.Close()
    End Sub

End Module

'Botão vermelho: limpar campos (inclusão), excluir dado (apenas leitura) e cancelar atualização (em atualização).
'Botão verde: incluir novo dado ou atualizar dado já existente (verificar existência).
'
'Modo edição: ao utilizar botão Localizar e clicar no campo.
'Modo inclusão: ao abrir a tela.
'Modo leitura: apenas ao utilizar botão Localizar.
'
'Para utilizar o modo edição, ele verifica se o dado existe no banco. Se existir, ele altera.
'
'Para utilizar o modo inclusão, basta abrir a tela.
'
'Para utilizar o modo leitura, ao clicar no botão Localizar, os controles desabilitam. O botão vermelho
'deleta o dado cadastrado.
