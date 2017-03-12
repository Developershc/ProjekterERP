Public Class frmInicial

#Region "Fechar Programa"
    Private Sub frmInicial_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If MessageBox.Show("Deseja finalizar esta sessão e fazer login com outro usuário?", "Finalizar sessão", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes = True Then
            frmLogin.Show()
            frmLogin.UsernameTextBox.Focus()
            frmLogin.UsernameTextBox.SelectAll()
            frmLogin.PasswordTextBox.Clear()
        Else
            End
        End If
    End Sub
#End Region

#Region "Finalizar Sessão"
    Private Sub FinalizarSessãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles btnFimSessao.Click
        If MessageBox.Show("Deseja finalizar esta sessão e fazer login com outro usuário?", "Finalizar sessão", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes = True Then
            Me.Close()
            frmLogin.Show()
            frmLogin.UsernameTextBox.Focus()
            frmLogin.UsernameTextBox.SelectAll()
            frmLogin.PasswordTextBox.Clear()
        Else
            Application.Exit()
        End If
    End Sub
#End Region

#Region "Calculadora"
    Private Sub CalculadoraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculadoraToolStripMenuItem.Click
        Shell("calc")
    End Sub
#End Region

#Region "Cadastro de Empresa"
    Private Sub btnCadEmp_Click(sender As Object, e As EventArgs) Handles btnCadEmp.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmCadastroEF As New frmCadastroEF
        frmCadastroEF.MdiParent = Me
        frmCadastroEF.Show()
        frmCadastroEF = Nothing
    End Sub
#End Region

#Region "Fornecedor"
    Private Sub FornceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FornceToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmCadastroEF As New frmCadastroEF
        frmCadastroEF.MdiParent = Me
        frmCadastroEF.Show()
        frmCadastroEF.GroupBox1.Text = "Informações do fornecedor"
        frmCadastroEF.Text = "Fornecedor"
        frmCadastroEF = Nothing
    End Sub
#End Region

#Region "Sobre"
    Private Sub SobreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SobreToolStripMenuItem.Click
        frmSobre.Show()
        frmSobre.Text = "Sobre"
    End Sub
#End Region

#Region "Peça"
    Private Sub PeçasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PeçasToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmPeca As New frmPeca
        frmPeca.MdiParent = Me
        frmPeca.Show()
        frmPeca = Nothing
    End Sub
#End Region

#Region "Montagem"
    Private Sub MontagensToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MontagensToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmMP As New frmMP
        frmMP.MdiParent = Me
        frmMP.Show()
        frmMP.Text = "Montagem"
        frmMP = Nothing
    End Sub
#End Region

#Region "Produto"
    Private Sub ProdutosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProdutosToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmMP As New frmMP
        frmMP.MdiParent = Me
        frmMP.Show()
        frmMP.Text = "Produto"
        frmMP = Nothing
    End Sub
#End Region

#Region "Manutenção"
    Private Sub btnUsu_Click(sender As Object, e As EventArgs) Handles btnUsu.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmManutencao As New frmManutencao
        frmManutencao.MdiParent = Me
        frmManutencao.Show()
        frmManutencao = Nothing
    End Sub
#End Region

#Region "Material"
    Private Sub MatériaPrimaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MatériaPrimaToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmMaterial As New frmMaterial
        frmMaterial.MdiParent = Me
        frmMaterial.Show()
        frmMaterial = Nothing
    End Sub
#End Region

#Region "Funcionário"
    Private Sub FuncionariosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FuncionariosToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmFuncionario As New frmFuncionario
        frmFuncionario.MdiParent = Me
        frmFuncionario.Show()
        frmFuncionario = Nothing
    End Sub
#End Region

#Region "Maquinário"
    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmSM As New frmSM
        frmSM.MdiParent = Me
        frmSM.Show()
        frmSM.Text = "Maquinário"
        frmSM.Label19.Text = "Maquinário:"
        frmSM.Label18.Text = "Maquinários já cadastrados:"
        frmSM.CheckBox1.Visible = True
        frmSM = Nothing
    End Sub
#End Region

#Region "Valores"
    Private Sub TabelasDeValoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TabelasDeValoresToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmValores As New frmValores
        frmValores.MdiParent = Me
        frmValores.Show()
        frmValores = Nothing
    End Sub
#End Region

#Region "Consumíveis"
    Private Sub ConsumiveisToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ConsumiveisToolStripMenuItem1.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmConsumivel As New frmConsumivel
        frmConsumivel.MdiParent = Me
        frmConsumivel.Show()
        frmConsumivel = Nothing
    End Sub
#End Region

#Region "Comparador"
    Private Sub ComparadorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ComparadorToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmComparados As New frmComparados
        frmComparados.MdiParent = Me
        frmComparados.Show()
        frmComparados = Nothing
    End Sub
#End Region

#Region "SM"
    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmSM As New frmSM
        frmSM.MdiParent = Me
        frmSM.Show()
        frmSM = Nothing
    End Sub
#End Region

#Region "Processo de Fabricação"
    Private Sub ProcessoDeFabricaçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcessoDeFabricaçãoToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmProducao As New frmProcessosFabr
        frmProducao.MdiParent = Me
        frmProducao.Show()
        frmProducao = Nothing
    End Sub
#End Region

#Region "CNC"
    Private Sub CNCToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CNCToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmCNC As New frmCNC
        frmCNC.MdiParent = Me
        frmCNC.Show()
        frmCNC = Nothing
    End Sub
#End Region

#Region "Orçamentos"
    Private Sub OrçamentosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrçamentosToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmOrcamento As New frmOrcamento
        frmOrcamento.MdiParent = Me
        frmOrcamento.Show()
        frmOrcamento = Nothing
    End Sub
#End Region

#Region "Estoque Inicial"
    Private Sub EstoqueInicialToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles EstoqueInicialToolStripMenuItem2.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmEI As New frmEI
        frmEI.MdiParent = Me
        frmEI.Show()
        frmEI = Nothing
    End Sub
#End Region

#Region "Estoque Atual"
    Private Sub EstoqueAtualToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EstoqueAtualToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmEA As New frmEA
        frmEA.MdiParent = Me
        frmEA.Show()
        frmEA = Nothing
    End Sub
#End Region

#Region "Consumos"
    Private Sub ConsumosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsumosToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmConsumos As New frmConsumos
        frmConsumos.MdiParent = Me
        frmConsumos.Show()
        frmConsumos = Nothing
    End Sub
#End Region

#Region "Impostos"
    Private Sub ImpostosEEncargosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImpostosEEncargosToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmImpostos As New frmImpostos
        frmImpostos.MdiParent = Me
        frmImpostos.Show()
        frmImpostos = Nothing
    End Sub
#End Region

#Region "Entrada"
    Private Sub EntradaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EntradaToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmES As New frmES
        frmES.MdiParent = Me
        frmES.Show()
        frmES.Text = "Entrada"
        frmES.Label8.Visible = True
        frmES.ComboBox6.Visible = True
        frmES = Nothing
    End Sub
#End Region

#Region "Saída"
    Private Sub SaídaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaídaToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmES As New frmES
        frmES.MdiParent = Me
        frmES.Show()
        frmES.Text = "Saída"
        frmES = Nothing
    End Sub
#End Region

#Region "Pedido de Venda"
    Private Sub PedidoDeVendaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PedidoDeVendaToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmPedidosCV As New frmPedidosCV
        frmPedidosCV.MdiParent = Me
        frmPedidosCV.Show()
        frmPedidosCV.Text = "Pedidos de vendas"
        frmGerarPedido.Label2.Text = "Cliente:"
        frmPedidosCV = Nothing
    End Sub
#End Region

#Region "Pedido de Compra"
    Private Sub PedidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PedidosToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmPedidosCV As New frmPedidosCV
        frmPedidosCV.Show()
        frmPedidosCV.Text = "Pedidos de compras"
        frmPedidosCV = Nothing
    End Sub
#End Region

#Region "Solicitação de Compra"
    Private Sub SolicitaçõesDeCompraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SolicitaçõesDeCompraToolStripMenuItem.Click
        For Each Form In Me.MdiChildren
            Form.Close()
        Next
        Dim frmPedidosCV As New frmPedidosCV
        frmPedidosCV.MdiParent = Me
        frmPedidosCV.Show()
        frmPedidosCV.Text = "Solicitações de compra"
        frmPedidosCV.Label1.Text = "Nº da solicitação:"
        frmPedidosCV.Button1.Text = "Visualizar Solicitação"
        frmPedidosCV.Button2.Text = "Nova Solicitação"
        frmPedidosCV.ComboBox1.Visible = True
        frmPedidosCV.TextBox3.Visible = False
        frmPedidosCV.Label3.Text = "Necessidade:"
        frmGerarPedido.Text = "Solicitação"
        frmGerarPedido.Label4.Text = "Nº da solicitação:"
        frmPedidosCV.Button5.Visible = True
        frmPedidosCV = Nothing
    End Sub
#End Region

#Region "Backup de banco de dados"
    Private Sub FazerBackupDeBancoDeDadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FazerBackupDeBancoDeDadosToolStripMenuItem.Click

    End Sub
#End Region

#Region "Restauração de banco de dados"
    Private Sub FazerRestauraçãoDeBancoDeDadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FazerRestauraçãoDeBancoDeDadosToolStripMenuItem.Click

    End Sub
#End Region

End Class

