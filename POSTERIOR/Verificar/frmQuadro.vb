Public Class frmQuadro

#Region "ToolTips"
    Private Sub Button2_MouseHover(sender As Object, e As EventArgs) Handles Button2.MouseHover
        ToolTip1.SetToolTip(Button2, "Gerar relatório de quadro")
    End Sub

    Private Sub Button3_MouseHover(sender As Object, e As EventArgs) Handles Button3.MouseHover
        ToolTip1.SetToolTip(Button3, "Excluir funcionário do quadro")
    End Sub
#End Region

#Region "Botão Vermelho"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show("Tem certeza que quer deletar este funcionário?", "Funcionário", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes = True Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("")
            'Executar()
            'Fechar()
            MessageBox.Show("Excluído com sucesso!", "Funcionário", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        End If
    End Sub
#End Region 'VERIFICAR

End Class

'Botão Verde seleciona os dados para aparecerem na tela que chamar esta.