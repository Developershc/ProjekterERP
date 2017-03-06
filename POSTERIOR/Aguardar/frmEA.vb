Public Class frmEA

#Region "ToolTip"
    Private Sub Button2_MouseHover(sender As Object, e As EventArgs) Handles Button2.MouseHover
        ToolTip1.SetToolTip(Button2, "Relatório de estoque atual")
    End Sub
#End Region

End Class

'Dados serão alterados diretamente no DataGrid.
'Apenas consulta.