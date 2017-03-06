Public Class frmGerarPedido

#Region "Calendário"
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        MonthCalendar1.Visible = False
    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick
        MonthCalendar1.Visible = True
    End Sub

    Private Sub ComboBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ComboBox1.MouseClick
        MonthCalendar1.Visible = False
    End Sub

    Private Sub ComboBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles ComboBox2.MouseClick
        MonthCalendar1.Visible = False
    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        TextBox1.Text = MonthCalendar1.SelectionRange.Start.ToShortDateString
    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox1.Text = MonthCalendar1.SelectionRange.Start.ToShortDateString
    End Sub

    Private Sub TextBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox2.MouseClick
        MonthCalendar1.Visible = False
    End Sub
#End Region

#Region "ToolTip"
    Private Sub Button2_MouseHover(sender As Object, e As EventArgs) Handles Button2.MouseHover
        ToolTip1.SetToolTip(Button2, "Gerar pedido")
    End Sub
#End Region

#Region "Botão Vermelho"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
#End Region 'FAZER CÓDIGO

#Region "Botão Verde"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
#End Region 'FAZER CÓDIGO

End Class

'Pedido pode ser alterado direto no DataGrid.
'DataGrid recebe todos os pedidos selecionados.
'Botão Gerar Pedido gera um relatório com os dados do pedido. 