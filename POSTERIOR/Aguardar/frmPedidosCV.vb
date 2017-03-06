Public Class frmPedidosCV

#Region "Calendário"
    Private Sub TextBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox2.MouseClick
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox3_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox3.MouseClick
        MonthCalendar1.Visible = False
    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick
        MonthCalendar1.Visible = False
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        MonthCalendar1.Visible = False
    End Sub
#End Region 'OK

#Region "Configurações"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frmGerarPedido.Show()
        frmGerarPedido.Text = "Pedido"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        frmGerarPedido.Show()
        frmGerarPedido.Text = "Pedido"
        frmGerarPedido.Label4.Text = "Nº do pedido:"
    End Sub
#End Region 'OK

End Class