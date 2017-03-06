Public Class frmES

#Region "Configurações"
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Material" Then
            GroupBox1.Text = "Informações do material:"
            Label2.Visible = True
            Label3.Visible = True
            Label6.Visible = True
            Label7.Visible = True
            ComboBox2.Visible = True
            ComboBox3.Visible = True
            ComboBox4.Visible = True
            ComboBox5.Visible = True
            Label2.Text = "Grupo de Material:"
            Label3.Text = "Qualidade:"
            Label6.Text = "Tipo:"
            Label7.Text = "Material:"
        ElseIf ComboBox1.Text = "Montagem" Then
            GroupBox1.Text = "Informações da montagem:"
            Label2.Visible = True
            Label3.Visible = True
            Label6.Visible = False
            Label7.Visible = False
            ComboBox2.Visible = True
            ComboBox3.Visible = True
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            Label2.Text = "Código:"
            Label3.Text = "Titulo:"
        ElseIf ComboBox1.Text = "Peça" Then
            GroupBox1.Text = "Informações da peça:"
            Label2.Visible = True
            Label3.Visible = True
            Label6.Visible = False
            Label7.Visible = False
            ComboBox2.Visible = True
            ComboBox3.Visible = True
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            Label2.Text = "Código:"
            Label3.Text = "Titulo:"
        ElseIf ComboBox1.Text = "Consumível" Then
            GroupBox1.Text = "Informações do consumível:"
            Label2.Visible = True
            Label3.Visible = True
            Label6.Visible = False
            Label7.Visible = False
            ComboBox2.Visible = True
            ComboBox3.Visible = True
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            Label2.Text = "Código:"
            Label3.Text = "Produto:"
        ElseIf ComboBox1.Text = "Produto" Then
            GroupBox1.Text = "Informações do produto:"
            Label2.Visible = True
            Label3.Visible = True
            Label6.Visible = False
            Label7.Visible = False
            ComboBox2.Visible = True
            ComboBox3.Visible = True
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            Label2.Text = "Código:"
            Label3.Text = "Titulo:"
        End If
    End Sub
#End Region 'OK

#Region "Selecionar NF"
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        OpenFileDialog1.ShowDialog()
        TextBox7.Text = OpenFileDialog1.FileName
    End Sub
#End Region 'OK

#Region "Exibir NF"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Process.Start(TextBox7.Text)
    End Sub
#End Region 'OK

#Region "Calendário"
    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        TextBox3.Text = MonthCalendar1.SelectionRange.Start.ToShortDateString
    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox3.Text = MonthCalendar1.SelectionRange.Start.ToShortDateString
    End Sub

    Private Sub TextBox3_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox3.MouseClick
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox4_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox4.MouseClick
        MonthCalendar1.Visible = False
    End Sub

    Private Sub TextBox7_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox7.MouseClick
        MonthCalendar1.Visible = False
    End Sub
#End Region 'OK

End Class