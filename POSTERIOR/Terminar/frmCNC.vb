Public Class frmCNC

    Dim resultado As Boolean

#Region "Escolher Arquivo"
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        OpenFileDialog1.ShowDialog()
        TextBox1.Text = OpenFileDialog1.FileName
    End Sub
#End Region 'OK

#Region "Exibir Relatório"
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start(TextBox1.Text)
    End Sub
#End Region 'OK

#Region "Botão Verde"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT * FROM TBCNCS")
        'Ler()
        'resultado = leitura.HasRows
        'If resultado = True Then
        'If MessageBox.Show("Os dados serão alterados. Deseja continuar?", "CNC", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes = True Then
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("UPDATE TBCNCS SET NOMEPF = '" & ComboBox1.Text & "', TITULOP = '" & ComboBox2.Text & "', RELATORIO = '" & TextBox1.Text & "'")
        'Executar()
        'Fechar()
        'MessageBox.Show("Alterado com sucesso!", "CNC", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'TextBox1.Clear()
        'ComboBox1.Text = ""
        'ComboBox2.Text = ""
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
        'Else
        'TextBox1.Focus()
        'TextBox1.SelectAll()
        'End If
        'Else
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("INSERT INTO TBCNCS (NOMEPF, TITULOP, RELATORIO) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & TextBox1.Text & "')")
        'Executar()
        'Fechar()
        MessageBox.Show("Inserido com sucesso!", "CNC", MessageBoxButtons.OK, MessageBoxIcon.Information)
        TextBox1.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
        'End If
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'VERIFICAR

#Region "Botão Vermelho"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show("Você deseja cancelar a operação?", "CNC", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes = True Then
            Me.Close()
        Else
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Clear()
                End If
            Next
            TextBox1.Focus()
        End If
    End Sub
#End Region 'VERIFICAR

#Region "Carregar ComboBoxes"
    Private Sub frmCNC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Try
        'Conectar()
        'Iniciar()
        'Comandar("SELECT NOME FROM TBPROCESSOSFABR")
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
        'Comandar("SELECT TITULOP FROM TBPECAS")
        'Ler()
        'While leitura.Read
        'ComboBox2.Items.Add(leitura("TITULO"))
        'End While
        'Fechar()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub
#End Region 'OK

End Class

'Haverá verificação para alterar os dados?
'O botão vermelho está ok ou deve exibir a MessageBox?