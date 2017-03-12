﻿Public Class frmMaterial

    Dim medida As Decimal
    Dim medidaprinc As Decimal
    Dim id As Integer

#Region "ToolTip"
    Private Sub Label5_MouseHover(sender As Object, e As EventArgs) Handles Label5.MouseHover
        ToolTip1.SetToolTip(Label5, "Comprimento")
    End Sub
#End Region

#Region "Exibir Imagens"
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            PictureBox2.Image = Image.FromFile(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            PictureBox3.Image = Image.FromFile(OpenFileDialog1.FileName)
        End If
    End Sub
#End Region

#Region "Medidas"
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        TextBox1.Visible = True
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        RadioButton6.Visible = False
        RadioButton7.Visible = False
        RadioButton8.Visible = False
        RadioButton9.Visible = False
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        CheckBox1.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox1.Enabled = True
        CheckBox2.Enabled = False
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        RadioButton6.Visible = True
        RadioButton6.Checked = True
        RadioButton7.Visible = False
        RadioButton8.Visible = False
        RadioButton9.Visible = False
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged

        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False

        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = False
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True

        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = False
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False

        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = False
        TextBox5.Visible = False

        RadioButton6.Visible = True
        RadioButton6.Checked = False
        RadioButton7.Visible = True
        RadioButton8.Visible = False
        RadioButton9.Visible = False

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged

        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox5.Checked = False

        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = False
        CheckBox5.Enabled = True

        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        Label12.Visible = True
        Label13.Visible = False
        Label14.Visible = False

        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        TextBox5.Visible = False

        RadioButton6.Visible = True
        RadioButton6.Checked = False
        RadioButton7.Visible = True
        RadioButton8.Visible = True
        RadioButton9.Visible = False

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged

        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False

        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = False

        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        Label12.Visible = True
        Label13.Visible = True
        Label14.Visible = True

        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        TextBox5.Visible = True

        RadioButton6.Visible = True
        RadioButton7.Visible = True
        RadioButton8.Visible = True
        RadioButton9.Visible = True

    End Sub
#End Region

#Region "Seleção/Alteração de Grupos de Material"
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        TextBox8.ReadOnly = True
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEGRUPOMAT, CAMINHOIMGGRUPOMAT FROM TB_GRUPOSMAT WHERE NOMEGRUPOMAT = '" & ListBox2.SelectedItem & "'")
            Ler()
            While leitura.Read
                If leitura("CAMINHOIMGGRUPOMAT") = "SEM IMAGEM" Then
                    TextBox8.Text = leitura("NOMEGRUPOMAT")
                    MessageBox.Show("Não há imagem cadastrada!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    TextBox8.Text = leitura("NOMEGRUPOMAT")
                    PictureBox1.Image = Image.FromFile(leitura("CAMINHOIMGGRUPOMAT"))
                End If
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        RemoveHandler Button4.Click, AddressOf Button4_Click
        AddHandler Button4.Click, AddressOf AlterarDadosGrupo
    End Sub
#End Region

#Region "Seleção/Alteração de Qualidade"
    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        TextBox9.ReadOnly = True
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEGRUPOMAT, NOMEQUALIDADE, CAMINHOIMGQUALIDADE FROM TB_QUALIDADES WHERE NOMEQUALIDADE = '" & ListBox3.SelectedItem & "'")
            Ler()
            While leitura.Read
                If leitura("CAMINHOIMGQUALIDADE") = "SEM IMAGEM" Then
                    ComboBox4.Text = leitura("NOMEGRUPOMAT")
                    TextBox9.Text = leitura("NOMEQUALIDADE")
                    MessageBox.Show("Não há imagem cadastrada!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ComboBox4.Text = leitura("NOMEGRUPOMAT")
                    TextBox9.Text = leitura("NOMEQUALIDADE")
                    PictureBox2.Image = Image.FromFile(leitura("CAMINHOIMGQUALIDADE"))
                End If
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        RemoveHandler Button4.Click, AddressOf Button4_Click
        AddHandler Button4.Click, AddressOf AlterarDadosQualidade
    End Sub
#End Region

#Region "Seleção/Alteração de Tipo"
    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        TextBox10.ReadOnly = True
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, CAMINHOIMGTIPO FROM TB_TIPOS WHERE NOMETIPO = '" & ListBox4.SelectedItem & "'")
            Ler()
            While leitura.Read
                If leitura("CAMINHOIMGTIPO") = "SEM IMAGEM" Then
                    ComboBox5.Text = leitura("NOMEGRUPOMAT")
                    ComboBox6.Text = leitura("NOMEQUALIDADE")
                    TextBox10.Text = leitura("NOMETIPO")
                    MessageBox.Show("Não há imagem cadastrada!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ComboBox6.Text = leitura("NOMEGRUPOMAT")
                    ComboBox5.Text = leitura("NOMEQUALIDADE")
                    TextBox10.Text = leitura("NOMETIPO")
                    PictureBox3.Image = Image.FromFile(leitura("CAMINHOIMGTIPO"))
                End If
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        RemoveHandler Button4.Click, AddressOf Button4_Click
        AddHandler Button4.Click, AddressOf AlterarDadosTipo
    End Sub
#End Region

#Region "Seleção/Alteração de Material"
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        TextBox6.ReadOnly = True
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox5.ReadOnly = True
        TextBox7.ReadOnly = True
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEMATERIAL, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINCIPAL, KG, NOMEGRUPOMAT, NOMEQUALIDADE, NOMETIPO FROM TB_MATERIAIS WHERE NOMEMATERIAL = '" & ListBox1.SelectedItem & "'")
            Ler()
            While leitura.Read
                ComboBox1.Text = leitura("NOMEGRUPOMAT")
                ComboBox2.Text = leitura("NOMEQUALIDADE")
                ComboBox3.Text = leitura("NOMETIPO")
                TextBox6.Text = leitura("NOMEMATERIAL")
                medida = leitura("QUANTMEDIDAS")
                medidaprinc = leitura("MEDIDAPRINCIPAL")
                If medida = 1 Then
                    CheckBox1.Checked = True
                    TextBox1.Text = leitura("MEDIDA1")
                    TextBox7.Text = leitura("KG")
                ElseIf medida = 2 Then
                    CheckBox2.Checked = True
                    RadioButton6.Checked = True
                    TextBox1.Text = leitura("MEDIDA1")
                    TextBox2.Text = leitura("MEDIDA2")
                    TextBox7.Text = leitura("KG")
                ElseIf medida = 3 Then
                    If medidaprinc = 2 Then
                        CheckBox3.Checked = True
                        RadioButton6.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox7.Text = leitura("KG")
                    ElseIf medidaprinc = 3 Then
                        CheckBox3.Checked = True
                        RadioButton7.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox7.Text = leitura("KG")
                    End If
                ElseIf medida = 4 Then
                    If medidaprinc = 2 Then
                        CheckBox4.Checked = True
                        RadioButton6.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox4.Text = leitura("MEDIDA4")
                        TextBox7.Text = leitura("KG")
                    ElseIf medidaprinc = 3 Then
                        CheckBox4.Checked = True
                        RadioButton7.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox4.Text = leitura("MEDIDA4")
                        TextBox7.Text = leitura("KG")
                    ElseIf medidaprinc = 4 Then
                        CheckBox4.Checked = True
                        RadioButton8.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox4.Text = leitura("MEDIDA4")
                        TextBox7.Text = leitura("KG")
                    End If
                ElseIf medida = 5 Then
                    If medidaprinc = 2 Then
                        CheckBox5.Checked = True
                        RadioButton6.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox4.Text = leitura("MEDIDA4")
                        TextBox5.Text = leitura("MEDIDA5")
                        TextBox7.Text = leitura("KG")
                    ElseIf medidaprinc = 3 Then
                        CheckBox5.Checked = True
                        RadioButton7.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox4.Text = leitura("MEDIDA4")
                        TextBox5.Text = leitura("MEDIDA5")
                        TextBox7.Text = leitura("KG")
                    ElseIf medidaprinc = 4 Then
                        CheckBox5.Checked = True
                        RadioButton8.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox4.Text = leitura("MEDIDA4")
                        TextBox5.Text = leitura("MEDIDA5")
                        TextBox7.Text = leitura("KG")
                    ElseIf medidaprinc = 5 Then
                        CheckBox5.Checked = True
                        RadioButton9.Checked = True
                        TextBox1.Text = leitura("MEDIDA1")
                        TextBox2.Text = leitura("MEDIDA2")
                        TextBox3.Text = leitura("MEDIDA3")
                        TextBox4.Text = leitura("MEDIDA4")
                        TextBox5.Text = leitura("MEDIDA5")
                        TextBox7.Text = leitura("KG")
                    End If
                End If
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        RemoveHandler Button1.Click, AddressOf Button1_Click
        AddHandler Button1.Click, AddressOf AlterarDadosMaterial
    End Sub
#End Region

#Region "Alterar Dados - Grupo de Material"
    Private Sub AlterarDadosGrupo()
        If PictureBox1.Image Is Nothing Then
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT WHERE NOMEGRUPOMAT = '" & ListBox2.SelectedIndex & "'")
                Ler()
                While leitura.Read
                    id = leitura("NOMEGRUPOMAT")
                End While
                Fechar()
                Conectar()
                Iniciar()
                Comandar("UPDATE TB_GRUPOSMAT SET NOMEGRUPOMAT = '" & TextBox8.Text & "' WHERE NOMEGRUPOMAT = '" & id & "'")
                Executar()
                Fechar()
                MessageBox.Show("Alterado com sucesso!", "Entrada", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox8.Clear()
                PictureBox1.Image = Nothing
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            ListBox2.Items.Clear()
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                Ler()
                While leitura.Read
                    ListBox2.Items.Add(leitura("NOMEGRUPOMAT"))
                End While
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            Try
                Conectar()
                Iniciar()
                Comandar("UPDATE TB_GRUPOSMAT SET NOMEGRUPOMAT = '" & TextBox8.Text & "' AND CAMINHOIMGTIPO = '" & OpenFileDialog1.FileName & "' WHERE NOMEGRUPOMAT = '" & ListBox2.SelectedIndex & "'")
                Executar()
                Fechar()
                MessageBox.Show("Alterado com sucesso!", "Entrada", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox8.Clear()
                PictureBox1.Image = Nothing
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            ListBox2.Items.Clear()
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                Ler()
                While leitura.Read
                    ListBox2.Items.Add(leitura("NOMEGRUPOMAT"))
                End While
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub
#End Region 'VERIFICAR

#Region "Alterar Dados - Qualidade"
    Private Sub AlterarDadosQualidade()
        Try
            Conectar()
            Iniciar()
            Comandar("UPDATE TBQUALIDADES SET NOMEGRUPOMAT = '" & ComboBox4.Text & "' AND NOMEQ = '" & TextBox9.Text & "' AND IMAGEM = '" & OpenFileDialog1.FileName & "' WHERE NOMEQ = '" & ListBox3.SelectedIndex & "'")
            Executar()
            Fechar()
            MessageBox.Show("Alterado com sucesso!", "Entrada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox4.Text = ""
            TextBox9.Clear()
            PictureBox2.Image = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region 'VERIFICAR

#Region "Alterar Dados - Tipo"
    Private Sub AlterarDadosTipo()
        Try
            Conectar()
            Iniciar()
            Comandar("UPDATE TBTIPOS SET NOMEGRUPOMAT = '" & ComboBox6.Text & "' AND NOMEQ = '" & ComboBox5.Text & "' AND NOMET = '" & TextBox10.Text & "' AND IMAGEM = '" & OpenFileDialog1.FileName & "' WHERE NOMET = '" & ListBox4.SelectedIndex & "'")
            Executar()
            Fechar()
            MessageBox.Show("Alterado com sucesso!", "Entrada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox6.Text = ""
            ComboBox5.Text = ""
            TextBox10.Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region 'VERIFICAR

#Region "Alterar Dados - Material"
    Private Sub AlterarDadosMaterial()
        If CheckBox1.Checked = True Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT = '" & ComboBox1.Text & "', NOMEQ = '" & ComboBox2.Text & "', NOMET = '" & ComboBox3.Text & "', NOMEMAT = '" & TextBox6.Text & "', QUANTMEDIDAS = '" & CheckBox1.Text & "', MEDIDA1 = '" & TextBox1.Text & "', MEDIDAPRINC = 1, KG = '" & TextBox7.Text & "'")
            'Executar()
            'Fechar()
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf CheckBox2.Checked = True Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT = '" & ComboBox1.Text & "', NOMEQ = '" & ComboBox2.Text & "', NOMET = '" & ComboBox3.Text & "', NOMEMAT = '" & TextBox6.Text & "', QUANTMEDIDAS = '" & CheckBox2.Text & "', MEDIDA1 = '" & TextBox1.Text & "', MEDIDA2 = '" & TextBox2.Text & "', MEDIDAPRINC = 2, KG = '" & TextBox7.Text & "'")
            'Executar()
            'Fechar()
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf CheckBox3.Checked = True Then
            If RadioButton6.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT = '" & ComboBox1.Text & "', NOMEQ = '" & ComboBox2.Text & "', NOMET = '" & ComboBox3.Text & "', NOMEMAT = '" & TextBox6.Text & "', QUANTMEDIDAS = '" & CheckBox3.Text & "', MEDIDA1 = '" & TextBox1.Text & "', MEDIDA2 = '" & TextBox2.Text & "', MEDIDA3 = '" & TextBox3.Text & "', MEDIDAPRINC = 2, KG = '" & TextBox7.Text & "'")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton7.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT = '" & ComboBox1.Text & "', NOMEQ = '" & ComboBox2.Text & "', NOMET = '" & ComboBox3.Text & "', NOMEMAT = '" & TextBox6.Text & "', QUANTMEDIDAS '" & CheckBox3.Text & "', MEDIDA1 = '" & TextBox1.Text & "', MEDIDA2 = '" & TextBox2.Text & "', MEDIDA3 = '" & TextBox3.Text & "', MEDIDAPRINC = 3, KG = '" & TextBox7.Text & "'")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton8.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT = '" & ComboBox1.Text & "', NOMEQ = '" & ComboBox2.Text & "', NOMET = '" & ComboBox3.Text & "', NOMEMAT = '" & TextBox6.Text & "', QUANTMEDIDAS = '" & CheckBox3.Text & "', MEDIDA1 = '" & TextBox1.Text & "', MEDIDA2 = '" & TextBox2.Text & "', MEDIDA3 = '" & TextBox3.Text & "', MEDIDAPRINC = 4, KG = '" & TextBox7.Text & "'")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton9.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT = '" & ComboBox1.Text & "', NOMEQ = '" & ComboBox2.Text & "', NOMET = '" & ComboBox3.Text & "', NOMEMAT = '" & TextBox6.Text & "', QUANTMEDIDAS = '" & CheckBox3.Text & "', MEDIDA1 = '" & TextBox1.Text & "', MEDIDA2 = '" & TextBox2.Text & "', MEDIDA3 = '" & TextBox3.Text & "', MEDIDAPRINC = 5, KG = '" & TextBox7.Text & "'")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            End If
        ElseIf CheckBox4.Checked = True Then
            If RadioButton6.Checked = True Then
                '----------------- PAREI AQUI

                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDAPRINC, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox4.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', 2, '" & TextBox7.Text & "')")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton7.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDAPRINC, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox4.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', 3, '" & TextBox7.Text & "')")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton8.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDAPRINC, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox4.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', 4, '" & TextBox7.Text & "')")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton9.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET (NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDAPRINC, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox4.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', 5, '" & TextBox7.Text & "')")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            End If
        ElseIf CheckBox5.Checked = True Then
            If RadioButton6.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINC, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox5.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "', 2, '" & TextBox7.Text & "')")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton7.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINC, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox5.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "', 3, '" & TextBox7.Text & "')")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton8.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINC, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox5.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "', 4, '" & TextBox7.Text & "')")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            ElseIf RadioButton9.Checked = True Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("UPDATE TBMATERIAIS SET NOMEGRUPOMAT, NOMEQ, NOMET, NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINC, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox5.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "', 5, '" & TextBox7.Text & "')")
                'Executar()
                'Fechar()
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
            End If
        End If
    End Sub
#End Region 'VERIFICAR

#Region "Botão Verde - Grupo de Material"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT WHERE NOMEGRUPOMAT = '" & TextBox8.Text & "'")
            Ler()
            If leitura.HasRows Then
                MessageBox.Show("Grupo de material já cadastrado!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TextBox8.Focus()
                TextBox8.SelectAll()
                Fechar()
            Else
                If PictureBox1.Image Is Nothing Then
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_GRUPOSMAT (NOMEGRUPOMAT, CAMINHOIMGGRUPOMAT) VALUES ('" & TextBox8.Text & "','SEM IMAGEM')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Grupo cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        PictureBox1.Image = Nothing
                        TextBox8.Clear()
                        TextBox8.Focus()
                        ListBox2.Items.Clear()
                        Try
                            Conectar()
                            Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                            Ler()
                            While leitura.Read
                                ListBox2.Items.Add(leitura("NOMEGRUPOMAT"))
                            End While
                            Fechar()
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_GRUPOSMAT (NOMEGRUPOMAT, CAMINHOIMGGRUPOMAT) VALUES ('" & TextBox8.Text & "','" & OpenFileDialog1.FileName & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Grupo cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        PictureBox1.Image = Nothing
                        TextBox8.Clear()
                        TextBox8.Focus()
                        ListBox2.Items.Clear()
                        Try
                            Conectar()
                            Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                            Ler()
                            While leitura.Read
                                ListBox2.Items.Add(leitura("NOMEGRUPOMAT"))
                            End While
                            Fechar()
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Botão Verde - Qualidade"
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEGRUPOMAT, NOMEQUALIDADE FROM TB_QUALIDADES WHERE NOMEQUALIDADE = '" & TextBox9.Text & "'")
            Ler()
            If leitura.HasRows Then
                MessageBox.Show("Qualidade já cadastrada!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TextBox8.Focus()
                TextBox8.SelectAll()
                Fechar()
            Else
                If PictureBox2.Image Is Nothing Then
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_QUALIDADES (NOMEGRUPOMAT, NOMEQUALIDADE, CAMINHOIMGQUALIDADE) VALUES ('" & ComboBox4.Text & "','" & TextBox9.Text & "','SEM IMAGEM')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Qualidade cadastrada com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        PictureBox2.Image = Nothing
                        ComboBox4.Text = "(Selecione...)"
                        TextBox9.Clear()
                        ComboBox4.Focus()
                        ListBox3.Items.Clear()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_QUALIDADES (NOMEGRUPOMAT, NOMEQUALIDADE, CAMINHOIMGQUALIDADE) VALUES ('" & ComboBox4.Text & "','" & TextBox9.Text & "','" & OpenFileDialog1.FileName & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Qualidade cadastrada com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        PictureBox2.Image = Nothing
                        ComboBox4.Text = "(Selecione...)"
                        TextBox9.Clear()
                        ComboBox4.Focus()
                        ListBox3.Items.Clear()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Botão Verde - Tipo"
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If PictureBox3.Image Is Nothing Then
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT, NOMEQUALIDADE, NOMETIPO FROM TB_TIPOS WHERE NOMETIPO = '" & TextBox10.Text & "'")
                Ler()
                If leitura.HasRows Then
                    MessageBox.Show("Tipo já cadastrado!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    TextBox8.Focus()
                    TextBox8.SelectAll()
                    Fechar()
                Else
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_TIPOS (NOMEGRUPOMAT, NOMEQUALIDADE, NOMETIPO, CAMINHOIMGTIPO) VALUES ('" & ComboBox6.Text & "','" & ComboBox5.Text & "','" & TextBox10.Text & "','SEM IMAGEM')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Tipo cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        PictureBox3.Image = Nothing
                        ComboBox6.Text = "(Selecione...)"
                        ComboBox5.Text = "(Selecione...)"
                        TextBox10.Clear()
                        ComboBox6.Focus()
                        ListBox4.Items.Clear()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT, NOMEQUALIDADE, NOMETIPO FROM TB_TIPOS WHERE NOMETIPO = '" & TextBox10.Text & "'")
                Ler()
                If leitura.HasRows Then
                    MessageBox.Show("Tipo já cadastrado!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    TextBox8.Focus()
                    TextBox8.SelectAll()
                    Fechar()
                Else
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_TIPOS (NOMEGRUPOMAT, NOMEQUALIDADE, NOMETIPO, CAMINHOIMGTIPO) VALUES ('" & ComboBox6.Text & "','" & ComboBox5.Text & "','" & TextBox10.Text & "','" & OpenFileDialog1.FileName & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Tipo cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        PictureBox3.Image = Nothing
                        ComboBox6.Text = "(Selecione...)"
                        ComboBox5.Text = "(Selecione...)"
                        TextBox10.Clear()
                        ComboBox6.Focus()
                        ListBox4.Items.Clear()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub
#End Region

#Region "Botão Verde - Material"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL FROM TB_MATERIAIS WHERE NOMEMATERIAL = '" & TextBox6.Text & "'")
            Ler()
            If leitura.HasRows Then
                MessageBox.Show("Material já cadastrado!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TextBox8.Focus()
                TextBox8.SelectAll()
                Fechar()
            Else
                If CheckBox1.Checked = True Then
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, QUANTMEDIDAS, MEDIDA1, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "', 1,'" & TextBox1.Text & "', 1, '" & TextBox7.Text & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ComboBox2.Text = "(Selecione...)"
                        ComboBox1.Text = "(Selecione...)"
                        ComboBox3.Text = "(Selecione...)"
                        TextBox6.Clear()
                        CheckBox1.Checked = False
                        TextBox7.Clear()
                        ComboBox1.Focus()
                        ListBox1.Items.Clear()
                        TextBox1.Clear()
                        Try
                            Conectar()
                            Iniciar()
                            Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                            Ler()
                            While leitura.Read
                                ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                            End While
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                ElseIf CheckBox2.Checked = True Then
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox2.Text & "','" & TextBox1.Text & "', 2, '" & TextBox7.Text & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ComboBox2.Text = "(Selecione...)"
                        ComboBox1.Text = "(Selecione...)"
                        ComboBox3.Text = "(Selecione...)"
                        TextBox6.Clear()
                        CheckBox2.Checked = False
                        RadioButton6.Checked = False
                        TextBox7.Clear()
                        ComboBox2.Focus()
                        ListBox1.Items.Clear()
                        Try
                            Conectar()
                            Iniciar()
                            Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                            Ler()
                            While leitura.Read
                                ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                            End While
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                ElseIf CheckBox3.Checked = True Then
                    If RadioButton6.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox3.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "', 2, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox3.Checked = False
                            RadioButton6.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    ElseIf RadioButton7.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox3.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "', 3, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox3.Checked = False
                            RadioButton7.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    End If
                ElseIf CheckBox4.Checked = True Then
                    If RadioButton6.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox3.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "', 2, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox4.Checked = False
                            RadioButton6.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    ElseIf RadioButton7.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox3.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "', 3, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox4.Checked = False
                            RadioButton7.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    ElseIf RadioButton8.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox3.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "', 4, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox4.Checked = False
                            RadioButton8.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    End If
                ElseIf CheckBox5.Checked = True Then
                    If RadioButton6.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox5.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', 2, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox5.Checked = False
                            RadioButton6.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    ElseIf RadioButton7.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox5.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', 3, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox5.Checked = False
                            RadioButton7.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    ElseIf RadioButton8.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEGRUPOMAT, NOMEQUALIDADE, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox5.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', 4, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox5.Checked = False
                            RadioButton8.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    ElseIf RadioButton9.Checked = True Then
                        Try
                            Fechar()
                            Conectar()
                            Iniciar()
                            Comandar("INSERT INTO TB_MATERIAIS (NOMEGRUPOMAT, NOMEQUALIDADE, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox5.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', 5, '" & TextBox7.Text & "')")
                            Executar()
                            Fechar()
                            MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ComboBox2.Text = "(Selecione...)"
                            ComboBox1.Text = "(Selecione...)"
                            ComboBox3.Text = "(Selecione...)"
                            TextBox6.Clear()
                            CheckBox5.Checked = False
                            RadioButton9.Checked = False
                            TextBox7.Clear()
                            ComboBox2.Focus()
                            ListBox1.Items.Clear()
                            Try
                                Conectar()
                                Iniciar()
                                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                                Ler()
                                While leitura.Read
                                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                                End While
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Botão Vermelho - Grupo de Material"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frmMessageBox.Text = "Materiais - Grupos"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region

#Region "Botão Vermelho - Qualidade"
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        frmMessageBox.Text = "Materiais - Qualidade"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region

#Region "Botão Vermelho - Tipo"
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        frmMessageBox.Text = "Materiais - Tipo"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region

#Region "Botão Vermelho - Material"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmMessageBox.Text = "Materiais"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region

#Region "TextBoxes Editáveis"
    Private Sub TextBox8_Click(sender As Object, e As EventArgs) Handles TextBox8.Click
        TextBox8.ReadOnly = False
    End Sub

    Private Sub TextBox9_Click(sender As Object, e As EventArgs) Handles TextBox9.Click
        TextBox9.ReadOnly = False
    End Sub

    Private Sub TextBox10_Click(sender As Object, e As EventArgs) Handles TextBox10.Click
        TextBox10.ReadOnly = False
    End Sub

    Private Sub TextBox6_Click(sender As Object, e As EventArgs) Handles TextBox6.Click
        TextBox6.ReadOnly = False
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        TextBox1.ReadOnly = False
    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        TextBox2.ReadOnly = False
    End Sub

    Private Sub TextBox3_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        TextBox3.ReadOnly = False
    End Sub

    Private Sub TextBox4_Click(sender As Object, e As EventArgs) Handles TextBox4.Click
        TextBox4.ReadOnly = False
    End Sub

    Private Sub TextBox5_Click(sender As Object, e As EventArgs) Handles TextBox5.Click
        TextBox5.ReadOnly = False
    End Sub

    Private Sub TextBox7_Click(sender As Object, e As EventArgs) Handles TextBox7.Click
        TextBox7.ReadOnly = False
    End Sub
#End Region

#Region "Verificação de Dados"
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            ListBox2.Items.Clear()
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                Ler()
                Dim resultado As Boolean = leitura.HasRows
                If resultado = True Then
                    While leitura.Read
                        ListBox2.Items.Add(leitura("NOMEGRUPOMAT"))
                    End While
                Else
                    MessageBox.Show("Ainda não há grupos de material cadastrados!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ListBox2.Items.Add("Não há grupos de material cadastrados!")
                End If
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf TabControl1.SelectedIndex = 1 Then
            ListBox3.Items.Clear()
            ComboBox4.Items.Clear()
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                Ler()
                While leitura.Read
                    ComboBox4.Items.Add(leitura("NOMEGRUPOMAT"))
                End While
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEQUALIDADE FROM TB_QUALIDADES")
                Ler()
                Dim resultado As Boolean = leitura.HasRows
                If resultado = False Then
                    MessageBox.Show("Ainda não há qualidades cadastradas!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ListBox1.Items.Add("Não há qualidades cadastradas!")
                End If
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf TabControl1.SelectedIndex = 2 Then
            ListBox4.Items.Clear()
            ComboBox6.Items.Clear()
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                Ler()
                While leitura.Read
                    ComboBox6.Items.Add(leitura("NOMEGRUPOMAT"))
                End While
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMETIPO FROM TB_TIPOS")
                Ler()
                Dim resultado As Boolean = leitura.HasRows
                If resultado = False Then
                    MessageBox.Show("Ainda não há tipos cadastrados!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ListBox1.Items.Add("Não há tipos cadastrados!")
                End If
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf TabControl1.SelectedIndex = 3 Then
            ListBox1.Items.Clear()
            ComboBox1.Items.Clear()
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                Ler()
                While leitura.Read
                    ComboBox1.Items.Add(leitura("NOMEGRUPOMAT"))
                End While
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
                Ler()
                Dim resultado As Boolean = leitura.HasRows
                If resultado = False Then
                    MessageBox.Show("Ainda não há materiais cadastrados!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ListBox1.Items.Add("Não há materiais cadastrados!")
                End If
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub
#End Region

#Region "Carregar Combo Qualidade - Tipo"
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        ListBox4.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEQUALIDADE FROM TB_QUALIDADES WHERE NOMEGRUPOMAT = '" & ComboBox6.Text & "'")
            Ler()
            While leitura.Read
                ComboBox5.Items.Add(leitura("NOMEQUALIDADE"))
                ListBox4.Items.Add(leitura("NOMEQUALIDADE"))
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Carregar Lista - Tipo"
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        ListBox4.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMETIPO FROM TB_TIPOS WHERE NOMEQUALIDADE = '" & ComboBox5.Text & "'")
            Ler()
            While leitura.Read
                ComboBox5.Items.Add(leitura("NOMETIPO"))
                ListBox4.Items.Add(leitura("NOMETIPO"))
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Carregar Lista - Qualidade"
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        ListBox3.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEQUALIDADE FROM TB_QUALIDADES WHERE NOMEGRUPOMAT = '" & ComboBox4.Text & "'")
            Ler()
            While leitura.Read
                ComboBox5.Items.Add(leitura("NOMEQUALIDADE"))
                ListBox3.Items.Add(leitura("NOMEQUALIDADE"))
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Correção de Erros"
    Private Sub ComboBox6_Click(sender As Object, e As EventArgs) Handles ComboBox6.Click
        ListBox4.Items.Clear()
        ComboBox5.Text = "(Selecione...)"
        TextBox10.Clear()
        PictureBox3.Image = Nothing
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        ComboBox2.Text = "(Selecione...)"
        ComboBox3.Text = "(Selecione...)"
        ListBox1.Items.Clear()
        TextBox6.Clear()
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
        RadioButton8.Checked = False
        RadioButton9.Checked = False
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox7.Clear()
    End Sub

    Private Sub ComboBox2_Click(sender As Object, e As EventArgs) Handles ComboBox2.Click
        ComboBox3.Text = "(Selecione...)"
        ListBox1.Items.Clear()
        TextBox6.Clear()
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
        RadioButton8.Checked = False
        RadioButton9.Checked = False
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox7.Clear()
    End Sub

    Private Sub ComboBox4_Click(sender As Object, e As EventArgs) Handles ComboBox4.Click
        TextBox9.Clear()
        PictureBox1.Image = Nothing
        ListBox3.Items.Clear()
    End Sub
#End Region

#Region "Carregar Combo Tipo - Material"
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ListBox1.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMETIPO FROM TB_TIPOS WHERE NOMEQUALIDADE = '" & ComboBox2.Text & "' AND NOMEGRUPOMAT = '" & ComboBox1.Text & "'")
            Ler()
            While leitura.Read
                ComboBox3.Items.Add(leitura("NOMETIPO"))
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Carregar Lista - Material"
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ListBox1.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS WHERE NOMEGRUPOMAT = '" & ComboBox1.Text & "' AND NOMEQUALIDADE = '" & ComboBox2.Text & "' AND NOMETIPO = '" & ComboBox3.Text & "'")
            Ler()
            While leitura.Read
                ListBox1.Items.Add(leitura("NOMEMATERIAL"))
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Carregar Combo Qualidade - Material"
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ListBox1.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEQUALIDADE FROM TB_QUALIDADES WHERE NOMEGRUPOMAT = '" & ComboBox1.Text & "'")
            Ler()
            While leitura.Read
                ComboBox2.Items.Add(leitura("NOMEQUALIDADE"))
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

End Class