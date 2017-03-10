Public Class frmMaterial

    'A lista deve ser carregada a partir do Grupo de Material/Tipo/Qualidade escolhida.

    Dim medida As Decimal
    Dim medidaprinc As Decimal

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

#Region "Carregar Listas e Combos"
    Private Sub frmMaterial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEMATERIAL FROM TB_MATERIAIS")
            Ler()
            Dim resultado As Boolean = leitura.HasRows
            If resultado = True Then
                While leitura.Read
                    ListBox1.Items.Add(leitura("NOMEMATERIAL"))
                End While
            Else
                MessageBox.Show("Ainda não há materiais cadastrados!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ListBox1.Items.Add("Não há materiais cadastrados!")
            End If
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        'ListBox2.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
            Ler()
            Dim resultado As Boolean = leitura.HasRows
            If resultado = True Then
                While leitura.Read
                    ListBox2.Items.Add(leitura("NOMEGRUPOMAT"))
                    ComboBox4.Items.Add(leitura("NOMEGRUPOMAT"))
                    ComboBox6.Items.Add(leitura("NOMEGRUPOMAT"))
                    ComboBox1.Items.Add(leitura("NOMEGRUPOMAT"))
                End While
            Else
                MessageBox.Show("Ainda não há grupos de material cadastrados!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        ListBox3.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEQUALIDADE FROM TB_QUALIDADES")
            Ler()
            Dim resultado As Boolean = leitura.HasRows
            If resultado = True Then
                While leitura.Read
                    ListBox3.Items.Add(leitura("NOMEQUALIDADE"))
                    ComboBox5.Items.Add(leitura("NOMEQUALIDADE"))
                    ComboBox2.Items.Add(leitura("NOMEQUALIDADE"))
                End While
            Else
                MessageBox.Show("Ainda não há qualidades cadastradas!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ListBox3.Items.Add("Não há qualidade cadastradas!")
            End If
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        ListBox4.Items.Clear()
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMETIPO FROM TB_TIPOS")
            Ler()
            Dim resultado As Boolean = leitura.HasRows
            If resultado = True Then
                While leitura.Read
                    ListBox4.Items.Add(leitura("NOMETIPO"))
                    ComboBox3.Items.Add(leitura("NOMETIPO"))
                End While
            Else
                ListBox4.Items.Add("Não há tipos cadastrados!")
            End If
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Seleção/Alteração de Grupos de Material"
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEGRUPOMAT, CAMINHOIMGGRUPOMAT FROM TB_GRUPOSMAT WHERE NOMEGRUPOMAT = '" & ListBox2.SelectedItem & "'")
            Ler()
            While leitura.Read
                TextBox8.Text = leitura("NOMEGRUPOMAT")
                PictureBox1.Image = Image.FromFile(leitura("CAMINHOIMGGRUPOMAT"))
            End While
            Fechar()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Seleção/Alteração de Qualidade"
    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEGRUPOMAT, NOMEQUALIDADE, CAMINHOIMGQUALIDADE FROM TB_QUALIDADES WHERE NOMEQUALIDADE = '" & ListBox3.SelectedItem & "'")
            Ler()
            While leitura.Read
                If leitura("NOMEGRUPOMAT").ToString Is DBNull.Value Then
                    ComboBox4.Text = "NÃO POSSUI"
                    TextBox9.Text = leitura("NOMEQUALIDADE")
                    PictureBox2.Image = Image.FromFile(leitura("CAMINHOIMGQUALIDADE"))
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
    End Sub
#End Region

#Region "Seleção/Alteração de Tipo"
    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        Try
            Conectar()
            Iniciar()
            Comandar("SELECT NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, CAMINHOIMGTIPO FROM TB_TIPOS WHERE NOMETIPO = '" & ListBox4.SelectedItem & "'")
            Ler()
            While leitura.Read
                If leitura("NOMEGRUPOMAT").ToString Is DBNull.Value And leitura("NOMEQUALIDADE").ToString Is DBNull.Value Then
                    ComboBox5.Text = "NÃO POSSUI"
                    ComboBox6.Text = "NÃO POSSUI"
                    TextBox10.Text = leitura("NOMETIPO")
                    PictureBox3.Image = Image.FromFile(leitura("CAMINHOIMGTIPO"))
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
    End Sub
#End Region

#Region "Alteração de Material"
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
            Comandar("SELECT NOMEMAT, QUANTMEDIDAS, MEDIDA1, MEDIDA2, MEDIDA3, MEDIDA4, MEDIDA5, MEDIDAPRINC, KG, NOMEGRUPOMAT, NOMEQ, NOMET FROM TBMATERIAIS WHERE NOMEMAT = '" & ListBox1.SelectedIndex & "'")
            Ler()
            While leitura.Read
                ComboBox1.Text = leitura("NOMEGRUPOMAT")
                ComboBox2.Text = leitura("NOMEQ")
                ComboBox3.Text = leitura("NOMET")
                TextBox6.Text = leitura("NOMEMAT")
                medida = leitura("QUANTMEDIDAS")
                medidaprinc = leitura("MEDIDAPRINC")
                If medida = 1 Then
                    CheckBox1.Checked = True
                    TextBox1.Text = leitura("MEDIDA1")
                    TextBox7.Text = leitura("KD")
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
#End Region 'TESTAR

#Region "Alterar Dados - Grupo de Material"
    Private Sub AlterarDadosGrupo()
        Try
            Conectar()
            Iniciar()
            Comandar("UPDATE TBGRUPOSMAT SET NOMEGRUPOMAT = '" & TextBox8.Text & "' AND IMAGEM = '" & OpenFileDialog1.FileName & "' WHERE NOMEGRUPOMAT = '" & ListBox2.SelectedIndex & "'")
            Executar()
            Fechar()
            MessageBox.Show("Alterado com sucesso!", "Entrada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox8.Clear()
            PictureBox1.Image = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region 'TESTAR

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
#End Region 'TESTAR

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
#End Region 'TESTAR

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
                Try
                    Fechar()
                    Conectar()
                    Iniciar()
                    Comandar("INSERT INTO TB_QUALIDADES (NOMEGRUPOMAT, NOMEQUALIDADE, CAMINHOIMGQUALIDADE) VALUES ('" & ComboBox4.Text & "','" & TextBox9.Text & "','" & OpenFileDialog1.FileName & "')")
                    Executar()
                    Fechar()
                    MessageBox.Show("Qualidade cadastrada com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    PictureBox2.Image = Nothing
                    ComboBox4.Text = ""
                    TextBox9.Clear()
                    ComboBox4.Focus()
                    ListBox3.Items.Clear()
                    Try
                        Conectar()
                        Iniciar()
                        Comandar("SELECT NOMEQUALIDADE FROM TB_QUALIDADES")
                        Ler()
                        While leitura.Read
                            ListBox3.Items.Add(leitura("NOMEQUALIDADE"))
                        End While
                        Fechar()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Botão Verde - Tipo"
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
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
                    ComboBox6.Text = ""
                    ComboBox5.Text = ""
                    TextBox10.Clear()
                    ComboBox6.Focus()
                    ListBox4.Items.Clear()
                    Try
                        Conectar()
                        Iniciar()
                        Comandar("SELECT NOMETIPO FROM TB_TIPOS")
                        Ler()
                        While leitura.Read
                            ListBox4.Items.Add(leitura("NOMETIPO"))
                        End While
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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
                        Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox1.Text & "', 1, '" & TextBox7.Text & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ComboBox2.Text = ""
                        ComboBox1.Text = ""
                        ComboBox3.Text = ""
                        TextBox6.Clear()
                        CheckBox1.Checked = False
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
                ElseIf CheckBox2.Checked = True Then
                    Try
                        Fechar()
                        Conectar()
                        Iniciar()
                        Comandar("INSERT INTO TB_MATERIAIS (NOMEQUALIDADE, NOMEGRUPOMAT, NOMETIPO, NOMEMATERIAL, MEDIDA1, MEDIDA2, MEDIDAPRINCIPAL, KG) VALUES ('" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox3.Text & "','" & TextBox6.Text & "','" & CheckBox2.Text & "','" & TextBox1.Text & "', 2, '" & TextBox7.Text & "')")
                        Executar()
                        Fechar()
                        MessageBox.Show("Material cadastrado com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ComboBox2.Text = ""
                        ComboBox1.Text = ""
                        ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
                            ComboBox2.Text = ""
                            ComboBox1.Text = ""
                            ComboBox3.Text = ""
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
#End Region 'VERIFICAR

#Region "Botão Vermelho - Qualidade"
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        frmMessageBox.Text = "Materiais - Qualidade"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region 'VERIFICAR

#Region "Botão Vermelho - Tipo"
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        frmMessageBox.Text = "Materiais - Tipo"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region 'VERIFICAR

#Region "Botão Vermelho - Material"
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmMessageBox.Text = "Materiais"
        frmMessageBox.Button1.Text = "Excluir"
        frmMessageBox.Button2.Text = "Cancelar"
        frmMessageBox.Label1.Text = "Você deseja excluir esses dados ou cancelar a operação?"
        frmMessageBox.Show()
    End Sub
#End Region 'VERIFICAR

#Region "TextBoxes Editáveis"
    Private Sub TextBox8_Click(sender As Object, e As EventArgs) Handles TextBox8.Click
        'TextBox8.ReadOnly = False
    End Sub

    Private Sub TextBox9_Click(sender As Object, e As EventArgs) Handles TextBox9.Click
        'TextBox9.ReadOnly = False
    End Sub

    Private Sub TextBox10_Click(sender As Object, e As EventArgs) Handles TextBox10.Click
        'TextBox10.ReadOnly = False
    End Sub

    Private Sub TextBox6_Click(sender As Object, e As EventArgs) Handles TextBox6.Click
        'TextBox6.ReadOnly = False
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        'TextBox1.ReadOnly = False
    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        'TextBox2.ReadOnly = False
    End Sub

    Private Sub TextBox3_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        'TextBox3.ReadOnly = False
    End Sub

    Private Sub TextBox4_Click(sender As Object, e As EventArgs) Handles TextBox4.Click
        'TextBox4.ReadOnly = False
    End Sub

    Private Sub TextBox5_Click(sender As Object, e As EventArgs) Handles TextBox5.Click
        'TextBox5.ReadOnly = False
    End Sub

    Private Sub TextBox7_Click(sender As Object, e As EventArgs) Handles TextBox7.Click
        'TextBox7.ReadOnly = False
    End Sub
#End Region 'VERIFICAR

    'RemoveHandler Button4.Click, AddressOf Button4_Click
    'AddHandler Button4.Click, AddressOf AlterarDadosGrupo

    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click

    End Sub

    Private Sub Label31_Click(sender As Object, e As EventArgs) Handles Label31.Click

    End Sub
End Class