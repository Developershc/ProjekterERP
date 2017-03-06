Public Class frmMessageBox

#Region "Excluir"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Text = "Produção" Then

            ' ------------------- FAZER

        ElseIf Me.Text = "Produto" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("DELETE FROM TBPRODUTOS WHERE CODIGOPROD = '" & frmMP.TextBox1.Text & "' AND REVISAO = '" & frmMP.TextBox11.Text & "' AND TITULOPROD = '" & frmMP.TextBox2.Text & "' AND VOLUME = '" & frmMP.TextBox3.Text & "' AND MASSA = '" & frmMP.TextBox4.Text & "' AND TRATACAB = '" & frmMP.ComboBox1.Text & "' AND COR = '" & frmMP.TextBox7.Text & "' AND CUSTOMAT = '" & frmMP.TextBox8.Text & "' AND GARANTIA = '" & frmMP.TextBox6.Text & "' AND CUSTOTRATACAB = '" & frmMP.TextBox9.Text & "' AND CUSTOMO = '" & frmMP.TextBox10.Text & "' AND DESENHO = '" & frmMP.TextBox13.Text & "' AND 3D = '" & frmMP.TextBox14.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Excluído com sucesso!", "Produto", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Montagem" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("DELETE FROM TBMONTAGENS WHERE CODIGOM = '" & frmMP.TextBox1.Text & "' AND REVISAO = '" & frmMP.TextBox11.Text & "' AND TITULOM = '" & frmMP.TextBox2.Text & "' AND VOLUME = '" & frmMP.TextBox3.Text & "' AND MASSA = '" & frmMP.TextBox4.Text & "' AND TRATACAB = '" & frmMP.ComboBox1.Text & "' AND COR = '" & frmMP.TextBox7.Text & "' AND CUSTOMAT = '" & frmMP.TextBox8.Text & "' AND GARANTIA = '" & frmMP.TextBox6.Text & "' AND CUSTOTRATACAB = '" & frmMP.TextBox9.Text & "' AND CUSTOMO = '" & frmMP.TextBox10.Text & "' AND DESENHO = '" & frmMP.TextBox13.Text & "' AND 3D = '" & frmMP.TextBox14.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Excluído com sucesso!", "Montagem", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Maquinário" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("DELETE FROM TBMAQUINARIOS WHERE NOME = '" & frmSM.TextBox8.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Excluído com sucesso!", "Maquinário", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Seção" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("DELETE FROM TBSECOES WHERE NOMESECAO = '" & frmSM.TextBox8.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Excluído com sucesso!", "Seção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Serviços" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("DELETE FROM TBSERVICOS WHERE NOMES = '" & frmServico.TextBox1.Text & "', BASECALCULO = '" & frmServico.ComboBox1.Text & "', UNIDADE = '" & frmServico.TextBox2.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Excluído com sucesso!", "Serviços", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Orçamento" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("DELETE FROM TBORCAMENTOS WHERE CODIGO = '" & frmOrcamento.TextBox1.Text & "', CLIENTE = '" & frmOrcamento.TextBox2.Text & "', REVISAO = '" & frmOrcamento.TextBox3.Text & "', CUSTOINTERNO = '" & frmOrcamento.TextBox4.Text & "', IMPOSTO = '" & frmOrcamento.TextBox5.Text & "', LUCRO = '" & frmOrcamento.TextBox6.Text & "', DESCONTO = '" & frmOrcamento.TextBox11.Text & "', DATAINCLUSAO = '" & frmOrcamento.TextBox8.Text & "', DATAALTERACAO = '" & frmOrcamento.TextBox9.Text & "', DATAAPROVACAO = '" & frmOrcamento.TextBox7.Text & "', DESENHO = '" & frmOrcamento.TextBox10.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Excluído com sucesso!", "Orçamento", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.message)
            'End Try
        ElseIf Me.Text = "Usuários" Then
            'Try
            'Conectar()
            'Iniciar()
            'Comandar("DELETE FROM TBUSUARIOS WHERE NOMEGRUPO = '" & frmManutencao.ComboBox1.Text & "', USUARIO = '" & frmManutencao.TextBox5.Text & "', SENHA = '" & frmManutencao.TextBox6.Text & "'")
            'Executar()
            'Fechar()
            Me.Close()
            MessageBox.Show("Excluído com sucesso!", "Usuários", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Catch ex As Exception
            'MessageBox.Show(ex.Message)
            'End Try
        ElseIf Me.Text = "Materiais - Grupos" Then
            Try
                Conectar()
                Iniciar()
                Comandar("DELETE FROM TB_GRUPOSMAT WHERE NOMEGRUPOMAT = '" & frmMaterial.TextBox8.Text & "'")
                Executar()
                Fechar()
                Me.Close()
                'MessageBox.Show("Excluído com sucesso!", "Materiais - Grupos", MessageBoxButtons.OK, MessageBoxIcon.Information)
                frmMaterial.ListBox2.Items.Clear()
                'Try
                '    Conectar()
                '    Comandar("SELECT NOMEGRUPOMAT FROM TB_GRUPOSMAT")
                '    Ler()
                '    While leitura.Read
                '        frmMaterial.ListBox2.Items.Add(leitura("NOMEGRUPOMAT"))
                '    End While
                '    Fechar()
                'Catch ex As Exception
                '    MessageBox.Show(ex.Message)
                'End Try
                frmMaterial.TextBox8.Clear()
                frmMaterial.TextBox8.Focus()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf Me.Text = "Materiais - Qualidade" Then
            Try
                Conectar()
                Iniciar()
                Comandar("DELETE FROM TB_QUALIDADES WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox4.Text & "', NOMEQUALIDADE = '" & frmMaterial.TextBox9.Text & "'")
                Executar()
                Fechar()
                Me.Close()
                MessageBox.Show("Excluído com sucesso!", "Materiais - Qualidade", MessageBoxButtons.OK, MessageBoxIcon.Information)
                frmMaterial.ListBox3.Items.Clear()
                frmMaterial.ComboBox4.Text = ""
                frmMaterial.TextBox9.Clear()
                frmMaterial.ComboBox4.Focus()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMEQUALIDADE FROM TB_QUALIDADES")
                Ler()
                While leitura.Read
                    frmMaterial.ListBox3.Items.Clear()
                    frmMaterial.ListBox3.Items.Add(leitura("NOMEQUALIDADE"))
                End While
                Fechar()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf Me.Text = "Materiais - Tipos" Then
            Try
                Conectar()
                Iniciar()
                Comandar("DELETE FROM TB_TIPOS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox6.Text & "', NOMEQUALIDADE = '" & frmMaterial.ComboBox5.Text & "', NOMETIPO = '" & frmMaterial.TextBox10.Text & "'")
                Executar()
                Fechar()
                Me.Close()
                MessageBox.Show("Excluído com sucesso!", "Materiais - Tipos", MessageBoxButtons.OK, MessageBoxIcon.Information)
                frmMaterial.ListBox4.Items.Clear()
                frmMaterial.ComboBox6.Text = ""
                frmMaterial.ComboBox5.Text = ""
                frmMaterial.TextBox10.Clear()
                frmMaterial.ComboBox6.Focus()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            Try
                Conectar()
                Iniciar()
                Comandar("SELECT NOMETIPO FROM TB_TIPOS")
                Ler()
                While leitura.Read
                    frmMaterial.ListBox4.Items.Clear()
                    frmMaterial.ListBox4.Items.Add(leitura("NOMETIPO"))
                End While
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf Me.Text = "Materiais" Then
                If frmMaterial.CheckBox1.Checked = True Then
                    'Try
                    'Conectar()
                    'Iniciar()
                    'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox1.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDAPRINC = 1, KG = '" & frmMaterial.TextBox7.Text & "'")
                    'Executar()
                    'Fechar()
                    Me.Close()
                    MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    'Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                    'End Try
                ElseIf frmMaterial.CheckBox2.Checked = True Then
                    'Try
                    'Conectar()
                    'Iniciar()
                    'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox1.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDAPRINC = 2, KG = '" & frmMaterial.TextBox7.Text & "'")
                    'Executar()
                    'Fechar()
                    Me.Close()
                    MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    'Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                    'End Try
                ElseIf frmMaterial.CheckBox3.Checked = True Then
                    If frmMaterial.RadioButton6.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDAPRINC = 2, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton7.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDAPRINC = 3, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton8.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDAPRINC = 4, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton9.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDAPRINC = 5, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    End If
                ElseIf frmMaterial.CheckBox4.Checked = True Then
                    If frmMaterial.RadioButton6.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDA4 = '" & frmMaterial.TextBox4.Text & "', MEDIDAPRINC = 2, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton7.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDA4 = '" & frmMaterial.TextBox4.Text & "', MEDIDAPRINC = 3, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton8.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDA4 = '" & frmMaterial.TextBox4.Text & "', MEDIDAPRINC = 4, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton9.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDA4 = '" & frmMaterial.TextBox4.Text & "', MEDIDAPRINC = 5, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    End If
                ElseIf frmMaterial.CheckBox5.Checked = True Then
                    If frmMaterial.RadioButton6.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDA4 = '" & frmMaterial.TextBox4.Text & "', MEDIDA5 = '" & frmMaterial.TextBox5.Text & "', MEDIDAPRINC = 2, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton7.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDA4 = '" & frmMaterial.TextBox4.Text & "', MEDIDA5 = '" & frmMaterial.TextBox5.Text & "', MEDIDAPRINC = 3, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton8.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        '("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDA4 = '" & frmMaterial.TextBox4.Text & "', MEDIDA5 = '" & frmMaterial.TextBox5.Text & "', MEDIDAPRINC = 4, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    ElseIf frmMaterial.RadioButton9.Checked = True Then
                        'Try
                        'Conectar()
                        'Iniciar()
                        'Comandar("DELETE FROM TBMATERIAIS WHERE NOMEGRUPOMAT = '" & frmMaterial.ComboBox1.Text & "', NOMEQ = '" & frmMaterial.ComboBox2.Text & "', NOMET = '" & frmMaterial.ComboBox3.Text & "', NOMEMAT = '" & frmMaterial.TextBox6.Text & "', QUANTMEDIDAS = '" & frmMaterial.CheckBox3.Text & "', MEDIDA1 = '" & frmMaterial.TextBox1.Text & "', MEDIDA2 = '" & frmMaterial.TextBox2.Text & "', MEDIDA3 = '" & frmMaterial.TextBox3.Text & "', MEDIDA4 = '" & frmMaterial.TextBox4.Text & "', MEDIDA5 = '" & frmMaterial.TextBox5.Text & "', MEDIDAPRINC = 5, KG = '" & frmMaterial.TextBox7.Text & "'")
                        'Executar()
                        'Fechar()
                        Me.Close()
                        MessageBox.Show("Excluído com sucesso!", "Materiais", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                        'End Try
                    End If
                End If
        ElseIf Me.Text = "Funcionário" Then
                'Try
                ' Conectar()
                'Iniciar()
                'Comandar("DELETE FROM TBFUNCIONARIOS WHERE NOCRACHA = '" & frmFuncionario.TextBox26.Text & "', NOME = '" & frmFuncionario.TextBox25.Text & "', CPF = '" & frmFuncionario.TextBox24.Text & "', ENDERECO = '" & frmFuncionario.TextBox23.Text & "', BAIRRO = '" & frmFuncionario.TextBox22.Text & "', CIDADE = '" & frmFuncionario.TextBox21.Text & "', ESTADO = '" & frmFuncionario.TextBox20.Text & "', CEP = '" & frmFuncionario.TextBox19.Text & "', DDD = '" & frmFuncionario.TextBox18.Text & "', TELEFONE = '" & frmFuncionario.TextBox17.Text & "', BANCO = '" & frmFuncionario.TextBox16.Text & "', AGENCIA = '" & frmFuncionario.TextBox15.Text & "', CONTA = '" & frmFuncionario.TextBox14.Text & "', FUNCAO = '" & frmFuncionario.TextBox1.Text & "', SALARIO = '" & frmFuncionario.TextBox2.Text & "', HORASSEMANA = '" & frmFuncionario.TextBox3.Text & "', PRECOHORA = '" & frmFuncionario.TextBox4.Text & "', DATAADMISSAO = '" & frmFuncionario.TextBox5.Text & "', CTPS = '" & frmFuncionario.TextBox6.Text & "', PISPASEP = '" & frmFuncionario.TextBox7.Text & "', DATANASC = '" & frmFuncionario.TextBox8.Text & "'")
                'Executar()
                'Fechar()
                Me.Close()
                MessageBox.Show("Excluído com sucesso!", "Funcionário", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
        ElseIf Me.Text = "Peça" Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("DELETE FROM TBPECAS WHERE CODIGOP = '" & frmPeca.TextBox1.Text & "', REVISAO = '" & frmPeca.TextBox11.Text & "', TITULOP = '" & frmPeca.TextBox2.Text & "', AREA = '" & frmPeca.TextBox3.Text & "', MASSA = '" & frmPeca.TextBox4.Text & "', RENDIMENTO = '" & frmPeca.TextBox5.Text & "', TRATACAB = '" & frmPeca.ComboBox4.Text & "', COR = '" & frmPeca.TextBox7.Text & "', NOMEGRUPOMAT = '" & frmPeca.ComboBox1.Text & "', NOMEQ = '" & frmPeca.ComboBox2.Text & "', NOMET = '" & frmPeca.ComboBox3.Text & "', NOMEMAT = '" & frmPeca.ComboBox5.Text & "', CUSTOMAT = '" & frmPeca.TextBox8.Text & "', GARANTIA = '" & frmPeca.TextBox6.Text & "', CUSTOTRATACAB = '" & frmPeca.TextBox9.Text & "', CUSTOMO = '" & frmPeca.TextBox10.Text & "', CAD = '" & frmPeca.TextBox12.Text & "', DESENHO = '" & frmPeca.TextBox13.Text & "', 3D = '" & frmPeca.TextBox14.Text & "'")
                'Executar()
                'Fechar()
                Me.Close()
                MessageBox.Show("Excluído com sucesso!", "Peça", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
        ElseIf Me.Text = "Cadastro da empresa" Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("DELETE FROM TBEMPRESAS WHERE RAZAOSOCIAL = '" & frmCadastroEF.TextBox1.Text & "', ENDERECO = '" & frmCadastroEF.TextBox2.Text & "', CIDADE = '" & frmCadastroEF.TextBox3.Text & "', ESTADO = '" & frmCadastroEF.TextBox4.Text & "', TELEFONE = '" & frmCadastroEF.TextBox5.Text & "', CEP = '" & frmCadastroEF.TextBox6.Text & "', CNPJ = '" & frmCadastroEF.TextBox7.Text & "'")
                'Executar()
                'Fechar()
                Me.Close()
                MessageBox.Show("Excluído com sucesso!", "Cadastro da empresa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
        ElseIf Me.Text = "Fornecedor" Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("DELETE FROM TBFORNECEDORES WHERE RAZAOSOCIAL = '" & frmCadastroEF.TextBox1.Text & "', ENDERECO = '" & frmCadastroEF.TextBox2.Text & "', CIDADE = '" & frmCadastroEF.TextBox3.Text & "', ESTADO = '" & frmCadastroEF.TextBox4.Text & "', TELEFONE = '" & frmCadastroEF.TextBox5.Text & "', CEP = '" & frmCadastroEF.TextBox6.Text & "', CNPJ = '" & frmCadastroEF.TextBox7.Text & "'")
                'Executar()
                'Fechar()
                Me.Close()
                MessageBox.Show("Excluído com sucesso!", "Fornecedor", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
        ElseIf Me.Text = "Consumível" Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("DELETE FROM TBCONSUMIVEIS WHERE CODINTERNO = '" & frmConsumivel.TextBox1.Text & "', PRODUTO = '" & frmConsumivel.TextBox2.Text & "', SECAO = '" & frmConsumivel.ComboBox1.Text & "', PRAZOVALIDADE = '" & frmConsumivel.TextBox4.Text & "', GARANTIA = '" & frmConsumivel.TextBox6.Text & "'")
                'Executar()
                'Fechar()
                Me.Close()
                MessageBox.Show("Excluído com sucesso!", "Consumível", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
        ElseIf Me.Text = "CNC" Then
                'Try
                'Conectar()
                'Iniciar()
                'Comandar("DELETE FROM TBCNCS WHERE NOME = '" & frmCNC.ComboBox1.Text & "', PECA = '" & frmCNC.ComboBox2.Text & "', RELATORIO = '" & frmCNC.TextBox1.Text & "'")
                'Executar()
                'Fechar()
                Me.Close()
                MessageBox.Show("Excluído com sucesso!", "CNC", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Catch ex As Exception
                'MessageBox.Show(ex.Message)
                'End Try
        End If
    End Sub
#End Region

#Region "Cancelar"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()

        If Me.Text = "Produto" Then
            frmMP.Close()
        ElseIf Me.Text = "Montagem" Then
            frmMP.Close()
        ElseIf Me.Text = "Maquinário" Then
            frmSM.Close()
        ElseIf Me.Text = "Seção" Then
            frmSM.Close()
        ElseIf Me.Text = "Orçamento" Then
            frmOrcamento.Close()
        ElseIf Me.Text = "Materiais - Grupos" Then
            frmMaterial.Close()
        ElseIf Me.Text = "Materiais - Qualidades" Then
            frmMaterial.Close()
        ElseIf Me.Text = "Materiais - Tipos" Then
            frmMaterial.Close()
        ElseIf Me.Text = "Materiais" Then
            frmMaterial.Close()
        ElseIf Me.Text = "Funcionário" Then
            frmFuncionario.Close()
        ElseIf Me.Text = "Peça" Then
            frmPeca.Close()
        ElseIf Me.Text = "Cadastro da empresa" Then
            frmCadastroEF.Close()
        ElseIf Me.Text = "Fornecedor" Then
            frmCadastroEF.Close()
        ElseIf Me.Text = "Entrada" Then
            frmES.Close()
        ElseIf Me.Text = "Saída" Then
            frmES.Close()
        ElseIf Me.Text = "Impostos" Then
            frmImpostos.Close()
        ElseIf Me.Text = "Consumível" Then
            frmConsumivel.Close()
        ElseIf Me.Text = "CNC" Then
            frmCNC.Close()
        ElseIf Me.Text = "Consumos" Then
            frmConsumos.Close()
        End If
    End Sub
#End Region

End Class