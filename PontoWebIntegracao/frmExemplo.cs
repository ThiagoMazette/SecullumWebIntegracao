using PontoWebIntegracaoExterna.Filtros;
using PontoWebIntegracaoExterna.Modelos;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PontoWebIntegracaoExterna
{
    public partial class frmExemplo : Form
    {
        IntegracaoPontoWeb integracao = new IntegracaoPontoWeb();

        public frmExemplo()
        {
            InitializeComponent();
        }

        private bool ConsistirDados()
        {
            if (string.IsNullOrEmpty(txtCS_Access_Token.Text))
            {
                MessageBox.Show("É necessário fazer login na Conta Secullum");

                tabControl1.SelectedIndex = 0;

                return false;
            }

            return true;
        }



        private void btnCS_Autenticar_Click(object sender, EventArgs e)
        {
            var resp = integracao.AutenticarContaSecullum(txtCS_Usuario.Text, txtCS_Senha.Text);

            if (resp.erro)
            {
                txtCS_Access_Token.Text = "";
                txtCS_Refresh_Token.Text = "";
                txtCS_NomeUsuario.Text = "";
                cboCS_Bancos.DataSource = null;
                MessageBox.Show(resp.mensagem);
            }
            else
            {
                txtCS_Access_Token.Text = resp.access_token;
                txtCS_Refresh_Token.Text = resp.refresh_token;
            }
        }

        private void btnCS_ListarDados_Click(object sender, EventArgs e)
        {
            var resp = integracao.BuscaDadosContaSecullum(txtCS_Access_Token.Text);

            if (resp.erro)
            {
                txtCS_NomeUsuario.Text = "";
                cboCS_Bancos.Items.Clear();
                MessageBox.Show(resp.mensagem);
            }
            else
            {
                txtCS_NomeUsuario.Text = resp.nome;

                cboCS_Bancos.DisplayMember = "nome";
                cboCS_Bancos.ValueMember = "identificador";
                cboCS_Bancos.DataSource = new BindingSource(resp.listaBancos, null);
            }
        }

        private void btnEmpresas_Listar_Click(object sender, EventArgs e)
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvEmpresas.DataSource = integracao.ListarEmpresas(string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ConsistirDados())
            {
                MessageBox.Show(integracao.ExcluirEmpresa(txtCnpjEmpresa.Text));
            }
        }

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvHorarios.DataSource = integracao.ListarHorarios(string.Empty);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    if (ConsistirDados())
        //    {
        //        MessageBox.Show(integracao.ExcluirHorario(txthorarioNumero.Text));
        //    }
        //}

        //private void button5_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvFuncoes.DataSource = integracao.ListarFuncoes(string.Empty);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button4_Click(object sender, EventArgs e)
        //{
        //    if (ConsistirDados())
        //    {
        //        MessageBox.Show(integracao.ExcluirFuncao(txtFuncaoDescricao.Text));
        //    }
        //}

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvDepartamentos.DataSource = integracao.ListarDepartamentos(string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (ConsistirDados())
            {
                MessageBox.Show(integracao.ExcluirDepartamento(txtDepartamentoDescricao.Text));
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvFuncionarios.DataSource = integracao.ListarFuncionarios(string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (ConsistirDados())
            {
                MessageBox.Show(integracao.ExcluirFuncionario(txtPis.Text));
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvEmpresas.DataSource = integracao.ListarEmpresas(txtCnpjEmpresa.Text);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //private void button13_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvHorarios.DataSource = integracao.ListarHorarios(txthorarioNumero.Text);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button12_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvFuncoes.DataSource = integracao.ListarFuncoes(txtFuncaoDescricao.Text);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvDepartamentos.DataSource = integracao.ListarDepartamentos(txtDepartamentoDescricao.Text);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvFuncionarios.DataSource = integracao.ListarFuncionarios(txtPis.Text);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //private void button17_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvMotivosDemissao.DataSource = integracao.ListarMotivosDemissao(string.Empty);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button15_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvMotivosDemissao.DataSource = integracao.ListarMotivosDemissao(txtMotivoDemissaoDescricao.Text);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button20_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvAfastamentos.DataSource = integracao.ListarAfastamentos();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button18_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvAfastamentos.DataSource = integracao.ListarAfastamentos(txtAfastamentoDataInicio.Text, txtAfastamentoDataFim.Text, txtAfastamentoFuncionarioPis.Text);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button23_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvJustificativas.DataSource = integracao.ListarJustificativas(string.Empty);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button21_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvJustificativas.DataSource = integracao.ListarJustificativas(txtJustificativaDescricao.Text);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button26_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvPerguntasAdicionais.DataSource = integracao.ListarPerguntasAdicionais();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void button24_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvPerguntasAdicionais.DataSource = integracao.ListarPerguntasAdicionais(txtPerguntaDescricao.Text, txtPerguntaGrupo.Text);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        private void button29_Click(object sender, EventArgs e)
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvBatidas.DataSource = integracao.ListarBatidas(new BatidaFiltro()
                    {
                        DataInicio = txtBatidasDataInicio.Text,
                        DataFim = txtBatidasDataFim.Text,
                        HoraInicio = txtBatidasHoraInicio.Text,
                        HoraFim = txtBatidasHoraFim.Text,
                        FuncionarioPis = txtBatidasFuncionarioPis.Text,
                        EmpresaDocumento = txtBatidasEmpresaDocumento.Text
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //private void button27_Click(object sender, EventArgs e)
        //{
        //    if (!ConsistirDados())
        //    {
        //        return;
        //    }

        //    var dados = new Afastamento()
        //    {
        //        NumeroPis = txtAfastamentoFuncionarioPis.Text,
        //        Inicio = Convert.ToDateTime(txtAfastamentoDataInicio.Text),
        //        Fim = Convert.ToDateTime(txtAfastamentoDataFim.Text),
        //        JustificativaNome = txtAfastamentoJustificativaNome.Text,
        //        Motivo = txtAfastamentoMotivo.Text
        //    };

        //    MessageBox.Show(integracao.SalvarAfastamento(dados));
        //}

        //private void button16_Click(object sender, EventArgs e)
        //{
        //    if (ConsistirDados())
        //    {
        //        MessageBox.Show(integracao.ExcluirMotivoDemissao(txtMotivoDemissaoDescricao.Text));
        //    }
        //}

        //private void button19_Click(object sender, EventArgs e)
        //{
        //    if (ConsistirDados())
        //    {
        //        MessageBox.Show(integracao.ExcluirAfastamento(txtAfastamentoDataInicio.Text, txtAfastamentoDataFim.Text, txtAfastamentoFuncionarioPis.Text));
        //    }
        //}

        //private void button22_Click(object sender, EventArgs e)
        //{
        //    if (ConsistirDados())
        //    {
        //        MessageBox.Show(integracao.ExcluirJustificativa(txtJustificativaDescricao.Text));
        //    }
        //}

        //private void button25_Click(object sender, EventArgs e)
        //{
        //    if (ConsistirDados())
        //    {
        //        MessageBox.Show(integracao.ExcluirPerguntaAdicional(txtPerguntaDescricao.Text, txtPerguntaGrupo.Text));
        //    }
        //}

        //private void btnLISTAR_PENDENCIAS_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (ConsistirDados())
        //        {
        //            dgvPENDENCIAS.DataSource = integracao.ListarPendencias();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        //private void btnSALVAR_PENDENCIA_Click(object sender, EventArgs e)
        //{
        //    if (!ConsistirDados())
        //    {
        //        return;
        //    }

        //    var dados = new PendenciaProcessada()
        //    {
        //        NomeComputador = txtNOME_COMPUTADOR.Text,
        //        Erro = txtERRO.Text
        //    };

        //    int.TryParse(txtPENDENCIA_ID.Text, out int pendenciaId);

        //    dados.Id = pendenciaId;

        //    MessageBox.Show(integracao.SalvarPendenciaProcessada(dados));
        //}

        private void cboCS_Bancos_SelectedIndexChanged(object sender, EventArgs e)
        {
            integracao.BancoPontoWebSelecionado = cboCS_Bancos.SelectedValue.ToString();
        }


        public List<string> ListaIDFuncionarios = new List<string>();

        public class ClasseListaNomeFuncionario
        {
            public string ID { get; set; }
            public string Nome { get; set; }
            public string Dpto { get; set; }

        }
        public List<ClasseListaNomeFuncionario> ListaNomeFuncionario = new List<ClasseListaNomeFuncionario>();

        public class ClasseListaNomeDepartamento
        {
            public string ID { get; set; }
            public string Nome { get; set; }
        }
        public List<ClasseListaNomeDepartamento> ListaNomeDepartamento = new List<ClasseListaNomeDepartamento>();



        // private void btnGerarXLSListar_Click(object sender, EventArgs e) //esse mostra somente quem veio
        // {
        //     //dgvEmpresas.AutoGenerateColumns=false;
        //     voidbtnAutenticar1();
        //     voidbtnListarEmpresas();
        //     txtBatidasDataInicio.Text = dateTimePicker1.Value.ToString("yyyy-MM-dd");
        //     txtBatidasDataFim.Text = dateTimePicker2.Value.ToString("yyyy-MM-dd");

        //     voidListarDeto();
        //     voidbtnListarFunc();
        //     voidListarBatidas();


        //     foreach (DataGridViewRow Dep in dgvDepartamentos.Rows)
        //     {

        //         string IdDep = dgvDepartamentos.Rows[Dep.Index].Cells["Id"].Value.ToString();
        //         string NomeDep = dgvDepartamentos.Rows[Dep.Index].Cells["Descricao"].Value.ToString();

        //         ListaNomeDepartamento.Add(new ClasseListaNomeDepartamento()
        //         {
        //             ID = IdDep,
        //             Nome = NomeDep

        //         });

        //     }

        //     foreach (DataGridViewRow func in dgvFuncionarios.Rows)
        //     {

        //         string IdFunc = dgvFuncionarios.Rows[func.Index].Cells["Id"].Value.ToString();
        //         string NomeFunc = dgvFuncionarios.Rows[func.Index].Cells["Nome"].Value.ToString();
        //         string DptoFunc = dgvFuncionarios.Rows[func.Index].Cells["DepartamentoId"].Value.ToString();

        //         ListaNomeFuncionario.Add(new ClasseListaNomeFuncionario()
        //              {
        //                  ID = IdFunc,
        //                  Nome = NomeFunc,
        //                  Dpto = DptoFunc

        //         });

        //     }


        //     //aparecer bonito
        //     //percorer dgvbatidas, verificar o dia e se tem pelo menos uma batida


        //     foreach (DataGridViewRow bat in dgvBatidas.Rows)
        //     {
        //         string Data = dgvBatidas.Rows[bat.Index].Cells["Data"].Value.ToString().Substring(0,10);
        //         string idFunc = dgvBatidas.Rows[bat.Index].Cells["FuncionarioId"].Value.ToString();
        //         string NomeFunc = "";
        //         string idDpto = "";
        //         string NomeDpto = "";
        //         string Falta = "1";
        //    //     string DataDentro = "";
        //    //     string idFuncDentro = "";
        //         string Batida1 = dgvBatidas.Rows[bat.Index].Cells["Entrada1"].Value.ToString();
        //         string Batida2 = dgvBatidas.Rows[bat.Index].Cells["Saida1"].Value.ToString();
        //         string Batida3 = dgvBatidas.Rows[bat.Index].Cells["Entrada2"].Value.ToString();
        //         string Batida4 = dgvBatidas.Rows[bat.Index].Cells["Saida2"].Value.ToString();
        //         string Batida5 = dgvBatidas.Rows[bat.Index].Cells["Entrada3"].Value.ToString();
        //         string Batida6 = dgvBatidas.Rows[bat.Index].Cells["Saida3"].Value.ToString();
        //         string Batida7 = dgvBatidas.Rows[bat.Index].Cells["Entrada4"].Value.ToString();
        //         string Batida8 = dgvBatidas.Rows[bat.Index].Cells["Saida4"].Value.ToString();
        //         string Batida9 = dgvBatidas.Rows[bat.Index].Cells["Entrada5"].Value.ToString();
        //         string Batida10 = dgvBatidas.Rows[bat.Index].Cells["Saida5"].Value.ToString();
        ////         bool TemAlgumaBatida = false;

        //         if(Batida1 !="" || Batida2 !="" || Batida3 != "" || Batida3 != "" || Batida4 != "" || Batida5 != "" || Batida6 != "" || Batida7 != "" || Batida8 != "" || Batida9 != "" || Batida10 != "" )
        //         {
        //             Falta = "0";
        //         }

        //         foreach (var Func in ListaNomeFuncionario)
        //         {
        //             if (Func.ID == idFunc)
        //             {
        //                 NomeFunc = Func.Nome;
        //                 idDpto = Func.Dpto;
        //                 foreach (var dep in ListaNomeDepartamento)
        //                 {
        //                     if (dep.ID == idDpto)
        //                     {
        //                         NomeDpto = dep.Nome;
        //                     }

        //                 }
        //             }

        //         }

        //         //if (!ListaIDFuncionarios.Contains(idFunc))
        //         //{
        //         //    ListaIDFuncionarios.Add(idFunc);
        //         //    Falta = "0";

        //         //}


        //         dataGridView1.ColumnCount = 4; //carregar a qdade de colunas
        //         dataGridView1.Columns[0].Name = "Data";
        //         dataGridView1.Columns[1].Name = "Nome";
        //         dataGridView1.Columns[2].Name = "Dpto";
        //         dataGridView1.Columns[3].Name = "Falta";

        //         dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
        //         dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        //         dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
        //         dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;


        //         dataGridView1.Rows.Add(Data.ToString(), NomeFunc.ToString(), NomeDpto.ToString(), Falta.ToString());





        //     }



        // }


        private void btnGerarXLSListar_Click(object sender, EventArgs e)
        {
            ListaIDFuncionarios.Clear();
            ListaNomeDepartamento.Clear();
            ListaNomeFuncionario.Clear();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            //dgvEmpresas.AutoGenerateColumns=false;
            voidbtnAutenticar1();
            voidbtnListarEmpresas();
            txtBatidasDataInicio.Text = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            txtBatidasDataFim.Text = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            voidListarDeto();
            voidbtnListarFunc();
            voidListarBatidas();

            #region Carrega lista de dep
            foreach (DataGridViewRow Dep in dgvDepartamentos.Rows)
            {

                string IdDep = dgvDepartamentos.Rows[Dep.Index].Cells["Id"].Value.ToString();
                string NomeDep = dgvDepartamentos.Rows[Dep.Index].Cells["Descricao"].Value.ToString();

                ListaNomeDepartamento.Add(new ClasseListaNomeDepartamento()
                {
                    ID = IdDep,
                    Nome = NomeDep

                });

            }
            #endregion

            #region Carrega lista de Func com Dpto
            foreach (DataGridViewRow func in dgvFuncionarios.Rows)
            {

                string IdFunc = dgvFuncionarios.Rows[func.Index].Cells["Id"].Value.ToString();
                string NomeFunc = dgvFuncionarios.Rows[func.Index].Cells["Nome"].Value.ToString();
                string DptoFunc = dgvFuncionarios.Rows[func.Index].Cells["DepartamentoId"].Value.ToString();
                bool Invisivel = Convert.ToBoolean(dgvFuncionarios.Rows[func.Index].Cells["Invisivel"].Value.ToString());

                if (Invisivel == false)
                {
                    ListaNomeFuncionario.Add(new ClasseListaNomeFuncionario()
                    {
                        ID = IdFunc,
                        Nome = NomeFunc,
                        Dpto = DptoFunc

                    });
                }

            }
            #endregion


            //aparecer bonito
            //percorer dgvbatidas, verificar o dia e se tem pelo menos uma batida

            string dataInicial = dateTimePicker1.Value.ToString("dd/MM/yyyy");
            string dataFinal = dateTimePicker2.Value.ToString("dd/MM/yyyy");
            int totalDias = (DateTime.Parse(dataFinal).Subtract(DateTime.Parse(dataInicial))).Days;
            DateTime DataDGV = DateTime.Parse(dataInicial);


            for (int i = 0; i <= totalDias; i++)
            {


                foreach (var Func in ListaNomeFuncionario)
                {
                    string NomeFunc = "";
                    string idDpto = "";
                    string NomeDpto = "";
                    string Falta = "1";
                    string idFunc = "";
                    DateTime HorasTrabalhadas = new DateTime();

                    foreach (DataGridViewRow bat in dgvBatidas.Rows)
                    {
                        DateTime Data = Convert.ToDateTime(dgvBatidas.Rows[bat.Index].Cells["Data"]
                            .Value.ToString().Substring(0, 10));
                        idFunc = dgvBatidas.Rows[bat.Index].Cells["FuncionarioId"].Value.ToString();


                        if (DataDGV == Data && Func.ID == idFunc)
                        {
                            Falta = "0";
                            
                        //}
                        //if (Func.ID == idFunc)
                        //{
                            DateTime Batida1 = Convert.ToDateTime(dgvBatidas.Rows[bat.Index].Cells["Entrada1"].Value.ToString());
                            string Batida2 = dgvBatidas.Rows[bat.Index].Cells["Saida1"].Value.ToString();
                            string Batida3 = dgvBatidas.Rows[bat.Index].Cells["Entrada2"].Value.ToString();
                            string Batida4 = dgvBatidas.Rows[bat.Index].Cells["Saida2"].Value.ToString();
                            string Batida5 = dgvBatidas.Rows[bat.Index].Cells["Entrada3"].Value.ToString();
                            string Batida6 = dgvBatidas.Rows[bat.Index].Cells["Saida3"].Value.ToString();
                            string Batida7 = dgvBatidas.Rows[bat.Index].Cells["Entrada4"].Value.ToString();
                            string Batida8 = dgvBatidas.Rows[bat.Index].Cells["Saida4"].Value.ToString();
                            string Batida9 = dgvBatidas.Rows[bat.Index].Cells["Entrada5"].Value.ToString();
                            string Batida10 = dgvBatidas.Rows[bat.Index].Cells["Saida5"].Value.ToString();

                            break;
                        }

                        TimeSpan horaIni = new TimeSpan(2, 45, 0);
                        TimeSpan horaFim = new TimeSpan(3, 20, 0);
                        //Mostrar o resultado no prompt
                        Console.WriteLine(horaFim.Subtract(horaIni).ToString());






                    }

                    if (Falta == "1")
                    {
                        idFunc = Func.ID;

                    }


                    NomeFunc = Func.Nome;
                    idDpto = Func.Dpto;
                    foreach (var dep in ListaNomeDepartamento)
                    {
                        if (dep.ID == idDpto)
                        {
                            NomeDpto = dep.Nome;
                            break;
                        }

                    }





                    dataGridView1.ColumnCount = 5; //carregar a qdade de colunas
                    dataGridView1.Columns[0].Name = "Data";
                    dataGridView1.Columns[1].Name = "Nome";
                    dataGridView1.Columns[2].Name = "Dpto";
                    dataGridView1.Columns[3].Name = "Falta";
                    dataGridView1.Columns[4].Name = "HorasTrab";

                    dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
                    dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
                    dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                    dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;


                    dataGridView1.Rows.Add(DataDGV.ToString().Substring(0, 10), NomeFunc.ToString(),
                        NomeDpto.ToString(), Falta.ToString(), HorasTrabalhadas.ToShortDateString());




                }

                DataDGV = DataDGV.AddDays(1);

            }


        }






        void voidbtnAutenticar1()
        {
            var resp = integracao.AutenticarContaSecullum(txtCS_Usuario.Text, txtCS_Senha.Text);

            if (resp.erro)
            {
                txtCS_Access_Token.Text = "";
                txtCS_Refresh_Token.Text = "";
                txtCS_NomeUsuario.Text = "";
                cboCS_Bancos.DataSource = null;
                MessageBox.Show(resp.mensagem);
            }
            else
            {
                txtCS_Access_Token.Text = resp.access_token;
                txtCS_Refresh_Token.Text = resp.refresh_token;
            }
            voidbtnListarDados();
        }

        void voidbtnListarDados()
        {
            var resp = integracao.BuscaDadosContaSecullum(txtCS_Access_Token.Text);

            if (resp.erro)
            {
                txtCS_NomeUsuario.Text = "";
                cboCS_Bancos.Items.Clear();
                MessageBox.Show(resp.mensagem);
            }
            else
            {
                txtCS_NomeUsuario.Text = resp.nome;

                cboCS_Bancos.DisplayMember = "nome";
                cboCS_Bancos.ValueMember = "identificador";
                cboCS_Bancos.DataSource = new BindingSource(resp.listaBancos, null);
            }


            cboCS_Bancos.SelectedIndex = 0;

        }



        void voidbtnListarEmpresas()
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvEmpresas.DataSource = integracao.ListarEmpresas(string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        void voidListarDeto()
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvDepartamentos.DataSource = integracao.ListarDepartamentos(string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        void voidbtnListarFunc()
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvFuncionarios.DataSource = integracao.ListarFuncionarios(string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void voidListarBatidas()
        {
            try
            {
                if (ConsistirDados())
                {
                    dgvBatidas.DataSource = integracao.ListarBatidas(new BatidaFiltro()
                    {
                        DataInicio = txtBatidasDataInicio.Text,
                        DataFim = txtBatidasDataFim.Text,
                        HoraInicio = txtBatidasHoraInicio.Text,
                        HoraFim = txtBatidasHoraFim.Text,
                        FuncionarioPis = txtBatidasFuncionarioPis.Text,
                        EmpresaDocumento = txtBatidasEmpresaDocumento.Text
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                try
                {
                    XcelApp.Application.Workbooks.Add(Type.Missing);

                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {
                        XcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    }
                    //
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            XcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    //
                    XcelApp.Columns.AutoFit();
                    //
                    XcelApp.Visible = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro : " + ex.Message);
                    XcelApp.Quit();
                }
            }

        }

        private void dgvBatidas_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
