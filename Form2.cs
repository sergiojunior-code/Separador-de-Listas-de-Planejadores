using DocumentFormat.OpenXml.Drawing.Charts;
using Separador_de_Listas_de_Planejadores;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Separador_de_Listas_de_Planejadores
{
    public partial class FormAtualizacoes : Form
    {
        // RichTextBox já definido no designer
        private RichTextBox rtbAtualizacoes;
        private Button btnFechar;

        public FormAtualizacoes()
        {
            InitializeComponent();
            InicializarComponentes();
            CarregarAtualizacoes();
        }

        private void InicializarComponentes()
        {
            var form1 = new Form1();
            string versao = form1.versao;
            this.Text = $"{versao} Notas de Atualização";
            this.Size = new System.Drawing.Size(500, 400);
            this.StartPosition = FormStartPosition.CenterParent;

            var alt = this.ClientSize.Height;
            var larg = this.ClientSize.Width;

            // RichTextBox
            rtbAtualizacoes = new RichTextBox();
            rtbAtualizacoes.Dock = DockStyle.Top;
            rtbAtualizacoes.Height = 370;
            rtbAtualizacoes.ReadOnly = true;
            rtbAtualizacoes.ScrollBars = RichTextBoxScrollBars.Vertical;
            rtbAtualizacoes.BackColor = Color.White;
            rtbAtualizacoes.Font = new Font("Consolas", 10);

            // Botão Fechar
            btnFechar = new Button();
            btnFechar.Text = "Fechar";
            btnFechar.Width = 100;
            btnFechar.Height = 30;
            btnFechar.Top = rtbAtualizacoes.Bottom + 5;
            btnFechar.Left = (this.ClientSize.Width - btnFechar.Width) / 2;
            btnFechar.Anchor = AnchorStyles.Bottom;
            btnFechar.Click += (s, e) => this.Close();

            this.Controls.Add(rtbAtualizacoes);
            this.Controls.Add(btnFechar);
        }

        private void CarregarAtualizacoes()
        {
            // Lista de atualizações
            var lista = new List<Atualizacao>
            {
                new Atualizacao // Versão 1.2
                {
                    Versao = "1.2",
                    Data = DateTime.Parse("2026-01-01"),
                    Mudancas = new List<string>
                    {
                        " * Melhoria: Gera lista de planejador dos pedidos que contem ancoragem em um unico PDF, eliminando a analise do usuario em todos os pedidos, focando somente nos que contem.\n" +
                        " * Melhoria: Gera um Log no final do processo com data e horário, mesmo quando contem erro esse log será gerado. Esse log nunca será sobrescrito, sempre acrescentado abaixo toda vez que roda o programa.\n" +
                        " * Melhoria: No painel, melhor a visualização do que está sendo feito e corrigido alguns detalhes de exibição.\n" +
                        " * Melhoria: Exceções inseridas em alguns metodos para tratamento de erros, não deixando o programa quebrar.\n",
                    }
                }

            };

            PreencherRichTextBox(lista);
        }
        private void rtbAtualizacoes_TextChanged(object sender, EventArgs e)
        {

        }

        private void PreencherRichTextBox(List<Atualizacao> lista)
        {
            rtbAtualizacoes.Clear();

            var ordenada = lista.OrderBy(a => a.Data).ToList();


            foreach (var at in ordenada)
            {
                // Título: Versão + Data
                rtbAtualizacoes.SelectionFont = new Font("Consolas", 10, FontStyle.Bold);
                rtbAtualizacoes.SelectionColor = Color.DarkBlue;
                rtbAtualizacoes.AppendText($"Versão {at.Versao} - {at.Data:dd/MM/yyyy}\n");

                // Separador
                rtbAtualizacoes.SelectionFont = new Font("Consolas", 9, FontStyle.Italic);
                rtbAtualizacoes.SelectionColor = Color.Gray;
                rtbAtualizacoes.AppendText("------------------------\n");

                // Mudanças
                rtbAtualizacoes.SelectionFont = new Font("Consolas", 9, FontStyle.Regular);
                rtbAtualizacoes.SelectionColor = Color.Black;

                foreach (var item in at.Mudancas)
                {
                    rtbAtualizacoes.AppendText($"{item}\n");
                }

                rtbAtualizacoes.AppendText("\n");
            }
        }

        private void FormAtualizacoes_Load(object sender, EventArgs e)
        {

        }
    }

    // Classe para manter cada atualização organizada
    public class Atualizacao
    {
        public string Versao { get; set; }
        public DateTime Data { get; set; }
        public List<string> Mudancas { get; set; } = new List<string>();
    }
}