using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Runtime.CompilerServices;

namespace Separador_de_Listas_de_Planejadores
{
    public partial class Form1 : Form
    {
        public string pdfPath;
        public string excelPath;
        public bool excelCarregado = false;
        public bool pdfCarregado = false;
        public string resumoFinal = "";

        private List<Armazenar> listaArmazenar = new List<Armazenar>();
        List<string> pastasNaoEncontradas = new List<string>();
        public Form1()
        {
            InitializeComponent();
            planejadores.Add("Planejador: 24 - OBRAS");
            planejadores.Add("Planejador: 25 - FABRICA FERRAGENS");
            planejadores.Add("Planejador: 26 - VIDROS");
            planejadores.Add("Planejador: 27 - PERFIL");
            planejadores.Add("Planejador: 28 - TAPEÇARIA");
            planejadores.Add("Planejador: 29 - SERRALHERIA");
            planejadores.Add("Planejador: 30 - PINTURA ACESSORIOS");
            planejadores.Add("Planejador: 31 - ILUMINACAO/ELETRICA");
            planejadores.Add("Planejador: 32 - ACRILICO");
            planejadores.Add("Planejador: 33 - ESPECIAIS");
            


    }
    public class Armazenar
        {
            // Propriedades públicas
            public string OrdemCompra { get; set; }
            public string SubPlanejadores { get; set; }
            public string Planejadores { get; set; }

            
            // Construtor
            public Armazenar(string ordemcompra, string subplanejadores, string planejadores)
            {
                OrdemCompra = ordemcompra;
                SubPlanejadores = subplanejadores;
                Planejadores = planejadores;
                
            }
        }

        public List<string> lista = new List<string>();
        public List<string> planejadores = new List<string>();
        List<string> LerExcel(string caminho)
        {
            var lista = new List<string>(); // garante que a lista começa vazia
            try
            {
                using (var wb = new XLWorkbook(caminho))
                {
                    var ws = wb.Worksheet(1);
                    foreach (var row in ws.RowsUsed().Skip(1)) // pula cabeçalho
                    {
                        try
                        {
                            string valor = row.Cell(4).GetString().Trim();
                            if (!string.IsNullOrEmpty(valor))
                                lista.Add(valor);
                        }
                        catch (Exception exCell)
                        {
                            // Captura erro de célula individual, mas continua o loop
                            MessageBox.Show($"Erro ao ler célula na linha {row.RowNumber()}: {exCell.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao ler o arquivo Excel: {ex.Message}");
            }

            return lista;
        }
        private void btnCarregarXLS_Click(object sender, EventArgs e)
        {
            lblAgExcel.Text = "Aguardando arquivo excel...";
            excelPath = string.Empty;
            excelCarregado = false;

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx;*.xls";


            if (ofd.ShowDialog() == DialogResult.OK)
            {
                lista = LerExcel(ofd.FileName);
                foreach (var item in lista)
                {
                    if (!item.Contains(" - "))
                    {
                        MessageBox.Show($"O ambiente '{item}' não está no formato esperado. Por favor, verifique o arquivo Excel.");
                        excelCarregado = false;
                        lblAgExcel.Text = "Aguardando arquivo excel...";
                        excelPath = string.Empty;
                        lista.Clear();
                        return;
                    }
                }
                if (lista.Count == 0)
                {
                    MessageBox.Show("O arquivo Excel selecionado não contém dados na coluna D. Por favor, selecione um arquivo válido.");
                    excelCarregado = false;
                    lblAgExcel.Text = "Aguardando arquivo excel...";
                    excelPath = string.Empty;
                    lista.Clear();
                    return;
                }


                int qntdAmbientes = 0;
                richtxtPainel.AppendText("Arquivo:" + ofd.FileName + "\n\n");
                foreach (var item in lista)
                {
                    qntdAmbientes++;
                    richtxtPainel.AppendText(item + "\n");
                    richtxtPainel.ScrollToCaret();
                }

                richtxtPainel.AppendText($"Quantidade de ambientes: {qntdAmbientes} \n");
                richtxtPainel.ScrollToCaret();
                excelCarregado = true;
                lblAgExcel.Text = "Excel carregado com sucesso!";
                excelPath = ofd.FileName;     
            }
            if (excelCarregado == true && pdfCarregado == true)
            {
                lblStatus.Text = "Processando...";
                Application.DoEvents();
                lblStatus.Text = "Processamento concluído!";
                lblStatus.Text = "Pronto para começar!";
            }
        }


        private void richtxtPainel_TextChanged(object sender, EventArgs e)
        {

        }

        private void CriarPDF(string ordemcompra)
        {
            string origem = pdfPath;
            string pastaOrigem = Path.GetDirectoryName(origem);
            string nomeArquivo = Path.GetFileNameWithoutExtension(pdfPath);

            // Extrair números do pedido
            string fourDigits = ordemcompra.Substring(0, 4);
            string fiveDigits = ordemcompra.Substring(0, 5);
            int firstDigits = 5;

            //int tamanhoPrefixo = ordemcompra.Substring(0, ordemcompra.IndexOf(" - ")).Length;
            //if (tamanhoPrefixo == 7) firstDigits = 5;
            //else if (tamanhoPrefixo == 6) firstDigits = 4;

            richtxtPainel.AppendText($"Prefixo detectado: {firstDigits}\n");
            richtxtPainel.ScrollToCaret();

            // Extrair numAmbiente corretamente
            int indiceHifen = ordemcompra.IndexOf("- ");
            string numAmbiente = ordemcompra.Substring(indiceHifen + 2, 2);

            richtxtPainel.AppendText($"Processando ordem de compra: {ordemcompra} (Ambiente: {numAmbiente})\n");
            richtxtPainel.ScrollToCaret();

            // Buscar diretórios de contratos filtrados
            string prefixo = firstDigits == 5 ? fiveDigits : fourDigits;
            var diretorioContratos = Directory.GetDirectories(@"J:\Pedidos 2026", prefixo + "*");

            var numContrato = ordemcompra.Substring(0, ordemcompra.IndexOf(" - ")).Trim();
            var nomeAmbiente = ordemcompra.Substring(ordemcompra.IndexOf(" - ") + 3).Trim();


            using (var pdfOrigem = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(origem)))
            {
                for (int i = 1; i <= pdfOrigem.GetNumberOfPages(); i++)
                {
                    var strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy();
                    string texto = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(pdfOrigem.GetPage(i), strategy);
                    if (texto.Substring(texto.IndexOf("Data de emissão: ", StringComparison.OrdinalIgnoreCase) + 6, 4) == "2024")
                        diretorioContratos = Directory.GetDirectories(@"J:\Pedidos 2024", prefixo + "*");
                    if (texto.Substring(texto.IndexOf("Data de emissão: ", StringComparison.OrdinalIgnoreCase) + 6, 4) == "2025")
                        diretorioContratos = Directory.GetDirectories(@"J:\Pedidos 2025", prefixo + "*");
                    if (texto.Substring(texto.IndexOf("Data de emissão: ", StringComparison.OrdinalIgnoreCase) + 6, 4) == "2026")
                        diretorioContratos = Directory.GetDirectories(@"J:\Pedidos 2026", prefixo + "*");
                    if (texto.Substring(texto.IndexOf("Data de emissão: ", StringComparison.OrdinalIgnoreCase) + 6, 4) == "2027")
                        diretorioContratos = Directory.GetDirectories(@"J:\Pedidos 2027", prefixo + "*");
                    if (texto.Substring(texto.IndexOf("Data de emissão: ", StringComparison.OrdinalIgnoreCase) + 6, 4) == "2028")
                        diretorioContratos = Directory.GetDirectories(@"J:\Pedidos 2028", prefixo + "*");
                    if (texto.Substring(texto.IndexOf("Data de emissão: ", StringComparison.OrdinalIgnoreCase) + 6, 4) == "2029")
                        diretorioContratos = Directory.GetDirectories(@"J:\Pedidos 2029", prefixo + "*");
                    if (texto.Substring(texto.IndexOf("Data de emissão: ", StringComparison.OrdinalIgnoreCase) + 6, 4) == "2030")
                        diretorioContratos = Directory.GetDirectories(@"J:\Pedidos 2030", prefixo + "*");
                }
            }
            

            bool pastaEncontrada = false;
            string destino = string.Empty;

            string armazenarPlanejadores = "";
            string armazenarSubPlanejadores = "";
            foreach (var pastacontrato in diretorioContratos)
            {
                foreach (var pastaambiente in Directory.GetDirectories(pastacontrato))
                {
                    string nomePastaAmbiente = Path.GetFileName(pastaambiente);
                    if (nomePastaAmbiente.StartsWith(numAmbiente))
                    {
                        destino = Path.Combine(pastaambiente, $"{ordemcompra} - Lista de Planejadores.pdf");
                        pastaEncontrada = true;

                        richtxtPainel.AppendText($"Pasta encontrada para a ordem de compra {ordemcompra}: {pastaambiente}\n");
                        richtxtPainel.ScrollToCaret();
                        break;
                    }
                }
                if (pastaEncontrada) break;
            }

            // Se não encontrou, cria pasta no mesmo diretório do PDF
            if (!pastaEncontrada)
            {
                string pastaDestino = Path.Combine(pastaOrigem, ordemcompra);
                if (!Directory.Exists(pastaDestino))
                    Directory.CreateDirectory(pastaDestino);

                destino = Path.Combine(pastaDestino, $"{ordemcompra} - Lista de Planejadores.pdf");
                richtxtPainel.AppendText($"Pasta não encontrada. Criando destino padrão: {destino}\n");
                richtxtPainel.ScrollToCaret();
            }
            if (File.Exists(destino))
                File.Delete(destino);

            // Criar PDF
            bool encontrou = false;
            using (var pdfOrigem = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(origem)))            
            using (var pdfDestino = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfWriter(destino)))
            {
                for (int i = 1; i <= pdfOrigem.GetNumberOfPages(); i++)
                {
                    var strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy();
                    string texto = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(pdfOrigem.GetPage(i), strategy);
                    if (texto.IndexOf(ordemcompra, StringComparison.OrdinalIgnoreCase) >= 0)
                    {

                        pdfOrigem.CopyPagesTo(i, i, pdfDestino);

                        foreach (var planejador in planejadores)
                        {
                            if (texto.IndexOf(planejador, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                richtxtPainel.AppendText($"Planejador '{planejador}' encontrado na página {i} da ordem de compra {ordemcompra}.\n");
                                richtxtPainel.ScrollToCaret();
                                if(planejador.Contains("PINTURA ACESSORIOS", StringComparison.OrdinalIgnoreCase) && !armazenarSubPlanejadores.Contains("P.A"))
                                    armazenarSubPlanejadores += "P.A | ";
                                else if (!armazenarSubPlanejadores.Contains(planejador.Substring(planejador.IndexOf(@"-") + 2, 1)))
                                    armazenarSubPlanejadores += planejador.Substring(planejador.IndexOf(@"-") + 2, 1) +  " | ";
                                if (!armazenarPlanejadores.Contains(planejador))
                                    armazenarPlanejadores += planejador + " | ";

                            }
                        }

                        richtxtPainel.AppendText($"Página {i} adicionada em: {ordemcompra}\nSalvo em: {destino}\n");
                        richtxtPainel.ScrollToCaret();
                        encontrou = true;

                    }
                    else if (texto.Contains(numContrato) && texto.Contains(nomeAmbiente))
                    {

                        pdfOrigem.CopyPagesTo(i, i, pdfDestino);

                        foreach (var planejador in planejadores)
                        {
                            if (texto.IndexOf(planejador, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                richtxtPainel.AppendText($"Planejador '{planejador}' encontrado na página {i} da ordem de compra {ordemcompra}.\n");
                                richtxtPainel.ScrollToCaret();
                                if (planejador.Contains("PINTURA ACESSORIOS", StringComparison.OrdinalIgnoreCase) && !armazenarSubPlanejadores.Contains("P.A"))
                                    armazenarSubPlanejadores += "P.A | ";
                                else if (!armazenarSubPlanejadores.Contains(planejador.Substring(planejador.IndexOf(@"-") + 2, 1)))
                                    armazenarSubPlanejadores += planejador.Substring(planejador.IndexOf(@"-") + 2, 1) + " | ";
                                if (!armazenarPlanejadores.Contains(planejador))
                                    armazenarPlanejadores += planejador + " | ";

                            }
                        }

                        richtxtPainel.AppendText($"Página {i} adicionada em: {ordemcompra}\nSalvo em: {destino}\n");
                        richtxtPainel.ScrollToCaret();
                        encontrou = true;

                    }
                }
            }

            Armazenar armazenar = new Armazenar(ordemcompra, armazenarSubPlanejadores, armazenarPlanejadores);
            listaArmazenar.Add(armazenar);
            if (!encontrou)
                pastasNaoEncontradas.Add(ordemcompra);
                
            // Se não encontrou nada, apaga PDF criado
            if (!encontrou && File.Exists(destino))
                File.Delete(destino);

            
        }

        private void btnCarregarPDF_Click(object sender, EventArgs e)
        {
            lblAgPDF.Text = "Aguardando arquivo PDF...";
            pdfPath = string.Empty;
            pdfCarregado = false;

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "PDF Files|*.pdf";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                pdfPath = ofd.FileName;
                lblAgPDF.Text = "PDF carregado com sucesso!";
                pdfCarregado = true;
                richtxtPainel.AppendText("Arquivo PDF selecionado: " + ofd.FileName + "\n");
            }
            else
            {
                MessageBox.Show("Nenhum arquivo PDF selecionado.");
            }

            if (excelCarregado == true && pdfCarregado == true)
            {
                lblStatus.Text = "Processando...";
                Application.DoEvents();
                lblStatus.Text = "Processamento concluído!";
                lblStatus.Text = "Pronto para começar!";
            }


        }

        private void lblAgExcel_Click(object sender, EventArgs e)
        {

        }

        private void lblAgPDF_Click(object sender, EventArgs e)
        {

        }

        public void btnIniciarSep_Click(object sender, EventArgs e)
        {
            if (excelCarregado != true && pdfCarregado != true)
            {
                lblStatus.Text = "Processando...";
                Application.DoEvents();                
                lblStatus.Text = "Processamento concluído!";
                lblStatus.Text = "Ocioso";
                MessageBox.Show("Por favor, carregue ambos os arquivos (Excel e PDF) antes de iniciar o processamento.");
                return;
            }
            if (excelCarregado != true && pdfCarregado == true)
            {
                lblStatus.Text = "Processando...";
                Application.DoEvents();
                lblStatus.Text = "Processamento concluído!";
                lblStatus.Text = "Ocioso";
                MessageBox.Show("Por favor, carregue o arquivo Excel antes de iniciar o processamento.");
                return;
            }
            if (excelCarregado == true && pdfCarregado != true)
            {
                lblStatus.Text = "Processando...";
                Application.DoEvents();
                lblStatus.Text = "Processamento concluído!";
                lblStatus.Text = "Ocioso";
                MessageBox.Show("Por favor, carregue o arquivo PDF antes de iniciar o processamento.");
                return;

            }
            foreach (var pedido in lista)
            {
                CriarPDF(pedido);
            }
            CriarExcel(listaArmazenar);
            
            if (pastasNaoEncontradas.Count() > 0)
            {
                richtxtPainel.AppendText("\n\n===== RESUMO FINAL =====\n\nPasta não localizada dos ambientes:\n\n");
                foreach(var i in pastasNaoEncontradas)
                {
                    richtxtPainel.AppendText($"\n{i}\n");
                    richtxtPainel.ScrollToCaret();
                }

                richtxtPainel.AppendText("\n\n===== FIM =====\n\nPor favor, validar e copiar para o local correto.\n\n");
                richtxtPainel.ScrollToCaret();
            }
            else
                richtxtPainel.AppendText("\n\nTodos os pedidos foram encontrado suas pastas. 🗸 \n\n");

            richtxtPainel.AppendText("\n\nProcesso finalizado! ✅\n\n");
            richtxtPainel.ScrollToCaret();
            lblStatus.Text = "Processo Concluído! ✅";


        }
        public void CriarExcel(List<Armazenar> listaArmazenar)
        {
            string caminhoExcel = Path.Combine(Path.GetDirectoryName(pdfPath), "Pedidos e Planejadores.xlsx");
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Pedidos");

                    // Cabeçalho
                    worksheet.Cell(1, 1).Value = "Ordem de Compra";
                    worksheet.Cell(1, 2).Value = "Sub Planejadores";
                    worksheet.Cell(1, 3).Value = "Planejadores";

                    // Preencher linhas
                    int linha = 2;
                    foreach (var item in listaArmazenar)
                    {
                        worksheet.Cell(linha, 1).Value = item.OrdemCompra;
                        worksheet.Cell(linha, 2).Value = item.SubPlanejadores; // corrigido índice
                        worksheet.Cell(linha, 3).Value = item.Planejadores;
                        linha++;
                    }

                    // Auto-ajustar colunas
                    worksheet.Columns().AdjustToContents();

                    // Salvar arquivo
                    workbook.SaveAs(caminhoExcel);
                    richtxtPainel.AppendText($"Excel de Pedidos e seus planejadores gerado em: {caminhoExcel}\n");
                    richtxtPainel.ScrollToCaret();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao gerar o Excel: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                richtxtPainel.AppendText($"Falha ao gerar Excel: {ex.Message}\n");
                richtxtPainel.ScrollToCaret();
            }

        }

        private void lblStatus_Click(object sender, EventArgs e)
        {

        }
    }
}
