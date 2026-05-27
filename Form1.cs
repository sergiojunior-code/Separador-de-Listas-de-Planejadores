using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Runtime.CompilerServices;
using Path = System.IO.Path;


namespace Separador_de_Listas_de_Planejadores
{
    public partial class Form1 : Form
    {
        public string pdfPath;
        public string excelPath;
        public bool excelCarregado = false;
        public bool pdfCarregado = false;
        public string resumoFinal = "";

        public string origem = "";
        public string pastaOrigem = "";
        public string nomeArquivo = "";
        public string destinoAncoragem = "";
        public string destinoTensionada = "";
        public string versao = "1.4";

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
            this.Text = $"[{versao}] - Separador de Lista de Planejadores";


        }
        public class Armazenar
        {
            // Propriedades públicas
            public string ERP { get; set; }
            public string OrdemCompra { get; set; }
            public string SubPlanejadores { get; set; }
            public string Planejadores { get; set; }


            // Construtor
            public Armazenar(string erp,string ordemcompra, string subplanejadores, string planejadores)
            {
                ERP = erp;
                OrdemCompra = ordemcompra;
                SubPlanejadores = subplanejadores;
                Planejadores = planejadores;

            }
        }

        public List<string> lista = new List<string>();
        public List<string> planejadores = new List<string>();
        public List<string> ListaERP = new List<string>();
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
                            string valorERP = row.Cell(2).GetString().Trim();
                            if (!string.IsNullOrEmpty(valorERP))
                                ListaERP.Add(valorERP);
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

            origem = pdfPath;
            pastaOrigem = Path.GetDirectoryName(origem);
            nomeArquivo = Path.GetFileNameWithoutExtension(pdfPath);

            destinoAncoragem = origem.Replace(nomeArquivo, "Lista de Planejador - Ancoragens");
            destinoTensionada = origem.Replace(nomeArquivo, "Lista de Planejador - Tela Tensionadas");

            // Extrair números do pedido
            //string fourDigits = ordemcompra.Substring(0, 4);
            //string fiveDigits = ordemcompra.Substring(0, 5);
            //int firstDigits = 5;

            //int tamanhoPrefixo = ordemcompra.Substring(0, ordemcompra.IndexOf(" - ")).Length;
            //if (tamanhoPrefixo == 7) firstDigits = 5;
            //else if (tamanhoPrefixo == 6) firstDigits = 4;



            // Extrair numAmbiente corretamente
            int indiceHifen = ordemcompra.IndexOf("- ");
            string numAmbiente = ordemcompra.Substring(indiceHifen + 2, 2);

            richtxtPainel.AppendText($"Processando ordem de compra: {ordemcompra} (Ambiente: {numAmbiente})\n");
            richtxtPainel.ScrollToCaret();

            // Buscar diretórios de contratos filtrados
            string prefixo = ordemcompra.Substring(0, ordemcompra.IndexOf(" - ")).Trim();
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

                        richtxtPainel.AppendText($"Pasta encontrada para a ordem de compra {ordemcompra}: {pastaambiente}\n\n");
                        richtxtPainel.ScrollToCaret();
                        break;
                    }
                }
                if (pastaEncontrada == true) break;
            }

            // Se não encontrou, cria pasta no mesmo diretório do PDF
            if (pastaEncontrada == false)
            {
                string pastaDestino = Path.Combine(pastaOrigem, ordemcompra);
                if (!Directory.Exists(pastaDestino))
                    Directory.CreateDirectory(pastaDestino);

                destino = Path.Combine(pastaDestino, $"{ordemcompra} - Lista de Planejadores.pdf");
                richtxtPainel.AppendText($"Pasta não encontrada. Criando destino padrão: {destino}\n\n");
                richtxtPainel.ScrollToCaret();
            }
            if (File.Exists(destino))
            {
                try
                {
                    richtxtPainel.AppendText($"Arquivo existente encontrado. Excluindo: {destino}\n\n");

                    File.Delete(destino);
                }
                catch (Exception ex)
                {
                    richtxtPainel.AppendText($"Erro ao excluir arquivo existente: {ex.Message}\n\n");
                }
                richtxtPainel.ScrollToCaret();
            }
            // Criar PDF
            bool encontrou = false;
            try
            {
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
                                    richtxtPainel.AppendText($"'{planejador}' encontrado na página {i} da ordem de compra {ordemcompra}.\n\n");
                                    richtxtPainel.ScrollToCaret();
                                    if (planejador.Contains("PINTURA ACESSORIOS", StringComparison.OrdinalIgnoreCase) && !armazenarSubPlanejadores.Contains("P.A"))
                                        armazenarSubPlanejadores += "P.A | ";
                                    else if (!armazenarSubPlanejadores.Contains(planejador.Substring(planejador.IndexOf(@"-") + 2, 1)))
                                        armazenarSubPlanejadores += planejador.Substring(planejador.IndexOf(@"-") + 2, 1) + " | ";
                                    if (!armazenarPlanejadores.Contains(planejador))
                                        armazenarPlanejadores += planejador + " | ";

                                }
                            }

                            richtxtPainel.AppendText($"Página {i} adicionada em: {ordemcompra}\n");
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
                                    richtxtPainel.AppendText($"'{planejador}' encontrado na página {i} da ordem de compra {ordemcompra}.\n\n");
                                    richtxtPainel.ScrollToCaret();
                                    if (planejador.Contains("PINTURA ACESSORIOS", StringComparison.OrdinalIgnoreCase) && !armazenarSubPlanejadores.Contains("P.A"))
                                        armazenarSubPlanejadores += "P.A | ";
                                    else if (!armazenarSubPlanejadores.Contains(planejador.Substring(planejador.IndexOf(@"-") + 2, 1)))
                                        armazenarSubPlanejadores += planejador.Substring(planejador.IndexOf(@"-") + 2, 1) + " | ";
                                    if (!armazenarPlanejadores.Contains(planejador))
                                        armazenarPlanejadores += planejador + " | ";

                                }
                            }

                            richtxtPainel.AppendText($"Página {i} adicionada em: {ordemcompra}\n");

                            encontrou = true;

                        }
                    }
                }
                richtxtPainel.AppendText($"Salvo em: {destino}\n\n");
                richtxtPainel.ScrollToCaret();
                var numERP = "";
                foreach (var erp in ListaERP)
                {

                    var index = lista.IndexOf(ordemcompra);
                    if (ListaERP.IndexOf(erp) == index)
                    {
                        numERP = erp; break;
                    }
                }
                Armazenar armazenar = new Armazenar(numERP, ordemcompra, armazenarSubPlanejadores, armazenarPlanejadores);
                listaArmazenar.Add(armazenar);
                if (!encontrou)
                    pastasNaoEncontradas.Add(ordemcompra);

                // Se não encontrou nada, apaga PDF criado
                if (!encontrou && File.Exists(destino))
                    File.Delete(destino);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao processar o PDF: {ex.Message}");
                richtxtPainel.AppendText($"Erro ao processar o PDF: {ex.Message}\n\n");
                richtxtPainel.ScrollToCaret();

                // Log
                string registro = $"{DateTime.Now}";
                richtxtPainel.AppendText("Horário do registro: " + registro);
                var textoLog = richtxtPainel.Text;
                string caminhoLog = Path.Combine(pastaOrigem, "Separador de Lista de Planejador_Log.txt");
                File.AppendAllText(caminhoLog, textoLog + Environment.NewLine);

            }


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

        public void CriarPDFTensionada()
        {
            bool encontrouTensionada = false;
            if (File.Exists(destinoTensionada))
            {
                try
                {
                    richtxtPainel.AppendText($"Arquivo de tela tensionada encontrado. Excluindo: {destinoTensionada}\n\n");
                    File.Delete(destinoTensionada);
                }
                catch (Exception ex)
                {
                    richtxtPainel.AppendText($"Erro ao excluir arquivo existente: {ex.Message}\n\n");
                }
                richtxtPainel.ScrollToCaret();
            }
            richtxtPainel.AppendText("Criando PDF de tela tensionada...\n");
            try
            {
                using (var pdfOrigem = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(origem)))
                using (var pdfTensionadas = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfWriter(destinoTensionada)))
                {
                    for (int i = 1; i <= pdfOrigem.GetNumberOfPages(); i++)
                    {
                        var strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy();
                        string texto = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(pdfOrigem.GetPage(i), strategy);

                        if (texto.IndexOf("TLTE002001", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            pdfOrigem.CopyPagesTo(i, i, pdfTensionadas);
                            richtxtPainel.AppendText($"Página com tela tensionada encontrada na página {i}.\n");
                            richtxtPainel.ScrollToCaret();

                            encontrouTensionada = true;
                        }
                    }
                }
                richtxtPainel.AppendText("\n");
                richtxtPainel.ScrollToCaret();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao processar o PDF: {ex.Message}");
                richtxtPainel.AppendText($"Erro ao processar o PDF: {ex.Message}\n\n");
                richtxtPainel.ScrollToCaret();

                // Log
                string registro = $"{DateTime.Now}";
                richtxtPainel.AppendText("Horário do registro: " + registro);
                var textoLog = richtxtPainel.Text;
                string caminhoLog = System.IO.Path.Combine(pastaOrigem, "Separador de Lista de Planejador_Log.txt");
                File.AppendAllText(caminhoLog, textoLog + Environment.NewLine);
            }

            if (!encontrouTensionada)
            {
                richtxtPainel.AppendText($"Nenhuma página com tela tensionada encontrada no PDF.\n\n");
                richtxtPainel.ScrollToCaret();
                File.Delete(destinoTensionada);
            }
            else
            {
                richtxtPainel.AppendText($"PDF de tela tensionada criado em: {destinoTensionada}\n\n");
                richtxtPainel.ScrollToCaret();
            }
        }

        public void CriarPDFAncoragem()
        {
            bool encontrouAncoragem = false;
            richtxtPainel.AppendText($"Criando PDF de ancoragem...\n");
            try
            {
                if (File.Exists(destinoAncoragem))
                {
                    try
                    {
                        richtxtPainel.AppendText($"Arquivo de ancoragem encontrado. Excluindo: {destinoAncoragem}\n\n");
                        File.Delete(destinoAncoragem);
                    }
                    catch (Exception ex)
                    {
                        richtxtPainel.AppendText($"Erro ao excluir arquivo existente: {ex.Message}\n\n");
                    }
                    richtxtPainel.ScrollToCaret();
                }

                using (var pdfOrigem = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(origem)))
                using (var pdfAncoragens = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfWriter(destinoAncoragem)))
                {
                    for (int i = 1; i <= pdfOrigem.GetNumberOfPages(); i++)
                    {
                        var strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy();
                        string texto = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(pdfOrigem.GetPage(i), strategy);

                        if (texto.IndexOf("ANAP999001", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            pdfOrigem.CopyPagesTo(i, i, pdfAncoragens);
                            richtxtPainel.AppendText($"Página de ancoragem encontrada na página {i}.\n");
                            

                            encontrouAncoragem = true;

                        }
                    }
                }
                richtxtPainel.AppendText("\n"); 
                richtxtPainel.ScrollToCaret();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao processar o PDF: {ex.Message}");
                richtxtPainel.AppendText($"Erro ao processar o PDF: {ex.Message}\n\n");
                richtxtPainel.ScrollToCaret();

                // Log
                string registro = $"{DateTime.Now}";
                richtxtPainel.AppendText("Horário do registro: " + registro);
                var textoLog = richtxtPainel.Text;
                string caminhoLog = Path.Combine(pastaOrigem, "Separador de Lista de Planejador_Log.txt");
                File.AppendAllText(caminhoLog, textoLog + Environment.NewLine);
            }
            if(!encontrouAncoragem)
            {
                richtxtPainel.AppendText($"Nenhuma página de ancoragem encontrada no PDF.\n\n");
                richtxtPainel.ScrollToCaret();
                File.Delete(destinoAncoragem);
            }
            else
            {
                richtxtPainel.AppendText($"PDF de ancoragem criado em: {destinoAncoragem}\n\n");
                richtxtPainel.ScrollToCaret();
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

            // Garante a limpeza da lista de ambientes toda vez que inicia o processo
            listaArmazenar.Clear();

            foreach (var pedido in lista)
            {
                CriarPDF(pedido);

            }
            CriarPDFAncoragem();
            CriarPDFTensionada();
            CriarExcel(listaArmazenar);

            if (pastasNaoEncontradas.Count() > 0)
            {
                richtxtPainel.AppendText("\n\n===== RESUMO FINAL =====\n\nPasta não localizada dos ambientes:\n\n");
                foreach (var i in pastasNaoEncontradas)
                {
                    richtxtPainel.AppendText($"\n{i}\n");
                    richtxtPainel.ScrollToCaret();
                }

                richtxtPainel.AppendText("\n\n===== FIM =====\n\nPor favor, validar e copiar para o local correto.\n\n");
                richtxtPainel.ScrollToCaret();
            }
            else
                richtxtPainel.AppendText("\n\nTodos os pedidos foram encontrado suas pastas. 🗸 \n\n");

            richtxtPainel.AppendText("\n\nProcesso finalizado! ✅\n\n * Log gerado na pasta de origem\n\n");
            richtxtPainel.ScrollToCaret();

            // Log
            string registro = $"{DateTime.Now}";
            richtxtPainel.AppendText("Horário do registro: " + registro);
            var textoLog = richtxtPainel.Text;
            string caminhoLog = Path.Combine(pastaOrigem, "Separador de Lista de Planejador_Log.txt");
            File.AppendAllText(caminhoLog, textoLog + Environment.NewLine);

            // Status finalizado
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
                    worksheet.Cell(1, 1).Value = "ERP";
                    worksheet.Cell(1, 2).Value = "Ordem de Compra";
                    worksheet.Cell(1, 3).Value = "Sub Planejadores";
                    worksheet.Cell(1, 4).Value = "Planejadores";

                    // Preencher linhas
                    int linha = 2;
                    foreach (var item in listaArmazenar)
                    {
                        worksheet.Cell(linha, 1).Value = item.ERP;
                        worksheet.Cell(linha, 2).Value = item.OrdemCompra;
                        worksheet.Cell(linha, 3).Value = item.SubPlanejadores; // corrigido índice
                        worksheet.Cell(linha, 4).Value = item.Planejadores;
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

        private void btnAtualizacoes_Click(object sender, EventArgs e)
        {
            FormAtualizacoes formAtualizacoes = new FormAtualizacoes();
            formAtualizacoes.ShowDialog();
        }
    }
}
