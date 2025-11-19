using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace ExcelToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // Configuração de licença do QuestPDF (Community License - grátis)
            QuestPDF.Settings.License = LicenseType.Community;

            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine("CONVERSOR EXCEL → PDF PROFISSIONAL");
            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine();

            string excelPath;
            string pdfPath;

            // Se passar argumentos na linha de comandos
            if (args.Length > 0)
            {
                excelPath = args[0];
                pdfPath = args.Length > 1 ? args[1] : Path.ChangeExtension(excelPath, ".pdf");
            }
            else
            {
                // Pedir ao utilizador
                Console.Write("Caminho do ficheiro Excel: ");
                excelPath = Console.ReadLine()?.Trim('"') ?? "";

                if (string.IsNullOrEmpty(excelPath))
                {
                    Console.WriteLine("Erro: Caminho não pode estar vazio!");
                    Console.WriteLine("Prima qualquer tecla para sair...");
                    Console.ReadKey();
                    return;
                }

                Console.Write("Nome do PDF (Enter para automático): ");
                string pdfInput = Console.ReadLine()?.Trim('"') ?? "";
                pdfPath = string.IsNullOrEmpty(pdfInput)
                    ? Path.ChangeExtension(excelPath, ".pdf")
                    : pdfInput;
            }

            try
            {
                ConvertExcelToPdf(excelPath, pdfPath);
                Console.WriteLine();
                Console.WriteLine("✓ PDF criado com sucesso!");
                Console.WriteLine($"✓ Localização: {Path.GetFullPath(pdfPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine($"✗ ERRO: {ex.Message}");
                Console.WriteLine($"Detalhes: {ex.InnerException?.Message}");
            }

            Console.WriteLine();
            Console.WriteLine("Prima qualquer tecla para sair...");
            Console.ReadKey();
        }

        static void ConvertExcelToPdf(string excelPath, string pdfPath)
        {
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("Ficheiro Excel não encontrado!", excelPath);

            Console.WriteLine($"A ler ficheiro: {excelPath}");

            // Ler Excel usando ClosedXML
            using (var workbook = new XLWorkbook(excelPath))
            {
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();

                if (range == null || range.RowCount() < 2)
                    throw new Exception("Excel vazio ou sem dados!");

                var rowCount = range.RowCount();
                var colCount = Math.Min(range.ColumnCount(), 3);

                Console.WriteLine($"Encontradas {rowCount - 1} linhas de dados");
                Console.WriteLine($"A criar PDF: {pdfPath}");

                // Extrair dados do Excel
                var headers = new string[colCount];
                var data = new List<string[]>();

                // Cabeçalhos (primeira linha)
                for (int col = 1; col <= colCount; col++)
                {
                    headers[col - 1] = range.Cell(1, col).Value.ToString();
                }

                // Dados (restantes linhas)
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new string[colCount];
                    for (int col = 1; col <= colCount; col++)
                    {
                        rowData[col - 1] = range.Cell(row, col).Value.ToString();
                    }
                    data.Add(rowData);
                }

                // Criar PDF usando QuestPDF
                Document.Create(container =>
                {
                    container.Page(page =>
                    {
                        page.Size(PageSizes.A4.Landscape());
                        page.Margin(40);
                        page.PageColor(Colors.White);
                        page.DefaultTextStyle(x => x.FontSize(10));

                        // Cabeçalho da página
                        page.Header()
                            .Column(column =>
                            {
                                column.Item().AlignCenter().Text("Inventário de Software por Categoria")
                                    .FontSize(18)
                                    .SemiBold()
                                    .FontColor(Colors.Grey.Darken3);


                                column.Item().PaddingVertical(10);
                            });

                        // Conteúdo - Tabela
                        page.Content()
                            .Table(table =>
                            {
                                // Definir colunas (25%, 25%, 50%)
                                table.ColumnsDefinition(columns =>
                                {
                                    columns.RelativeColumn(1);  // Programa
                                    columns.RelativeColumn(1);  // Categoria
                                    columns.RelativeColumn(2);  // Descrição
                                });

                                // Cabeçalho da tabela
                                table.Header(header =>
                                {
                                    foreach (var headerText in headers)
                                    {
                                        header.Cell()
                                            .Background(Colors.Blue.Darken3)
                                            .Padding(10)
                                            .AlignCenter()
                                            .AlignMiddle()
                                            .Text(text =>
                                            {
                                                text.Span(headerText)
                                                    .FontColor(Colors.White)
                                                    .FontSize(11)
                                                    .SemiBold();
                                            });
                                    }
                                });

                                // Linhas de dados
                                int rowIndex = 0;
                                foreach (var row in data)
                                {
                                    var isAlternate = rowIndex % 2 == 1;
                                    var bgColor = isAlternate ? Colors.Grey.Lighten4 : Colors.White;

                                    foreach (var cell in row)
                                    {
                                        table.Cell()
                                            .Background(bgColor)
                                            .Border(0.5f)
                                            .BorderColor(Colors.Grey.Lighten1)
                                            .Padding(8)
                                            .Text(cell)
                                            .FontSize(9);
                                    }

                                    rowIndex++;
                                }
                            });

                        // Rodapé da página
                        page.Footer()
                            .AlignCenter()
                            .Text(text =>
                            {
                                text.DefaultTextStyle(x => x.FontSize(8).FontColor(Colors.Grey.Medium));
                                text.Span($"Total de programas: {data.Count} | ");
                                text.Span("Realizado por Gonçalo Russo, aluno n.º 259");
                            });
                    });
                })
                .GeneratePdf(pdfPath);
            }
        }
    }
}
