using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Runtime.Intrinsics.X86;
using System.Text;
using UglyToad.PdfPig;

class Program
{
    static void Main()
    {
        string pasta = @"C:\Folder";   // insert folder location | insira diretorio da pasta
        string textoProcurado = "Keyword"; // insert keyword you're looking for | insira a palavra chave que está procurando

        foreach (string arquivo in Directory.GetFiles(pasta))
        {
            string texto = ""; 

            try
            {
                if (arquivo.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) // pdf block | bloco pdf
                {
                    texto = LerPdf(arquivo);
                }
                else if (arquivo.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) // word block | bloco word
                {
                    texto = LerDocx(arquivo);
                }
                else if (arquivo.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) // txt block | bloco txt
                {
                    texto = File.ReadAllText(arquivo);
                }

                if (!string.IsNullOrWhiteSpace(texto) &&
                    texto.IndexOf(textoProcurado, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"Encontrado em: {Path.GetFileName(arquivo)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao ler {Path.GetFileName(arquivo)}: {ex.Message}"); // error
            }
        }

        Console.WriteLine("Busca finalizada."); // success | sucesso
        Console.ReadKey();
    }

    static string LerPdf(string caminho)
    {
        StringBuilder texto = new StringBuilder();

        using (var pdf = PdfDocument.Open(caminho))
        {
            foreach (var pagina in pdf.GetPages())
            {
                texto.AppendLine(pagina.Text);
            }
        }

        return texto.ToString();
    }

    static string LerDocx(string caminho)
    {
        StringBuilder texto = new StringBuilder();

        using (WordprocessingDocument doc =
               WordprocessingDocument.Open(caminho, false))
        {
            Body body = doc.MainDocumentPart.Document.Body;

            foreach (var paragrafo in body.Elements<Paragraph>())
            {
                texto.AppendLine(paragrafo.InnerText);
            }
        }

        return texto.ToString();
    }
}

