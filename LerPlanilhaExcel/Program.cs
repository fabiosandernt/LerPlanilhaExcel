using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using System.Text;
using System.Threading.Tasks;
using System.IO; // A BIBLIOTECA DE ENTRADA E SAIDA DE ARQUIVOS
using iTextSharp; //E A BIBLIOTECA ITEXTSHARP E SUAS EXTENÇÕES
using iTextSharp.text; //ESTENSAO 1 (TEXT)
using iTextSharp.text.pdf;//ESTENSAO 2 (PDF)


namespace LerPlanilhaExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            
            IXLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            wb.SaveAs(@"C:\Projects\LerPlanilhaExcel\LerPlanilhaExcel\excel\MeuExcel.xlsx");
                        
            var workbook = new XLWorkbook(@"C:\Projects\LerPlanilhaExcel\LerPlanilhaExcel\excel\MEI ATIVO.xlsx");
            StreamWriter logCnpj = new StreamWriter(@"C:\Projects\LerPlanilhaExcel\LerPlanilhaExcel\log\logsCnpj.txt");

            int contador = 0;
            int cont = 2;
            
            int praca = workbook.Worksheets.Count;
            Console.WriteLine(praca);
            
            for (int i = 1; i <= praca; i++)
            {
                var sheet = workbook.Worksheet(i);

                var linha = 2;

                while (true)
                {

                    var email = sheet.Cell("O" + linha).Value.ToString();
                    var cnpj = sheet.Cell("C" + linha).Value.ToString();
                    var emailLider = sheet.Cell("P" + linha).Value.ToString();
                    var dataContrato = sheet.Cell("Q" + linha).Value.ToString();
                    var nome = sheet.Cell("B" + linha).Value.ToString();
                    var endereco = sheet.Cell("J" + linha).Value.ToString();
                    var cep = sheet.Cell("L" + linha).Value.ToString();

                    if (contador == 5) break;
                    if (string.IsNullOrWhiteSpace(cnpj) && (string.IsNullOrWhiteSpace(email)))
                    {
                        contador++;
                    }
                    else
                    {
                        Console.WriteLine(linha + ": " + sheet.ToString() + " - " + cnpj + " - " + email + " - " + emailLider + " - " + dataContrato);
                        
                        logCnpj.Write(linha + " - "+ sheet + " - " + cnpj + " - " + email + " - " + emailLider + " - " + dataContrato + "\r\n");
                        ws.Cell(cont, 1).Value = cnpj;
                        ws.Cell(cont, 2).Value = nome;
                        ws.Cell(cont, 3).Value = email;
                        ws.Cell(cont, 4).Value = endereco;
                        ws.Cell(cont, 5).Value = cep;
                        ws.Cell(cont, 6).Value = dataContrato;
                        ws.Cell(cont, 7).Value = emailLider;

                        wb.SaveAs(@"C:\Projects\LerPlanilhaExcel\LerPlanilhaExcel\excel\MeuExcel.xlsx");
                        cont++;
                    }

                    //System.Threading.Thread.Sleep(100);
                    linha++;

                }
                contador = 0;
            }

            workbook.Dispose();
            logCnpj.Close();
            Console.WriteLine("Total de Praças lidas: " + praca);
            
        }
               
    }
  
}





