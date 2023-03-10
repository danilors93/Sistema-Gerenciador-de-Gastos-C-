using System;
using System.IO;
using OfficeOpenXml;

namespace GerenciadorFinanceiro
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("Bem-vindo ao seu gerenciador financeiro!\n");
            Console.Write("Informe o mês atual: ");
            string mes = Console.ReadLine();

            Console.Write("Informe o saldo inicial disponível: ");
            double saldoinicial = double.Parse(Console.ReadLine());
            double saldo = saldoinicial;

            string caminhoArquivo = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{mes}.xlsx");
            FileInfo arquivo = new FileInfo(caminhoArquivo);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage pacote = new ExcelPackage(arquivo);
            ExcelWorksheet planilhaMes = pacote.Workbook.Worksheets.Add(mes);

            planilhaMes.Cells[1, 1].Value = "Tipo de Gasto";
            planilhaMes.Cells[1, 2].Value = "Valor";

            int numLinhas = 1;
            while (true)
            {
                Console.Write("\nNome do gasto (ou 'fim' para encerrar): ");
                string nomeGasto = Console.ReadLine();

                if (nomeGasto.ToUpper() == "FIM")
                    break;

                Console.Write("Valor do gasto: ");
                double valorGasto = double.Parse(Console.ReadLine());

                planilhaMes.Cells[numLinhas + 1, 1].Value = nomeGasto;
                planilhaMes.Cells[numLinhas + 1, 2].Value = valorGasto;
                numLinhas++;

                saldo -= valorGasto;
                Console.WriteLine($"\nGasto de {valorGasto:C} registrado com sucesso!");
                Console.WriteLine($"Saldo atual: {saldo:C}");
            }

            planilhaMes.Cells[numLinhas + 2, 1].Value = "Saldo final:";
            planilhaMes.Cells[numLinhas + 2, 2].Value = saldo;

            pacote.Save();

            Console.WriteLine($"\nGastos registrados com sucesso na planilha {mes}.xlsx!");
            Console.WriteLine($"Valor total gasto no mês: {(saldoinicial - saldo):C}");
            Console.WriteLine($"Saldo final disponível: {saldo:C}");

            Console.Write("\nDeseja abrir a planilha agora? (S/N): ");
            string resposta = Console.ReadLine().ToUpper();

            if (resposta == "s")
                System.Diagnostics.Process.Start(caminhoArquivo);

            Console.WriteLine("\nObrigado por utilizar o seu gerenciador financeiro!");
        }
    }
}
