
// Autor: Leonardo Perosa Zago.

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;



public class Cliente
{
    public int Id { get; set; }
    public int IntervaloDeChegada { get; set; } // OK Intervalo
    public int MomChegada { get; set; } // OK 1 Momento de chegada = Intervalo + Intervalo anterior
    public int TempoDeAtendimento { get; set; } // 3 OK Duração
    public int InicioAtendimento { get; set; } // 2 OK Começa a ser atendido = 
    public int FimAtendimento { get; set; } // OK 4 Fim de atendimento = InicioAtendimento + TempoDeAtendimento ou Duração
    public int IniFila { get; set; } // 5 OK Momchegada caso ele seja < FimAtendimento do anterior
    public int FimFila {  get; set; } // 6 OK Caso haja fila = IniFila + Duração
    public int TempodeFila { get; set; } // OK 7 = FimFila - IniFila
}

class Program
{
    static void Main(string[] args)
    {
        List<Cliente> listaDeClientes = new List<Cliente>();

        string input;
        int idsum = 1;
        int anteriorMomChegada = 0;
        int anteriorFim = 0;

        while (true)
        {
            Cliente cliente = new Cliente();

            cliente.Id = idsum;

            Console.Write("Digite o intervalo de chegada do cliente número " + cliente.Id + ": ");
            string intervalo = Console.ReadLine();

            if (int.TryParse(intervalo, out int inte))
            {
                if (inte >= 0)
                {
                    cliente.IntervaloDeChegada = inte;
                    Console.WriteLine($"O intervalo de chegada: {cliente.IntervaloDeChegada}");
                }
                else
                {
                    Console.WriteLine("O intervalo de chegada deve ser um número positivo.");
                }
            }
            else
            {
                Console.WriteLine("Valor inválido. Por favor, digite um número inteiro.");
            }

            Console.Write("Digite a duração do atendimento: ");
            string dura = Console.ReadLine();

            if (int.TryParse(dura, out int dur))
            {
                if (dur > 0)
                {
                    cliente.TempoDeAtendimento = dur;
                    Console.WriteLine($"A duração do atendimento é: {cliente.TempoDeAtendimento}");
                }
                else
                {
                    Console.WriteLine("A duração do atendimento deve ser maior que zero.");
                }
            }
            else
            {
                Console.WriteLine("Valor inválido. Por favor, digite um número inteiro.");
            }

            // Momento de chegada:
            if (cliente.Id == 1)
            {
                cliente.MomChegada = cliente.IntervaloDeChegada;
            }
            else
            {
                cliente.MomChegada = cliente.IntervaloDeChegada + anteriorMomChegada;
            }

            anteriorMomChegada = cliente.MomChegada;

            // Inicio do atendimento:
            if (cliente.MomChegada >= anteriorFim)
            {
                cliente.InicioAtendimento = cliente.MomChegada;
            }
            else
            {
                cliente.InicioAtendimento = anteriorFim;
            }

            // IniFila:
            if (cliente.MomChegada < anteriorFim)
            {
                cliente.IniFila = cliente.MomChegada;
            }
            else
            {
                cliente.IniFila = 0;
            }

            // FimFila
            if (cliente.IniFila != 0)
            {
                cliente.FimFila = anteriorFim;
            }
            else
            {
                cliente.FimFila = 0;
            }

            // Fim do atendimento:
            cliente.FimAtendimento = cliente.InicioAtendimento + cliente.TempoDeAtendimento;

            anteriorFim = cliente.FimAtendimento;

            //Tempo de fila
            cliente.TempodeFila = cliente.FimFila - cliente.IniFila;


            listaDeClientes.Add(cliente);

            idsum++;

            Console.Write("Caso tenha terminado a lista digite 'ok'.");
            input = Console.ReadLine();
            if (input.ToLower() == "ok")
            {
                break;
            }
        }

        foreach (Cliente cliente in listaDeClientes)
        {
            Console.WriteLine($"Cliente ID: {cliente.Id}, Intervalo de chegada: {cliente.IntervaloDeChegada}, Tempo de atendimento: {cliente.TempoDeAtendimento}");
        }

        // Intervalo médio entre as chegadas
        double somaIntervalos = listaDeClientes.Sum(cliente => cliente.IntervaloDeChegada);
        double mediaIntervalos = somaIntervalos / listaDeClientes.Count;

        // Duração média dos atendimentos
        double somaDuracoes = listaDeClientes.Sum(cliente => cliente.TempoDeAtendimento);
        double mediaDuracoes = somaDuracoes / listaDeClientes.Count;

        // Tamanho médio da fila
        double somaFila = listaDeClientes.Sum(cliente => cliente.TempodeFila);
        double tempoMedioFila = somaFila / listaDeClientes.Count;

        // Número médio na fila
        Cliente ultimoCliente = listaDeClientes.Last();
        double numeroMedioFila = somaFila / ultimoCliente.FimAtendimento;

        // Utiliza o NuGet EPPlus para exportar em planilha excel
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string diretorioProjeto = Directory.GetCurrentDirectory();
        string pastaPlanilha = Path.Combine(diretorioProjeto, "Planilha");
        Directory.CreateDirectory(pastaPlanilha);
        string nomeArquivo = "Relatorio_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
        string caminhoCompleto = Path.Combine(pastaPlanilha, nomeArquivo);

        using (ExcelPackage package = new ExcelPackage())
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Clientes");

            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Intervalo de Chegada";
            worksheet.Cells[1, 3].Value = "Momento de Chegada";
            worksheet.Cells[1, 4].Value = "Início Atendimento";
            worksheet.Cells[1, 5].Value = "Duração / Tempo de atendimento";
            worksheet.Cells[1, 6].Value = "Fim do atendimento";
            worksheet.Cells[1, 7].Value = "Início Fila";
            worksheet.Cells[1, 8].Value = "Fim Fila";
            worksheet.Cells[1, 9].Value = "Tempo de fila";
            worksheet.Cells[1, 12].Value = "Intervalo médio entre chegadas";
            worksheet.Cells[2, 12].Value = mediaIntervalos;
            worksheet.Cells[1, 13].Value = "Duração média dos atendimentos";
            worksheet.Cells[2, 13].Value = mediaDuracoes;
            worksheet.Cells[1, 14].Value = "Tempo médio de fila";
            worksheet.Cells[2, 14].Value = tempoMedioFila;
            worksheet.Cells[1, 15].Value = "Número médio na fila";
            worksheet.Cells[2, 15].Value = numeroMedioFila;


            int row = 2;
            foreach (var cliente in listaDeClientes)
            {
                worksheet.Cells[row, 1].Value = cliente.Id;
                worksheet.Cells[row, 2].Value = cliente.IntervaloDeChegada;
                worksheet.Cells[row, 3].Value = cliente.MomChegada;
                worksheet.Cells[row, 4].Value = cliente.InicioAtendimento;
                worksheet.Cells[row, 5].Value = cliente.TempoDeAtendimento;
                worksheet.Cells[row, 6].Value = cliente.FimAtendimento;
                worksheet.Cells[row, 7].Value = cliente.IniFila;
                worksheet.Cells[row, 8].Value = cliente.FimFila;
                worksheet.Cells[row, 9].Value = cliente.TempodeFila;
                row++;
            }

            FileInfo excelFile = new FileInfo(caminhoCompleto);
            package.SaveAs(excelFile);
        }
        Console.WriteLine("Planilha salva com sucesso em: " + caminhoCompleto);
    }
}