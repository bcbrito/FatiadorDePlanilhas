using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // Definir o contexto da licença do EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Configuração de serviços e logging
        var serviceProvider = ConfigureServices();
        var logger = serviceProvider.GetService<ILogger<Program>>();

        try
        {
            var config = serviceProvider.GetService<IConfiguration>();
            string pastaEntrada = config["PastaEntrada"];
            string pastaSaida = config["PastaSaida"];
            string prefixoNomeArqSaida = config["prefixoNomeSaida"];
            string nomeArquivoEntrada = Path.Combine(pastaEntrada, "input.xlsx");
            int linhasPorDivisao = int.Parse(config["LinhasPorDivisao"]);

            // Criar as pastas de entrada e saída se não existirem
            CriarPastasSeNecessario(pastaEntrada);
            CriarPastasSeNecessario(pastaSaida);

            // Fazer backup dos arquivos antigos na pasta de saída
            FazerBackupArquivosAntigos(pastaSaida);

            ValidarArquivoEntrada(nomeArquivoEntrada);
            DividirArquivoExcel(nomeArquivoEntrada, linhasPorDivisao, pastaSaida, logger, prefixoNomeArqSaida);

            logger.LogInformation("Arquivo Excel foi dividido em vários arquivos com sucesso.");
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Ocorreu um erro durante o processamento.");
        }
    }

    static void FazerBackupArquivosAntigos(string pastaSaida)
    {
        if (Directory.Exists(pastaSaida))
        {
            var arquivosNaRaiz = Directory.GetFiles(pastaSaida, "*.xlsx");
            if (arquivosNaRaiz.Length > 0)
            {
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string pastaBackup = Path.Combine(pastaSaida, $"backup_{timestamp}");

                Directory.CreateDirectory(pastaBackup);

                foreach (var arquivo in arquivosNaRaiz)
                {
                    string nomeArquivo = Path.GetFileName(arquivo);
                    string destino = Path.Combine(pastaBackup, nomeArquivo);
                    File.Move(arquivo, destino);
                }
            }
        }
    }

    static void CriarPastasSeNecessario(string pasta)
    {
        if (!Directory.Exists(pasta))
        {
            Directory.CreateDirectory(pasta);
        }
    }

    static void ValidarArquivoEntrada(string nomeArquivoEntrada)
    {
        if (!File.Exists(nomeArquivoEntrada))
        {
            throw new FileNotFoundException($"O arquivo {nomeArquivoEntrada} não foi encontrado.");
        }
    }

    static void DividirArquivoExcel(string nomeArquivoEntrada, int linhasPorDivisao, string pastaSaida, ILogger logger, string prefixoNomeArqSaida)
    {
        var fileInfo = new FileInfo(nomeArquivoEntrada);
        using (var package = new ExcelPackage(fileInfo))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet?.Dimension == null)
            {
                throw new InvalidOperationException("A planilha está vazia ou não foi carregada corretamente.");
            }

            int totalLinhas = worksheet.Dimension.Rows;
            int totalColunas = worksheet.Dimension.Columns;
            int indiceArquivo = 1;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            for (int linhaInicial = 2; linhaInicial <= totalLinhas; linhaInicial += linhasPorDivisao)
            {
                var nomeNovoArquivo = Path.Combine(pastaSaida, $"{prefixoNomeArqSaida}_{timestamp}_Parte{indiceArquivo}.xlsx");
                var novoArquivo = new FileInfo(nomeNovoArquivo);
                using (var novoPacote = new ExcelPackage(novoArquivo))
                {
                    var novaPlanilha = novoPacote.Workbook.Worksheets.Add("Sheet1");

                    // Copiar a linha de cabeçalho
                    for (int col = 1; col <= totalColunas; col++)
                    {
                        novaPlanilha.Cells[1, col].Value = worksheet.Cells[1, col].Value;
                    }

                    // Copiar as linhas de dados
                    for (int linha = 0; linha < linhasPorDivisao && (linhaInicial + linha) <= totalLinhas; linha++)
                    {
                        for (int col = 1; col <= totalColunas; col++)
                        {
                            novaPlanilha.Cells[linha + 2, col].Value = worksheet.Cells[linhaInicial + linha, col].Value;
                        }
                    }

                    novoPacote.Save();
                }

                indiceArquivo++;
            }
        }
    }

    static ServiceProvider ConfigureServices()
    {
        var serviceCollection = new ServiceCollection();

        // Configuração de logging
        serviceCollection.AddLogging(configure => configure.AddConsole());

        // Configuração de parâmetros
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        serviceCollection.AddSingleton<IConfiguration>(configuration);

        return serviceCollection.BuildServiceProvider();
    }

    static void RegistrarErro(Exception ex)
    {
        // Aqui você pode implementar um sistema de logging mais robusto
        Console.WriteLine($"Erro: {ex.Message}");
    }
}