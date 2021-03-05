using System;
using System.Data.OleDb;
using System.IO;
using System.Diagnostics;
using System.Collections;

namespace agrupamento2
{
    class Program
    {
        public static string cWay;
        public static string cName;
        public static string cName2;
        public static string cName3;
        public static float sum_M0 = 0;
        public static float sum_0_1 = 0;
        public static float sum_2_7 = 0;
        public static float sum_8_14 = 0;
        public static float sum_M15 = 0;
        public static float sum_total = 0;
        public static float sum_faixa1 = 0;

        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Clear();
            Console.WriteLine("CSIS Software 2021");
            Console.WriteLine("Agrupamento.exe versão 2.0");
            Console.WriteLine("----------------------------------------------");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(DateTime.Now + "-Inicio do processo.");
            Console.WriteLine(DateTime.Now + "-Excluindo arquivos auxiliares e arquivos desatualizados.");

            String arquivo1 = @"C:\agrupamento2\tmp\consolidado_valores.xlsx";
            String arquivo2 = @"C:\agrupamento2\tmp\consolidado_valores.xlsx";
            String arquivo3 = @"C:\agrupamento2\tmp\DBF_files.txt";
            String arquivo4 = @"C:\agrupamento2\tmp\consolidado_percentual.xlsx";
            String arquivo5 = @"C:\agrupamento2\data\sraghosp2.bak";
            String arquivo6 = @"C:\agrupamento2\data\consolidado_percentual.txt";
            String arquivo7 = @"C:\agrupamento2\data\consolidado_percentual.csv";

            File.Delete(arquivo1);
            File.Delete(arquivo2);
            File.Delete(arquivo3);
            File.Delete(arquivo4);
            File.Delete(arquivo5);
            File.Delete(arquivo6);
            File.Delete(arquivo7);

            System.Diagnostics.Process processCMD = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfoCMD = new System.Diagnostics.ProcessStartInfo();
            startInfoCMD.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfoCMD.FileName = "cmd.exe";
            startInfoCMD.Arguments = "/C del c:\\agrupamento2\\tab\\*.dbf";
            processCMD.StartInfo = startInfoCMD;
            processCMD.Start();
            startInfoCMD.Arguments = "/C del c:\\agrupamento2\\data\\*.dbf";
            processCMD.StartInfo = startInfoCMD;
            processCMD.Start();
            startInfoCMD.Arguments = "/C del c:\\agrupamento2\\tmp\\*.dbf";
            processCMD.StartInfo = startInfoCMD;
            processCMD.Start();

            Console.WriteLine(DateTime.Now + "-Rodando o script 'unzipping_srag.exe'.");
            // Roda o script "unzipping_srag.exe", que descompacta o(s) arquivo(s) de SRAG que estiver(em)
            // na subpasta "data".

            ProcessStartInfo startInfo_unzip = new ProcessStartInfo();
            startInfo_unzip.CreateNoWindow = false;
            startInfo_unzip.UseShellExecute = true;

            startInfo_unzip.FileName = "unzipping_srag.exe";
            startInfo_unzip.Arguments = "run";
            startInfo_unzip.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo_unzip.WorkingDirectory = @"c:\agrupamento2\bin\";

            var process_unzip = Process.Start(startInfo_unzip);
            process_unzip.WaitForExit();

            if (process_unzip.ExitCode != 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do 'unzipping_srag.exe'.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Pressione qualquer tecla.");
                Console.Read();
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Fim da execução do script 'unzipping_srag.exe'.");
            }

            Console.WriteLine(DateTime.Now + "-Rodando o script 'merge_srag.exe'.");
            // Roda o script "merge_srag.exe", que funde os arquivos descompactados pelo script
            // "unzipping_srag" em um só arquivo.

            ProcessStartInfo startInfo_merge = new ProcessStartInfo();
            startInfo_merge.CreateNoWindow = false;
            startInfo_merge.UseShellExecute = true;

            startInfo_merge.FileName = "merge_srag.exe";
            startInfo_merge.Arguments = "run";
            startInfo_merge.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo_merge.WorkingDirectory = @"c:\agrupamento2\bin\";

            var process_merge = Process.Start(startInfo_merge);
            process_merge.WaitForExit();

            if (process_merge.ExitCode != 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do 'merge_srag.exe'.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Pressione qualquer tecla.");
                Console.Read();
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Fim da execução do script 'merge_srag.exe'.");
            }

            Console.WriteLine(DateTime.Now + "-Tentando abrir conexao com o arquivo sraghosp.dbf. ");
            // Estabelece a conexão com os arquivos dbf disponíveis.
            string strConn0 = String.Format("Provider=VFPOLEDB.1; Data Source=C:\\agrupamento2\\data; MultipleActiveResultSets=true;");

            OleDbConnection oConn0 = new OleDbConnection(strConn0);
            OleDbCommand oCmd0 = new OleDbCommand();

            {

                oCmd0.Connection = oConn0;
                try
                {
                    oCmd0.Connection.Open(); // Abre a conexão.
                }
                catch
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(DateTime.Now + "Erro! Não foi aberta conexão com o VFPOLEDB.");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Pressione qualquer tecla.");
                    Console.Read();
                    Environment.Exit(1);
                }

                Console.WriteLine(DateTime.Now + "-Conexao ok.");
                Console.WriteLine(DateTime.Now + "-Marcando registros cuja evolucao nao seja obito.");
                oCmd0.CommandText = "DELETE FROM c:\\agrupamento2\\data\\sraghosp.dbf WHERE sraghosp.evolucao <> '2'";
                oCmd0.ExecuteNonQuery();
                Console.WriteLine(DateTime.Now + "-Marcando registros cuja class.final nao seja confirmado p/Covid.");
                oCmd0.CommandText = "DELETE FROM c:\\agrupamento2\\data\\sraghosp.dbf WHERE sraghosp.classi_fin <> '5'";
                oCmd0.ExecuteNonQuery();
                oConn0.Close();
                oConn0.Dispose();
                oCmd0.Dispose();
            }

            Console.WriteLine(DateTime.Now + "-Excluindo registros que não satisfazem os criterios evolucao/confirmacao.");
            // Estabelece a conexão com os arquivos dbf disponíveis.
            string strConn0a = String.Format("Provider=VFPOLEDB.1; Data Source=C:\\agrupamento2\\data; MultipleActiveResultSets=true;");

            OleDbConnection oConn0a = new OleDbConnection(strConn0a);
            OleDbCommand oCmd0a = new OleDbCommand();

            {
                oCmd0a.Connection = oConn0a;
                try
                {
                    oCmd0a.Connection.Open(); // Abre a conexão.
                }
                catch
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(DateTime.Now + "Erro! Não foi aberta conexão com o VFPOLEDB.");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Pressione qualquer tecla.");
                    Console.Read();
                    Environment.Exit(1);
                }

                Console.WriteLine(DateTime.Now + "-Conexao ok.");
                oCmd0a.CommandText = "PACK c:\\agrupamento2\\data\\sraghosp.dbf";
                oCmd0a.ExecuteNonQuery();

                oConn0a.Close();
                oConn0a.Dispose();
                oCmd0a.Dispose();
            }

            Console.WriteLine(DateTime.Now + "-Rodando o script 'conv_dt_evoluc.exe'.");
            // Roda o script "conv_dt_evoluc.exe", que altera o tipo do campo "dt_evoluca", originalmente um campo do tipo caracter após a exportação no
            // SIVEP Gripe, para um campo do tipo data. Após a finalização, é gerado o campo "dt_evoluc2".

            ProcessStartInfo startInfo_zero = new ProcessStartInfo();
            startInfo_zero.CreateNoWindow = false;
            startInfo_zero.UseShellExecute = true;

            startInfo_zero.FileName = "conv_dt_evoluc.exe";
            startInfo_zero.Arguments = "run";
            startInfo_zero.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo_zero.WorkingDirectory = @"c:\agrupamento2\bin\";

            var process_zero = Process.Start(startInfo_zero);
            process_zero.WaitForExit();

            if (process_zero.ExitCode != 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do 'conv_dt_evoluc.exe'.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Pressione qualquer tecla.");
                Console.Read();
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Fim da execução do script 'conv_dt_evoluc.exe'.");
            }
                                                                        
            Console.WriteLine(DateTime.Now + "-Tentando abrir conexao com o arquivo sraghosp2.dbf.");
            // Estabelece a conexão com os arquivos dbf disponíveis.
            string strConn = String.Format("Provider=VFPOLEDB.1; Data Source=C:\\agrupamento2\\data; MultipleActiveResultSets=true;");

            OleDbConnection oConn = new OleDbConnection(strConn);
            OleDbCommand oCmd = new OleDbCommand();

            {
                oCmd.Connection = oConn;
                try
                {
                    oCmd.Connection.Open(); // Abre a conexão.
                }
                catch
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(DateTime.Now + "Erro! Não foi aberta conexão com o VFPOLEDB.");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Pressione qualquer tecla.");
                    Console.Read();
                    Environment.Exit(1);
                }

                Console.WriteLine(DateTime.Now + "-Conexao ok.");
                Console.WriteLine(DateTime.Now + "-Adicionando campo 'intervalo' no arquivo 'sraghosp2.dbf'.");
                oCmd.CommandText = "ALTER TABLE c:\\agrupamento2\\data\\sraghosp2.dbf ADD COLUMN intervalo n(5)";
                oCmd.ExecuteNonQuery();
                Console.WriteLine(DateTime.Now + "-Adicionando campo 'escala' no arquivo 'sraghosp2.dbf'.");
                oCmd.CommandText = "ALTER TABLE c:\\agrupamento2\\data\\sraghosp2.dbf ADD COLUMN escala C(35)";
                oCmd.ExecuteNonQuery();

                oConn.Close();
                oConn.Dispose();
                oCmd.Dispose();
            }

            Console.WriteLine(DateTime.Now + "-Rodando o script 'intervalo.exe'.");
            // Roda o arquivo "intervalo.exe", que calcula a diferença entre a data de digitação e a data de evolucao.
            // Registros em que não é possível realizar o cálculo, são excluídos pelo script.
            ProcessStartInfo startInfo_alfa = new ProcessStartInfo();
            startInfo_alfa.CreateNoWindow = false;
            startInfo_alfa.UseShellExecute = true;

            startInfo_alfa.FileName = "intervalo.exe";
            startInfo_alfa.Arguments = "run";
            startInfo_alfa.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo_alfa.WorkingDirectory = @"c:\agrupamento2\bin\";

            var process_alfa = Process.Start(startInfo_alfa);
            process_alfa.WaitForExit();

            if (process_alfa.ExitCode != 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do 'intervalo.exe'.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Pressione qualquer tecla.");
                Console.Read();
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Fim da execução do script 'intervalo.exe'.");
            }

            Console.WriteLine(DateTime.Now + "-Rodando o script 'separador.exe'.");
            // Roda o script "separador.exe", que segrega as semanas epidemiológicas do arquivo 'sraghosp2.dbf' em
            // diversos arquivos distintos.

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = true;

            startInfo.FileName = "separador.exe";
            startInfo.Arguments = "run";
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.WorkingDirectory = @"c:\agrupamento2\bin\";

            var process = Process.Start(startInfo);
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do 'separador.exe'.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Pressione qualquer tecla.");
                Console.Read();
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Fim da execução do script 'separador.exe'.");
            }
            
            // Estabelece a conexão com os arquivos dbf disponíveis.
            string strConn2 = String.Format("Provider=VFPOLEDB.1; Data Source=C:\\agrupamento2\\tmp; MultipleActiveResultSets=true;");

            OleDbConnection oConn2 = new OleDbConnection(strConn2);
            OleDbCommand oCmd2 = new OleDbCommand();

            {

                oCmd2.Connection = oConn2;
                try
                {
                    oCmd2.Connection.Open(); // Abre a conexão.
                }
                catch
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(DateTime.Now + "-Erro! Não foi aberta conexão com o VFPOLEDB.");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Pressione qualquer tecla.");
                    Console.Read();
                    Environment.Exit(1);
                }

                Console.WriteLine(DateTime.Now + "-Conexao ok.");
                Console.WriteLine(DateTime.Now + "-Criando novas tabelas...");
                oCmd2.CommandText = "CREATE TABLE escala (id_municip C(27), co_uni_not C(7), id_unidade C(60), escala C(35), qtd N(10))";
                oCmd2.ExecuteNonQuery();
                oCmd2.CommandText = "CREATE TABLE consolidado_valores (periodo C(20), oportuno F(7,2), d_2_e_7 F(7,2), d_8_e_14 F(7,2), d_15_mais F(7,2), total F(7,2))";
                oCmd2.ExecuteNonQuery();
                oCmd2.CommandText = "CREATE TABLE consolidado_percent (periodo C(20), oportuno C(5), d_2_e_7 C(5), d_8_e_14 C(5), d_15_mais C(5), total C(5))";
                oCmd2.ExecuteNonQuery();

                oConn2.Close();
                oConn2.Dispose();
                oCmd2.Dispose();

                if (File.Exists(@"c:\agrupamento2\tmp\escala.dbf"))
                {
                    if (File.Exists(@"c:\agrupamento2\tmp\consolidado_valores.dbf"))
                    {
                        if (File.Exists(@"c:\agrupamento2\tmp\consolidado_valores.dbf"))
                        {
                            Console.WriteLine(DateTime.Now + "-Tabelas cridas.");
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(DateTime.Now + "-Erro! Falha na criação das tabelas.");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Pressione qualquer tecla.");
                            Console.Read();
                            Environment.Exit(1);
                        }
                    }
                }
            }
            
            Console.WriteLine(DateTime.Now + "-Rodando o script 'classificador.exe'.");
            // Verifica um a um os arquivos DBF criados pelo 'separador.exe'.
            // Se possuirem registros invalidos, exclui esses registros.
            // Classifica na tabela 'file_manager.dbf' se os arquivos vao ser processados na proxima etapa ou nao.
            
            ProcessStartInfo startInfo_beta = new ProcessStartInfo();
            startInfo_beta.CreateNoWindow = false;
            startInfo_beta.UseShellExecute = true;

            startInfo_beta.FileName = "classificador.exe";
            startInfo_beta.Arguments = "run";
            startInfo_beta.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo_beta.WorkingDirectory = @"c:\agrupamento2\bin\";

            var process_beta = Process.Start(startInfo_beta);
            process_beta.WaitForExit();

            if (process_beta.ExitCode != 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do 'classificador.exe'.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Pressione qualquer tecla.");
                Console.Read();
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Fim da execução do script 'classificador.exe'.");
            }

            Console.WriteLine(DateTime.Now + "-Rodando o script '2txt.exe'.");
            // Usa o arquivo 'file_manager.dbf' para gerar um arquivo txt dos arquivos que serao processados pelo 'agrupamento2.exe'.
            // Os arquivos DBF que não tem registros ficam de fora da listagem no arquivo txt.  

            ProcessStartInfo startInfo_gama = new ProcessStartInfo();
            startInfo_gama.CreateNoWindow = false;
            startInfo_gama.UseShellExecute = true;

            startInfo_gama.FileName = "2txt.exe";
            startInfo_gama.Arguments = "run";
            startInfo_gama.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo_gama.WorkingDirectory = @"c:\agrupamento2\bin\";

            var process_gama = Process.Start(startInfo_gama);
            process_gama.WaitForExit();

            if (process_gama.ExitCode != 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do '2txt.exe'.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Pressione qualquer tecla.");
                Console.Read();
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Fim da execução do script '2txt.exe'.");
            }

            Console.WriteLine(DateTime.Now + "-Armazena o nome dos arquivos gerados em um array'.");
            ArrayList arlist = new ArrayList();
            StreamReader x;
               string Caminho = "C:\\agrupamento2\\tmp\\DBF_files.txt";
               x = File.OpenText(Caminho);

                while (x.EndOfStream != true)                                             
                {
                    string linha = x.ReadLine();
                // Procura por linhas que não contém um arquivo dbf.
                int tipoDBF = linha.IndexOf(".dbf");
                if (tipoDBF > 0)
                {
                    arlist.Add(linha);
                }

                }
                x.Close();

            Console.WriteLine(DateTime.Now + "-Processando os arquivos gerados, um a um.");
            for (int i = 0; i < arlist.Count; i++)
            {
                Program.sum_M0 = 0;
                Program.sum_0_1 = 0;
                Program.sum_2_7 = 0;
                Program.sum_8_14 = 0;
                Program.sum_M15 = 0;
                Program.sum_faixa1 = 0;
                Program.sum_total = 0;

                Program.cWay = "c:\\agrupamento2\\tab\\" + arlist[i];
                Program.cName = "" + arlist[i];
                Program.cName2 = cName.Replace(".dbf", "");
                int CharSem = cName2.IndexOf("sem");
                Program.cName3 = Program.cName2.Substring(CharSem, 9);
                Console.WriteLine(DateTime.Now + "-Processando." + (Program.cName));

                // Estabelece a conexão com os arquivos dbf disponíveis.
                string strConn3a = String.Format("Provider=VFPOLEDB.1; Data Source=C:\\agrupamento2\\tmp; MultipleActiveResultSets=true;");

                OleDbConnection oConn3a = new OleDbConnection(strConn3a);
                OleDbCommand oCmd3a = new OleDbCommand();

                {

                    oCmd3a.Connection = oConn3a;
                    try
                    {
                        oCmd3a.Connection.Open(); // Abre a conexão.
                    }
                    catch
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(DateTime.Now + "-Erro! Não foi aberta conexão com o VFPOLEDB.");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Pressione qualquer tecla.");
                        Console.Read();
                        Environment.Exit(1);
                    }

                    Console.WriteLine(DateTime.Now + "-Criando tabelas auxiliares...");
                    Console.WriteLine(DateTime.Now + "-scale_M0.dbf...");
                    Console.WriteLine(cName);
                    Console.WriteLine(cWay);
                    Console.WriteLine(cName2);
                    Console.WriteLine("----------------------------------");
                    //oCmd3a.CommandText = "SELECT sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala, COUNT (sraghosp2.id_unidade) as qtd GROUP BY sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala where sraghosp2.intervalo <0 FROM c:\\agrupamento2\\data\\sraghosp2.dbf INTO TABLE scale_M0.dbf";
                    oCmd3a.CommandText = String.Format("SELECT {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala, COUNT ({0}.id_unidade) as qtd GROUP BY {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala where {0}.intervalo <0 FROM {1} INTO TABLE scale_M0.dbf", cName2, cWay);
                    oCmd3a.ExecuteNonQuery();
                    Console.WriteLine(DateTime.Now + "-OK");
                    Console.WriteLine(DateTime.Now + "-scale_0_1.dbf...");
                    //oCmd3a.CommandText = "SELECT sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala, COUNT (sraghosp2.id_unidade) as qtd GROUP BY sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala where (sraghosp2.intervalo >=0  .and. sraghosp2.intervalo <= 1) FROM c:\\agrupamento2\\data\\sraghosp2.dbf INTO TABLE scale_0_1.dbf";
                    oCmd3a.CommandText = String.Format("SELECT {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala, COUNT ({0}.id_unidade) as qtd GROUP BY {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala where ({0}.intervalo >=0  .and. {0}.intervalo <= 1) FROM {1} INTO TABLE scale_0_1.dbf", cName2, cWay);
                    oCmd3a.ExecuteNonQuery();
                    Console.WriteLine(DateTime.Now + "-OK");
                    Console.WriteLine(DateTime.Now + "-scale_2_7.dbf...");
                    //oCmd3a.CommandText = "SELECT sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala, COUNT (sraghosp2.id_unidade) as qtd GROUP BY sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala where (sraghosp2.intervalo >=2 .and. sraghosp2.intervalo <= 7) FROM c:\\agrupamento2\\data\\sraghosp2.dbf INTO TABLE scale_2_7.dbf";
                    oCmd3a.CommandText = String.Format("SELECT {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala, COUNT ({0}.id_unidade) as qtd GROUP BY {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala where ({0}.intervalo >=2  .and. {0}.intervalo <= 7) FROM {1} INTO TABLE scale_2_7.dbf", cName2, cWay);
                    oCmd3a.ExecuteNonQuery();
                    Console.WriteLine(DateTime.Now + "-OK");
                    Console.WriteLine(DateTime.Now + "-scale_8_14.dbf...");
                    //oCmd3a.CommandText = "SELECT sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala, COUNT (sraghosp2.id_unidade) as qtd GROUP BY sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala where (sraghosp2.intervalo >= 8 .and. sraghosp2.intervalo <= 14) FROM c:\\agrupamento2\\data\\sraghosp2.dbf INTO TABLE scale_8_14.dbf";
                    oCmd3a.CommandText = String.Format("SELECT {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala, COUNT ({0}.id_unidade) as qtd GROUP BY {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala where ({0}.intervalo >=8  .and. {0}.intervalo <= 14) FROM {1} INTO TABLE scale_8_14.dbf", cName2, cWay);
                    oCmd3a.ExecuteNonQuery();
                    Console.WriteLine(DateTime.Now + "-OK");
                    Console.WriteLine(DateTime.Now + "-scale_M15.dbf...");
                    //oCmd3a.CommandText = "SELECT sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala, COUNT (sraghosp2.id_unidade) as qtd GROUP BY sraghosp2.id_municip,sraghosp2.co_uni_not,sraghosp2.id_unidade,sraghosp2.escala where sraghosp2.intervalo >= 15 FROM c:\\agrupamento2\\data\\sraghosp2.dbf INTO TABLE scale_M15.dbf";
                    oCmd3a.CommandText = String.Format("SELECT {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala, COUNT ({0}.id_unidade) as qtd GROUP BY {0}.id_municip,{0}.co_uni_not,{0}.id_unidade,{0}.escala where {0}.intervalo >=15 FROM {1} INTO TABLE scale_M15.dbf", cName2, cWay);
                    oCmd3a.ExecuteNonQuery();
                    Console.WriteLine(DateTime.Now + "-OK");

                    Console.WriteLine(DateTime.Now + "-Verificando se arquivos foram efetivamente criados...");
                    if (!File.Exists(@"c:\agrupamento2\tmp\scale_M0.dbf"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(DateTime.Now + "-Erro! Arquivo 'scale_M0.dbf' não foi encontrado.");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Pressione qualquer tecla.");
                        Console.Read();
                        Environment.Exit(1);
                    }

                    if (!File.Exists(@"c:\agrupamento2\tmp\scale_0_1.dbf"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(DateTime.Now + "-Erro! Arquivo 'scale_0_1.dbf' não foi encontrado.");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Pressione qualquer tecla.");
                        Console.Read();
                        Environment.Exit(1);
                    }

                    if (!File.Exists(@"c:\agrupamento2\tmp\scale_2_7.dbf"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(DateTime.Now + "-Erro! Arquivo 'scale_2_7.dbf' não foi encontrado.");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Pressione qualquer tecla.");
                        Console.Read();
                        Environment.Exit(1);
                    }

                    if (!File.Exists(@"c:\agrupamento2\tmp\scale_8_14.dbf"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(DateTime.Now + "-Erro! Arquivo 'scale_8_14.dbf' não foi encontrado.");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Pressione qualquer tecla.");
                        Console.Read();
                        Environment.Exit(1);
                    }

                    if (!File.Exists(@"c:\agrupamento2\tmp\scale_M15.dbf"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(DateTime.Now + "-Erro! Arquivo 'scale_M15.dbf' não foi encontrado.");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Pressione qualquer tecla.");
                        Console.Read();
                        Environment.Exit(1);
                    }

                    oConn3a.Close();
                    oConn3a.Dispose();
                    oCmd3a.Dispose();
                }

                Console.WriteLine(DateTime.Now + "-Rodando o script 'filler.exe'.");
                // Roda o arquivo "filler.exe", que preenche o campo 'escala' do arquivo DBF criado.
                // Agrega os dados gerados no arquivo 'escala.dbf'.
                ProcessStartInfo startInfo2 = new ProcessStartInfo();
                startInfo2.CreateNoWindow = false;
                startInfo2.UseShellExecute = true;

                startInfo2.FileName = "filler.exe";
                startInfo2.Arguments = "run";
                startInfo2.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo2.WorkingDirectory = @"c:\agrupamento2\bin\";

                var process2 = Process.Start(startInfo2);
                process2.WaitForExit();

                if (process2.ExitCode != 0)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do 'filler.exe'.");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Pressione qualquer tecla.");
                    Console.Read();
                    Environment.Exit(1);
                }
                else
                {
                    Console.WriteLine(DateTime.Now + "-Fim da execução do script 'filler.exe'.");
                }

                // Estabelece a conexão com os arquivos dbf disponíveis.
                string strConn4 = String.Format("Provider=VFPOLEDB.1; Data Source=C:\\agrupamento2\\tmp; MultipleActiveResultSets=true;");

                OleDbConnection oConn4 = new OleDbConnection(strConn4);
                OleDbCommand oCmd4 = new OleDbCommand();

                {

                    oCmd4.Connection = oConn4;
                    try
                    {
                        oCmd4.Connection.Open(); // Abre a conexão.
                    }
                    catch
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(DateTime.Now + "Erro! Não foi aberta conexão com o VFPOLEDB.");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Pressione qualquer tecla.");
                        Console.Read();
                        Environment.Exit(1);
                    }

                    Console.WriteLine(DateTime.Now + "-Conexao ok.");
                    Console.WriteLine(DateTime.Now + "-Somando as notificacoes de cada agrupamento...");
                    try
                    {
                        oCmd4.CommandText = "SELECT SUM(scale_m0.qtd) FROM scale_m0";
                        Program.sum_M0 = Convert.ToInt32(oCmd4.ExecuteScalar());
                    }
                    catch
                    {
                        Program.sum_M0 = 0;
                    }
                    try
                    {
                        oCmd4.CommandText = "SELECT SUM(scale_0_1.qtd) FROM scale_0_1";
                        Program.sum_0_1 = Convert.ToInt32(oCmd4.ExecuteScalar());
                    }
                    catch
                    {
                        Program.sum_0_1 = 0;
                    }
                    try
                    {
                        oCmd4.CommandText = "SELECT SUM(scale_2_7.qtd) FROM scale_2_7";
                        Program.sum_2_7 = Convert.ToInt32(oCmd4.ExecuteScalar());
                    }
                    catch
                    {
                        Program.sum_2_7 = 0;
                    }
                    try
                    {
                        oCmd4.CommandText = "SELECT SUM(scale_8_14.qtd) FROM scale_8_14";
                        Program.sum_8_14 = Convert.ToInt32(oCmd4.ExecuteScalar());
                    }
                    catch
                    {
                        Program.sum_8_14 = 0;
                    }
                    try
                    {
                        oCmd4.CommandText = "SELECT SUM(scale_m15.qtd) FROM scale_m15";
                        Program.sum_M15 = Convert.ToInt32(oCmd4.ExecuteScalar());
                    }
                    catch
                    {
                        Program.sum_M15 = 0;
                    }

                    Program.sum_total = (Program.sum_M0 + Program.sum_0_1 + Program.sum_2_7 + Program.sum_8_14 + sum_M15);
                    Program.sum_faixa1 = (Program.sum_M0 + Program.sum_0_1);

                    Console.WriteLine(DateTime.Now + "-Notificações totalizadas.");
                    Console.WriteLine("-------------------------------------------------------");
                    Console.WriteLine("Digitacao oportuna:" + (Program.sum_faixa1));
                    Console.WriteLine("Digitacao entre 2 e 7 dias:" + Program.sum_2_7);
                    Console.WriteLine("Digitacao entre 8 e 14 dias:" + Program.sum_8_14);
                    Console.WriteLine("Digitacao 15 dias ou mais:" + Program.sum_M15);
                    Console.WriteLine("total de casos:" + Program.sum_total);

                    Console.WriteLine(DateTime.Now + "-Gravando o resultado em 'consolidado_valores.dbf'...");

                    oCmd4.CommandText = string.Format("INSERT INTO c:\\agrupamento2\\tmp\\consolidado_valores.dbf VALUES ('{0}',{1},{2},{3},{4},{5})", Program.cName3, Program.sum_faixa1, Program.sum_2_7, Program.sum_8_14, Program.sum_M15, Program.sum_total);
                    oCmd4.ExecuteNonQuery();

                    Console.WriteLine(DateTime.Now + "-Calculando os valores percentuais dos casos. ");

                    float percent_faixa1 = 0;
                    float percent_2_7 = 0;
                    float percent_8_14 = 0;
                    float percent_M15 = 0;
                    float percent_total = 0;

                    percent_faixa1 = (float)Math.Round(((Program.sum_faixa1) * 100) / (Program.sum_total), 2);
                    percent_2_7 = (float)Math.Round(((Program.sum_2_7) * 100) / (Program.sum_total), 2);
                    percent_8_14 = (float)Math.Round(((Program.sum_8_14) * 100) / (Program.sum_total), 2);
                    percent_M15 = (float)Math.Round(((Program.sum_M15) * 100) / (Program.sum_total), 2);
                    percent_total = (float)Math.Round(percent_faixa1 + percent_2_7 + percent_8_14 + percent_M15, 0);

                    string str_percent_faixa1 = string.Format("{0:N2}", percent_faixa1);
                    string str_percent_2_7 = string.Format("{0:N2}", percent_2_7);
                    string str_percent_8_14 = string.Format("{0:N2}", percent_8_14);
                    string str_percent_M15 = string.Format("{0:N2}", percent_M15);
                    string str_percent_total = string.Format("{0}", percent_total);

                    Console.WriteLine("-------------------------------------------------------");
                    Console.WriteLine("Digitacao oportuna (em %):" + (percent_faixa1));
                    Console.WriteLine("Digitacao entre 2 e 7 dias (em %):" + (percent_2_7));
                    Console.WriteLine("Digitacao entre 8 e 14 dias (em %):" + (percent_8_14));
                    Console.WriteLine("Digitacao 15 dias ou mais (em %):" + (percent_M15));

                    Console.WriteLine(DateTime.Now + "-Gravando o resultado em 'consolidado_percent.dbf'...");

                    oCmd4.CommandText = string.Format("INSERT INTO c:\\agrupamento2\\tmp\\consolidado_percent.dbf VALUES ('{0}','{1}','{2}','{3}','{4}','{5}')", Program.cName3, str_percent_faixa1, str_percent_2_7, str_percent_8_14, str_percent_M15, str_percent_total);
                    oCmd4.ExecuteNonQuery();

                    oConn4.Close();
                    oConn4.Dispose();
                    oCmd4.Dispose();

                }

            }

            Console.WriteLine(DateTime.Now + "-Rodando o script converter.exe...");
            ProcessStartInfo startInfoConvert = new ProcessStartInfo();
            startInfoConvert.CreateNoWindow = false;
            startInfoConvert.UseShellExecute = true;
            startInfoConvert.FileName = "converter.exe";
            startInfoConvert.Arguments = "run";
            startInfoConvert.WindowStyle = ProcessWindowStyle.Hidden;
            startInfoConvert.WorkingDirectory = @"c:\agrupamento2\bin\";
            var process40 = Process.Start(startInfoConvert);
            process40.WaitForExit();
            var process50 = process40.ExitCode.ToString();
            var process60 = startInfoConvert.FileName;
            var process70 = startInfoConvert.Arguments;

            if (process_gama.ExitCode != 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(DateTime.Now + "-Erro! Falha na execução do 'converter.exe'.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Pressione qualquer tecla.");
                Console.Read();
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Fim da execução do script 'converter.exe'.");
            }

            System.Diagnostics.Process processCMD2 = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfoCMD2 = new System.Diagnostics.ProcessStartInfo();
            startInfoCMD2.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfoCMD2.FileName = "cmd.exe";
            startInfoCMD2.Arguments = "/C cd\\&&cd c:\\agrupamento2\\bin&&sed  -e s/,/;/g c:\\agrupamento2\\tmp\\consolidado_percentual.txt > c:\\agrupamento2\\tmp\\consolidado_percentual.csv";
            processCMD2.StartInfo = startInfoCMD2;
            processCMD2.Start();

            Console.WriteLine(DateTime.Now + "-Convertendo arquivo com os resultados para xlsx...");
            ProcessStartInfo startInfo1 = new ProcessStartInfo();
            startInfo1.CreateNoWindow = false;
            startInfo1.UseShellExecute = true;
            startInfo1.FileName = "soffice.exe";
            startInfo1.Arguments = "--headless --convert-to xlsx --outdir \"c:\\agrupamento2\\tmp\" \"c:\\agrupamento2\\tmp\\consolidado_valores.dbf\"";
            startInfo1.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo1.WorkingDirectory = @"c:\Program Files\LibreOffice\program";
            var process4 = Process.Start(startInfo1);
            process4.WaitForExit();
            var process5 = process4.ExitCode.ToString();
            var process6 = startInfo1.FileName;
            var process7 = startInfo1.Arguments;

            if (!File.Exists(@"C:\agrupamento2\tmp\consolidado_valores.xlsx"))
            {
                Console.WriteLine(DateTime.Now + "-Erro! Falha ao criar arquivo 'consolidado_valores.xlsx'.");
            }
            else
            {
                Console.WriteLine(DateTime.Now + "-Arquivo 'consolidado_valores.xlsx' criado com sucesso.");
            }

        }
    }
}