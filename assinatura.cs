using Milord.Dados.Comercial;
using Milord.Dados.Sistema;
using Milord.Negocio.Comercial;
using Milord.Negocio.Sistema;
using Milord.Negocio.Seguranca;
using Milord.Dados.Seguranca;
using System;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.ServiceProcess;
using System.Timers;
using System.Collections.Generic;
using Milord.Negocio.Tecnico;
using Milord.Dados.Tecnico;
using System.Collections;
using System.Web;
using milord.negocio.Tecnico;
using System.Text;
using System.IO;
using MilordAssinatura;

namespace AssinaturaService
{
    public partial class assinatura : ServiceBase
    {
        #region varioaveis globais
        private int eventId = 1;
        private Timer timer;
        private EventLog eventLog1;

        public assinatura()
        {
            InitializeComponent();

            this.AutoLog = false;
            this.CanStop = true;
            this.CanPauseAndContinue = true;

            string MySource = "Milord",
            MyLog = "Application";

            eventLog1 = new EventLog();
            timer = new Timer();

            if (!EventLog.SourceExists(MySource))
                EventLog.CreateEventSource(MySource, MyLog);

            eventLog1.Source = MySource;
            eventLog1.Log = MyLog;
            

            //ServicoAssinatura(ServicosAssinatura.Calibracao, false);


            /*
             bool ambienteteste = false;
             if (ConfigurationManager.AppSettings["ambiente"] != null)
             {
                 if (ConfigurationManager.AppSettings["ambiente"] == "T")
                     ambienteteste = true;
             }


             if (ambienteteste)
             {
                 ServicoAssinatura(ServicosAssinatura.QualificacaoTermica, false);
                 ServicoAssinatura(ServicosAssinatura.Calibracao, false);
                 ServicoAssinatura(ServicosAssinatura.Manutencao, false);
                 ServicoAssinatura(ServicosAssinatura.Ensaio, false);


                 ServicoAssinatura(ServicosAssinatura.Calibracao, true);
                 ServicoAssinatura(ServicosAssinatura.Manutencao, true);
                 ServicoAssinatura(ServicosAssinatura.QualificacaoTermica, true);
                 ServicoAssinatura(ServicosAssinatura.Ensaio, true);
             }*/

        }
        #endregion

        #region Manipuladores de tempo

        protected override void OnStart(string[] args)
        {
            try
            {
                base.OnStart(args);

                logeventos.RegistraEventoLog(eventLog1, "Serviço Iniciado", EventLogEntryType.Information);

                // Configuração do timer
                timer.Interval = 60000 * 2;
                timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
                timer.Start();
            }
            catch (Exception ex)
            {
                // Registre a exceção 
                logeventos.RegistraEventoLog(eventLog1, $"Erro ao iniciar serviço: {ex.Message}", EventLogEntryType.Error);

                this.Stop();
            }
        }

        protected override void OnStop()
        {
            try
            {
                logeventos.RegistraEventoLog(eventLog1, "Serviço Interrompido", EventLogEntryType.Information);

                // Certifique-se de parar o timer de maneira adequada
                if (timer != null)
                {
                    timer.Stop();
                    timer.Dispose();
                }
            }
            catch (Exception ex)
            {
                // Registre a exceção 
                logeventos.RegistraEventoLog(eventLog1, $"Erro no processo de parada do serviço: {ex.Message}", EventLogEntryType.Error);
            }
        }

        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            try
            {
                // Pausa o timer enquanto executa a lógica do serviço
                timer.Stop();

                // Lógica do serviço
                ServicoAssinatura(ServicosAssinatura.Calibracao, false);
                ServicoAssinatura(ServicosAssinatura.Manutencao, false);
                ServicoAssinatura(ServicosAssinatura.QualificacaoTermica, false);
                ServicoAssinatura(ServicosAssinatura.Ensaio, false);

                ServicoAssinatura(ServicosAssinatura.Calibracao, true);
                ServicoAssinatura(ServicosAssinatura.Manutencao, true);
                ServicoAssinatura(ServicosAssinatura.QualificacaoTermica, true);
                ServicoAssinatura(ServicosAssinatura.Ensaio, true);
            }
            catch (Exception ex)
            {
                // Registra a exceção ou toma alguma ação adequada
                logeventos.RegistraEventoLog(eventLog1, $"Erro na execução do timer: {ex.Message}", EventLogEntryType.Error);
            }
            finally
            {
                // Reinicia o timer independentemente de sucesso ou falha
                timer.Start();
            }
        }

        #endregion

        private void ServicoAssinatura(ServicosAssinatura servicoselecionado, bool erro)
        {
            // Criação da instância do serviço específico
            IServicoAssinaturaDAO dao = CriarDAO(servicoselecionado);
            DataTable dtgeral = dao.RetornaDadosConsultaAgrupado(erro);

            // Lógica de processamento dos dados
            listaservico listaservico = ProcessarDados(dtgeral, servicoselecionado);

            // Se houver itens a serem assinados, realiza a assinatura
            if (listaservico.agrupado.Count > 0)
            {
                listaservico.eventLog1 = eventLog1;
                listaservico.ServicoSelecionado = servicoselecionado;

                // Realiza a assinatura
                servicos servicos = new servicos();
                servicos.ServicoAssinar = assinar.Instancia(servicoselecionado);
                servicos.ServicoAssinar.Assinar(listaservico, erro);
            }
        }

        private IServicoAssinaturaDAO CriarDAO(ServicosAssinatura servicoselecionado)
        {
            // Cria a instância do DAO correspondente ao tipo de serviço
            switch (servicoselecionado)
            {
                case ServicosAssinatura.Calibracao:
                    return new calibracaoassinaturaDAO();
                case ServicosAssinatura.Manutencao:
                    return new manutencaoassinaturaDAO();
                case ServicosAssinatura.QualificacaoTermica:
                    return new qualificacaotermicaassinaturaDAO();
                case ServicosAssinatura.Ensaio:
                    return new ensaioassinaturaDAO();
                default:
                    throw AssinarException.TipoServicoNaoReconhecido();
            }
        }

        private listaservico ProcessarDados(DataTable dtgeral, ServicosAssinatura servicoselecionado)
        {
            listaservico listaservico = new listaservico();

            foreach (DataRow row in dtgeral.Rows)
            {
                // Lógica comum para todas as opções de serviço
                MilordAssinatura.listaservico.servicoagrupado servicoagrupado = new MilordAssinatura.listaservico.servicoagrupado
                {
                    DataAssinatura = Convert.ToDateTime(row["dataassinatura"]),
                    CodigoGerente = Convert.ToInt32(row["CODIGOGERENTE"]),
                    HoraAssinatura = row["hora"].ToString()
                };

                string nomecampo = ObterNomeCampo(servicoselecionado);

                DataTable dt = ObterDadosConsulta(servicoselecionado, row, nomecampo);

                foreach (DataRow linha in dt.Rows)
                {
                    MilordAssinatura.listaservico.servicoagrupado.item item = new MilordAssinatura.listaservico.servicoagrupado.item
                    {
                        CodigoServico = Convert.ToInt32(linha[nomecampo])
                    };
                    servicoagrupado.items.Add(item);
                }

                listaservico.agrupado.Add(servicoagrupado);
            }

            return listaservico;
        }

        private string ObterNomeCampo(ServicosAssinatura servicoselecionado)
        {
            // Obtém o nome do campo correspondente ao tipo de serviço
            switch (servicoselecionado)
            {
                case ServicosAssinatura.Calibracao: return "CODIGOCALIBRACAO";
                case ServicosAssinatura.Manutencao: return "CODIGOMANUTENCAOPREVENTIVA";
                case ServicosAssinatura.QualificacaoTermica: return "CODIGOQUALIFICACAOTERMICA";
                case ServicosAssinatura.Ensaio: return "CODIGOENSAIO";
                default: throw AssinarException.TipoServicoNaoReconhecido();
            }
        }

        private DataTable ObterDadosConsulta(ServicosAssinatura servicoselecionado, DataRow row, string nomecampo)
        {
            // Obtém os dados da consulta correspondentes ao tipo de serviço
            switch (servicoselecionado)
            {
                case ServicosAssinatura.Calibracao:
                    calibracaoassinatura calibracaoassinatura = new calibracaoassinatura
                    {
                        Dataassinatura = Convert.ToDateTime(row["dataassinatura"]),
                        Hora = row["hora"].ToString(),
                        Codigogerente = Convert.ToInt32(row["CODIGOGERENTE"])
                    };
                    return new calibracaoassinaturaDAO().RetornaDadosConsulta(calibracaoassinatura);

                case ServicosAssinatura.Manutencao:
                    manutencaoassinatura manutencaoassinatura = new manutencaoassinatura
                    {
                        Dataassinatura = Convert.ToDateTime(row["dataassinatura"]),
                        Hora = row["hora"].ToString(),
                        Codigogerente = Convert.ToInt32(row["CODIGOGERENTE"])
                    };
                    return new manutencaoassinaturaDAO().RetornaDadosConsulta(manutencaoassinatura);

                case ServicosAssinatura.QualificacaoTermica:
                    qualificacaotermicaassinatura qualificacaoassinatura = new qualificacaotermicaassinatura
                    {
                        Dataassinatura = Convert.ToDateTime(row["dataassinatura"]),
                        Hora = row["hora"].ToString(),
                        Codigogerente = Convert.ToInt32(row["CODIGOGERENTE"])
                    };
                    return new qualificacaotermicaassinaturaDAO().RetornaDadosConsulta(qualificacaoassinatura);

                case ServicosAssinatura.Ensaio:
                    ensaioassinatura ensaioassinatura = new ensaioassinatura
                    {
                        Dataassinatura = Convert.ToDateTime(row["dataassinatura"]),
                        Hora = row["hora"].ToString(),
                        Codigogerente = Convert.ToInt32(row["CODIGOGERENTE"])
                    };
                    return new ensaioassinaturaDAO().RetornaDadosConsulta(ensaioassinatura);

                default:
                    throw AssinarException.TipoServicoNaoReconhecido();
            }
        }


    }
}
