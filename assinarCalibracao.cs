using milord.negocio.Tecnico;
using Milord.Dados.Comercial;
using Milord.Dados.Producao;
using Milord.Dados.Seguranca;
using Milord.Dados.Sistema;
using Milord.Dados.Tecnico;
using Milord.Negocio.Comercial;
using Milord.Negocio.Producao;
using Milord.Negocio.Seguranca;
using Milord.Negocio.Sistema;
using Milord.Negocio.Tecnico;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;

namespace MilordAssinatura
{
    internal sealed class assinarCalibracao : Iassinar
    {
        internal static Lazy<Iassinar> Instance { get; } = new Lazy<Iassinar>(() => new assinarCalibracao());


        public void Assinar(listaservico listaServico, bool erro)
        {
            parametrogeral parametrogeral = new parametrogeral();
            parametrogeralDAO parametrogeralDAO = new parametrogeralDAO();
            idiomaDAO idiomaDAO = new idiomaDAO();
            empresaidiomaDAO empresaidiomaDAO = new empresaidiomaDAO();

            parametrogeral = parametrogeralDAO.RetornaDadosConsultaid(parametrogeral);
            string diretorioServidor = parametrogeral.Diretoriocertificado;

            if (!System.IO.Directory.Exists(diretorioServidor))
            {
                logeventos.RegistraEventoLog(listaServico.eventLog1, "Diretório inexistente: " + diretorioServidor, EventLogEntryType.Error);
                return;
            }

            //Registros são agrupados conforme solicitação realizada pelo usuário
            foreach (var agrupado in listaServico.agrupado)
            {
                usuario usuario = new usuario();
                usuarioDAO usuarioDAO = new usuarioDAO();
                usuario.Codigousuario = agrupado.CodigoGerente;
                usuario = usuarioDAO.RetornaDadosConsultaid(usuario);
                List<int> ListaServicos = new List<int>();


                //Percorre cada serviço dentro do agrupamento realizado
                foreach (var item in agrupado.items)
                {

                    calibracao calibracao = new calibracao();
                    qualificacaotermica qualificacaotermica = new qualificacaotermica();
                    manutencaopreventiva manutencao = new manutencaopreventiva();
                    ensaio ensaio = new ensaio();

                    Int32? _codigoservicoselecionado = 0, _codigoempresa = 0;

                    if (listaServico.ServicoSelecionado == ServicosAssinatura.Calibracao)
                    {

                        calibracaoDAO calibracaoDAO = new calibracaoDAO();
                        calibracao.Codigocalibracao = item.CodigoServico;
                        calibracao = calibracaoDAO.RetornaDadosConsultaid(calibracao);
                        _codigoservicoselecionado = calibracao.Codigocalibracao;
                        _codigoempresa = calibracao.Instrumento.Empresa.Codigoempresa;
                    }
                    else
                     if (listaServico.ServicoSelecionado == ServicosAssinatura.Manutencao)
                    {

                        manutencaopreventivaDAO manutencaopreventivaDAO = new manutencaopreventivaDAO();
                        manutencao.Codigomanutencaopreventiva = item.CodigoServico;
                        manutencao = manutencaopreventivaDAO.RetornaDadosConsultaid(manutencao);
                        _codigoservicoselecionado = manutencao.Codigomanutencaopreventiva;
                        _codigoempresa = manutencao.Instrumento.Empresa.Codigoempresa;
                    }
                    else
                    if (listaServico.ServicoSelecionado == ServicosAssinatura.QualificacaoTermica)
                    {

                        qualificacaotermicaDAO qualificacaotermicaDAO = new qualificacaotermicaDAO();
                        qualificacaotermica.Codigoqualificacaotermica = item.CodigoServico;
                        qualificacaotermica = qualificacaotermicaDAO.RetornaDadosConsultaid(qualificacaotermica);
                        _codigoservicoselecionado = qualificacaotermica.Codigoqualificacaotermica;
                        _codigoempresa = qualificacaotermica.Instrumento.Empresa.Codigoempresa;
                    }
                    else
                    if (listaServico.ServicoSelecionado == ServicosAssinatura.Ensaio)
                    {

                        ensaioDAO ensaioDAO = new ensaioDAO();
                        ensaio.Codigoensaio = item.CodigoServico;
                        ensaio = ensaioDAO.RetornaDadosConsultaid(ensaio);
                        _codigoservicoselecionado = ensaio.Codigoensaio;
                        _codigoempresa = ensaio.Instrumento.Empresa.Codigoempresa;
                    }

                    if (erro == false) //Relação de serviços que possuem o numero de tentativas inferior a 3
                    {
                        if (usuario.Certificadao != null && _codigoservicoselecionado > 0)
                        {
                            empresaidioma empresaIdioma = new empresaidioma();
                            empresaIdioma.Empresa.Codigoempresa = _codigoempresa;
                            DataTable dtIdioma = empresaidiomaDAO.RetornaDadosConsulta(empresaIdioma);

                            bool operacaoOK = true;

                            foreach (DataRow row in dtIdioma.Rows)
                            {
                                int idiomaCodigo = Convert.ToInt32(row["codigoidioma"].ToString());
                                idioma idiomaAux = new idioma();
                                idiomaAux.Codigoidioma = idiomaCodigo;
                                idiomaAux = idiomaDAO.RetornaDadosConsultaid(idiomaAux);

                                ArrayList lista = new ArrayList();
                                lista.Add(item.CodigoServico);

                                try
                                {

                                    if (listaServico.ServicoSelecionado == ServicosAssinatura.Calibracao)
                                    {
                                        certificadoDAO2 certificadoDAO2 = new certificadoDAO2();
                                        certificadoDAO2.exibirCalibracao(diretorioServidor, idiomaAux, false, lista, usuario, false);
                                        string mensagemErro = "";

                                        if (!certificadoDAO2.GerarAssinatura(diretorioServidor, idiomaAux, usuario, false, lista, ref mensagemErro))
                                        {
                                            operacaoOK = false;
                                            logeventos.RegistraEventoLog(listaServico.eventLog1, "Erro no processo de assinatura: " + mensagemErro, EventLogEntryType.Error);
                                        }

                                        if (!certificadoDAO2.GerarAssinatura(diretorioServidor, idiomaAux, usuario, true, lista, ref mensagemErro))
                                        {
                                            operacaoOK = false;
                                            logeventos.RegistraEventoLog(listaServico.eventLog1, "Erro no processo de assinatura: " + usuario.Nomeusuario, EventLogEntryType.Error);
                                        }

                                    }
                                    else if (listaServico.ServicoSelecionado == ServicosAssinatura.Manutencao)
                                    {
                                        certificadomanutencaoDAO certificadomanutencaoDAO = new certificadomanutencaoDAO();

                                        certificadomanutencaoDAO.exibirCalibracao(string.Empty, idiomaAux, false, lista, usuario, false);
                                        string mensagemerro = "";

                                        if (!certificadomanutencaoDAO.GerarAssinatura(idiomaAux, usuario, false, lista, ref mensagemerro))
                                        {
                                            operacaoOK = false;
                                            logeventos.RegistraEventoLog(listaServico.eventLog1, "Erro no processo de assinatura: " + usuario.Nomeusuario, EventLogEntryType.Error);
                                        }
                                    }
                                    else if (listaServico.ServicoSelecionado == ServicosAssinatura.QualificacaoTermica)
                                    {

                                        certificadoqualificacaotermicaDAO certificadoqualificacaotermicaDAO = new certificadoqualificacaotermicaDAO();
                                        string arquivoretorno = string.Empty;
                                        certificadoqualificacaotermicaDAO.exibirqualificacaotermica("", idiomaAux, false, lista, usuario, false,true);

                                        certificadoqualificacaotermicaDAO.gerararquivo(qualificacaotermica, false, true, idiomaAux, false, ref arquivoretorno, false);
                                        string mensagemerro = "";

                                        if (!certificadoqualificacaotermicaDAO.GerarAssinatura(idiomaAux, usuario, false, lista, ref mensagemerro))
                                        {
                                            operacaoOK = false;
                                            logeventos.RegistraEventoLog(listaServico.eventLog1, "Erro no processo de assinatura: " + usuario.Nomeusuario, EventLogEntryType.Error);
                                        }
                                    }
                                    else if (listaServico.ServicoSelecionado == ServicosAssinatura.Ensaio)
                                    {
                                        certificadoensaioDAO certificadoensaioDAO = new certificadoensaioDAO();

                                        certificadoensaioDAO.exibirCalibracao("", idiomaAux, false, lista, usuario, false);

                                        string arquivoretorno = string.Empty;
                                        certificadoensaioDAO.gerararquivo(ensaio, false, true, idiomaAux, false, ref arquivoretorno, false);
                                        string mensagemerro = "";


                                        if (string.IsNullOrEmpty(arquivoretorno))
                                        {
                                            logeventos.RegistraEventoLog(listaServico.eventLog1, "Erro no processo de assinatura: " + usuario.Nomeusuario, EventLogEntryType.Error);
                                        }

                                        if (!certificadoensaioDAO.GerarAssinatura("", idiomaAux, usuario, false, lista, ref mensagemerro))
                                        {
                                            operacaoOK = false;
                                            logeventos.RegistraEventoLog(listaServico.eventLog1, "Erro no processo de assinatura: " + usuario.Nomeusuario, EventLogEntryType.Error);
                                        }

                                    }

                                }
                                catch (Exception ex)
                                {
                                    operacaoOK = false;
                                    logeventos.RegistraEventoLog(listaServico.eventLog1, "Erro no processo de assinatura: " + ex.ToString(), EventLogEntryType.Error);

                                }
                            }

                            if (operacaoOK)
                            {

                                if (!ExcluirRegistroAssinatura(listaServico.ServicoSelecionado, usuario, item.CodigoServico))
                                {
                                    logeventos.RegistraEventoLog(listaServico.eventLog1, AssinarException.ExcluirRegistroAssinatura().ToString() + " - " + item.CodigoServico, EventLogEntryType.Error);
                                    throw AssinarException.ExcluirRegistroAssinatura();
                                }
                                else
                                {
                                    ListaServicos.Add((int)item.CodigoServico);
                                }
                            }
                            else
                            {
                                if (!RegistraErroAssinatura(listaServico.ServicoSelecionado, usuario, item.CodigoServico))
                                {
                                    logeventos.RegistraEventoLog(listaServico.eventLog1, AssinarException.ExcluirRegistroAssinatura().ToString() + " - " + item.CodigoServico, EventLogEntryType.Error);
                                    throw AssinarException.ExcluirRegistroAssinatura();
                                }
                            }
                        }
                        else if (_codigoservicoselecionado == 0)
                        {
                            logeventos.RegistraEventoLog(listaServico.eventLog1, "Serviço Inexistente: " + item.CodigoServico, EventLogEntryType.Error);
                            calibracaoassinatura auxiliar = new calibracaoassinatura();
                            auxiliar.Codigocalibracao = item.CodigoServico;

                            if (!RegistraErroAssinatura(listaServico.ServicoSelecionado, usuario, item.CodigoServico))
                            {
                                logeventos.RegistraEventoLog(listaServico.eventLog1, AssinarException.ExcluirRegistroAssinatura().ToString() + " - " + item.CodigoServico, EventLogEntryType.Error);
                                throw AssinarException.ExcluirRegistroAssinatura();
                            }
                        }
                        else
                        {
                            logeventos.RegistraEventoLog(listaServico.eventLog1, "Certificado Digital Inexistente: ", EventLogEntryType.Error);
                            calibracaoassinatura auxiliar = new calibracaoassinatura();
                            auxiliar.Codigocalibracao = item.CodigoServico;

                            if (!RegistraErroAssinatura(listaServico.ServicoSelecionado, usuario, item.CodigoServico))
                            {
                                logeventos.RegistraEventoLog(listaServico.eventLog1, AssinarException.ExcluirRegistroAssinatura().ToString() + " - " + item.CodigoServico, EventLogEntryType.Error);
                                throw AssinarException.ExcluirRegistroAssinatura();
                            }
                        }
                    }
                    else //gera lista de erros
                    {
                        if (!ExcluirRegistroAssinatura(listaServico.ServicoSelecionado, usuario, item.CodigoServico))
                        {
                            logeventos.RegistraEventoLog(listaServico.eventLog1, AssinarException.ExcluirRegistroAssinatura().ToString() + " - " + item.CodigoServico, EventLogEntryType.Error);
                            throw AssinarException.ExcluirRegistroAssinatura();
                        }
                        else
                        {
                            ListaServicos.Add((int)item.CodigoServico);
                        }

                    }

                }

                if (ListaServicos.Count > 0)
                {
                    EnvioEmail(listaServico.ServicoSelecionado, ListaServicos, usuario, listaServico.eventLog1, agrupado.DataAssinatura, agrupado.HoraAssinatura, erro, false);

                    ListaServicos.Clear();

                    foreach (var item in agrupado.items)
                    {
                        int? _codigoempresa = 0;
                        string _numeroinstrumento = "";

                        if (listaServico.ServicoSelecionado == ServicosAssinatura.Calibracao)
                        {
                            calibracao calibracao = new calibracao();
                            calibracaoDAO calibracaoDAO = new calibracaoDAO();
                            calibracao.Codigocalibracao = item.CodigoServico;
                            calibracao = calibracaoDAO.RetornaDadosConsultaid(calibracao);
                            _codigoempresa = calibracao.Instrumento.Empresa.Codigoempresa;
                            _numeroinstrumento = calibracao.Instrumento.Numeroinstrumento;
                        }
                        else if (listaServico.ServicoSelecionado == ServicosAssinatura.Manutencao)
                        {
                            manutencaopreventiva manutencao = new manutencaopreventiva();
                            manutencaopreventivaDAO manutencaopreventivaDAO = new manutencaopreventivaDAO();
                            manutencao.Codigomanutencaopreventiva = item.CodigoServico;
                            manutencao = manutencaopreventivaDAO.RetornaDadosConsultaid(manutencao);
                            _codigoempresa = manutencao.Instrumento.Empresa.Codigoempresa;
                            _numeroinstrumento = manutencao.Instrumento.Numeroinstrumento;
                        }
                        else if (listaServico.ServicoSelecionado == ServicosAssinatura.QualificacaoTermica)
                        {
                            qualificacaotermica qualificacaotermica = new qualificacaotermica();
                            qualificacaotermicaDAO qualificacaotermicaDAO = new qualificacaotermicaDAO();
                            qualificacaotermica.Codigoqualificacaotermica = item.CodigoServico;
                            qualificacaotermica = qualificacaotermicaDAO.RetornaDadosConsultaid(qualificacaotermica);
                            _codigoempresa = qualificacaotermica.Instrumento.Empresa.Codigoempresa;
                            _numeroinstrumento = qualificacaotermica.Instrumento.Numeroinstrumento;
                        }
                        else if (listaServico.ServicoSelecionado == ServicosAssinatura.Ensaio)
                        {
                            ensaio ensaio = new ensaio();
                            ensaioDAO ensaioDAO = new ensaioDAO();
                            ensaio.Codigoensaio = item.CodigoServico;
                            ensaio = ensaioDAO.RetornaDadosConsultaid(ensaio);
                            _codigoempresa = ensaio.Instrumento.Empresa.Codigoempresa;
                            _numeroinstrumento = ensaio.Instrumento.Numeroinstrumento;
                        }


                        /*  empresaeffort empresaeffort = new empresaeffort();
                          empresaeffortDAO empresaeffortDAO = new empresaeffortDAO();
                          empresaeffort.Empresa.Codigoempresa = _codigoempresa;
                          empresaeffort = empresaeffortDAO.RetornaDadosConsultaid(empresaeffort);

                         if (string.IsNullOrEmpty(empresaeffort.Codigoempresaeffortfilho))
                          {
                              ordemservicoeffort ordemservicoeffort = new ordemservicoeffort();
                              ordemservicoeffortDAO ordemservicoeffortDAO = new ordemservicoeffortDAO();
                              ordemservicoeffort.EmpresaEffort = empresaeffort;
                              ordemservicoeffort.Numeroinstrumento = _numeroinstrumento;

                              ordemservicoeffort = ordemservicoeffortDAO.RetornaDadosConsultaid(ordemservicoeffort);

                              if (ordemservicoeffort.Codigoordemservico > 0)
                              {
                                  //Sincronização
                              }
                              else
                              {
                                  ListaServicos.Add((int)item.CodigoServico);
                              }
                          }
                          */
                      
                    }

                    if (ListaServicos.Count > 0)
                    {
                        EnvioEmail(listaServico.ServicoSelecionado, ListaServicos, usuario, listaServico.eventLog1, agrupado.DataAssinatura, agrupado.HoraAssinatura, erro, true);
                    }
                }
            }
        }

        private bool ExcluirRegistroAssinatura(ServicosAssinatura servicoselecionado, usuario usuario, int _codigoservico)
        {
            if (servicoselecionado == ServicosAssinatura.Calibracao)
            {
                calibracaoassinatura calibracaoassinatura = new calibracaoassinatura();
                calibracaoassinatura.Codigocalibracao = _codigoservico;
                calibracaoassinaturaDAO calibracaoassinaturaDAO = new calibracaoassinaturaDAO();
                return calibracaoassinaturaDAO.excluir(usuario, calibracaoassinatura);
            }
            else if (servicoselecionado == ServicosAssinatura.Manutencao)
            {
                manutencaoassinatura manutencaoassinatura = new manutencaoassinatura();
                manutencaoassinatura.CodigoManutencaopreventiva = _codigoservico;
                manutencaoassinaturaDAO manutencaoassinaturaDAO = new manutencaoassinaturaDAO();
                return manutencaoassinaturaDAO.excluir(usuario, manutencaoassinatura);
            }
            else
                if (servicoselecionado == ServicosAssinatura.QualificacaoTermica)
            {
                qualificacaotermicaassinatura qualificacaotermicaassinatura = new qualificacaotermicaassinatura();
                qualificacaotermicaassinatura.CodigoQualificacaotermica = _codigoservico;
                qualificacaotermicaassinaturaDAO qualificacaotermicaassinaturaDAO = new qualificacaotermicaassinaturaDAO();
                return qualificacaotermicaassinaturaDAO.excluir(usuario, qualificacaotermicaassinatura);
            }
            else
            if (servicoselecionado == ServicosAssinatura.Ensaio)
            {
                ensaioassinatura ensaioassinatura = new ensaioassinatura();
                ensaioassinatura.Codigoensaio = _codigoservico;
                ensaioassinaturaDAO ensaioassinaturaDAO = new ensaioassinaturaDAO();
                return ensaioassinaturaDAO.excluir(usuario, ensaioassinatura);
            }
            else
            {
                return false;
            }
        }

        private bool RegistraErroAssinatura(ServicosAssinatura servicoselecionado, usuario usuario, int _codigoservico)
        {
            if (servicoselecionado == ServicosAssinatura.Calibracao)
            {
                calibracaoassinatura calibracaoassinatura = new calibracaoassinatura();
                calibracaoassinatura.Codigocalibracao = _codigoservico;

                calibracaoassinaturaDAO calibracaoassinaturaDAO = new calibracaoassinaturaDAO();
                return calibracaoassinaturaDAO.registraerroassinatura(usuario, calibracaoassinatura);

            }
            else if (servicoselecionado == ServicosAssinatura.Manutencao)
            {
                manutencaoassinatura manutencaoassinatura = new manutencaoassinatura();
                manutencaoassinatura.CodigoManutencaopreventiva = _codigoservico;

                manutencaoassinaturaDAO manutencaoassinaturaDAO = new manutencaoassinaturaDAO();
                return manutencaoassinaturaDAO.registraerroassinatura(usuario, manutencaoassinatura);
            }
            else
                if (servicoselecionado == ServicosAssinatura.QualificacaoTermica)
            {
                qualificacaotermicaassinatura qualificacaotermicaassinatura = new qualificacaotermicaassinatura();
                qualificacaotermicaassinatura.CodigoQualificacaotermica = _codigoservico;

                qualificacaotermicaassinaturaDAO qualificacaotermicaassinaturaDAO = new qualificacaotermicaassinaturaDAO();
                return qualificacaotermicaassinaturaDAO.registraerroassinatura(usuario, qualificacaotermicaassinatura);
            }
            else
            if (servicoselecionado == ServicosAssinatura.Ensaio)
            {
                ensaioassinatura ensaioassinatura = new ensaioassinatura();
                ensaioassinatura.Codigoensaio = _codigoservico;

                ensaioassinaturaDAO ensaioassinaturaDAO = new ensaioassinaturaDAO();
                return ensaioassinaturaDAO.registraerroassinatura(usuario, ensaioassinatura);
            }
            else
            {
                return false;
            }
        }

        public bool EnvioEmail(ServicosAssinatura servicesign, List<int> ListaServico, usuario usuario, EventLog eventLog1, DateTime dataasinatura, string horaassinatura, bool listaerro, bool effort)
        {

            //Teste - AJustar linha que ta fixo Calibracao
            try
            {
                if (ListaServico.Count == 0 || string.IsNullOrEmpty(usuario.Email))
                {
                    return false; // Não há necessidade de enviar e-mail
                }

                var parametrogeralDAO = new parametrogeralDAO();
                parametrogeral parametrogeral = new parametrogeral();
                parametrogeral = parametrogeralDAO.RetornaDadosConsultaid(parametrogeral);

                var emailDAO = new emailDAO();
                var email = emailDAO.RetornaDadosConsultaid(new email { Codigoemail = parametrogeral.CodigoEmailSistema });

                var envioEmailDAO = new envioemailDAO();

                string _assunto;


                if (listaerro)
                {
                    _assunto = "Milord MX – Inspeções de serviço não realizadas";
                }
                else if (effort)
                {
                    _assunto = "Milord MX – Inspeções de serviço instrumentos sem O.S Effort";
                }
                else
                {
                    _assunto = "Milord MX – Resumo da Inspeção de Serviços";
                }

                bool ambienteteste = false;
                if (ConfigurationManager.AppSettings["ambiente"] != null)
                {
                    if (ConfigurationManager.AppSettings["ambiente"] == "T")
                    {
                        ambienteteste = true;
                    }
                }

                string _emailpara = string.Empty;

                //   if (ambienteteste)
                //      _emailpara = "leandronleal@yahoo.com.br";
                // else
                _emailpara = usuario.Email;


                var envioEmail = new envioemail
                {
                    Contaemail = email,
                    Emailpara = _emailpara,
                    Assunto = _assunto,
                    Mensagem = emailformatado.FormatarMensagem(servicesign, ListaServico, usuario, dataasinatura, horaassinatura, listaerro)
                };

                envioEmailDAO.enviarMensagem(envioEmail, true);

                logeventos.RegistraEventoLog(eventLog1, "Envio de Email: " + usuario.Email, EventLogEntryType.Information);

                return true; // Envio de e-mail bem-sucedido???? funcao nao retorna true




            }
            catch (Exception ex)
            {
                logeventos.RegistraEventoLog(eventLog1, "Erro no processo de envio de email: " + ex.Message, EventLogEntryType.Error);
                return false; // Envio de e-mail falhou
            }
        }




    }
}
