using CrystalDecisions.CrystalReports.Engine;
using iTextSharpSign;
using Milord.Dados.Comercial;
using Milord.Dados.Producao;
using Milord.Dados.Sistema;
using Milord.Enum;
using Milord.Funcoes;
using Milord.Negocio.Comercial;
using Milord.Negocio.Producao;
using Milord.Negocio.Seguranca;
using Milord.Negocio.Sistema;
using Milord.Negocio.Tecnico;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;

namespace Milord.Dados.Tecnico
{
    public class certificadoqualificacaotermicaDAO : conexao
    {

        public bool exibirqualificacaotermica(string caminhoservidor, idioma idioma, bool conferencia, ArrayList calibracoes, usuario Log, bool ambienteteste, bool projetocertificado)
        {
            bool erro = false;

            qualificacaotermicaDAO qualificacaotermicaDAO = new qualificacaotermicaDAO();
            DataTable dt = qualificacaotermicaDAO.retornarcalibracoes(calibracoes);

            if (dt.Rows.Count > 0)
            {
                for (int ct = 0; ct < dt.Rows.Count; ct++)
                {
                    try
                    {
                        //Atualiza o status da inspecao.
                        if (conferencia == false)
                        {
                            qualificacaotermica calaux = new qualificacaotermica();
                            calaux.Codigoqualificacaotermica = Convert.ToInt32(dt.Rows[ct]["CODIGOqualificacaotermica"].ToString());
                            calaux.Gerente.Codigousuario = Log.Codigousuario;

                            calaux.Statusinspecao = "APROVADO";
                            qualificacaotermicaDAO.Inspecao(Log, calaux);
                        }

                        qualificacaotermica qualificacaotermica = new qualificacaotermica();
                        qualificacaotermica.Codigoqualificacaotermica = Convert.ToInt32(dt.Rows[ct]["codigoqualificacaotermica"].ToString());
                        qualificacaotermica.Instrumento.Empresa.Codigoempresa = Convert.ToInt32(dt.Rows[ct]["CODIGOEMPRESA"].ToString());

                        if (DateTime.TryParse(dt.Rows[0]["Dataafericao"].ToString(), out DateTime _dataafericao))
                        {
                            qualificacaotermica.Dataafericao = _dataafericao;
                        }
                        string arquivogerado = string.Empty;

                        gerararquivo(qualificacaotermica, conferencia, true, idioma, false, ref arquivogerado, ambienteteste);

                    }
                    catch (Exception er)
                    {
                        erro = true;
                        throw new Exception("Ocorreu um erro no processo de assinatura.[6]  " + er.Message.ToString());

                    }
                }
            }
            return erro;


        }



        public string retornatraducaocampo(idioma idioma, string tipodocumento, string nomecampo)
        {
            traducaocertificado traducaocertificado = new traducaocertificado();
            traducaocertificado.Idioma.Codigoidioma = idioma.Codigoidioma;
            traducaocertificado.Campocertificado.Tipodocumento = tipodocumento;
            traducaocertificado.Campocertificado.Nomecampo = nomecampo;
            traducaocertificadoDAO traducaocertificadoDAO = new traducaocertificadoDAO();

            traducaocertificado = traducaocertificadoDAO.RetornaDadosConsultaid(traducaocertificado);

            return traducaocertificado.Descricao;
        }


        public bool gerararquivo(qualificacaotermica qualiaux, bool conferencia, bool exibirimagemfundo, idioma idioma, bool ficha, ref string arquivoretorno, bool ambienteteste)
        {

            try
            {
                //Muda globalization
                System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-US");
                System.Threading.Thread.CurrentThread.CurrentCulture = ci;
                System.Threading.Thread.CurrentThread.CurrentUICulture = ci;


                bool printhistory = false;
                int codigoqualificacao = (int)qualiaux.Codigoqualificacaotermica, revisao = 0;

                parametrogeral parametrogeral = new parametrogeral();
                parametrogeralDAO parametrogeralDAO = new parametrogeralDAO();
                parametrogeral = parametrogeralDAO.RetornaDadosConsultaid(parametrogeral);

                CabecalhoQualificacao CabecalhoQualificacao = new CabecalhoQualificacao();

                CabecalhoQualificacao.TituloRelatorio = retornatraducaocampo(idioma, "QUALIFICACAO", "txtnumerorelatorio");
                CabecalhoQualificacao.TituloDataRelatorio = retornatraducaocampo(idioma, "QUALIFICACAO", "txtdatacalibracao");
                CabecalhoQualificacao.TituloDataValidade = retornatraducaocampo(idioma, "CALIBRACAO", "txtdatavalidade");
                CabecalhoQualificacao.TituloOrdemServico = retornatraducaocampo(idioma, "CALIBRACAO", "txtos");


                if (codigoqualificacao > 0)
                {
                    qualificacaotermica quali = new qualificacaotermica();
                    qualificacaotermicaDAO qualiDAO = new qualificacaotermicaDAO();

                    quali.Codigoqualificacaotermica = codigoqualificacao;
                    quali = qualiDAO.RetornaDadosConsultaid(quali);

                    idiomaDAO idiomaDAO = new idiomaDAO();
                    idioma = idiomaDAO.RetornaDadosConsultaid(idioma);

                    suplementoqualificacaotermicaDAO suplementoqualificacaotermicaDAO = new suplementoqualificacaotermicaDAO();
                    suplementoqualificacaotermica suplementoqualificacaotermica = new suplementoqualificacaotermica();
                    suplementoqualificacaotermica.Qualificacaotermica = quali;

                    if (revisao == 0)
                    {
                        revisao = suplementoqualificacaotermicaDAO.UltimaRevisao(suplementoqualificacaotermica);
                    }
                    else
                    {
                        revisao = revisao - 1;
                    }

                    string diretorio = parametrogeral.Diretoriocertificado;

                    qualificacaotermicaDAO qualificacaotermicaDAO = new qualificacaotermicaDAO();
                    qualificacaotermica qualificacaotermica = new qualificacaotermica();
                    qualificacaotermica.Codigoqualificacaotermica = codigoqualificacao;
                    qualificacaotermica = qualificacaotermicaDAO.RetornaDadosConsultaid(qualificacaotermica);

                    certificadoqualificacaotermicaDAO certificadoqualificacaotermicaDAO = new certificadoqualificacaotermicaDAO();



                    DataSet ds = new DataSet();

                    try
                    {
                        ds.Tables.Add(RetornaProcedimento(qualificacaotermica, revisao, printhistory, ""));
                        ds.Tables[0].TableName = "procedimento";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [procedimento]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(RetornaPadrao(qualificacaotermica, revisao, printhistory, ""));
                        ds.Tables[1].TableName = "padrao";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [padrão]" + ex);
                    }

                    try
                    {
                        titulogradecalibracao titulogradecalibracao = new titulogradecalibracao();
                        titulogradecalibracao.Tipo = "Q";
                        ds.Tables.Add(retornodadosrelatorio(qualificacaotermica, revisao, printhistory, false, "", idioma, titulogradecalibracao));
                        ds.Tables[2].TableName = "leitura";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [leitura]" + ex);
                    }
                    try
                    {
                        ds.Tables.Add(certificadoqualificacaotermicaDAO.retornadadoscabecalhorelatorio(qualificacaotermica, conferencia, revisao, printhistory, "", false, idioma, CabecalhoQualificacao));
                        ds.Tables[3].TableName = "cabecalho";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [cabeçalho]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(RetornaImagem(qualificacaotermica, revisao, printhistory, "", "I")); //Ilustracoes
                        ds.Tables[4].TableName = "imagem";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagem]" + ex);
                    }


                    try
                    {
                        ds.Tables.Add(certificadoqualificacaotermicaDAO.retornaImagemFundo(qualificacaotermica, true, "", ambienteteste, idioma)); //Alterar true 20-02-2022
                        ds.Tables[5].TableName = "imagemfundo";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagemfundo]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(certificadoqualificacaotermicaDAO.RetornaParametroCiclo(qualificacaotermica, idioma));
                        ds.Tables[6].TableName = "parametrociclos";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagemfundo]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(RetornaGrafico(qualificacaotermica, revisao, printhistory, "", false, idioma));
                        ds.Tables[7].TableName = "grafico";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagem]" + ex);
                    }


                    try
                    {
                        titulogradecalibracao titulogradecalibracao = new titulogradecalibracao();
                        titulogradecalibracao.Tipo = "Q";

                        ds.Tables.Add(certificadoqualificacaotermicaDAO.RetornaParametroLeitura(qualificacaotermica, revisao, printhistory, false, "", idioma, titulogradecalibracao));
                        ds.Tables[8].TableName = "parametroleitura";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagemfundo]" + ex);
                    }


                    try
                    {
                        ds.Tables.Add(RetornaImagem(qualificacaotermica, revisao, printhistory, "", "E")); //Ilustracoes
                        ds.Tables[9].TableName = "imageminstrumento";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagem]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(RetornaImagem(qualificacaotermica, revisao, printhistory, "", "S")); //Ilustracoes
                        ds.Tables[10].TableName = "imagemsensor";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagem]" + ex);
                    }


                    try
                    {
                        titulogradecalibracao titulogradecalibracao = new titulogradecalibracao();
                        titulogradecalibracao.Tipo = "Q";
                        ds.Tables.Add(certificadoqualificacaotermicaDAO.RetornaItemCalibracao(qualificacaotermica, revisao, printhistory, false, "", idioma, titulogradecalibracao));
                        ds.Tables[11].TableName = "itemqualificacao";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [leitura]" + ex);
                    }


                    try
                    {
                        ds.Tables.Add(RetornaImagem(qualificacaotermica, revisao, printhistory, "", "Q"));
                        ds.Tables[12].TableName = "imagemqualificacaobiologica";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagem]" + ex);
                    }


                    try
                    {
                        ds.Tables.Add(certificadoqualificacaotermicaDAO.RetornaQualificacaoBiologica(qualificacaotermica, idioma));
                        ds.Tables[13].TableName = "qualificacaobiologica";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [QualificacaoBiologica]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(certificadoqualificacaotermicaDAO.RetornaParametroQualificacaoBiologica(qualificacaotermica, idioma));
                        ds.Tables[14].TableName = "parametroqualificacaobiologica";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [QualificacaoBiologica]" + ex);
                    }


                    try
                    {
                        ds.Tables.Add(RetornaGrafico(qualificacaotermica, revisao, printhistory, "", true, idioma));
                        ds.Tables[15].TableName = "graficomedia";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [grafico media]" + ex);
                    }


                    try
                    {
                        ds.Tables.Add(RetornaGraficoLetalidade(qualificacaotermica, revisao, printhistory, ""));
                        ds.Tables[16].TableName = "graficoletalidade";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [grafico media]" + ex);
                    }

                    try
                    {
                        titulogradecalibracao titulogradecalibracao = new titulogradecalibracao();
                        titulogradecalibracao.Tipo = "Q";
                        ds.Tables.Add(RetornaLetalidade(qualificacaotermica, revisao, printhistory, "", idioma, titulogradecalibracao));
                        ds.Tables[17].TableName = "letalidade";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [Tabela Letalidade]" + ex);
                    }


                    try
                    {
                        titulogradecalibracao titulogradecalibracao = new titulogradecalibracao();
                        titulogradecalibracao.Tipo = "Q";

                        ds.Tables.Add(certificadoqualificacaotermicaDAO.RetornaParametrosStatus(qualificacaotermica, idioma));
                        ds.Tables[18].TableName = "STATUS";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [Tabela Letalidade]" + ex);
                    }


                    try
                    {
                        ds.Tables.Add(certificadoqualificacaotermicaDAO.RetornaParametroEnsaio(qualificacaotermica, idioma));
                        ds.Tables[19].TableName = "parametroensaio";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [parametroensaio]" + ex);
                    }


                    ReportDocument rptcertificadocalibracao;

                    rptcertificadocalibracao = new ReportDocument();
                    rptcertificadocalibracao = ReportFactory.GetReport(rptcertificadocalibracao.GetType());

                    //TODO: Avaliar estre trecho do código
                    string pastaexportar = parametrogeral.Diretoriocertificado + "\\exportar\\";

                    ReportDocument rptcapa;
                    rptcapa = new ReportDocument();
                    rptcapa = ReportFactory.GetReport(rptcapa.GetType());

                    DataTable dtcapa = new DataTable();
                    DataSet dscapa = new DataSet();
                    libfuncoes.AdicionarColuna(dtcapa, "CODIGOIDIOMA", typeof(int));

                    DataRow row = dtcapa.NewRow();
                    row["CODIGOIDIOMA"] = idioma.Codigoidioma;

                    dtcapa.Rows.Add(row);

                    dscapa.Tables.Add(dtcapa);
                    dscapa.Tables[0].TableName = "CAPA";

                    string strrptcapa = string.Empty;

                    if (ambienteteste)
                        strrptcapa = parametrogeral.Diretoriorpttecnicoteste + "\\rptcapa.rpt";
                    else
                        strrptcapa = parametrogeral.Diretoriorpttecnico + "\\rptcapa.rpt";

                    String pathcapa = pastaexportar;


                    if (System.IO.File.Exists(strrptcapa))
                    {
                        rptcapa.Load(strrptcapa, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy);

                        rptcapa.SetDataSource(dscapa);


                        pathcapa = pathcapa + Guid.NewGuid().ToString() + ".pdf";
                        rptcapa.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, pathcapa);

                        rptcapa.Close();
                        rptcapa.Dispose();
                        GC.Collect();

                    }


                    string patharquivocabecalho = pastaexportar + Guid.NewGuid().ToString() + ".pdf";


                    string arquivoorigem = string.Empty;//, arquivodestino = string.Empty;

                    if (ambienteteste)
                        arquivoorigem = parametrogeral.Diretoriorpttecnicoteste + "RptQualificacaoTermica.rpt";
                    else
                        arquivoorigem = parametrogeral.Diretoriorpttecnico + "RptQualificacaoTermica.rpt";



                    if (System.IO.File.Exists(arquivoorigem))
                    {
                        rptcertificadocalibracao.Load(arquivoorigem, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy); //04-08-2021

                        String patharquivo = pastaexportar;
                        patharquivo = patharquivo + Guid.NewGuid().ToString() + ".pdf";

                        String patharquivoimageminstrumento = pastaexportar;
                        patharquivoimageminstrumento = patharquivoimageminstrumento + Guid.NewGuid().ToString() + ".pdf";


                        traducaocertificado traducaocertificadocali = new traducaocertificado();
                        traducaocertificadoDAO traducaocertificadoDAO = new traducaocertificadoDAO();
                        traducaocertificadocali.Idioma.Codigoidioma = idioma.Codigoidioma;
                        traducaocertificadocali.Campocertificado.Tipodocumento = "CALIBRACAO";

                        DataTable dtcali = traducaocertificadoDAO.RetornaDadosConsulta(traducaocertificadocali);

                        for (int ct = 0; ct < dtcali.Rows.Count; ct++)
                        {
                            libfuncoes.TraduzirCampo(rptcertificadocalibracao, dtcali.Rows[ct]["nomecampo"].ToString(), dtcali.Rows[ct]["descricao"].ToString());
                        }

                        traducaocertificado traducaocertificado = new traducaocertificado();
                        traducaocertificado.Idioma.Codigoidioma = idioma.Codigoidioma;
                        traducaocertificado.Campocertificado.Tipodocumento = "QUALIFICACAO";
                        DataTable dt = traducaocertificadoDAO.RetornaDadosConsulta(traducaocertificado);

                        for (int ct = 0; ct < dt.Rows.Count; ct++)
                        {
                            libfuncoes.TraduzirCampo(rptcertificadocalibracao, dt.Rows[ct]["nomecampo"].ToString(), dt.Rows[ct]["descricao"].ToString());
                        }

                        rptcertificadocalibracao.SetDataSource(ds);




                        rptcertificadocalibracao.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, patharquivo);



                        rptcertificadocalibracao.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, patharquivoimageminstrumento);

                        string arquivotextobase = "", arquivosumario = "", arquivosinteseprocedimento = "";


                        if (qualificacaotermica.Textobase2 != null)
                            arquivotextobase = libfuncoes.ExportarRTFPDF(pastaexportar, qualificacaotermica.Textobase2);//, imagemcabecalho);

                        if (qualificacaotermica.Sumario2 != null)
                            arquivosumario = libfuncoes.ExportarRTFPDF(pastaexportar, qualificacaotermica.Sumario2); //, imagemcabecalho);

                        if (qualificacaotermica.Sinteseprocedimento2 != null)
                            arquivosinteseprocedimento = libfuncoes.ExportarRTFPDF(pastaexportar, qualificacaotermica.Sinteseprocedimento2);//, imagemcabecalho);


                        rptcertificadocalibracao.Close();
                        rptcertificadocalibracao.Dispose();
                        GC.Collect();


                        //Muda globalization
                        System.Globalization.CultureInfo ci2 = new System.Globalization.CultureInfo("pt-BR");
                        System.Threading.Thread.CurrentThread.CurrentCulture = ci2;
                        System.Threading.Thread.CurrentThread.CurrentUICulture = ci2;


                        string path = parametrogeral.Diretoriocertificado + "\\certificados\\" + qualificacaotermica.Instrumento.Empresa.Codigoempresa.ToString() + "\\";
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }

                        string patharquivofinal = string.Empty;

                        if (idioma.Codigoidioma == 0)
                        {
                            patharquivofinal = path + "\\sign\\";
                        }
                        else
                        {
                            patharquivofinal = path + "\\sign-" + idioma.Siglaidioma + "\\";
                        }


                        if (!Directory.Exists(patharquivofinal))
                        {
                            Directory.CreateDirectory(patharquivofinal);
                        }

                        if (conferencia)
                        {
                            patharquivofinal += Guid.NewGuid().ToString() + ".pdf";
                        }
                        else
                        {
                            patharquivofinal += "Q" + qualificacaotermica.Codigoqualificacaotermica.ToString() + ".pdf";
                        }


                        string _nomearquivo = RetornaImagemLogoTipo((Int32)qualificacaotermica.Filial.Codigofilial, idioma);

                        libfuncoes.CombineMultiplePDFsQualificacao(pathcapa, arquivosumario, arquivotextobase, arquivosinteseprocedimento, patharquivo, patharquivofinal, pastaexportar, CabecalhoQualificacao, parametrogeral.Diretoriorpttecnico, _nomearquivo);

                        libfuncoes.ExcluirArquivo(pathcapa);
                        libfuncoes.ExcluirArquivo(arquivosumario);
                        libfuncoes.ExcluirArquivo(arquivotextobase);
                        // libfuncoes.ExcluirArquivo(arquivosinteseprocedimento);
                        libfuncoes.ExcluirArquivo(arquivosinteseprocedimento);
                        libfuncoes.ExcluirArquivo(patharquivo);
                        libfuncoes.ExcluirArquivo(patharquivocabecalho);
                        arquivoretorno = patharquivofinal;

                        return System.IO.File.Exists(patharquivofinal);
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                //Muda globalization
                System.Globalization.CultureInfo ci2 = new System.Globalization.CultureInfo("pt-BR");
                System.Threading.Thread.CurrentThread.CurrentCulture = ci2;
                System.Threading.Thread.CurrentThread.CurrentUICulture = ci2;

                throw new Exception("Ocorre um erro na tentativa gerar relatório " + ex.Message.ToString());
            }
        }

        public bool GerarAssinatura(idioma idioma, usuario usuario, bool ficha, ArrayList listaqualificacao, ref string mensagemerro)
        {
            qualificacaotermicaDAO qualificacaotermicaDAO = new qualificacaotermicaDAO();
            DataTable dt = qualificacaotermicaDAO.retornarcalibracoes(listaqualificacao);
            bool retornook = true;

            parametrogeral parametrogeral = new parametrogeral();
            parametrogeralDAO parametrogeralDAO = new parametrogeralDAO();
            parametrogeral = parametrogeralDAO.RetornaDadosConsultaid(parametrogeral);

            string tsaUrl = string.Empty;

            if (!string.IsNullOrEmpty(parametrogeral.Urltimestamp))
            {
                tsaUrl = parametrogeral.Urltimestamp;
            }
            else
            {
                tsaUrl = "http://timestamp.comodoca.com/authenticod";
            }

            Cert myCert = null;


            if (usuario.Certificadao != null)
            {
                myCert = new Cert(usuario.Certificadao, usuario.Senhacertificado.Trim().ToString(), tsaUrl, "", "");
            }
            else
            {
                mensagemerro = "Atenção. Certificado digital está invalido e/ou a senha informada está incorreta. Acesse o cadastro de usuários e vincule novamente o certificado.";
                return false;
            }


            if (dt.Rows.Count > 0)
            {
                for (Int32 ct = 0; ct < dt.Rows.Count; ct++)
                {
                    qualificacaotermica qualificacaotermica = new qualificacaotermica();
                    qualificacaotermica.Codigoqualificacaotermica = Convert.ToInt32(dt.Rows[ct]["Codigoqualificacaotermica"].ToString());
                    qualificacaotermica.Instrumento.Empresa.Codigoempresa = Convert.ToInt32(dt.Rows[ct]["CODIGOEMPRESA"].ToString());
                    qualificacaotermica = qualificacaotermicaDAO.RetornaDadosConsultaid(qualificacaotermica);

                    try
                    {
                        string caminho = string.Empty;
                        string path = string.Empty;
                        string diretorioidioma = string.Empty;

                        if (idioma.Codigoidioma == 0)
                        {
                            diretorioidioma = "\\sign\\";
                        }
                        else
                        {
                            diretorioidioma = "\\sign-" + idioma.Siglaidioma + "\\";
                        }

                        path = parametrogeral.Diretoriocertificado + "\\certificados\\" + qualificacaotermica.Instrumento.Empresa.Codigoempresa.ToString() + diretorioidioma;

                        if (!ficha)
                        {
                            caminho = path + "Q" + qualificacaotermica.Codigoqualificacaotermica.ToString() + ".pdf";
                        }
                        else
                        {
                            caminho = path + "FQ" + qualificacaotermica.Codigoqualificacaotermica.ToString() + ".pdf";
                        }


                        assinar(parametrogeral.Diretoriocertificado, myCert, caminho, qualificacaotermica, usuario, idioma, ficha);
                    }
                    catch (Exception er)
                    {
                        throw new Exception("Ocorreu um erro no processo de assinatura.[7]  " + er.Message.ToString());
                    }
                }
            }
            return retornook;
        }

        private void assinar(string caminhoservidor, Cert myCert, String caminhoarquivo, qualificacaotermica qualificacaotermica, usuario usuario, idioma idioma, bool ficha)
        {
            try
            {

                String arquivodestino = String.Empty;
                String localizacao = String.Empty;
                string path = string.Empty;

                Int32 revisao = 0;
                suplementoqualificacaotermicaDAO suplementoqualificacaotermicaDAO = new suplementoqualificacaotermicaDAO();
                suplementoqualificacaotermica suplementoqualificacaotermica = new suplementoqualificacaotermica();
                suplementoqualificacaotermica.Qualificacaotermica = qualificacaotermica;

                revisao = suplementoqualificacaotermicaDAO.UltimaRevisao(suplementoqualificacaotermica);

                path = caminhoservidor + "\\certificados\\" + qualificacaotermica.Instrumento.Empresa.Codigoempresa.ToString() + "\\";

                string diretorioidioma = string.Empty;

                if (idioma.Codigoidioma == 0)
                {
                    diretorioidioma = "sign\\";
                }
                else
                {
                    diretorioidioma = "sign-" + idioma.Siglaidioma + "\\";
                }

                path += diretorioidioma;

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                if (!ficha)
                {
                    arquivodestino = path + "Q" + qualificacaotermica.Codigoqualificacaotermica.ToString() + "-" + revisao.ToString() + "-sign.pdf";
                }
                else
                {
                    arquivodestino = path + "FQ" + qualificacaotermica.Codigoqualificacaotermica.ToString() + "-" + revisao.ToString() + "-sign.pdf";
                }

                PDFEncryption pdfEnc = new PDFEncryption();
                pdfEnc.UserPwd = "";
                pdfEnc.OwnerPwd = "";
                pdfEnc.Encryption = true;
                pdfEnc.Permissions.Add(permissionType.Copy);
                pdfEnc.Permissions.Add(permissionType.DegradedPrinting);
                pdfEnc.Permissions.Add(permissionType.Printing);
                pdfEnc.Permissions.Add(permissionType.ScreenReaders);
                pdfEnc.Permissions.Add(permissionType.FillIn);


                PDFSigner pdfs = new PDFSigner(caminhoarquivo, arquivodestino, myCert);
                PDFSignatureAP sigAp = new PDFSignatureAP();

                filialDAO filialDAO = new filialDAO();
                qualificacaotermica.Filial = filialDAO.RetornaDadosConsultaid(qualificacaotermica.Filial);


                sigAp.Visible = true;
                sigAp.Multi = true;
                sigAp.Page = 1; //será utilizado sempre a ultima pagina


                if (!ficha)
                {
                    sigAp.CustomText = "Assinado Digitalmente por: " + usuario.Nomeusuario + "\n" + "Data: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                }



                pdfs.Sign(sigAp, false, pdfEnc);

                myCert = null;
                pdfs = null;
                pdfEnc = null;
                sigAp = null;


                if (System.IO.File.Exists(caminhoarquivo))
                {
                    System.IO.File.Delete(caminhoarquivo);
                }

                if (!ficha)
                {

                    if (!System.IO.File.Exists(arquivodestino))
                    {
                        //Atualiza o status da inspecao.
                        qualificacaotermica calaux = new qualificacaotermica();
                        qualificacaotermicaDAO qualificacaotermicaDAO = new qualificacaotermicaDAO();
                        calaux.Codigoqualificacaotermica = qualificacaotermica.Codigoqualificacaotermica;
                        usuario Log = new usuario();
                        Log.Codigousuario = usuario.Codigousuario;
                        calaux.Statusinspecao = "PENDENTE";
                        qualificacaotermicaDAO.Inspecao(Log, calaux);
                    }
                }

            }
            catch (Exception er)
            {
                //Atualiza o status da inspecao.
                if (!ficha)
                {
                    manutencaopreventiva calaux = new manutencaopreventiva();
                    manutencaopreventivaDAO manutencaopreventivaDAO = new manutencaopreventivaDAO();
                    calaux.Codigomanutencaopreventiva = qualificacaotermica.Codigoqualificacaotermica;
                    usuario Log = new usuario();
                    Log.Codigousuario = usuario.Codigousuario;
                    calaux.Statusinspecao = "PENDENTE";
                    manutencaopreventivaDAO.Inspecao(Log, calaux);
                    throw new Exception("Atenção. Certificado digital está invalido e/ou a senha informada está incorreta.  " + er.Message);
                }
            }

        }





    }
}
