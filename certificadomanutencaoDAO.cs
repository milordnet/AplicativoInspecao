//using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.CrystalReports.Engine;
using iTextSharpSign;
using Milord.Dados.Comercial;
using Milord.Dados.Producao;
using Milord.Dados.Sistema;
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
    public class certificadomanutencaoDAO : conexao
    {

        public bool GerarAssinatura(idioma idioma, usuario usuario, bool ficha, ArrayList listamanutencaopreventiva, ref string mensagemerro)
        {
            manutencaopreventivaDAO manutencaopreventivaDAO = new manutencaopreventivaDAO();
            DataTable dt = manutencaopreventivaDAO.retornarcalibracoes(listamanutencaopreventiva);
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
                    manutencaopreventiva manutencaopreventiva = new manutencaopreventiva();
                    manutencaopreventiva.Codigomanutencaopreventiva = Convert.ToInt32(dt.Rows[ct]["CODIGOmanutencaopreventiva"].ToString());

                    manutencaopreventiva.Instrumento.Empresa.Codigoempresa = Convert.ToInt32(dt.Rows[ct]["CODIGOEMPRESA"].ToString());
                    //manutencaopreventiva.Codigocredenciamento = dt.Rows[ct]["codigocredenciamento"].ToString();

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

                        path = parametrogeral.Diretoriocertificado + "\\certificados\\" + manutencaopreventiva.Instrumento.Empresa.Codigoempresa.ToString() + diretorioidioma;

                        if (!ficha)
                        {
                            caminho = path + "M" + manutencaopreventiva.Codigomanutencaopreventiva.ToString() + ".pdf";
                        }
                        else
                        {
                            caminho = path + "FE" + manutencaopreventiva.Codigomanutencaopreventiva.ToString() + ".pdf";
                        }


                        assinar(parametrogeral.Diretoriocertificado, myCert, caminho, manutencaopreventiva, usuario, idioma, ficha);
                    }
                    catch (Exception er)
                    {
                        throw new Exception("Ocorreu um erro no processo de assinatura.[7]  " + er.Message.ToString());
                    }
                }
            }
            return retornook;
        }

        public void assinar(string caminhoservidor, Cert myCert, String caminhoarquivo, manutencaopreventiva manutencaopreventiva, usuario usuario, idioma idioma, bool ficha)
        {
            try
            {

                String arquivodestino = String.Empty;
                String localizacao = String.Empty;


                Int32 revisao = 0;
                suplementomanutencaoDAO suplementomanutencaoDAO = new suplementomanutencaoDAO();
                suplementomanutencao suplementomanutencao = new suplementomanutencao();
                suplementomanutencao.Manutencaopreventiva = manutencaopreventiva;

                revisao = suplementomanutencaoDAO.UltimaRevisao(suplementomanutencao);

                parametrogeral parametrogeral = new parametrogeral();
                parametrogeralDAO parametrogeralDAO = new parametrogeralDAO();
                parametrogeral = parametrogeralDAO.RetornaDadosConsultaid(parametrogeral);
                string diretorioidioma = "";
                string path = parametrogeral.Diretoriocertificado + "\\certificados\\" + manutencaopreventiva.Instrumento.Empresa.Codigoempresa.ToString() + "\\";

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
                    arquivodestino = path + "M" + manutencaopreventiva.Codigomanutencaopreventiva.ToString() + "-" + revisao.ToString() + "-sign.pdf";
                }
                else
                {
                    arquivodestino = path + "FM" + manutencaopreventiva.Codigomanutencaopreventiva.ToString() + "-" + revisao.ToString() + "-sign.pdf";
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
                manutencaopreventiva.Filial = filialDAO.RetornaDadosConsultaid(manutencaopreventiva.Filial);


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
                        manutencaopreventiva calaux = new manutencaopreventiva();
                        manutencaopreventivaDAO manutencaopreventivaDAO = new manutencaopreventivaDAO();
                        calaux.Codigomanutencaopreventiva = manutencaopreventiva.Codigomanutencaopreventiva;
                        usuario Log = new usuario();
                        Log.Codigousuario = usuario.Codigousuario;
                        calaux.Statusinspecao = "PENDENTE";
                        manutencaopreventivaDAO.Inspecao(Log, calaux);
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
                    calaux.Codigomanutencaopreventiva = manutencaopreventiva.Codigomanutencaopreventiva;
                    usuario Log = new usuario();
                    Log.Codigousuario = usuario.Codigousuario;
                    calaux.Statusinspecao = "PENDENTE";
                    manutencaopreventivaDAO.Inspecao(Log, calaux);
                    throw new Exception("Atenção. Certificado digital está invalido e/ou a senha informada está incorreta.  " + er.Message);
                }
            }

        }

        public bool exibirCalibracao(string caminhoservidor, idioma idioma, bool conferencia, ArrayList calibracoes, usuario Log, bool ambienteteste)
        {
            bool erro = false;

            manutencaopreventivaDAO manutencaopreventivaDAO = new manutencaopreventivaDAO();
            DataTable dt = manutencaopreventivaDAO.retornarcalibracoes(calibracoes);

            if (dt.Rows.Count > 0)
            {
                for (int ct = 0; ct < dt.Rows.Count; ct++)
                {
                    try
                    {
                        //Atualiza o status da inspecao.
                        if (conferencia == false)
                        {
                            manutencaopreventiva calaux = new manutencaopreventiva();
                            calaux.Codigomanutencaopreventiva = Convert.ToInt32(dt.Rows[ct]["CODIGOmanutencaopreventiva"].ToString());
                            calaux.Gerente.Codigousuario = Log.Codigousuario;

                            calaux.Statusinspecao = "APROVADO";
                            manutencaopreventivaDAO.Inspecao(Log, calaux);
                        }

                        manutencaopreventiva manutencaopreventiva = new manutencaopreventiva();
                        manutencaopreventiva.Codigomanutencaopreventiva = Convert.ToInt32(dt.Rows[ct]["codigomanutencaopreventiva"].ToString());
                        manutencaopreventiva.Instrumento.Empresa.Codigoempresa = Convert.ToInt32(dt.Rows[ct]["CODIGOEMPRESA"].ToString());
                        //manutencaopreventiva.Codigocredenciamento = dt.Rows[ct]["codigocredenciamento"].ToString();

                        if (DateTime.TryParse(dt.Rows[0]["Dataverificacao"].ToString(), out DateTime _dataafericao))
                        {
                            manutencaopreventiva.Dataverificacao = _dataafericao;
                        }
                        string arquivogerado = string.Empty;

                        gerararquivo(manutencaopreventiva, conferencia, true, idioma, false, ref arquivogerado, ambienteteste);

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

        public bool gerararquivo(manutencaopreventiva manutencaopreventivaaux, bool conferencia, bool exibirimagemfundo, idioma idioma, bool ficha, ref string arquivodestino, bool ambienteteste)
        {
            //Muda globalization
            try
            {
                System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-US");
                System.Threading.Thread.CurrentThread.CurrentCulture = ci;
                System.Threading.Thread.CurrentThread.CurrentUICulture = ci;
                //bool printhistory = false;

                int codigomanutencaopreventiva = 0, revisao = 0;

                parametrogeral parametrogeral = new parametrogeral();
                parametrogeralDAO parametrogeralDAO = new parametrogeralDAO();
                parametrogeral = parametrogeralDAO.RetornaDadosConsultaid(parametrogeral);
                codigomanutencaopreventiva = (int)manutencaopreventivaaux.Codigomanutencaopreventiva;


                if (codigomanutencaopreventiva > 0)
                {
                    manutencaopreventiva manutencaopreventiva = new manutencaopreventiva();
                    manutencaopreventiva manut = new manutencaopreventiva();
                    manutencaopreventivaDAO manutencaopreventivaDAO = new manutencaopreventivaDAO();
                    idiomaDAO idiomaDAO = new idiomaDAO();

                    manutencaopreventiva.Codigomanutencaopreventiva = codigomanutencaopreventiva;
                    manut = manutencaopreventivaDAO.RetornaDadosConsultaid(manutencaopreventiva);

                    manutencaopreventiva.Filial.Codigofilial = manut.Filial.Codigofilial;
                    manutencaopreventiva.Instrumento.Empresa.Codigoempresa = manut.Instrumento.Empresa.Codigoempresa;

                    idioma = idiomaDAO.RetornaDadosConsultaid(idioma);

                    suplementomanutencaoDAO suplementomanutencaoDAO = new suplementomanutencaoDAO();
                    suplementomanutencao suplementomanutencao = new suplementomanutencao();
                    suplementomanutencao.Manutencaopreventiva = manutencaopreventiva;

                    if (revisao == 0)
                        revisao = suplementomanutencaoDAO.UltimaRevisao(suplementomanutencao);
                    else
                        revisao = revisao - 1;

                    string diretorio = parametrogeral.Diretoriocertificado;

                    // certificadomanutencaopreventivaDAO certificadomanutencaopreventivaDAO = new certificadomanutencaopreventivaDAO();

                    DataSet ds = new DataSet();

                    try
                    {
                        ds.Tables.Add(manutencaopreventivaDAO.RetornaItemManutencao(manutencaopreventiva, idioma));
                        ds.Tables[0].TableName = "ITEMMANUTENCAOPREVENTIVA";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir Impressão" + ex);

                    }

                    try
                    {
                        ds.Tables.Add(manutencaopreventivaDAO.RetornaManutencaoPreventiva(manutencaopreventiva, conferencia, idioma));
                        ds.Tables[1].TableName = "MANUTENCAOPREVENTIVA";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir Impressão" + ex);

                    }

                    try
                    {
                        ds.Tables.Add(manutencaopreventivaDAO.RetornaProcedimento(manutencaopreventiva));
                        ds.Tables[2].TableName = "procedimento";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagemfundo]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(manutencaopreventivaDAO.RetornaQuestionario(manutencaopreventiva, idioma));
                        ds.Tables[3].TableName = "QUESTIONARIOMANUTENCAOPREVENTIVA";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir Impressão" + ex);

                    }



                    manutencaopreventiva = manutencaopreventivaDAO.RetornaDadosConsultaid(manutencaopreventiva);

                    try
                    {
                        ds.Tables.Add(manutencaopreventivaDAO.retornaImagemFundo(manutencaopreventiva, exibirimagemfundo, ambienteteste, idioma));
                        ds.Tables[4].TableName = "imagemfundo";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagemfundo]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(manutencaopreventivaDAO.RetornaImagem(manutencaopreventiva));
                        ds.Tables[5].TableName = "imagem";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagem]" + ex);
                    }

                    try
                    {
                        ds.Tables.Add(manutencaopreventivaDAO.RetornaPadrao(manutencaopreventiva, idioma));
                        ds.Tables[6].TableName = "padrao";
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Ocorre um erro na tentativa exibir o certificado [imagem]" + ex);
                    }



                    ReportDocument rptcertificadocalibracao;
                    rptcertificadocalibracao = new ReportDocument();
                    rptcertificadocalibracao = ReportFactory.GetReport(rptcertificadocalibracao.GetType());

                    string pastaexportar = parametrogeral.DiretorioExportar;


                    string arquivorpt = string.Empty;

                    if (ambienteteste)
                        arquivorpt = parametrogeral.Diretoriorpttecnicoteste + "rptmanutencaopreventiva.rpt";
                    else
                        arquivorpt = parametrogeral.Diretoriorpttecnico + "rptmanutencaopreventiva.rpt";

                    if (System.IO.File.Exists(arquivorpt))
                    {
                        rptcertificadocalibracao.Load(arquivorpt, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy); //04-08-2021

                        traducaocertificado traducaocertificadocali = new traducaocertificado();
                        traducaocertificadoDAO traducaocertificadoDAO = new traducaocertificadoDAO();
                        traducaocertificadocali.Idioma.Codigoidioma = idioma.Codigoidioma;
                        traducaocertificadocali.Campocertificado.Tipodocumento = "CALIBRACAO";

                        DataTable dtcali = traducaocertificadoDAO.RetornaDadosConsulta(traducaocertificadocali);

                        for (int ct = 0; ct < dtcali.Rows.Count; ct++)
                            libfuncoes.TraduzirCampo(rptcertificadocalibracao, dtcali.Rows[ct]["nomecampo"].ToString(), dtcali.Rows[ct]["descricao"].ToString());

                        traducaocertificado traducaocertificado = new traducaocertificado();
                        traducaocertificado.Idioma.Codigoidioma = idioma.Codigoidioma;
                        traducaocertificado.Campocertificado.Tipodocumento = "MANUTENCAO";

                        DataTable dt = traducaocertificadoDAO.RetornaDadosConsulta(traducaocertificado);

                        for (int ct = 0; ct < dt.Rows.Count; ct++)
                            libfuncoes.TraduzirCampo(rptcertificadocalibracao, dt.Rows[ct]["nomecampo"].ToString(), dt.Rows[ct]["descricao"].ToString());

                        string patharquivo = string.Empty;

                        if (!conferencia)
                        {
                            patharquivo = parametrogeral.Diretoriocertificado + "\\certificados\\" + manutencaopreventiva.Instrumento.Empresa.Codigoempresa.ToString() + "\\";

                            if (idioma.Codigoidioma == 0)
                            {
                                patharquivo += "sign\\";
                            }
                            else
                            {
                                patharquivo += "sign-" + idioma.Siglaidioma + "\\";
                            }

                            if (!Directory.Exists(patharquivo))
                            {
                                Directory.CreateDirectory(patharquivo);
                            }

                            if (!ficha)
                            {
                                patharquivo = patharquivo + "M" + manutencaopreventiva.Codigomanutencaopreventiva.ToString() + ".pdf";
                            }
                            else
                            {
                                patharquivo = patharquivo + "FM" + manutencaopreventiva.Codigomanutencaopreventiva.ToString() + ".pdf";
                            }
                        }
                        else
                        {
                            patharquivo = parametrogeral.DiretorioExportar;
                            patharquivo = patharquivo + Guid.NewGuid().ToString() + ".pdf";
                        }

                        @patharquivo = @patharquivo.Replace("/", "\\");

                        rptcertificadocalibracao.SetDataSource(ds);
                        rptcertificadocalibracao.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, patharquivo);

                        arquivodestino = patharquivo;

                        rptcertificadocalibracao.Close();
                        rptcertificadocalibracao.Dispose();
                        GC.Collect();


                        //Muda globalization
                        System.Globalization.CultureInfo ci2 = new System.Globalization.CultureInfo("pt-BR");
                        System.Threading.Thread.CurrentThread.CurrentCulture = ci2;
                        System.Threading.Thread.CurrentThread.CurrentUICulture = ci2;

                        return System.IO.File.Exists(patharquivo);
                    }
                    else
                        return false;
                }
                else
                    return false;

            }
            catch (Exception ex)
            {
                //path = "";

                //Muda globalization
                System.Globalization.CultureInfo ci2 = new System.Globalization.CultureInfo("pt-BR");
                System.Threading.Thread.CurrentThread.CurrentCulture = ci2;
                System.Threading.Thread.CurrentThread.CurrentUICulture = ci2;

                throw new Exception("Ocorre um erro na tentativa gerar relatório " + ex.Message.ToString());
            }
        }





    }
}
