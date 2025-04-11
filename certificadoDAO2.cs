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
    public class certificadoDAO2 : conexao
    {
        public bool exibirCalibracao(string caminhoservidor, idioma idioma, bool conferencia, ArrayList calibracoes, usuario Log, bool ambienteteste)
        {
            bool erro = false;

            calibracaoDAO calibracaoDAO = new calibracaoDAO();
            DataTable dt = calibracaoDAO.retornarcalibracoes(calibracoes);

            if (dt.Rows.Count > 0)
            {
                for (int ct = 0; ct < dt.Rows.Count; ct++)
                {
                    try
                    {
                        if (conferencia == false)
                        {
                            if (idioma.Codigoidioma == 0) //31-10-2022
                            {
                                calibracao calaux = new calibracao();
                                calaux.Codigocalibracao = Convert.ToInt32(dt.Rows[ct]["CODIGOCALIBRACAO"].ToString());
                                calaux.Gerente.Codigousuario = Log.Codigousuario;
                                calaux.Statusinspecao = "APROVADO";
                                calibracaoDAO.Inspecao(Log, calaux);
                            }
                        }

                        calibracao calibracao = new calibracao();
                        calibracao.Codigocalibracao = Convert.ToInt32(dt.Rows[ct]["codigocalibracao"].ToString());
                        calibracao.Instrumento.Empresa.Codigoempresa = Convert.ToInt32(dt.Rows[ct]["CODIGOEMPRESA"].ToString());
                        calibracao.Codigocredenciamento = dt.Rows[ct]["codigocredenciamento"].ToString();


                        if (DateTime.TryParse(dt.Rows[0]["dataafericao"].ToString(), out DateTime _dataafericao))
                        {
                            calibracao.Dataafericao = _dataafericao;
                        }

                        if (DateTime.TryParse(dt.Rows[0]["datainspecao"].ToString(), out DateTime _datainspecao))
                        {
                            calibracao.Datainspecao = _datainspecao;
                        }
                        string arquivoretorno = string.Empty;
                        gerararquivo(caminhoservidor, calibracao, conferencia, true, idioma, false, ambienteteste, ref arquivoretorno);

                        //Ficha
                       // gerararquivo(caminhoservidor, calibracao, conferencia, true, idioma, true, ambienteteste, ref arquivoretorno);
                     //   GerarEtiquetaNet(calibracao, idioma, true);
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


        public void assinar(string caminhoservidor, Cert myCert, String caminhoarquivo, calibracao calibracao, usuario usuario, idioma idioma, bool ficha)
        {
            try
            {

                String arquivodestino = String.Empty;
                String localizacao = String.Empty;
                string path = string.Empty;


                Int32 revisao = 0;

                /*suplementocertificadoDAO suplementocertificadoDAO = new suplementocertificadoDAO();
                suplementocertificado suplementocertificado = new suplementocertificado();
                suplementocertificado.Calibracao = calibracao;
                
                revisao = suplementocertificadoDAO.UltimaRevisao(suplementocertificado);
                */

                path = caminhoservidor + "\\certificados\\" + calibracao.Instrumento.Empresa.Codigoempresa.ToString() + "\\";

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
                    arquivodestino = path + calibracao.Codigocalibracao.ToString() + "-" + revisao.ToString() + "-sign.pdf";
                }
                else
                {
                    arquivodestino = path + "F" + calibracao.Codigocalibracao.ToString() + "-" + revisao.ToString() + "-sign.pdf";
                }

                // throw new Exception(arquivodestino);


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

                //filialDAO filialDAO = new filialDAO();
                //calibracao.Filial = filialDAO.RetornaDadosConsultaid(calibracao.Filial);

                sigAp.Visible = true;
                sigAp.Multi = true;
                sigAp.Page = 1; //será utilizado sempre a ultima pagina

                if (!ficha)
                {
                    sigAp.CustomText = "Assinado Digitalmente por: " + usuario.Nomeusuario + "\n" + "Data: " + calibracao.Datainspecao.ToString("dd/MM/yyyy HH:mm");/*DateTime.Now.ToString("dd/MM/yyyy HH:MM");*/
                }
                else
                {
                    sigAp.CustomText = "Assinado Digitalmente por: " + usuario.Nomeusuario + "\n" + "Data: " + calibracao.Datainspecao.ToString("dd/MM/yyyy HH:mm");/*DateTime.Now.ToString("dd/MM/yyyy HH:MM");*/
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
                        /*calibracao calaux = new calibracao();
                        calibracaoDAO calibracaoDAO = new calibracaoDAO();
                        calaux.Codigocalibracao = calibracao.Codigocalibracao;
                        usuario Log = new usuario();
                        Log.Codigousuario = usuario.Codigousuario;
                        calaux.Statusinspecao = "PENDENTE";
                        calibracaoDAO.Inspecao(Log, calaux);
                        */
                    }
                }

            }
            catch (Exception er)
            {
                //Atualiza o status da inspecao.
                if (!ficha)
                {
                    /*
                    calibracao calaux = new calibracao();
                    calibracaoDAO calibracaoDAO = new calibracaoDAO();
                    calaux.Codigocalibracao = calibracao.Codigocalibracao;
                    usuario Log = new usuario();
                    Log.Codigousuario = usuario.Codigousuario;
                    calaux.Statusinspecao = "PENDENTE";
                    calibracaoDAO.Inspecao(Log, calaux);
                    */
                    throw new Exception("Atenção. Certificado digital está invalido e/ou a senha informada está incorreta.  " + er.Message);
                }
            }

        }
        public bool GerarAssinatura(string caminhoservidor, idioma idioma, usuario usuario, bool ficha, ArrayList listacalibracao, ref string mensagemerro)
        {
            bool retornook = true;
            DataTable dt = new DataTable();
            /*calibracaoDAO calibracaoDAO = new calibracaoDAO();
            DataTable dt = calibracaoDAO.retornarcalibracoes(listacalibracao);
            
*/
            parametrogeral parametrogeral = new parametrogeral();
            parametrogeralDAO parametrogeralDAO = new parametrogeralDAO();
            parametrogeral = parametrogeralDAO.RetornaDadosConsultaid(parametrogeral);

            string tsaUrl = string.Empty;
            /*
            if (!string.IsNullOrEmpty(parametrogeral.Urltimestamp))
            {
                tsaUrl = parametrogeral.Urltimestamp;
            }
            else
            {*/
                tsaUrl = "http://timestamp.comodoca.com/authenticod";
            //}

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
                    calibracao calibracao = new calibracao();
                    calibracao.Codigocalibracao = Convert.ToInt32(dt.Rows[ct]["CODIGOCALIBRACAO"].ToString());

                    calibracao.Instrumento.Empresa.Codigoempresa = Convert.ToInt32(dt.Rows[ct]["CODIGOEMPRESA"].ToString());
                    calibracao.Codigocredenciamento = dt.Rows[ct]["codigocredenciamento"].ToString();

                    if (DateTime.TryParse(dt.Rows[0]["datainspecao"].ToString(), out DateTime _datainspecao))
                    {
                        calibracao.Datainspecao = _datainspecao;
                    }

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

                        path = parametrogeral.Diretoriocertificado + "\\certificados\\" + calibracao.Instrumento.Empresa.Codigoempresa.ToString() + diretorioidioma;

                        if (!ficha)
                        {
                            caminho = path + calibracao.Codigocalibracao.ToString() + ".pdf";
                        }
                        else
                        {
                            caminho = path + "F" + calibracao.Codigocalibracao.ToString() + ".pdf";
                        }


                        assinar(parametrogeral.Diretoriocertificado, myCert, caminho, calibracao, usuario, idioma, ficha);
                    }
                    catch (Exception er)
                    {
                        throw new Exception("Ocorreu um erro no processo de assinatura.[7]  " + er.Message.ToString());
                    }
                }
            }
            return retornook;
        }

    }
}
