using milord.negocio.Tecnico;
using Milord.Negocio.Seguranca;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Milord.Dados.Tecnico;
using System.Threading.Tasks;

namespace MilordAssinatura
{
    public static class emailformatado
    {


        public static string FormatarMensagem(ServicosAssinatura servicoselecionado,List<int> listaservicos, usuario usuario, DateTime dataAssinatura, string horaAssinatura, bool listaerro)
        {
            DataTable dt = new DataTable();
            string campodata=string.Empty;

            if (servicoselecionado == ServicosAssinatura.Calibracao)
            {
                campodata = "DATAAFERICAO";

                calibracaoassinaturaDAO calibracaoAssinaturaDAO = new calibracaoassinaturaDAO();
                dt = calibracaoAssinaturaDAO.RetornaDadosEnvioEmail(listaservicos);
            }
           
            else if (servicoselecionado == ServicosAssinatura.Manutencao)
            {
                campodata = "DATAVERIFICACAO";

                manutencaoassinaturaDAO manutencaoassinaturaDAO = new manutencaoassinaturaDAO();
                dt = manutencaoassinaturaDAO.RetornaDadosEnvioEmail(listaservicos);
            }
            else if (servicoselecionado == ServicosAssinatura.QualificacaoTermica)
            {
                campodata = "DATAAFERICAO";

                qualificacaotermicaassinaturaDAO qualificacaotermicaassinaturaDAO = new qualificacaotermicaassinaturaDAO();
                dt = qualificacaotermicaassinaturaDAO.RetornaDadosEnvioEmail(listaservicos);
            }
            else if (servicoselecionado == ServicosAssinatura.Ensaio)
            {
                campodata = "DATAAFERICAO";

                ensaioassinaturaDAO ensaioassinaturaDAO = new ensaioassinaturaDAO();
                dt = ensaioassinaturaDAO.RetornaDadosEnvioEmail(listaservicos);
            }

            string arquivo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "template_html.txt");

            if (!File.Exists(arquivo))
            {
                return string.Empty;
            }

            string texto = File.ReadAllText(arquivo);

            StringBuilder tabela = new StringBuilder();

            int ct = 0;

            string _assinado;

            if (listaerro)
                _assinado = "Não";
            else
                _assinado = "Sim";

            foreach (DataRow row in dt.Rows)
            {
                string rowClass = ct % 2 == 0 ? string.Empty : "class=\"active-row\"";

                tabela.Append($@"
                                <tr {rowClass}>
                                <td>{row["CODIGOEMPRESA"]} - {row["NOMEEMPRESA"]}</td>
                                <td>{row["NUMEROINSTRUMENTO"]}</td>
                                <td>{row["DESCRICAOINSTRUMENTO"]}</td>
                                <td>{row["NUMEROCERTIFICADO"]}</td>
                                <td>{Convert.ToDateTime(row[campodata]).ToString("dd/MM/yyyy")}</td>
                                <td>{_assinado}</td>
                                <td>{row["CODIGOORDEMSERVICO"]} - {row["LETRAOS"]}</td>
                                </tr>"
                                );

                ct++;
            }

            string retorno = texto.Replace("[htmlreplace]", tabela.ToString())
                                  .Replace("[nomeusuario]", usuario.Nomeusuario)
                                  .Replace("[datasolicitacao]", dataAssinatura.ToString("dd/MM/yyyy"))
                                  .Replace("[horasolicitacao]", horaAssinatura);

            return retorno;
        }
        

    }
}
