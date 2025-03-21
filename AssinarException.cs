using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MilordAssinatura
{
    public sealed class AssinarException : Exception
    {
        private AssinarException(string message)
            : base(message)
        {

        }

        public static AssinarException ServicoNaoImplementado(int codigoservico)
           => new AssinarException($"Serviço de Assinatura não implementado: {codigoservico}");

        public static AssinarException ExcluirRegistroAssinatura()
         => new AssinarException($"Ocorreu um erro na tentativa de excluir o registro de assinatura de fila." 
             );
        public static AssinarException ErroEnvioEmail()
         => new AssinarException($"Ocorreu um erro no envio de e-mail."
             );


        public static AssinarException TipoServicoNaoReconhecido()
        => new AssinarException($"Tipo de serviço não reconhecido"
            );




    }
}
