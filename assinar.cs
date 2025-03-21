using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MilordAssinatura
{
    public static class assinar
    {
        private static readonly Dictionary<int, Lazy<Iassinar>> Servicos = new Dictionary<int, Lazy<Iassinar>>
        {
            [0] = assinarCalibracao.Instance,
            [1] = assinarCalibracao.Instance,
            [2] = assinarCalibracao.Instance,
            [3] = assinarCalibracao.Instance

        };

        public static Iassinar Instancia(int codigotiposervico)
       => (Servicos.ContainsKey(codigotiposervico) ? Servicos[codigotiposervico] : throw AssinarException.ServicoNaoImplementado(codigotiposervico)).Value;


        public static Iassinar Instancia(ServicosAssinatura codigotiposervico)
            => Instancia((int)codigotiposervico);

  

    }
}
