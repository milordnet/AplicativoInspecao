using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MilordAssinatura
{
    [Serializable]
    [Browsable(false)]

    public class servico
    {
        public Iassinar Servico { get; set; }

        public servico(Iassinar servicoassinatura)
        {
            Servico = servicoassinatura;

        }

       

    }
}